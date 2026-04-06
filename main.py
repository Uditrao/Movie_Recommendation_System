"""
Movie Recommender API — 100% Local Dataset
==========================================
- Uses movies_metadata.csv + pkl files (no TMDB API key required)
- Poster images served via TMDB CDN (public, no auth needed)
- TF-IDF + Genre-based recommendations
"""

import ast
import os
import pickle
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
from fastapi import FastAPI, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel

# ─────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

TMDB_IMG_BASE = "https://image.tmdb.org/t/p/w500"   # public CDN — no key needed
TMDB_BACKDROP = "https://image.tmdb.org/t/p/w1280"

FALLBACK_POSTER = "https://via.placeholder.com/500x750/1f1f1f/ffffff?text=No+Poster"

# ─────────────────────────────────────────
# FASTAPI
# ─────────────────────────────────────────
app = FastAPI(title="Movie Recommender API (Local)", version="4.0")

# Serve the frontend at root
@app.get("/", include_in_schema=False)
def root():
    html_path = Path(BASE_DIR) / "index.html"
    if html_path.exists():
        return FileResponse(str(html_path), media_type="text/html")
    return {"message": "index.html not found"}


app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ─────────────────────────────────────────
# GLOBALS
# ─────────────────────────────────────────
META: Optional[pd.DataFrame] = None   # movies_metadata.csv (cleaned)
df: Optional[pd.DataFrame] = None     # tfidf df.pkl
indices_obj: Any = None
tfidf_matrix: Any = None
tfidf_obj: Any = None
TITLE_TO_IDX: Dict[str, int] = {}
ID_TO_META: Dict[int, dict] = {}      # tmdb_id -> metadata row


# ─────────────────────────────────────────
# PYDANTIC MODELS
# ─────────────────────────────────────────
class MovieCard(BaseModel):
    tmdb_id: int
    title: str
    poster_url: str
    backdrop_url: Optional[str] = None
    release_date: Optional[str] = None
    vote_average: Optional[float] = None
    popularity: Optional[float] = None


class Genre(BaseModel):
    id: Optional[int] = None
    name: str


class MovieDetails(BaseModel):
    tmdb_id: int
    title: str
    overview: Optional[str] = None
    release_date: Optional[str] = None
    poster_url: str
    backdrop_url: Optional[str] = None
    genres: List[Genre] = []
    vote_average: Optional[float] = None
    runtime: Optional[int] = None
    tagline: Optional[str] = None


class TFIDFRecItem(BaseModel):
    title: str
    score: float
    movie: Optional[MovieCard] = None


class SearchBundle(BaseModel):
    query: str
    movie_details: MovieDetails
    tfidf_recommendations: List[TFIDFRecItem]
    genre_recommendations: List[MovieCard]


# ─────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────
def _norm(t: str) -> str:
    return str(t).strip().lower()


def _poster(path: Optional[str]) -> str:
    if path and str(path) not in ("nan", "None", ""):
        p = str(path).strip()
        if not p.startswith("/"):
            p = "/" + p
        return f"{TMDB_IMG_BASE}{p}"
    return FALLBACK_POSTER


def _backdrop(path: Optional[str]) -> Optional[str]:
    if path and str(path) not in ("nan", "None", ""):
        p = str(path).strip()
        if not p.startswith("/"):
            p = "/" + p
        return f"{TMDB_BACKDROP}{p}"
    return None


def _parse_genres(raw) -> List[Genre]:
    """Parse genres from JSON-like string in CSV."""
    if not raw or str(raw) in ("nan", "None", ""):
        return []
    try:
        if isinstance(raw, list):
            items = raw
        else:
            items = ast.literal_eval(str(raw))
        return [Genre(id=g.get("id"), name=g["name"]) for g in items if "name" in g]
    except Exception:
        return []


def _parse_int(val, default=0) -> int:
    try:
        v = int(float(str(val)))
        return v if v > 0 else default
    except Exception:
        return default


def _parse_float(val, default=0.0) -> float:
    try:
        return float(str(val))
    except Exception:
        return default


def _row_to_card(row: dict) -> MovieCard:
    return MovieCard(
        tmdb_id=_parse_int(row.get("id", 0)),
        title=str(row.get("title", "Untitled")),
        poster_url=_poster(row.get("poster_path")),
        backdrop_url=_backdrop(row.get("backdrop_path")),
        release_date=str(row.get("release_date", ""))[:10] or None,
        vote_average=_parse_float(row.get("vote_average")),
        popularity=_parse_float(row.get("popularity")),
    )


def _row_to_details(row: dict) -> MovieDetails:
    return MovieDetails(
        tmdb_id=_parse_int(row.get("id", 0)),
        title=str(row.get("title", "Untitled")),
        overview=str(row.get("overview", "")) or None,
        release_date=str(row.get("release_date", ""))[:10] or None,
        poster_url=_poster(row.get("poster_path")),
        backdrop_url=_backdrop(row.get("backdrop_path")),
        genres=_parse_genres(row.get("genres")),
        vote_average=_parse_float(row.get("vote_average")),
        runtime=_parse_int(row.get("runtime"), default=None),
        tagline=str(row.get("tagline", "")) or None,
    )


def _find_meta_by_title(title: str) -> Optional[dict]:
    """Fuzzy title lookup in META dataframe."""
    global META
    if META is None:
        return None
    t = _norm(title)
    # exact match first
    exact = META[META["_norm_title"] == t]
    if not exact.empty:
        return exact.iloc[0].to_dict()
    # contains match
    contains = META[META["_norm_title"].str.contains(re.escape(t), na=False)]
    if not contains.empty:
        return contains.iloc[0].to_dict()
    return None


def _build_title_to_idx(indices: Any) -> Dict[str, int]:
    out: Dict[str, int] = {}
    try:
        for k, v in indices.items():
            out[_norm(str(k))] = int(v)
    except Exception as e:
        raise RuntimeError(f"indices.pkl parse error: {e}")
    return out


def _tfidf_recommend(query_title: str, top_n: int = 12) -> List[Tuple[str, float]]:
    global df, tfidf_matrix, TITLE_TO_IDX
    if df is None or tfidf_matrix is None:
        return []
    key = _norm(query_title)
    if key not in TITLE_TO_IDX:
        return []
    idx = TITLE_TO_IDX[key]
    qv = tfidf_matrix[idx]
    scores = (tfidf_matrix @ qv.T).toarray().ravel()
    order = np.argsort(-scores)
    out: List[Tuple[str, float]] = []
    for i in order:
        if int(i) == int(idx):
            continue
        try:
            t = str(df.iloc[int(i)]["title"])
        except Exception:
            continue
        out.append((t, float(scores[int(i)])))
        if len(out) >= top_n:
            break
    return out


# ─────────────────────────────────────────
# STARTUP
# ─────────────────────────────────────────
@app.on_event("startup")
def startup():
    global META, df, indices_obj, tfidf_matrix, tfidf_obj, TITLE_TO_IDX, ID_TO_META

    # 1) Load movies_metadata.csv
    meta_path = os.path.join(BASE_DIR, "movies_metadata.csv")
    if os.path.exists(meta_path):
        raw = pd.read_csv(meta_path, low_memory=False)

        # Keep only rows with a real tmdb id (numeric)
        raw = raw[pd.to_numeric(raw["id"], errors="coerce").notna()].copy()
        raw["id"] = raw["id"].apply(lambda x: _parse_int(x))
        raw = raw[raw["id"] > 0]

        # Drop obvious garbage rows (test entries, etc.)
        raw = raw[raw["title"].notna() & (raw["title"].str.len() > 0)]

        raw["_norm_title"] = raw["title"].apply(_norm)
        raw["vote_average"] = pd.to_numeric(raw["vote_average"], errors="coerce").fillna(0)
        raw["vote_count"] = pd.to_numeric(raw["vote_count"], errors="coerce").fillna(0)
        raw["popularity"] = pd.to_numeric(raw["popularity"], errors="coerce").fillna(0)

        META = raw.reset_index(drop=True)

        # Build id -> row dict for fast lookups
        for _, row in META.iterrows():
            mid = _parse_int(row.get("id", 0))
            if mid > 0:
                ID_TO_META[mid] = row.to_dict()

        print(f"[startup] Loaded {len(META)} movies from metadata CSV")
    else:
        print("[startup] WARNING: movies_metadata.csv not found — search/home feed limited")

    # 2) Load pkl files
    for name, path_key in [
        ("df", "df.pkl"),
        ("indices_obj", "indices.pkl"),
        ("tfidf_matrix", "tfidf_matrix.pkl"),
        ("tfidf_obj", "tfidf.pkl"),
    ]:
        fp = os.path.join(BASE_DIR, path_key)
        if os.path.exists(fp):
            with open(fp, "rb") as f:
                loaded = pickle.load(f)
            if name == "df":
                df = loaded
            elif name == "indices_obj":
                indices_obj = loaded
            elif name == "tfidf_matrix":
                tfidf_matrix = loaded
            elif name == "tfidf_obj":
                tfidf_obj = loaded
            print(f"[startup] Loaded {path_key}")
        else:
            print(f"[startup] WARNING: {path_key} not found")

    # 3) Build TITLE_TO_IDX
    if indices_obj is not None:
        try:
            TITLE_TO_IDX = _build_title_to_idx(indices_obj)
            print(f"[startup] Title index built: {len(TITLE_TO_IDX)} entries")
        except Exception as e:
            print(f"[startup] WARNING: Could not build title index: {e}")


# ─────────────────────────────────────────
# ROUTES
# ─────────────────────────────────────────

@app.get("/health")
def health():
    return {
        "status": "ok",
        "movies_loaded": len(META) if META is not None else 0,
        "tfidf_ready": tfidf_matrix is not None,
        "title_index_size": len(TITLE_TO_IDX),
    }


# ── HOME FEED ──────────────────────────────────────────────────────────────────
@app.get("/home", response_model=List[MovieCard])
def home(
    category: str = Query("trending"),
    limit: int = Query(24, ge=1, le=50),
):
    """
    Returns movie cards from local dataset by category:
    - trending    → highest popularity score
    - popular     → same as trending
    - top_rated   → highest vote_average (min 100 votes)
    - now_playing → randomly sampled from recent releases
    - upcoming    → movies with future or recent release dates
    """
    if META is None:
        raise HTTPException(status_code=503, detail="Dataset not loaded")

    base = META.copy()

    if category in ("trending", "popular"):
        result = base.sort_values("popularity", ascending=False).head(limit * 2)

    elif category == "top_rated":
        # Bayesian average: only include movies with decent vote counts
        qualified = base[base["vote_count"] >= 100]
        result = qualified.sort_values("vote_average", ascending=False).head(limit * 2)

    elif category == "now_playing":
        # Sample from movies that have posters, sorted by popularity
        has_poster = base[base["poster_path"].notna() & (base["poster_path"] != "")]
        result = has_poster.sort_values("popularity", ascending=False).head(200).sample(
            n=min(limit, len(has_poster)), random_state=42
        )

    elif category == "upcoming":
        # Movies released 2015+, sorted by release date desc
        base["_year"] = pd.to_numeric(
            base["release_date"].str[:4], errors="coerce"
        ).fillna(0)
        result = (
            base[base["_year"] >= 2015]
            .sort_values("_year", ascending=False)
            .head(limit * 2)
        )

    else:
        raise HTTPException(status_code=400, detail=f"Unknown category: {category}")

    cards = []
    for _, row in result.iterrows():
        card = _row_to_card(row.to_dict())
        if card.tmdb_id > 0:
            cards.append(card)
        if len(cards) >= limit:
            break

    return cards


# ── SEARCH ─────────────────────────────────────────────────────────────────────
@app.get("/search", response_model=List[MovieCard])
def search(
    query: str = Query(..., min_length=1),
    limit: int = Query(24, ge=1, le=50),
):
    """Search movies by title from local dataset."""
    if META is None:
        raise HTTPException(status_code=503, detail="Dataset not loaded")

    q = _norm(query)
    
    # Priority: title starts with query, then contains
    starts = META[META["_norm_title"].str.startswith(q, na=False)]
    contains = META[META["_norm_title"].str.contains(re.escape(q), na=False)]
    
    combined = pd.concat([starts, contains]).drop_duplicates(subset=["id"])
    combined = combined.sort_values("popularity", ascending=False)

    cards = []
    for _, row in combined.head(limit * 2).iterrows():
        card = _row_to_card(row.to_dict())
        if card.tmdb_id > 0 and card.title.strip():
            cards.append(card)
        if len(cards) >= limit:
            break

    return cards


# ── MOVIE DETAILS BY ID ────────────────────────────────────────────────────────
@app.get("/movie/id/{movie_id}", response_model=MovieDetails)
def movie_details(movie_id: int):
    """Get full details of a movie by its TMDB ID."""
    if movie_id in ID_TO_META:
        return _row_to_details(ID_TO_META[movie_id])

    if META is None:
        raise HTTPException(status_code=503, detail="Dataset not loaded")

    row_df = META[META["id"] == movie_id]
    if row_df.empty:
        raise HTTPException(status_code=404, detail=f"Movie ID {movie_id} not found")

    return _row_to_details(row_df.iloc[0].to_dict())


# ── BUNDLE: Details + TF-IDF + Genre recs ─────────────────────────────────────
@app.get("/movie/search", response_model=SearchBundle)
def search_bundle(
    query: str = Query(..., min_length=1),
    tfidf_top_n: int = Query(12, ge=1, le=30),
    genre_limit: int = Query(12, ge=1, le=30),
):
    """
    All-in-one: movie details + TF-IDF recommendations + genre recommendations.
    Uses local dataset only — no external API.
    """
    if META is None:
        raise HTTPException(status_code=503, detail="Dataset not loaded")

    # Find movie in local dataset
    meta_row = _find_meta_by_title(query)
    if meta_row is None:
        raise HTTPException(status_code=404, detail=f"Movie not found: '{query}'")

    details = _row_to_details(meta_row)
    movie_id = details.tmdb_id

    # 1) TF-IDF recommendations
    tfidf_items: List[TFIDFRecItem] = []
    recs = _tfidf_recommend(details.title, top_n=tfidf_top_n)
    if not recs:
        recs = _tfidf_recommend(query, top_n=tfidf_top_n)

    for title, score in recs:
        meta = _find_meta_by_title(title)
        card = _row_to_card(meta) if meta else None
        tfidf_items.append(TFIDFRecItem(title=title, score=score, movie=card))

    # 2) Genre-based recommendations
    genre_recs: List[MovieCard] = []
    if details.genres and META is not None:
        genre_name = details.genres[0].name
        q = _norm(genre_name)

        def _has_genre(raw_genres) -> bool:
            for g in _parse_genres(raw_genres):
                if _norm(g.name) == q:
                    return True
            return False

        genre_movies = META[META["genres"].apply(_has_genre)]
        genre_movies = genre_movies[genre_movies["id"] != movie_id]
        genre_movies = genre_movies.sort_values("vote_average", ascending=False)

        for _, row in genre_movies.head(genre_limit * 2).iterrows():
            card = _row_to_card(row.to_dict())
            if card.tmdb_id > 0:
                genre_recs.append(card)
            if len(genre_recs) >= genre_limit:
                break

    return SearchBundle(
        query=query,
        movie_details=details,
        tfidf_recommendations=tfidf_items,
        genre_recommendations=genre_recs,
    )


# ── GENRE RECOMMENDATIONS ──────────────────────────────────────────────────────
@app.get("/recommend/genre", response_model=List[MovieCard])
def recommend_genre(
    movie_id: int = Query(...),
    limit: int = Query(18, ge=1, le=50),
):
    """Recommend movies by the same genre as the given movie ID."""
    if META is None:
        raise HTTPException(status_code=503, detail="Dataset not loaded")

    row = ID_TO_META.get(movie_id)
    if not row:
        raise HTTPException(status_code=404, detail=f"Movie ID {movie_id} not found")

    genres = _parse_genres(row.get("genres"))
    if not genres:
        return []

    genre_name = _norm(genres[0].name)

    def _has_genre(raw) -> bool:
        return any(_norm(g.name) == genre_name for g in _parse_genres(raw))

    same_genre = META[META["genres"].apply(_has_genre)]
    same_genre = same_genre[same_genre["id"] != movie_id]
    same_genre = same_genre.sort_values("vote_average", ascending=False)

    cards = []
    for _, r in same_genre.head(limit * 2).iterrows():
        c = _row_to_card(r.to_dict())
        if c.tmdb_id > 0:
            cards.append(c)
        if len(cards) >= limit:
            break
    return cards


# ── TF-IDF ONLY ────────────────────────────────────────────────────────────────
@app.get("/recommend/tfidf")
def recommend_tfidf(
    title: str = Query(..., min_length=1),
    top_n: int = Query(12, ge=1, le=50),
):
    recs = _tfidf_recommend(title, top_n=top_n)
    return [{"title": t, "score": s} for t, s in recs]
