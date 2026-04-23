# 🎬 Movie Recommendation System

A professional, **Netflix-style** movie recommendation engine built using **Content-Based Filtering**. The system analyzes movie metadata (overviews, genres, and taglines) to suggest the most semantically similar titles from a dataset of over **45,000 movies**.

![Netflix Style UI Preview](https://img.shields.io/badge/UI-Netflix--Style-red?style=for-the-badge)
![FastAPI](https://img.shields.io/badge/Backend-FastAPI-009688?style=for-the-badge&logo=fastapi&logoColor=white)
![Scikit-Learn](https://img.shields.io/badge/ML-Scikit--Learn-F7931E?style=for-the-badge&logo=scikitlearn&logoColor=white)

---

## 🚀 Key Features

- **🎯 Zero API Dependency**: Runs 100% locally using the TMDB 45K dataset. No API keys required for posters or metadata.
- **📱 Netflix-Style Frontend**: Modern, responsive UI with a hero banner, horizontal scrolling carousels, and an interactive movie detail modal.
- **🧠 4-Model Comparison**: Implements and compares four different ML approaches:
    - Cosine Similarity (TF-IDF)
    - K-Nearest Neighbors (KNN)
    - Sigmoid Kernel
    - **K-Means Clustering** (Best Performing)
- **⚡ High Performance**: Powered by **FastAPI** and pre-computed **Pickle (.pkl)** files for near-instant recommendations.
- **🔍 Advanced Search**: Intelligent search that fetches plot summaries, ratings, and genre-based recommendations.

---

## 🛠️ Tech Stack

- **Machine Learning**: Python, `scikit-learn`, `pandas`, `numpy`, `NLTK`.
- **Backend**: FastAPI, Uvicorn.
- **Frontend**: Vanilla HTML5, CSS3 (Modern Flexbox/Grid), JavaScript (Fetch API).
- **Dataset**: TMDB 5000 Movies + 45K Movies Metadata (Kaggle).

---

## 📊 Model Performance

We evaluated the system using a **Multi-Genre Overlap** methodology. **K-Means Clustering** emerged as the superior model for capturing broad thematic similarities.

| Model | Accuracy | Precision | Recall | F1-Score |
|---|---|---|---|---|
| Cosine Similarity | 70.0% | 90.0% | 64.29% | 75.00% |
| KNN | 70.0% | 90.0% | 64.29% | 75.00% |
| Sigmoid Kernel | 70.0% | 90.0% | 64.29% | 75.00% |
| **K-Means Clustering** | **72.0%** | **90.0%** | **66.18%** | **76.27%** |

---

## ⚙️ Installation & Setup

### 1. Clone the Repository
```bash
git clone https://github.com/Uditrao/Movie_Recommendation_System.git
cd Movie_Recommendation_System
```

### 2. Install Dependencies
```bash
pip install -r requirements.txt
```

### 3. Run the Backend
```bash
python -m uvicorn main:app --reload
```
The server will start at `http://127.0.0.1:8000`.

### 4. Access the UI
Simply open `index.html` in your browser. The frontend will automatically connect to the FastAPI backend.

---

## 📁 Project Structure

```text
├── main.py              # FastAPI Backend API
├── index.html           # Netflix-style Frontend
├── movies.ipynb         # Data Analysis & Model Training
├── requirements.txt     # Python Dependencies
├── .gitignore           # Git ignore rules for large files/venv
├── df.pkl               # Preprocessed Movie Data (Local only)
├── tfidf_matrix.pkl     # TF-IDF Sparse Matrix (Local only)
└── Project.html         # Academic Project Poster
```

---

## 💡 How it Works

1.  **Preprocessing**: Combines `overview`, `genres`, and `tagline` into a single `tags` field.
2.  **Vectorization**: Uses `TfidfVectorizer` (50,000 features, N-grams 1-2) to convert text into numeric vectors.
3.  **Similarity**: Computes the distance between movie vectors to find the closest matches.
4.  **UI**: Fetches and displays results in a beautiful, interactive layout.

---

## 📜 License
This project was developed by **Udit Yadav** for academic purposes. Feel free to use and modify for your own learning!
