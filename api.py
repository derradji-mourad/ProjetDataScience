"""
API REST FastAPI pour prédire l'orientation politique d'un article de presse.

Endpoints :
  GET  /health    : statut du service
  GET  /metadata  : info sur le modèle (nom, classes, métriques)
  POST /predict   : { "text": "..." } -> { orientation, label, probabilities }

Lancement local :
  set NEWSAPI_KEY=...   # optionnel
  uvicorn api:app --reload --port 8000
"""

from __future__ import annotations

import os
import pickle
import re
from pathlib import Path
from typing import Dict, List

import numpy as np
import pandas as pd
import scipy.sparse as sp
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field

import nltk
from nltk.corpus import stopwords
from nltk.stem import SnowballStemmer
from textblob import TextBlob
from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
MODEL_PATH = DATA / "best_model.pkl"
RESULTS_PATH = DATA / "resultats_modeles.csv"

# ── Préparation des ressources NLTK (silencieux si déjà téléchargés) ─────────
for pkg in ("stopwords", "punkt", "punkt_tab"):
    try:
        nltk.download(pkg, quiet=True)
    except Exception:
        pass

STEMMER = SnowballStemmer("french")
STOPS = set(stopwords.words("french")) | set(stopwords.words("english")) | {
    "les", "des", "une", "est", "que", "qui", "par", "sur", "cette",
    "pour", "dans", "avec", "plus", "tout", "leur", "aux", "ces", "ont",
}
SIA = SentimentIntensityAnalyzer()

EMO_TERMS = ["menace", "danger", "crise", "urgence", "catastrophe",
             "alarmant", "solidarite", "justice"]

VOCAB = {
    "Gauche":  ["droit", "justice", "solidarit", "egalit", "social", "lutt", "opprim"],
    "Centre":  ["analys", "expert", "rapport", "confirm", "strateg", "negoci"],
    "Droite":  ["securit", "souverain", "national", "frontier", "identit", "menac", "ordre"],
}

ARTIFACTS: Dict = {}


def clean_text(text: str) -> str:
    """Pipeline de nettoyage NLP identique à NB2."""
    if not isinstance(text, str):
        return ""
    text = text.lower()
    text = re.sub(r"http\S+|www\.\S+", "", text)
    text = re.sub(r"<[^>]+>", "", text)
    text = re.sub(r"[^a-zàâçéèêëîïôûùüÿñæœ\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    tokens = [STEMMER.stem(t) for t in text.split() if t not in STOPS and len(t) > 2]
    return " ".join(tokens)


def vocab_score(text: str, orientation: str) -> float:
    tokens = str(text).lower().split()
    if not tokens:
        return 0.0
    score = sum(1 for t in tokens if any(kw in t for kw in VOCAB[orientation]))
    return round(score / len(tokens), 4)


def numerical_features(raw_text: str, cleaned_text: str) -> List[float]:
    """Calcule les 11 features numériques attendues par le modèle, dans l'ordre."""
    polarity = SIA.polarity_scores(raw_text)
    try:
        subjectivity = TextBlob(raw_text).sentiment.subjectivity
    except Exception:
        subjectivity = 0.0
    words = cleaned_text.split()
    n_words = max(len(words), 1)
    richesse = round(len(set(words)) / n_words, 4)
    densite_emo = round(
        sum(1 for w in raw_text.lower().split() if any(e in w for e in EMO_TERMS))
        / max(len(raw_text.split()), 1),
        4,
    )
    sg = vocab_score(raw_text, "Gauche")
    sc = vocab_score(raw_text, "Centre")
    sd = vocab_score(raw_text, "Droite")
    polar = round(max(sg, sd) - sc, 4)
    return [
        round(polarity["compound"], 4),     # sentiment_vader
        round(polarity["pos"], 4),          # positivite_vader
        round(polarity["neg"], 4),          # negativite_vader
        round(subjectivity, 4),             # subjectivite
        richesse,                           # richesse_lexicale
        densite_emo,                        # densite_emotionnel
        len(raw_text.split()),              # nb_mots
        polar,                              # indice_polarisation
        sg, sc, sd,                         # scores vocab
    ]


# ── App ──────────────────────────────────────────────────────────────────────

app = FastAPI(
    title="Polarisation Médiatique — API",
    description="Prédit l'orientation politique d'un article (Gauche / Centre / Droite).",
    version="1.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.on_event("startup")
def load_artifacts() -> None:
    if not MODEL_PATH.exists():
        raise RuntimeError(
            f"Modèle introuvable : {MODEL_PATH}. "
            "Exécutez d'abord 03_modeling.ipynb pour le générer."
        )
    with open(MODEL_PATH, "rb") as f:
        ARTIFACTS.update(pickle.load(f))


# ── Schemas ──────────────────────────────────────────────────────────────────

class PredictIn(BaseModel):
    text: str = Field(..., min_length=10, description="Corps de l'article à classer")


class PredictOut(BaseModel):
    orientation: str
    label: int
    probabilities: Dict[str, float]
    confidence: float


# ── Endpoints ────────────────────────────────────────────────────────────────

@app.get("/health")
def health() -> Dict[str, str]:
    return {"status": "ok", "model_loaded": str(bool(ARTIFACTS))}


@app.get("/metadata")
def metadata() -> Dict:
    if not ARTIFACTS:
        raise HTTPException(503, "Modèle non chargé")
    info: Dict = {
        "labels": ARTIFACTS["label_map"],
        "num_features": ARTIFACTS["num_features"],
        "tfidf_vocab_size": len(ARTIFACTS["tfidf"].get_feature_names_out()),
        "model_class": type(ARTIFACTS["model"]).__name__,
    }
    if RESULTS_PATH.exists():
        df = pd.read_csv(RESULTS_PATH)
        info["leaderboard"] = df.to_dict(orient="records")
    return info


@app.post("/predict", response_model=PredictOut)
def predict(payload: PredictIn) -> PredictOut:
    if not ARTIFACTS:
        raise HTTPException(503, "Modèle non chargé")

    cleaned = clean_text(payload.text)
    if not cleaned:
        raise HTTPException(400, "Texte trop court ou vide après nettoyage")

    tfidf = ARTIFACTS["tfidf"]
    scaler = ARTIFACTS["scaler"]
    selector = ARTIFACTS["selector"]
    model = ARTIFACTS["model"]
    label_map = ARTIFACTS["label_map"]

    x_tfidf = tfidf.transform([cleaned])
    x_num = np.array([numerical_features(payload.text, cleaned)])
    x_num_scaled = scaler.transform(x_num)
    x_combined = sp.hstack([x_tfidf, sp.csr_matrix(x_num_scaled)])
    x_selected = selector.transform(x_combined)

    pred = int(model.predict(x_selected)[0])

    if hasattr(model, "predict_proba"):
        probs = model.predict_proba(x_selected)[0]
    else:
        # Fallback : SVM ou modèle sans proba -> softmax sur decision_function
        if hasattr(model, "decision_function"):
            scores = model.decision_function(x_selected)[0]
            exp = np.exp(scores - scores.max())
            probs = exp / exp.sum()
        else:
            probs = np.zeros(len(label_map))
            probs[pred] = 1.0

    proba_by_label = {label_map[i]: round(float(probs[i]), 4) for i in range(len(probs))}
    return PredictOut(
        orientation=label_map[pred],
        label=pred,
        probabilities=proba_by_label,
        confidence=round(float(probs[pred]), 4),
    )
