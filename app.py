"""
Application Streamlit — Polarisation Médiatique Géopolitique.

Démarrage :
  1) Lancer l'API dans un terminal :   uvicorn api:app --port 8000
  2) Lancer ce dashboard dans un autre : streamlit run app.py
"""

from __future__ import annotations

from pathlib import Path

import httpx
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
API_URL = "http://127.0.0.1:8000"
COLORS = {"Gauche": "#E74C3C", "Centre": "#27AE60", "Droite": "#2980B9"}

st.set_page_config(
    page_title="Polarisation Médiatique",
    page_icon=None,
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Helpers ──────────────────────────────────────────────────────────────────

@st.cache_data(ttl=60)
def load_dataset() -> pd.DataFrame:
    path = DATA / "articles_preprocesses.csv"
    if not path.exists():
        return pd.DataFrame()
    df = pd.read_csv(path)
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    return df


@st.cache_data(ttl=60)
def load_leaderboard() -> pd.DataFrame:
    path = DATA / "resultats_modeles.csv"
    return pd.read_csv(path) if path.exists() else pd.DataFrame()


def call_predict(text: str) -> dict:
    try:
        r = httpx.post(f"{API_URL}/predict", json={"text": text}, timeout=15)
        r.raise_for_status()
        return r.json()
    except httpx.HTTPStatusError as e:
        return {"error": f"HTTP {e.response.status_code} : {e.response.text}"}
    except httpx.RequestError as e:
        return {"error": f"API injoignable sur {API_URL} ({e}). Lancez : uvicorn api:app --port 8000"}


def call_metadata() -> dict | None:
    try:
        r = httpx.get(f"{API_URL}/metadata", timeout=5)
        r.raise_for_status()
        return r.json()
    except Exception:
        return None


def confidence_label(conf: float) -> tuple[str, str]:
    """Renvoie (libellé, couleur de fond) selon le niveau de confiance."""
    if conf >= 0.80:
        return "Très haute", "#27AE60"
    if conf >= 0.60:
        return "Modérée", "#F39C12"
    return "Faible — résultat à interpréter avec prudence", "#C0392B"


# ── Sidebar ──────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("## Polarisation Médiatique")
    st.caption("Bloc 4 — Data Science & IA")

    meta = call_metadata()
    if meta is None:
        st.error("API hors ligne")
        st.code("uvicorn api:app --port 8000", language="bash")
    else:
        st.success(f"API en ligne — modèle : **{meta['model_class']}**")
        st.caption(f"Vocabulaire TF-IDF : {meta['tfidf_vocab_size']} termes")

    st.divider()
    page = st.radio(
        "Navigation",
        ["Prédiction en direct", "Vue d'ensemble du dataset", "Performance des modèles"],
        index=0,
    )


# ════════════════════════════════════════════════════════════════════════════
# PAGE 1 — PRÉDICTION EN DIRECT
# ════════════════════════════════════════════════════════════════════════════

if page == "Prédiction en direct":
    st.title("Prédiction d'orientation politique")
    st.caption("Collez un article ou un paragraphe. Le modèle retourne l'orientation prédite et la confiance par classe.")

    with st.expander("Comment fonctionne le modèle ?"):
        st.markdown(
            """
            **Pipeline en 5 étapes** appliqué à votre texte côté API :
            1. **Nettoyage NLP** : minuscules, suppression URLs/HTML/ponctuation, tokenisation, suppression des stop-words FR + EN, *stemming* français.
            2. **Vectorisation TF-IDF** sur 5 000 termes (uni-grammes + bi-grammes).
            3. **Features numériques** (11) : sentiment VADER, subjectivité TextBlob, richesse lexicale, densité émotionnelle, scores de vocabulaire G/C/D, indice de polarisation.
            4. **Sélection** des 3 000 features les plus informatives (mutual information).
            5. **Classification** par XGBoost — choix issu de la comparaison de 5 algorithmes (cf. page Performance).
            """
        )

    examples = {
        "Choisir un exemple…": "",
        "Exemple — ton à gauche": (
            "Face à la crise migratoire, la solidarité internationale est essentielle. "
            "Les droits humains des réfugiés doivent être garantis face à l'injustice "
            "des politiques restrictives qui frappent les plus vulnérables. La justice "
            "sociale et l'égalité doivent guider les décisions politiques."
        ),
        "Exemple — ton centriste": (
            "Selon les dernières données officielles publiées par l'OCDE, la situation "
            "économique reste complexe. Les experts indiquent que les négociations "
            "commerciales se poursuivent dans un climat d'incertitude prolongée. "
            "Les analystes confirment que la situation reste à surveiller."
        ),
        "Exemple — ton à droite": (
            "La sécurité nationale et la souveraineté de nos frontières sont menacées. "
            "Le gouvernement doit défendre l'ordre républicain et l'identité nationale "
            "face à la pression migratoire incontrôlée. La menace est réelle et il faut "
            "défendre nos valeurs identitaires."
        ),
    }
    chosen = st.selectbox("Exemples préchargés", list(examples.keys()))
    default_text = examples[chosen]

    text = st.text_area(
        "Texte de l'article",
        value=default_text,
        height=180,
        placeholder="Collez ici un extrait d'article (au moins une centaine de caractères)…",
    )

    col_btn, col_info = st.columns([1, 3])
    with col_btn:
        clicked = st.button("Prédire", type="primary", use_container_width=True)
    with col_info:
        st.caption(f"{len(text.split())} mots · {len(text)} caractères")

    if clicked:
        if len(text.strip()) < 30:
            st.warning("Le texte est trop court (au moins 30 caractères).")
        else:
            with st.spinner("Appel de l'API…"):
                res = call_predict(text)
            if "error" in res:
                st.error(res["error"])
            else:
                color = COLORS.get(res["orientation"], "#666")
                st.markdown(
                    f"<div style='padding:18px;border-radius:8px;background:{color};color:white;"
                    f"font-size:22px;font-weight:600;text-align:center;'>"
                    f"Orientation prédite : {res['orientation']}"
                    f" — confiance {res['confidence']*100:.1f}%</div>",
                    unsafe_allow_html=True,
                )

                st.markdown("##### Probabilités par classe")
                probs = res["probabilities"]
                bars = pd.DataFrame({
                    "orientation": list(probs.keys()),
                    "probabilité": list(probs.values()),
                })
                fig = px.bar(
                    bars, x="orientation", y="probabilité",
                    color="orientation", color_discrete_map=COLORS,
                    text=bars["probabilité"].apply(lambda v: f"{v*100:.1f}%"),
                )
                fig.update_layout(yaxis=dict(range=[0, 1], tickformat=".0%"),
                                  showlegend=False, height=320)
                fig.update_traces(textposition="outside")
                st.plotly_chart(fig, use_container_width=True)

                # ── Analyse de la prédiction ─────────────────────────────────
                st.markdown("### Analyse du résultat")
                lvl, lvl_color = confidence_label(res["confidence"])
                cA, cB = st.columns([1, 2])
                with cA:
                    st.markdown(
                        f"<div style='padding:12px;border-radius:6px;background:{lvl_color};"
                        f"color:white;text-align:center;font-weight:600;'>"
                        f"Niveau de confiance : {lvl}</div>",
                        unsafe_allow_html=True,
                    )
                    p_sorted = sorted(probs.items(), key=lambda x: x[1], reverse=True)
                    margin = p_sorted[0][1] - p_sorted[1][1]
                    st.metric("Marge avec la 2ᵉ classe", f"{margin*100:.1f} pts")

                with cB:
                    interpretation = []
                    if res["confidence"] >= 0.80:
                        interpretation.append(
                            "Le modèle est **très sûr** de son choix : les signaux lexicaux et "
                            "numériques pointent fortement vers une seule orientation."
                        )
                    elif res["confidence"] >= 0.60:
                        interpretation.append(
                            "Confiance **modérée** : le texte mélange des marqueurs typiques "
                            "de plusieurs orientations. À considérer comme une indication, pas "
                            "une vérité absolue."
                        )
                    else:
                        interpretation.append(
                            "Confiance **faible** : le texte est ambigu pour le modèle. "
                            "Soit le vocabulaire est trop neutre, soit il combine des marqueurs "
                            "de plusieurs orientations en proportions équilibrées."
                        )
                    if margin < 0.10:
                        interpretation.append(
                            f"⚠ La 2ᵉ classe ({p_sorted[1][0]}) n'est qu'à {margin*100:.1f} pts "
                            "— ce texte se situe à la frontière entre deux orientations."
                        )
                    st.markdown(" ".join(interpretation))

                with st.expander("Limites de cette prédiction"):
                    st.markdown(
                        """
                        - **Le modèle est entraîné majoritairement sur de l'anglais** (Kaggle = 50 % du dataset). Les textes français courts peuvent être moins bien classés.
                        - **Sur des phrases courtes (< 50 mots)**, les features numériques ont peu de signal et le modèle a tendance à s'appuyer trop sur quelques mots-clés. Pour des résultats fiables, fournir des paragraphes complets.
                        - **Le label est statistique, pas factuel** : il indique la *tonalité éditoriale dominante*, pas une vérité sur le sujet traité.
                        """
                    )


# ════════════════════════════════════════════════════════════════════════════
# PAGE 2 — VUE D'ENSEMBLE DU DATASET
# ════════════════════════════════════════════════════════════════════════════

elif page == "Vue d'ensemble du dataset":
    df = load_dataset()
    st.title("Vue d'ensemble du dataset")
    st.caption("Statistiques descriptives et patterns observés dans les 3 000 articles utilisés pour l'entraînement.")

    if df.empty:
        st.warning("Pas de dataset trouvé. Exécutez d'abord 02_preprocessing_eda.ipynb.")
        st.stop()

    # ── KPIs ─────────────────────────────────────────────────────────────────
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Articles", f"{len(df):,}".replace(",", " "))
    c2.metric("Sources", df["source"].nunique())
    c3.metric("Événements", df["evenement"].nunique())
    c4.metric("Période", f"{df['date'].min():%Y-%m} → {df['date'].max():%Y-%m}")

    # ── Provenance ───────────────────────────────────────────────────────────
    if "provenance" in df.columns:
        with st.container(border=True):
            st.markdown("##### Origine des articles")
            prov = df["provenance"].value_counts().reset_index()
            prov.columns = ["provenance", "n"]
            prov["%"] = (prov["n"] / prov["n"].sum() * 100).round(1)
            cP1, cP2 = st.columns([2, 3])
            with cP1:
                st.dataframe(prov, hide_index=True, use_container_width=True)
            with cP2:
                fig = px.pie(prov, names="provenance", values="n", hole=0.45)
                fig.update_layout(height=260, margin=dict(t=10, b=10))
                st.plotly_chart(fig, use_container_width=True)
            st.caption(
                "**Lecture :** quatre sources combinées — Kaggle (annoté manuellement) "
                "fournit la majorité du signal d'apprentissage, RSS et NewsAPI assurent "
                "l'ancrage temporel récent, et le générateur synthétique équilibre les classes "
                "et garantit un volume suffisant (3 000) pour un entraînement stable."
            )

    st.divider()

    # ── Distribution des classes + événements ───────────────────────────────
    left, right = st.columns(2)
    with left:
        st.markdown("##### Répartition des orientations")
        dist = df["orientation"].value_counts().reset_index()
        dist.columns = ["orientation", "n"]
        dist["%"] = (dist["n"] / dist["n"].sum() * 100).round(1)
        fig = px.pie(dist, names="orientation", values="n",
                     color="orientation", color_discrete_map=COLORS, hole=0.4)
        st.plotly_chart(fig, use_container_width=True)

        majo = dist.iloc[0]
        mino = dist.iloc[-1]
        ratio = majo["n"] / mino["n"]
        st.info(
            f"**Observation** : les classes sont **modérément déséquilibrées** "
            f"({majo['orientation']} {majo['%']}% vs {mino['orientation']} {mino['%']}%, "
            f"ratio {ratio:.1f}×). C'est pour cette raison que le modèle est évalué "
            f"avec le **F1 macro** (moyenne non pondérée) et que les splits "
            f"train/test sont **stratifiés**."
        )

    with right:
        st.markdown("##### Articles par événement géopolitique")
        ev = df["evenement"].value_counts().reset_index()
        ev.columns = ["evenement", "n"]
        fig = px.bar(ev, x="n", y="evenement", orientation="h", height=380,
                     color="n", color_continuous_scale="Blues")
        fig.update_layout(yaxis=dict(autorange="reversed"), coloraxis_showscale=False)
        st.plotly_chart(fig, use_container_width=True)

        top = ev.iloc[0]
        flat = ev.iloc[-1]
        st.info(
            f"**Observation** : « {top['evenement']} » domine avec {top['n']} articles, "
            f"contre {flat['n']} pour « {flat['evenement']} ». La couverture médiatique "
            f"est **fortement biaisée vers les événements américains et le conflit ukrainien**, "
            f"reflétant l'agenda éditorial des sources anglophones majoritaires."
        )

    st.divider()

    # ── Sentiment par orientation ───────────────────────────────────────────
    if "sentiment_vader" in df.columns:
        st.markdown("##### Sentiment moyen (VADER) par orientation")
        sent = df.groupby("orientation").agg(
            sentiment_moyen=("sentiment_vader", "mean"),
            sentiment_median=("sentiment_vader", "median"),
            ecart_type=("sentiment_vader", "std"),
        ).round(3).reset_index()

        cS1, cS2 = st.columns([3, 2])
        with cS1:
            fig = px.bar(sent, x="orientation", y="sentiment_moyen",
                         color="orientation", color_discrete_map=COLORS,
                         text=sent["sentiment_moyen"].apply(lambda v: f"{v:+.3f}"))
            fig.update_traces(textposition="outside")
            fig.add_hline(y=0, line_dash="dash", line_color="gray",
                          annotation_text="Neutre", annotation_position="right")
            fig.update_layout(showlegend=False, height=350)
            st.plotly_chart(fig, use_container_width=True)
        with cS2:
            st.dataframe(sent, hide_index=True, use_container_width=True)

        most_neg = sent.loc[sent["sentiment_moyen"].idxmin(), "orientation"]
        most_pos = sent.loc[sent["sentiment_moyen"].idxmax(), "orientation"]
        st.info(
            f"**Observation** : la classe la plus **négative** est **{most_neg}**, "
            f"la plus **positive** est **{most_pos}**. C'est cohérent avec la littérature : "
            f"les médias d'opposition (qu'ils soient à gauche ou à droite) emploient "
            f"un vocabulaire émotionnellement plus chargé que les médias dits centristes, "
            f"qui adoptent un ton plus factuel et donc plus neutre."
        )

    st.divider()

    # ── Évolution temporelle ────────────────────────────────────────────────
    df["mois"] = df["date"].dt.to_period("M").astype(str)
    monthly = (df.groupby(["mois", "orientation"]).size()
                 .unstack(fill_value=0))
    if not monthly.empty:
        monthly_pct = monthly.div(monthly.sum(axis=1), axis=0) * 100
        st.markdown("##### Évolution mensuelle des orientations (% des articles)")
        fig = go.Figure()
        for orient, color in COLORS.items():
            if orient in monthly_pct.columns:
                fig.add_trace(go.Scatter(
                    x=monthly_pct.index, y=monthly_pct[orient],
                    mode="lines+markers", name=orient,
                    line=dict(color=color, width=2.5),
                ))
        fig.update_layout(yaxis=dict(range=[0, 100]), height=360,
                          legend_orientation="h")
        st.plotly_chart(fig, use_container_width=True)
        st.info(
            "**Observation** : les proportions oscillent fortement d'un mois à l'autre, "
            "reflétant le caractère événementiel du dataset (un sujet géopolitique chaud "
            "peut faire basculer la couverture). Les pics d'articles « Gauche » coïncident "
            "généralement avec les mois où l'actualité humanitaire (réfugiés, climat) domine, "
            "tandis que les pics « Droite » suivent les épisodes sécuritaires (élections, conflits)."
        )

    # ── Synthèse globale ────────────────────────────────────────────────────
    st.divider()
    with st.container(border=True):
        st.markdown("### Synthèse — qualité du dataset pour l'entraînement")
        st.markdown(
            f"""
            - **Volume suffisant** : 3 000 articles permettent une cross-validation 5-fold robuste (chaque fold contient 600 articles).
            - **Déséquilibre maîtrisé** : ratio max/min ≈ {ratio:.1f}×, traité par stratification systématique.
            - **Diversité des sources** : {df['source'].nunique()} sources distinctes empêchent le modèle de mémoriser une signature particulière.
            - **Couverture événementielle** : 10 sujets géopolitiques différents — le modèle apprend à séparer **l'orientation** de **l'événement**.
            - **Limite identifiée** : prédominance anglophone (Kaggle + RSS/NewsAPI internationaux). Performance probablement inférieure sur des médias français exclusifs.
            """
        )


# ════════════════════════════════════════════════════════════════════════════
# PAGE 3 — PERFORMANCE DES MODÈLES
# ════════════════════════════════════════════════════════════════════════════

else:
    st.title("Performance des modèles ML")
    st.caption("Comparaison des 5 algorithmes testés sur le test set (600 articles, split stratifié 80/20).")

    lb = load_leaderboard()
    if lb.empty:
        st.warning("Pas de leaderboard. Exécutez 03_modeling.ipynb.")
        st.stop()

    # Tableau coloré
    st.markdown("##### Classement détaillé")
    st.dataframe(
        lb.style.format(precision=4)
              .background_gradient(subset=["Test F1 Macro"], cmap="Greens")
              .background_gradient(subset=["Test Accuracy"], cmap="Blues"),
        use_container_width=True,
        hide_index=True,
    )

    # ── Verdict ────────────────────────────────────────────────────────────
    winner = lb.sort_values("Test F1 Macro", ascending=False).iloc[0]
    runner = lb.sort_values("Test F1 Macro", ascending=False).iloc[1]
    fastest = lb.sort_values("Temps (s)").iloc[0]
    gap_winner = winner["Test F1 Macro"] - runner["Test F1 Macro"]
    speed_ratio = winner["Temps (s)"] / fastest["Temps (s)"]

    cV1, cV2, cV3 = st.columns(3)
    cV1.metric("Meilleur F1 macro", f"{winner['Test F1 Macro']:.3f}",
               delta=f"+{gap_winner:.3f} vs 2ᵉ")
    cV2.metric("Modèle vainqueur", winner["Modèle"])
    cV3.metric("Modèle le plus rapide", fastest["Modèle"],
               delta=f"{fastest['Temps (s)']:.1f}s")

    st.divider()

    # ── Radar ─────────────────────────────────────────────────────────────
    st.markdown("##### Comparaison radar des 4 métriques")
    metrics = ["Test Accuracy", "Test F1 Macro", "Precision", "Recall"]
    fig = go.Figure()
    for _, row in lb.iterrows():
        fig.add_trace(go.Scatterpolar(
            r=[row[m] for m in metrics] + [row[metrics[0]]],
            theta=metrics + [metrics[0]],
            fill="toself", name=row["Modèle"], opacity=0.55,
        ))
    fig.update_layout(polar=dict(radialaxis=dict(range=[0, 1], tickformat=".0%")),
                      height=480, legend_orientation="h")
    st.plotly_chart(fig, use_container_width=True)

    # ── Trade-off vitesse / qualité ──────────────────────────────────────
    st.markdown("##### Compromis qualité / vitesse")
    fig = px.scatter(
        lb, x="Temps (s)", y="Test F1 Macro",
        size="Test Accuracy", color="Modèle",
        text="Modèle", size_max=40, log_x=True,
        hover_data={"CV F1 (mean)": True, "CV F1 (std)": True},
    )
    fig.update_traces(textposition="top center")
    fig.update_layout(height=420, showlegend=False,
                      xaxis_title="Temps d'entraînement (s, échelle log)",
                      yaxis_title="F1 macro (test)")
    st.plotly_chart(fig, use_container_width=True)
    st.caption(
        "**Lecture** : haut-gauche = idéal (rapide ET performant). "
        "Le SVM linéaire offre le meilleur compromis si la latence importe. "
        "XGBoost paie sa précision en étant ~10× plus lent."
    )

    # ── Stabilité CV ────────────────────────────────────────────────────
    st.markdown("##### Robustesse — stabilité en cross-validation")
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=lb["Modèle"], y=lb["CV F1 (mean)"],
        error_y=dict(type="data", array=lb["CV F1 (std)"]),
        marker_color="steelblue",
        text=lb["CV F1 (mean)"].apply(lambda v: f"{v:.3f}"),
        textposition="outside",
    ))
    fig.add_hline(y=0.33, line_dash="dot", line_color="gray",
                  annotation_text="Baseline aléatoire (33%)",
                  annotation_position="bottom right")
    fig.update_layout(yaxis=dict(range=[0, 1]), height=380, showlegend=False)
    st.plotly_chart(fig, use_container_width=True)
    most_stable = lb.sort_values("CV F1 (std)").iloc[0]
    st.caption(
        f"Tous les modèles dépassent largement la baseline aléatoire (33%). "
        f"Le plus **stable** est **{most_stable['Modèle']}** "
        f"(σ = {most_stable['CV F1 (std)']:.3f}) — variance la plus faible entre les 5 folds."
    )

    st.divider()

    # ── Analyse approfondie ─────────────────────────────────────────────
    st.markdown("### Analyse approfondie")

    tab1, tab2, tab3, tab4 = st.tabs([
        "Pourquoi XGBoost gagne",
        "Pourquoi Random Forest est en retrait",
        "Forces du SVM",
        "Risques & limites",
    ])

    with tab1:
        st.markdown(
            f"""
            **XGBoost obtient F1 = {winner['Test F1 Macro']:.3f}**, devançant le 2ᵉ
            ({runner['Modèle']}) de **{gap_winner*100:.1f} points**.

            Trois raisons :

            1. **Mix texte + numérique** — le pipeline concatène 5 000 features TF-IDF *et*
               11 features numériques (sentiment, subjectivité, scores de vocabulaire G/C/D,
               densité émotionnelle). Les arbres boostés capturent les **interactions
               non-linéaires** entre ces deux familles de signaux. Exemple : un texte avec
               densité émotionnelle élevée + vocabulaire « sécurité » → indicateur Droite
               renforcé. Les modèles linéaires ne voient pas ces interactions.

            2. **Gestion native du déséquilibre** — XGBoost pondère implicitement les
               exemples mal classés à chaque itération de boosting. Le F1 macro en bénéficie.

            3. **Régularisation L1+L2** — le modèle évite le surapprentissage malgré les
               3 000 features sélectionnées (cf. CV F1 = {winner['CV F1 (mean)']:.3f} ± {winner['CV F1 (std)']:.3f},
               proche du score test = {winner['Test F1 Macro']:.3f}).
            """
        )

    with tab2:
        rf = lb[lb["Modèle"].str.contains("Random Forest", na=False)]
        if not rf.empty:
            rf_row = rf.iloc[0]
            st.markdown(
                f"""
                **Random Forest = {rf_row['Test F1 Macro']:.3f}** (5ᵉ position).

                Surprenant ? Pas vraiment :

                - Hyperparamètres bridés : `max_depth=20` et `min_samples_split=5`
                  limitent la profondeur des arbres. Sur 3 000 features, c'est insuffisant
                  pour apprendre les interactions fines que XGBoost capture.
                - Bagging vs Boosting : Random Forest combine des arbres **indépendants**
                  qui votent. XGBoost ajoute des arbres qui **corrigent les erreurs des
                  précédents** — plus efficace pour la classification de texte.
                - Avec 200 arbres profonds non bridés, RF rejoindrait probablement
                  XGBoost mais perdrait son avantage de simplicité.
                """
            )

    with tab3:
        svm = lb[lb["Modèle"].str.contains("SVM", na=False)]
        if not svm.empty:
            svm_row = svm.iloc[0]
            st.markdown(
                f"""
                **SVM Linéaire = {svm_row['Test F1 Macro']:.3f}** (2ᵉ position),
                temps d'entraînement = {svm_row['Temps (s)']:.1f}s soit
                **~{winner['Temps (s)']/svm_row['Temps (s)']:.0f}× plus rapide qu'XGBoost**.

                Le SVM linéaire est le **meilleur choix opérationnel** pour :

                - **API à faible latence** — entraînement et prédiction quasi-instantanés.
                - **Interprétabilité** — les coefficients par classe indiquent directement
                  quels termes pèsent pour chaque orientation.
                - **Stabilité** — CV F1 = {svm_row['CV F1 (mean)']:.3f} ± {svm_row['CV F1 (std)']:.3f},
                  faible variance entre les folds.

                Si le projet devait passer en production sur un grand volume, le SVM
                serait probablement préféré à XGBoost pour son rapport qualité/coût.
                """
            )

    with tab4:
        st.markdown(
            """
            **Limites du protocole d'évaluation actuel :**

            - **Test set unique** (600 articles) → les écarts de quelques points entre
              modèles sont possiblement dans le bruit. La cross-validation 5-fold confirme
              cependant le classement général.
            - **Distribution Kaggle dominante** (50 % du dataset) → le modèle a probablement
              appris des spécificités d'écriture journalistique américaine. Performance non
              garantie sur de la presse française exclusive.
            - **Articles courts en production** → sur des prédictions « live » de phrases
              isolées, le modèle a tendance à sur-prédire la classe dont les marqueurs
              lexicaux sont les plus discriminants individuellement (souvent Droite, à
              cause des mots « sécurité » / « national »).
            - **Pas de dérive temporelle évaluée** — un modèle entraîné en 2026 sur des
              articles 2022–2024 verra ses scores chuter en 2027 si le vocabulaire politique
              évolue (nouveaux événements, nouveaux termes).

            **Pistes d'amélioration mesurables :**

            - Augmenter la diversité francophone (ajouter Mediapart, Le Figaro, Libération
              via NewsAPI ou scraping) → re-mesurer le F1 macro.
            - Calibrer les probabilités d'XGBoost (Platt / isotonic) → améliore la fiabilité
              du « niveau de confiance » affiché.
            - Tester un modèle de langage pré-entraîné multilingue (XLM-R, CamemBERT)
              comme baseline → benchmark pour mesurer la valeur ajoutée du TF-IDF.
            """
        )
