import streamlit as st
import pandas as pd

# ======================================================
# CONFIG PAGE
# ======================================================
st.set_page_config(
    page_title="Efficience des pointages OR",
    layout="wide"
)

st.title("📊 Analyse d’efficience des pointages OR")

# ======================================================
# UPLOAD FICHIER
# ======================================================
uploaded_file = st.file_uploader(
    "📥 Charger le fichier Excel (Pointage + BASE_BO)",
    type=["xlsx"]
)

if not uploaded_file:
    st.info("⬆️ Charge le fichier Excel pour démarrer l’analyse.")
    st.stop()

# ======================================================
# LECTURE DES FEUILLES
# ======================================================
pointage = pd.read_excel(uploaded_file, sheet_name="Pointage")
bo = pd.read_excel(uploaded_file, sheet_name="BASE_BO")

# ======================================================
# FONCTION NORMALISATION OR (clé métier robuste)
# ======================================================
def normalize_or(x):
    if pd.isna(x):
        return None
    x = str(x).strip()
    x = x.split("-")[0]
    x = x.split("/")[0]
    x = x.replace(".0", "")
    x = "".join(filter(str.isdigit, x))
    return x if x else None

# ======================================================
# NORMALISATION POINTAGE
# ======================================================
pointage["OR_KEY"] = pointage["OR (Numéro)"].apply(normalize_or)
pointage["Technicien"] = pointage["Salarié - Nom"]
pointage["Equipe"] = pointage["Salarié - Equipe(Nom)"]
pointage["Heures"] = pointage["Hr_travaillée"]

pointage["Date"] = pd.to_datetime(
    pointage["Saisie heures - Date"],
    errors="coerce"
)
pointage["Annee"] = pointage["Date"].dt.year

# ======================================================
# NORMALISATION BO
# ======================================================
bo["OR_KEY"] = bo["N° OR (Segment)"].apply(normalize_or)

bo["Temps_reference_OR"] = bo["Temps vendu (OR)"].fillna(
    bo["Temps prévu devis (OR)"]
)

# ======================================================
# FILTRES SIDEBAR
# ======================================================
st.sidebar.header("🎯 Filtres d’analyse")

# ---- Filtre année
annees = sorted(pointage["Annee"].dropna().unique())
annees_sel = st.sidebar.multiselect(
    "Année",
    options=annees,
    default=annees
)

pointage = pointage[pointage["Annee"].isin(annees_sel)]

# ======================================================
# AGRÉGATION POINTAGE → 1 OR = 1 LIGNE
# ======================================================
agg_or = (
    pointage
    .groupby("OR_KEY")
    .agg(
        Heures_totales_OR=("Heures", "sum"),
        Nb_techniciens=("Technicien", "nunique")
    )
    .reset_index()
)

# ======================================================
# TECHNICIEN & ÉQUIPE PRINCIPALE (max heures)
# ======================================================
tech_principal = (
    pointage
    .sort_values("Heures", ascending=False)
    .drop_duplicates("OR_KEY")
    [["OR_KEY", "Technicien", "Equipe"]]
    .rename(columns={
        "Technicien": "Technicien_principal",
        "Equipe": "Equipe_principale"
    })
)

pointage_or = agg_or.merge(
    tech_principal,
    on="OR_KEY",
    how="left"
)

pointage_or["OR_multi_tech"] = pointage_or["Nb_techniciens"].apply(
    lambda x: "OUI" if x > 1 else "NON"
)

# ======================================================
# MERGE AVEC BO
# ======================================================
bo_or = bo[["OR_KEY", "Temps_reference_OR"]]

df = pointage_or.merge(
    bo_or,
    on="OR_KEY",
    how="left"
)

# ======================================================
# CALCULS D’EFFICIENCE
# ======================================================
# Ratio brut (diagnostic)
df["Taux_couverture"] = df["Heures_totales_OR"] / df["Temps_reference_OR"]

# 🔥 Efficience (%) — peut être > 100 %
df["Efficience_%"] = (
    df["Temps_reference_OR"] / df["Heures_totales_OR"]
) * 100

df.loc[df["Heures_totales_OR"] <= 0, "Efficience_%"] = None

# ======================================================
# FILTRE ÉQUIPE (APRÈS MERGE)
# ======================================================
equipes = sorted(df["Equipe_principale"].dropna().unique())

equipes_sel = st.sidebar.multiselect(
    "Équipe",
    options=equipes,
    default=equipes
)

df = df[df["Equipe_principale"].isin(equipes_sel)]

# ======================================================
# KPI GLOBAUX
# ======================================================
st.subheader("📌 Indicateurs globaux")

c1, c2, c3, c4 = st.columns(4)

c1.metric("OR analysés", df.shape[0])
c2.metric("OR multi-techniciens", df[df["OR_multi_tech"] == "OUI"].shape[0])
c3.metric("Heures pointées", round(df["Heures_totales_OR"].sum(), 1))
c4.metric("OR sans BO", df["Temps_reference_OR"].isna().sum())

st.divider()

# ======================================================
# 📊 GRAPHIQUES – PILOTAGE
# ======================================================
st.subheader("📊 Pilotage de l’efficience")

col1, col2 = st.columns(2)

# Heures par équipe
heures_equipe = (
    df.groupby("Equipe_principale")["Heures_totales_OR"]
    .sum()
    .sort_values(ascending=False)
)
col1.bar_chart(heures_equipe)

# 🔥 Efficience moyenne par équipe
efficience_equipe = (
    df.groupby("Equipe_principale")["Efficience_%"]
    .mean()
    .sort_values(ascending=False)
)
col2.bar_chart(efficience_equipe)

st.divider()

# ======================================================
# Efficience par technicien principal
# ======================================================
st.subheader("👷‍♂️ Efficience par technicien (principal)")

efficience_tech = (
    df.groupby("Technicien_principal")["Efficience_%"]
    .mean()
    .sort_values()
)

st.bar_chart(efficience_tech)

st.divider()

# ======================================================
# TABLE EXPORT / AUDIT
# ======================================================
st.subheader("📋 Table OR agrégée (export)")

st.dataframe(
    df.sort_values("Efficience_%"),
    use_container_width=True
)
