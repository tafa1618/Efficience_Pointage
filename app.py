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
# UPLOAD
# ======================================================
uploaded_file = st.file_uploader(
    "📥 Charger le fichier Excel (Pointage + BASE_BO)",
    type=["xlsx"]
)

if not uploaded_file:
    st.info("⬆️ Charge le fichier Excel pour démarrer l’analyse.")
    st.stop()

# ======================================================
# LECTURE FICHIERS
# ======================================================
pointage = pd.read_excel(uploaded_file, sheet_name="Pointage")
bo = pd.read_excel(uploaded_file, sheet_name="BASE_BO")

# ======================================================
# FONCTION NORMALISATION OR (CLÉ MÉTIER ROBUSTE)
# ======================================================
def normalize_or(x):
    if pd.isna(x):
        return None
    x = str(x).strip()
    x = x.split("-")[0]
    x = x.split("/")[0]
    x = x.replace(".0", "")
    x = "".join(filter(str.isdigit, x))
    return x if x != "" else None

# ======================================================
# NORMALISATION POINTAGE
# ======================================================
pointage["OR_RAW"] = pointage["OR (Numéro)"].astype(str)
pointage["OR_KEY"] = pointage["OR_RAW"].apply(normalize_or)

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
bo["OR_RAW"] = bo["N° OR (Segment)"].astype(str)
bo["OR_KEY"] = bo["OR_RAW"].apply(normalize_or)

bo["Temps_reference_OR"] = bo["Temps vendu (OR)"].fillna(
    bo["Temps prévu devis (OR)"]
)

# ======================================================
# FILTRE ANNÉE (ANALYSE UNIQUEMENT)
# ======================================================
annees_disponibles = sorted(pointage["Annee"].dropna().unique())

annees_selectionnees = st.multiselect(
    "📅 Filtrer par année",
    options=annees_disponibles,
    default=annees_disponibles
)

pointage = pointage[pointage["Annee"].isin(annees_selectionnees)]

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
# TECHNICIEN PRINCIPAL (MAX HEURES)
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
# PRÉPARATION BO POUR MERGE
# ======================================================
bo_or = bo[[
    "OR_KEY",
    "Temps_reference_OR",
    "Durée pointage agents productifs (OR)"
]]

# ======================================================
# MERGE FINAL (CLÉ ROBUSTE)
# ======================================================
df = pointage_or.merge(
    bo_or,
    on="OR_KEY",
    how="left"
)

# ======================================================
# INDICATEURS
# ======================================================
df["Taux_couverture_OR"] = (
    df["Heures_totales_OR"] / df["Temps_reference_OR"]
)

df["Ecart_heures"] = (
    df["Heures_totales_OR"] - df["Temps_reference_OR"]
)

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
# GRAPHIQUES – PILOTAGE
# ======================================================
st.subheader("📊 Lecture rapide – Efficience")

col1, col2 = st.columns(2)

# Heures par équipe
heures_equipe = (
    df.groupby("Equipe_principale")["Heures_totales_OR"]
    .sum()
    .sort_values(ascending=False)
)
col1.bar_chart(heures_equipe)

# Taux couverture moyen par équipe
taux_equipe = (
    df.groupby("Equipe_principale")["Taux_couverture_OR"]
    .mean()
    .sort_values(ascending=False)
)
col2.bar_chart(taux_equipe)

st.divider()

# ======================================================
# Efficience par technicien principal
# ======================================================
st.subheader("👷‍♂️ Efficience par technicien (principal)")

taux_tech = (
    df.groupby("Technicien_principal")["Taux_couverture_OR"]
    .mean()
    .sort_values()
)

st.bar_chart(taux_tech)

st.divider()

# ======================================================
# Pareto OR en dérive
# ======================================================
st.subheader("⚠️ Pareto des OR en surconsommation")

pareto = (
    df[df["Ecart_heures"] > 0]
    .sort_values("Ecart_heures", ascending=False)
    .set_index("OR_KEY")["Ecart_heures"]
)

st.bar_chart(pareto)

st.divider()

# ======================================================
# TABLES (EXPORT / AUDIT)
# ======================================================
st.subheader("📋 Table OR agrégée (export)")

st.dataframe(
    df.sort_values("Heures_totales_OR", ascending=False),
    use_container_width=True
)
