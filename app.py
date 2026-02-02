import streamlit as st
import pandas as pd

# ======================================================
# CONFIG
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
    st.stop()

# ======================================================
# LECTURE
# ======================================================
pointage = pd.read_excel(uploaded_file, sheet_name="Pointage")
bo = pd.read_excel(uploaded_file, sheet_name="BASE_BO")

# ======================================================
# NORMALISATION
# ======================================================
pointage["OR"] = pointage["OR (Numéro)"].astype(str).str.strip()
pointage["Technicien"] = pointage["Salarié - Nom"]
pointage["Equipe"] = pointage["Salarié - Equipe(Nom)"]
pointage["Heures"] = pointage["Hr_travaillée"]

pointage["Date"] = pd.to_datetime(
    pointage["Saisie heures - Date"],
    errors="coerce"
)
pointage["Annee"] = pointage["Date"].dt.year

bo["OR"] = (
    bo["N° OR (Segment)"]
    .astype(str)
    .str.strip()
    .str.split("-")
    .str[0]
)

bo["Temps_reference_OR"] = bo["Temps vendu (OR)"].fillna(
    bo["Temps prévu devis (OR)"]
)

# ======================================================
# FILTRE ANNÉE
# ======================================================
annees = sorted(pointage["Annee"].dropna().unique())
annees_sel = st.multiselect(
    "📅 Filtrer par année",
    options=annees,
    default=annees
)

pointage = pointage[pointage["Annee"].isin(annees_sel)]

# ======================================================
# AGRÉGATION OR
# ======================================================
agg_or = (
    pointage
    .groupby("OR")
    .agg(
        Heures_totales_OR=("Heures", "sum"),
        Nb_techniciens=("Technicien", "nunique")
    )
    .reset_index()
)

tech_principal = (
    pointage
    .sort_values("Heures", ascending=False)
    .drop_duplicates("OR")
    [["OR", "Technicien", "Equipe"]]
    .rename(columns={
        "Technicien": "Technicien_principal",
        "Equipe": "Equipe_principale"
    })
)

pointage_or = agg_or.merge(tech_principal, on="OR", how="left")
pointage_or["OR_multi_tech"] = pointage_or["Nb_techniciens"].apply(
    lambda x: "OUI" if x > 1 else "NON"
)

# ======================================================
# MERGE BO
# ======================================================
bo_or = bo[[
    "OR",
    "Temps_reference_OR",
    "Durée pointage agents productifs (OR)"
]]

df = pointage_or.merge(bo_or, on="OR", how="left")
df["Taux_couverture_OR"] = df["Heures_totales_OR"] / df["Temps_reference_OR"]
df["Ecart_heures"] = df["Heures_totales_OR"] - df["Temps_reference_OR"]

# ======================================================
# KPI
# ======================================================
c1, c2, c3, c4 = st.columns(4)
c1.metric("OR analysés", df.shape[0])
c2.metric("OR multi-tech", df[df["OR_multi_tech"] == "OUI"].shape[0])
c3.metric("Heures pointées", round(df["Heures_totales_OR"].sum(), 1))
c4.metric("OR sans BO", df["Temps_reference_OR"].isna().sum())

st.divider()

# ======================================================
# 📊 GRAPHIQUES
# ======================================================
st.subheader("📊 Lecture rapide – Pilotage")

col_g1, col_g2 = st.columns(2)

# 1️⃣ Heures par équipe
heures_equipe = (
    df.groupby("Equipe_principale")["Heures_totales_OR"]
    .sum()
    .sort_values(ascending=False)
)

col_g1.bar_chart(heures_equipe)

# 2️⃣ Taux de couverture moyen par équipe
taux_equipe = (
    df.groupby("Equipe_principale")["Taux_couverture_OR"]
    .mean()
    .sort_values(ascending=False)
)

col_g2.bar_chart(taux_equipe)

st.divider()

# 3️⃣ Top / Flop techniciens
st.subheader("👷‍♂️ Efficience par technicien (technicien principal)")

taux_tech = (
    df.groupby("Technicien_principal")["Taux_couverture_OR"]
    .mean()
    .sort_values()
)

st.bar_chart(taux_tech)

st.divider()

# 4️⃣ Pareto OR non couverts
st.subheader("⚠️ Pareto des OR en dérive")

pareto = (
    df[df["Ecart_heures"] > 0]
    .sort_values("Ecart_heures", ascending=False)
    .set_index("OR")["Ecart_heures"]
)

st.bar_chart(pareto)

st.divider()

# ======================================================
# TABLES (EXPORT)
# ======================================================
st.subheader("📋 Table OR agrégée")
st.dataframe(df, use_container_width=True)
