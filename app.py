import streamlit as st
import pandas as pd
import numpy as np

# ======================================================
# CONFIG
# ======================================================
st.set_page_config(page_title="Efficience OR", layout="wide")
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
# NORMALISATION OR
# ======================================================
def normalize_or(x):
    if pd.isna(x):
        return None
    x = str(x)
    x = x.split("-")[0]
    x = x.split("/")[0]
    x = x.replace(".0", "")
    x = "".join(filter(str.isdigit, x))
    return x if x else None

# ================= POINTAGE =================
pointage["OR_KEY"] = pointage["OR (Numéro)"].apply(normalize_or)
pointage["Technicien"] = pointage["Salarié - Nom"]
pointage["Equipe"] = pointage["Salarié - Equipe(Nom)"]
pointage["Heures"] = pointage["Hr_travaillée"]

pointage["Date"] = pd.to_datetime(
    pointage["Saisie heures - Date"], errors="coerce"
)
pointage["Annee"] = pointage["Date"].dt.year

# ================= BO =================
bo["OR_KEY"] = bo["N° OR (Segment)"].apply(normalize_or)

bo["Temps_reference_OR"] = bo["Temps vendu (OR)"].fillna(
    bo["Temps prévu devis (OR)"]
)

# 🔴 FIX MAJEUR : BO → 1 OR = 1 ligne
bo_or = (
    bo.groupby("OR_KEY", as_index=False)
    .agg(
        Temps_reference_OR=("Temps_reference_OR", "max")
    )
)

# ======================================================
# FILTRE ANNÉE
# ======================================================
annees = sorted(pointage["Annee"].dropna().unique())
annees_sel = st.sidebar.multiselect(
    "Année",
    options=annees,
    default=annees
)

pointage = pointage[pointage["Annee"].isin(annees_sel)]

# ======================================================
# AGRÉGATION POINTAGE → 1 OR
# ======================================================
agg_or = (
    pointage
    .groupby("OR_KEY", as_index=False)
    .agg(
        Heures_totales_OR=("Heures", "sum"),
        Nb_techniciens=("Technicien", "nunique")
    )
)

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

pointage_or = agg_or.merge(tech_principal, on="OR_KEY", how="left")
pointage_or["OR_multi_tech"] = np.where(
    pointage_or["Nb_techniciens"] > 1, "OUI", "NON"
)

# ======================================================
# MERGE FINAL (PROPRE)
# ======================================================
df = pointage_or.merge(
    bo_or,
    on="OR_KEY",
    how="left"
)

# ======================================================
# CALCUL EFFICIENCE (SAIN)
# ======================================================
df["Efficience_%"] = np.where(
    (df["Heures_totales_OR"] > 0) & (df["Temps_reference_OR"] > 0),
    (df["Temps_reference_OR"] / df["Heures_totales_OR"]) * 100,
    np.nan
)

# ======================================================
# KPI
# ======================================================
st.subheader("📌 Indicateurs globaux")

c1, c2, c3, c4 = st.columns(4)

c1.metric("OR analysés", df.shape[0])
c2.metric("OR multi-tech", df[df["OR_multi_tech"] == "OUI"].shape[0])
c3.metric("Heures pointées", round(df["Heures_totales_OR"].sum(), 1))
c4.metric("OR sans BO", df["Temps_reference_OR"].isna().sum())

st.divider()

# ======================================================
# GRAPHIQUES
# ======================================================
st.subheader("📊 Efficience par équipe")

efficience_equipe = (
    df.dropna(subset=["Efficience_%"])
    .groupby("Equipe_principale")["Efficience_%"]
    .mean()
    .sort_values(ascending=False)
)

st.bar_chart(efficience_equipe)

st.divider()

# ======================================================
# TABLE FINALE
# ======================================================
st.subheader("📋 Table OR agrégée")

st.dataframe(
    df.sort_values("Efficience_%", ascending=False),
    use_container_width=True
)
