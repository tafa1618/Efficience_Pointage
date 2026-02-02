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
# LECTURE DES DONNÉES
# ======================================================
pointage = pd.read_excel(uploaded_file, sheet_name="Pointage")
bo = pd.read_excel(uploaded_file, sheet_name="BASE_BO")

# ======================================================
# HARMONISATION DES COLONNES (STRICTEMENT SELON TES NOMS)
# ======================================================
# Pointage
pointage["OR"] = pointage["OR (Numéro)"].astype(str)
pointage["Technicien"] = pointage["Salarié - Nom"]
pointage["Equipe"] = pointage["Salarié - Equipe(Nom)"]
pointage["Heures"] = pointage["Hr_travaillée"]

# BO
bo["OR"] = bo["N° OR (Segment)"].astype(str)
bo["Temps_reference_OR"] = bo["Temps vendu (OR)"].fillna(
    bo["Temps prévu devis (OR)"]
)

# ======================================================
# TABLE 1 — POINTAGE OR AGRÉGÉ (1 OR = 1 LIGNE)
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

# ======================================================
# TECHNICIEN PRINCIPAL (celui qui a le + d’heures)
# ======================================================
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

pointage_or = agg_or.merge(
    tech_principal,
    on="OR",
    how="left"
)

pointage_or["OR_multi_tech"] = pointage_or["Nb_techniciens"].apply(
    lambda x: "OUI" if x > 1 else "NON"
)

# ======================================================
# TABLE 2 — POINTAGE OR x TECH (détail opérationnel)
# ======================================================
pointage_or_tech = (
    pointage
    .groupby(["OR", "Technicien", "Equipe"])
    .agg(
        Heures_technicien=("Heures", "sum")
    )
    .reset_index()
)

pointage_or_tech = pointage_or_tech.merge(
    pointage_or[["OR", "Heures_totales_OR"]],
    on="OR",
    how="left"
)

pointage_or_tech["Part_OR_%"] = (
    pointage_or_tech["Heures_technicien"]
    / pointage_or_tech["Heures_totales_OR"]
) * 100

# ======================================================
# MERGE AVEC BO (APRÈS calculs pointage)
# ======================================================
bo_or = bo[[
    "OR",
    "Temps_reference_OR",
    "Durée pointage agents productifs (OR)"
]]

df_final = pointage_or.merge(
    bo_or,
    on="OR",
    how="left"
)

# ======================================================
# INDICATEURS
# ======================================================
df_final["Taux_couverture_OR"] = (
    df_final["Heures_totales_OR"] / df_final["Temps_reference_OR"]
)

# ======================================================
# KPI GLOBAUX
# ======================================================
st.subheader("📌 Indicateurs globaux")

c1, c2, c3, c4 = st.columns(4)

c1.metric("OR analysés", df_final.shape[0])
c2.metric("OR multi-techniciens", df_final[df_final["OR_multi_tech"] == "OUI"].shape[0])
c3.metric("Heures pointées totales", round(df_final["Heures_totales_OR"].sum(), 1))
c4.metric("OR sans temps BO", df_final["Temps_reference_OR"].isna().sum())

st.divider()

# ======================================================
# FILTRES
# ======================================================
st.subheader("🎯 Filtres")

equipes = st.multiselect(
    "Filtrer par équipe",
    options=sorted(df_final["Equipe_principale"].dropna().unique())
)

if equipes:
    df_final = df_final[df_final["Equipe_principale"].isin(equipes)]
    pointage_or_tech = pointage_or_tech[
        pointage_or_tech["Equipe"].isin(equipes)
    ]

# ======================================================
# AFFICHAGE TABLES
# ======================================================
st.subheader("📋 Vue OR agrégée (pilotage)")

st.dataframe(
    df_final.sort_values("Heures_totales_OR", ascending=False),
    use_container_width=True
)

st.subheader("👥 Détail OR × Technicien (opérationnel)")

st.dataframe(
    pointage_or_tech.sort_values("Heures_technicien", ascending=False),
    use_container_width=True
)

st.subheader("⚠️ OR multi-techniciens")

st.dataframe(
    df_final[df_final["OR_multi_tech"] == "OUI"],
    use_container_width=True
)
