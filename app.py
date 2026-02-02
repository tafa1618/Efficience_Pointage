import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="Efficience Pointage OR",
    layout="wide"
)

st.title("📊 Analyse d'efficience des pointages OR")

# =========================
# Upload fichier
# =========================
uploaded_file = st.file_uploader(
    "📥 Charger le fichier Excel (Pointage + BO)",
    type=["xlsx"]
)

if uploaded_file:

    # =========================
    # Lecture des données
    # =========================
    pointage = pd.read_excel(uploaded_file, sheet_name="Pointage")
    bo = pd.read_excel(uploaded_file, sheet_name="BASE_BO")

    # Nettoyage OR
    pointage["OR"] = pointage["OR"].astype(str)
    bo["N° OR"] = bo["N° OR"].astype(str)

    # =========================
    # Agrégation POINTAGE
    # =========================
    agg_or = (
        pointage
        .groupby("OR")
        .agg(
            Heures_totales_OR=("Hr_travaillée", "sum"),
            Nb_techniciens=("Salarié - Nom", "nunique")
        )
        .reset_index()
    )

    # =========================
    # Technicien principal
    # =========================
    tech_principal = (
        pointage
        .sort_values("Hr_travaillée", ascending=False)
        .drop_duplicates("OR")
        [["OR", "Salarié - Nom", "Salarié - Equipe"]]
        .rename(columns={
            "Salarié - Nom": "Technicien_principal",
            "Salarié - Equipe": "Equipe_principale"
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

    # =========================
    # Préparation BO
    # =========================
    bo["Temps_reference_OR"] = bo["Temps vendu (OR)"].fillna(
        bo["Temps prévu devis (OR)"]
    )

    bo_or = bo[[
        "N° OR",
        "Temps_reference_OR",
        "Durée pointage agents productifs (OR)"
    ]].rename(columns={"N° OR": "OR"})

    # =========================
    # Merge final
    # =========================
    df_final = pointage_or.merge(
        bo_or,
        on="OR",
        how="left"
    )

    # =========================
    # Indicateurs
    # =========================
    df_final["Taux_couverture_OR"] = (
        df_final["Heures_totales_OR"] / df_final["Temps_reference_OR"]
    )

    # =========================
    # KPI globaux
    # =========================
    col1, col2, col3, col4 = st.columns(4)

    col1.metric("OR analysés", df_final.shape[0])
    col2.metric("OR multi-techniciens", df_final[df_final["OR_multi_tech"] == "OUI"].shape[0])
    col3.metric("Heures pointées totales", round(df_final["Heures_totales_OR"].sum(), 1))
    col4.metric(
        "OR sans temps BO",
        df_final["Temps_reference_OR"].isna().sum()
    )

    st.divider()

    # =========================
    # Filtres
    # =========================
    equipe = st.multiselect(
        "Filtrer par équipe",
        options=df_final["Equipe_principale"].dropna().unique()
    )

    if equipe:
        df_final = df_final[df_final["Equipe_principale"].isin(equipe)]

    # =========================
    # Tables
    # =========================
    st.subheader("📋 Vue OR agrégée")
    st.dataframe(
        df_final.sort_values("Heures_totales_OR", ascending=False),
        use_container_width=True
    )

    st.subheader("🔍 Détail OR multi-techniciens")
    st.dataframe(
        df_final[df_final["OR_multi_tech"] == "OUI"],
        use_container_width=True
    )

else:
    st.info("⬆️ Merci de charger le fichier Excel pour démarrer l'analyse.")
