import streamlit as st
import pandas as pd
import duckdb

# ======================================================
# CONFIG & STYLE
# ======================================================
st.set_page_config(page_title="Efficience OR - SQL Engine", layout="wide")
st.title("📊 Analyse d’efficience (Moteur SQL)")

# ======================================================
# UPLOAD
# ======================================================
uploaded_file = st.file_uploader(
    "📥 Charger le fichier Excel (doit contenir 'Pointage' et 'BASE_BO')",
    type=["xlsx"]
)

if not uploaded_file:
    st.info("Veuillez charger un fichier Excel pour démarrer l'analyse.")
    st.stop()

# ======================================================
# CHARGEMENT INITIAL (RAW)
# ======================================================
@st.cache_data
def load_raw_data(file):
    df_p = pd.read_excel(file, sheet_name="Pointage")
    df_b = pd.read_excel(file, sheet_name="BASE_BO")
    return df_p, df_b

df_pointage_raw, df_bo_raw = load_raw_data(uploaded_file)

# ======================================================
# MOTEUR SQL (DUCKDB)
# ======================================================
# Connexion à une base en mémoire
con = duckdb.connect(database=':memory:')

# Enregistrement des DataFrames comme tables SQL
con.register('raw_pointage', df_pointage_raw)
con.register('raw_bo', df_bo_raw)

# ------------------------------------------------------
# SQL : NETTOYAGE ET NORMALISATION
# ------------------------------------------------------
# On crée des vues SQL pour nettoyer les clés OR
con.execute("""
    CREATE OR REPLACE VIEW view_pointage AS
    SELECT 
        regexp_replace(split_part(split_part(CAST("OR (Numéro)" AS VARCHAR), '-', 1), '/', 1), '[^0-9]', '', 'g') AS OR_KEY,
        "Salarié - Nom" AS Technicien,
        "Salarié - Equipe(Nom)" AS Equipe,
        Hr_travaillée AS Heures,
        strptime(CAST("Saisie heures - Date" AS VARCHAR), '%Y-%m-%d %H:%M:%S') AS Date_Saisie
    FROM raw_pointage;

    CREATE OR REPLACE VIEW view_bo AS
    SELECT 
        regexp_replace(split_part(split_part(CAST("N° OR (Segment)" AS VARCHAR), '-', 1), '/', 1), '[^0-9]', '', 'g') AS OR_KEY,
        COALESCE("Temps vendu (OR)", "Temps prévu devis (OR)") AS Temps_Ref
    FROM raw_bo;
""")

# ------------------------------------------------------
# SQL : AGRÉGATION ET JOINTURE (Le coeur du problème)
# ------------------------------------------------------
query = """
WITH agg_pointage AS (
    -- On calcule le total d'heures par OR et on identifie le tech principal (celui qui a fait le plus d'heures)
    SELECT 
        OR_KEY,
        SUM(Heures) as Heures_Totales,
        COUNT(DISTINCT Technicien) as Nb_Tech,
        -- Astuce SQL pour chopper le tech principal
        ARGMAX(Technicien, Heures) as Tech_Principal,
        ARGMAX(Equipe, Heures) as Equipe_Principale,
        MIN(Date_Saisie) as Date_Debut
    FROM view_pointage
    WHERE OR_KEY IS NOT NULL AND OR_KEY != ''
    GROUP BY OR_KEY
),
agg_bo AS (
    -- On s'assure d'avoir une seule ligne par OR dans le BO (on prend le temps max si doublon)
    SELECT 
        OR_KEY,
        MAX(Temps_Ref) as Temps_Reference
    FROM view_bo
    WHERE OR_KEY IS NOT NULL AND OR_KEY != ''
    GROUP BY OR_KEY
)
SELECT 
    p.*,
    b.Temps_Reference,
    CASE WHEN b.Temps_Reference IS NULL THEN 'OUI' ELSE 'NON' END as Manquant_BO,
    (b.Temps_Reference / NULLIF(p.Heures_Totales, 0)) * 100 as Efficience_Pct
FROM agg_pointage p
LEFT JOIN agg_bo b ON p.OR_KEY = b.OR_KEY
"""

df_final = con.execute(query).df()

# ======================================================
# FILTRES SIDEBAR (Sur le résultat SQL)
# ======================================================
df_final['Annee'] = pd.to_datetime(df_final['Date_Debut']).dt.year
annees = sorted(df_final['Annee'].dropna().unique().astype(int))
sel_annees = st.sidebar.multiselect("Filtrer par Année", options=annees, default=annees)

df_filtered = df_final[df_final['Annee'].isin(sel_annees)]

# ======================================================
# AFFICHAGE
# ======================================================
c1, c2, c3, c4 = st.columns(4)
c1.metric("OR Total", len(df_filtered))
c2.metric("Heures Total", f"{df_filtered['Heures_Totales'].sum():.1f}h")
c3.metric("Sans BO", df_filtered['Temps_Reference'].isna().sum())
c4.metric("Efficience Moy.", f"{df_filtered['Efficience_Pct'].mean():.1f}%")

st.divider()

col_left, col_right = st.columns(2)

with col_left:
    st.subheader("📊 Top 10 Équipes (Efficience)")
    eff_eq = df_filtered.groupby("Equipe_Principale")["Efficience_Pct"].mean().sort_values(ascending=False).head(10)
    st.bar_chart(eff_eq)

with col_right:
    st.subheader("⚠️ Top Erreurs (Heures sans Temps Ref)")
    manquants = df_filtered[df_filtered['Temps_Reference'].isna()].sort_values("Heures_Totales", ascending=False).head(10)
    st.write(manquants[['OR_KEY', 'Equipe_Principale', 'Heures_Totales']])

st.subheader("📋 Analyse détaillée par OR")
st.dataframe(df_filtered, use_container_width=True)
