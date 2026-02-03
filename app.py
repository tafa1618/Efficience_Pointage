import streamlit as st
import pandas as pd
import duckdb
import numpy as np

# ======================================================
# CONFIGURATION & STYLE
# ======================================================
st.set_page_config(page_title="Efficience OR - Engine v2", layout="wide")
st.title("📊 Analyse d’efficience des pointages OR")
st.markdown("---")

# ======================================================
# CHARGEMENT DES DONNÉES
# ======================================================
uploaded_file = st.file_uploader(
    "📥 Charger le fichier Excel (Onglets: Pointage + BASE_BO)",
    type=["xlsx"]
)

if not uploaded_file:
    st.info("En attente du fichier Excel pour démarrer l'analyse.")
    st.stop()

@st.cache_data
def get_data(file):
    df_p = pd.read_excel(file, sheet_name="Pointage")
    df_b = pd.read_excel(file, sheet_name="BASE_BO")
    return df_p, df_b

try:
    df_p_raw, df_b_raw = get_data(uploaded_file)
except Exception as e:
    st.error(f"Erreur lors de la lecture des onglets : {e}")
    st.stop()

# ======================================================
# MOTEUR DE TRANSFORMATION SQL (DUCKDB)
# ======================================================
con = duckdb.connect(database=':memory:')
con.register('raw_p', df_p_raw)
con.register('raw_b', df_b_raw)

# Normalisation agressive des clés OR
# 1. On garde uniquement les chiffres
# 2. On enlève les zéros au début (ex: 00123 -> 123)
# 3. On gère les extensions type /01 ou -A
con.execute("""
    CREATE OR REPLACE VIEW v_pointage AS
    SELECT 
        ltrim(regexp_replace(split_part(split_part(CAST("OR (Numéro)" AS VARCHAR), '-', 1), '/', 1), '[^0-9]', '', 'g'), '0') AS OR_KEY_CLEAN,
        "Salarié - Nom" AS Technicien,
        "Salarié - Equipe(Nom)" AS Equipe,
        Hr_travaillée AS Heures,
        CAST("Saisie heures - Date" AS DATE) AS Date_Pointage
    FROM raw_p
    WHERE "OR (Numéro)" IS NOT NULL;

    CREATE OR REPLACE VIEW v_bo AS
    SELECT 
        ltrim(regexp_replace(split_part(split_part(CAST("N° OR (Segment)" AS VARCHAR), '-', 1), '/', 1), '[^0-9]', '', 'g'), '0') AS OR_KEY_CLEAN,
        COALESCE("Temps vendu (OR)", "Temps prévu devis (OR)") AS Temps_Ref
    FROM raw_b
    WHERE "N° OR (Segment)" IS NOT NULL;
""")

# Agrégation et Jointure
query = """
WITH p_agg AS (
    SELECT 
        OR_KEY_CLEAN,
        SUM(Heures) as Heures_Totales,
        COUNT(DISTINCT Technicien) as Nb_Tech,
        ARGMAX(Technicien, Heures) as Tech_Principal,
        ARGMAX(Equipe, Heures) as Equipe_Principale,
        YEAR(MIN(Date_Pointage)) as Annee
    FROM v_pointage
    GROUP BY OR_KEY_CLEAN
),
b_agg AS (
    SELECT 
        OR_KEY_CLEAN,
        MAX(Temps_Ref) as Temps_Reference
    FROM v_bo
    GROUP BY OR_KEY_CLEAN
)
SELECT 
    p.*,
    b.Temps_Reference,
    CASE 
        WHEN b.Temps_Reference IS NULL THEN 0 
        WHEN p.Heures_Totales = 0 THEN 0
        ELSE (b.Temps_Reference / p.Heures_Totales) * 100 
    END as Efficience_Pct
FROM p_agg p
LEFT JOIN b_agg b ON p.OR_KEY_CLEAN = b.OR_KEY_CLEAN
"""

df_final = con.execute(query).df()

# ======================================================
# FILTRES & DASHBOARD
# ======================================================
# Sidebar
annees = sorted(df_final['Annee'].dropna().unique().astype(int))
sel_annees = st.sidebar.multiselect("Années", options=annees, default=annees)

# Application du filtre
df_view = df_final[df_final['Annee'].isin(sel_annees)].copy()

# Calcul des indicateurs
total_or = len(df_view)
matched_or = df_view['Temps_Reference'].notna().sum()
unmatched_or = total_or - matched_or

# Affichage des KPIs
st.subheader("📌 Indicateurs clés")
c1, c2, c3, c4 = st.columns(4)
c1.metric("OR Analysés", total_or)
c2.metric("Heures Pointées", f"{df_view['Heures_Totales'].sum():.1f}h")
c3.metric("OR avec Temps BO", matched_or)
c4.metric("Efficience Moyenne", f"{df_view[df_view['Temps_Reference'].notna()]['Efficience_Pct'].mean():.1f}%")

# Mode Debug si rien ne matche
if matched_or == 0 and total_or > 0:
    st.error("⚠️ AUCUNE CORRESPONDANCE TROUVÉE")
    st.warning("Vérifiez les formats de vos numéros d'OR ci-dessous :")
    db_col1, db_col2 = st.columns(2)
    with db_col1:
        st.write("Exemple clés Pointage :")
        st.write(con.execute("SELECT DISTINCT OR_KEY_CLEAN FROM v_pointage LIMIT 5").df())
    with db_col2:
        st.write("Exemple clés BO :")
        st.write(con.execute("SELECT DISTINCT OR_KEY_CLEAN FROM v_bo LIMIT 5").df())
    st.stop()

st.divider()

# ======================================================
# VISUALISATION
# ======================================================
col_left, col_right = st.columns(2)

with col_left:
    st.subheader("📊 Efficience par équipe")
    eff_eq = df_view.dropna(subset=['Temps_Reference']).groupby("Equipe_Principale")["Efficience_Pct"].mean().sort_values()
    st.bar_chart(eff_eq, horizontal=True)

with col_right:
    st.subheader("📋 Répartition Multi-Tech")
    multi_counts = df_view['Nb_Tech'].value_counts()
    st.bar_chart(multi_counts)

st.subheader("🔍 Détail des calculs par OR")
st.dataframe(
    df_view.sort_values("Efficience_Pct", ascending=False),
    use_container_width=True,
    column_config={
        "Efficience_Pct": st.column_config.ProgressColumn("Efficience %", format="%.1f%%", min_value=0, max_value=200),
        "Temps_Reference": "Temps Vendu (h)",
        "Heures_Totales": "Heures Réelles (h)"
    }
)

# Export option
st.download_button(
    "📥 Télécharger les résultats (CSV)",
    df_view.to_csv(index=False).encode('utf-8'),
    "analyse_efficience.csv",
    "text/csv"
)
