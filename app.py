import streamlit as st
import pandas as pd
import duckdb

# ======================================================
# CONFIGURATION
# ======================================================
st.set_page_config(page_title="Efficience OR - Final Fix", layout="wide")
st.title("📊 Analyse d’efficience (Mode Extraction Partielle)")

uploaded_file = st.file_uploader("📥 Charger le fichier Excel", type=["xlsx"])

if not uploaded_file:
    st.info("Téléchargez votre fichier pour commencer.")
    st.stop()

@st.cache_data
def load_data(file):
    return pd.read_excel(file, sheet_name="Pointage"), pd.read_excel(file, sheet_name="BASE_BO")

df_p_raw, df_b_raw = load_data(uploaded_file)

# ======================================================
# MOTEUR SQL : STRATÉGIE DE CORRESPONDANCE DE FIN DE CHAÎNE
# ======================================================
con = duckdb.connect(database=':memory:')
con.register('raw_p', df_p_raw)
con.register('raw_b', df_b_raw)

# On nettoie et on ne garde que les 5 DERNIERS CHIFFRES pour le match
con.execute("""
    CREATE OR REPLACE VIEW v_p AS
    SELECT 
        regexp_replace(CAST("OR (Numéro)" AS VARCHAR), '[^0-9]', '', 'g') as OR_FULL,
        right(regexp_replace(CAST("OR (Numéro)" AS VARCHAR), '[^0-9]', '', 'g'), 5) AS OR_KEY,
        "Salarié - Nom" AS Technicien,
        "Salarié - Equipe(Nom)" AS Equipe,
        Hr_travaillée AS Heures,
        CAST("Saisie heures - Date" AS DATE) AS Date_Pointage
    FROM raw_p
    WHERE "OR (Numéro)" IS NOT NULL;

    CREATE OR REPLACE VIEW v_b AS
    SELECT 
        regexp_replace(CAST("N° OR (Segment)" AS VARCHAR), '[^0-9]', '', 'g') as OR_FULL,
        right(regexp_replace(CAST("N° OR (Segment)" AS VARCHAR), '[^0-9]', '', 'g'), 5) AS OR_KEY,
        COALESCE("Temps vendu (OR)", "Temps prévu devis (OR)") AS Temps_Ref
    FROM raw_b
    WHERE "N° OR (Segment)" IS NOT NULL;
""")

query = """
WITH p_agg AS (
    SELECT 
        OR_KEY,
        OR_FULL,
        SUM(Heures) as Heures_Totales,
        ARGMAX(Equipe, Heures) as Equipe_Principale,
        YEAR(MIN(Date_Pointage)) as Annee
    FROM v_p
    GROUP BY OR_KEY, OR_FULL
),
b_agg AS (
    SELECT 
        OR_KEY,
        MAX(Temps_Ref) as Temps_Reference
    FROM v_b
    GROUP BY OR_KEY
)
SELECT 
    p.*,
    b.Temps_Reference,
    (b.Temps_Reference / NULLIF(p.Heures_Totales, 0)) * 100 as Efficience_Pct
FROM p_agg p
LEFT JOIN b_agg b ON p.OR_KEY = b.OR_KEY
"""

df_final = con.execute(query).df()

# ======================================================
# AFFICHAGE
# ======================================================
st.sidebar.header("Filtres")
annees = sorted(df_final['Annee'].dropna().unique().astype(int))
sel_annees = st.sidebar.multiselect("Années", options=annees, default=annees)
df_view = df_final[df_final['Annee'].isin(sel_annees)]

c1, c2, c3 = st.columns(3)
c1.metric("OR Pointage", len(df_view))
c2.metric("Matches trouvés", df_view['Temps_Reference'].notna().sum())
c3.metric("Échecs", df_view['Temps_Reference'].isna().sum())

if df_view['Temps_Reference'].notna().sum() > 0:
    st.success(f"✅ {df_view['Temps_Reference'].notna().sum()} OR ont été associés via les 5 derniers chiffres.")
    st.dataframe(df_view.sort_values("Efficience_Pct", ascending=False), use_container_width=True)
else:
    st.error("❌ Même la correspondance partielle a échoué.")
    st.write("Comparaison des suffixes (5 derniers chiffres) :")
    colA, colB = st.columns(2)
    colA.write("Suffixes Pointage :")
    colA.write(con.execute("SELECT DISTINCT OR_KEY FROM v_p LIMIT 10").df())
    colB.write("Suffixes BO :")
    colB.write(con.execute("SELECT DISTINCT OR_KEY FROM v_b LIMIT 10").df())
