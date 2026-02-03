import streamlit as st
import pandas as pd
import duckdb

# ======================================================
# CONFIGURATION
# ======================================================
st.set_page_config(page_title="Efficience OR - Force Numeric", layout="wide")
st.title("📊 Analyse d’efficience (Conversion Numérique)")

uploaded_file = st.file_uploader("📥 Charger le fichier Excel", type=["xlsx"])

if not uploaded_file:
    st.stop()

@st.cache_data
def load_data(file):
    return pd.read_excel(file, sheet_name="Pointage"), pd.read_excel(file, sheet_name="BASE_BO")

df_p_raw, df_b_raw = load_data(uploaded_file)

# ======================================================
# MOTEUR SQL AVEC CONVERSION NUMÉRIQUE FORCÉE
# ======================================================
con = duckdb.connect(database=':memory:')
con.register('raw_p', df_p_raw)
con.register('raw_b', df_b_raw)

# La logique ici :
# 1. Extraire uniquement les chiffres via regexp_replace
# 2. Convertir le résultat en BIGINT pour supprimer les '000' et uniformiser le type
con.execute("""
    CREATE OR REPLACE VIEW v_pointage AS
    SELECT 
        CAST(regexp_replace(split_part(CAST("OR (Numéro)" AS VARCHAR), '-', 1), '[^0-9]', '', 'g') AS BIGINT) AS OR_ID,
        "Salarié - Nom" AS Technicien,
        "Salarié - Equipe(Nom)" AS Equipe,
        Hr_travaillée AS Heures,
        CAST("Saisie heures - Date" AS DATE) AS Date_Pointage
    FROM raw_p
    WHERE "OR (Numéro)" IS NOT NULL 
      AND regexp_replace(CAST("OR (Numéro)" AS VARCHAR), '[^0-9]', '', 'g') != '';

    CREATE OR REPLACE VIEW v_bo AS
    SELECT 
        CAST(regexp_replace(split_part(CAST("N° OR (Segment)" AS VARCHAR), '-', 1), '[^0-9]', '', 'g') AS BIGINT) AS OR_ID,
        COALESCE("Temps vendu (OR)", "Temps prévu devis (OR)") AS Temps_Ref
    FROM raw_b
    WHERE "N° OR (Segment)" IS NOT NULL
      AND regexp_replace(CAST("N° OR (Segment)" AS VARCHAR), '[^0-9]', '', 'g') != '';
""")

# Jointure sur les IDs numériques
query = """
WITH p_agg AS (
    SELECT 
        OR_ID,
        SUM(Heures) as Heures_Totales,
        ARGMAX(Equipe, Heures) as Equipe_Principale,
        YEAR(MIN(Date_Pointage)) as Annee
    FROM v_pointage
    GROUP BY OR_ID
),
b_agg AS (
    SELECT 
        OR_ID,
        MAX(Temps_Ref) as Temps_Reference
    FROM v_bo
    GROUP BY OR_ID
)
SELECT 
    p.*,
    b.Temps_Reference,
    (b.Temps_Reference / NULLIF(p.Heures_Totales, 0)) * 100 as Efficience_Pct
FROM p_agg p
LEFT JOIN b_agg b ON p.OR_ID = b.OR_ID
"""

df_final = con.execute(query).df()

# ======================================================
# AFFICHAGE & FILTRES
# ======================================================
st.sidebar.header("Filtres")
annees = sorted(df_final['Annee'].dropna().unique().astype(int))
sel_annees = st.sidebar.multiselect("Années", options=annees, default=annees)
df_view = df_final[df_final['Annee'].isin(sel_annees)]

# Métriques de contrôle
c1, c2, c3 = st.columns(3)
c1.metric("OR total (Pointage)", len(df_view))
c2.metric("OR répertoriés dans BO", df_view['Temps_Reference'].notna().sum())
c3.metric("Échecs de correspondance", df_view['Temps_Reference'].isna().sum())

if df_view['Temps_Reference'].notna().sum() == 0:
    st.error("❌ Toujours aucune correspondance.")
    st.info("Vérification des 5 premières clés numériques générées :")
    st.write("Côté Pointage :", con.execute("SELECT DISTINCT OR_ID FROM v_pointage LIMIT 5").df())
    st.write("Côté BO :", con.execute("SELECT DISTINCT OR_ID FROM v_bo LIMIT 5").df())
else:
    st.success("✅ Correspondance établie !")
    st.dataframe(df_view.sort_values("Efficience_Pct", ascending=False), use_container_width=True)
