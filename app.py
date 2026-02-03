import streamlit as st
import pandas as pd
import duckdb
import plotly.express as px

# ======================================================
# CONFIGURATION
# ======================================================
st.set_page_config(page_title="Efficience OR - Full Match", layout="wide")
st.title("📊 Analyse d'Efficience (Correction 8 Chiffres)")

uploaded_file = st.file_uploader("📥 Charger le fichier Excel", type=["xlsx"])

if not uploaded_file:
    st.stop()

@st.cache_data
def load_data(file):
    return pd.read_excel(file, sheet_name="Pointage"), pd.read_excel(file, sheet_name="BASE_BO")

df_p_raw, df_b_raw = load_data(uploaded_file)

con = duckdb.connect(database=':memory:')
con.register('raw_p', df_p_raw)
con.register('raw_b', df_b_raw)

# ======================================================
# LOGIQUE DE NETTOYAGE RADICAL
# ======================================================
# 1. On transforme en texte
# 2. On enlève tout ce qui n'est pas un chiffre (TRIM inclus)
# 3. On garde les 8 chiffres exacts
con.execute("""
    CREATE OR REPLACE VIEW v_p AS
    SELECT 
        regexp_replace(CAST("OR (Numéro)" AS VARCHAR), '[^0-9]', '', 'g') AS OR_KEY,
        "Salarié - Nom" AS Technicien,
        "Salarié - Equipe(Nom)" AS Equipe,
        Hr_travaillée AS Heures,
        CAST("Saisie heures - Date" AS DATE) AS Date_Pointage,
        strftime(CAST("Saisie heures - Date" AS DATE), '%Y-%m') AS Mois_Annee
    FROM raw_p
    WHERE "OR (Numéro)" IS NOT NULL;

    CREATE OR REPLACE VIEW v_b AS
    SELECT 
        regexp_replace(CAST("N° OR (Segment)" AS VARCHAR), '[^0-9]', '', 'g') AS OR_KEY,
        COALESCE("Temps vendu (OR)", "Temps prévu devis (OR)") AS Temps_Ref
    FROM raw_b
    WHERE "N° OR (Segment)" IS NOT NULL;
""")

# Jointure sur l'égalité stricte des 8 chiffres nettoyés
query = """
WITH p_agg AS (
    SELECT 
        OR_KEY,
        Mois_Annee,
        SUM(Heures) as Heures_Totales,
        ARGMAX(Equipe, Heures) as Equipe_Principale
    FROM v_p
    GROUP BY OR_KEY, Mois_Annee
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
    CASE 
        WHEN b.Temps_Reference IS NULL THEN 0 
        ELSE (b.Temps_Reference / NULLIF(p.Heures_Totales, 0)) * 100 
    END as Efficience_Pct
FROM p_agg p
LEFT JOIN b_agg b ON p.OR_KEY = b.OR_KEY
"""

df_final = con.execute(query).df()

# ======================================================
# VISU & KPI
# ======================================================
matched = df_final['Temps_Reference'].notna().sum()
total = len(df_final)

c1, c2, c3 = st.columns(3)
c1.metric("OR Total", total)
c2.metric("OR Matchés", matched)
c3.metric("Taux de Match", f"{(matched/total)*100:.1f}%")

if matched == 0:
    st.error("❌ Zéro correspondance. Analyse des premiers caractères détectés :")
    colA, colB = st.columns(2)
    # Debug pour voir s'il y a des caractères cachés ou des longueurs différentes
    colA.write("Clés Pointage (Brut) :")
    colA.write(con.execute("SELECT OR_KEY, length(OR_KEY) as len FROM v_p LIMIT 5").df())
    colB.write("Clés BO (Brut) :")
    colB.write(con.execute("SELECT OR_KEY, length(OR_KEY) as len FROM v_b LIMIT 5").df())
else:
    st.success(f"✅ {matched} OR associés avec succès.")
    
    # Graphique d'efficience par équipe
    df_equipe = df_final.dropna(subset=['Temps_Reference']).groupby("Equipe_Principale")["Efficience_Pct"].mean().reset_index().sort_values("Efficience_Pct")
    
    st.subheader("📉 Qui est à la traîne ? (Efficience par équipe)")
    fig = px.bar(df_equipe, x="Efficience_Pct", y="Equipe_Principale", orientation='h', 
                 color="Efficience_Pct", color_continuous_scale="RdYlGn",
                 labels={"Efficience_Pct": "Efficience (%)", "Equipe_Principale": "Équipe"})
    st.plotly_chart(fig, use_container_width=True)

    # Évolution Mensuelle
    st.subheader("📅 Évolution Mensuelle")
    df_mois = df_final.dropna(subset=['Temps_Reference']).groupby(["Mois_Annee", "Equipe_Principale"])["Efficience_Pct"].mean().reset_index()
    fig_line = px.line(df_mois, x="Mois_Annee", y="Efficience_Pct", color="Equipe_Principale", markers=True)
    st.plotly_chart(fig_line, use_container_width=True)
