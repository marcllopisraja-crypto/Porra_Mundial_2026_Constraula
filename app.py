import streamlit as st
import pandas as pd
import altair as alt

st.set_page_config(
    page_title="Porra Mundial",
    layout="wide"
)

# --------------------------------------------------
# LOAD DATA
# --------------------------------------------------
EXCEL_FILE = "Porra_Mundial_Final_Definitiva.xlsx"

df_ranking = pd.read_excel(EXCEL_FILE, sheet_name="Gràfics")
df_porra = pd.read_excel(EXCEL_FILE, sheet_name="Porra")

# Netejar noms de columnes
df_ranking.columns = df_ranking.columns.astype(str).str.strip()
df_porra.columns = df_porra.columns.astype(str).str.strip()

# --------------------------------------------------
# DETECTAR COLUMNES
# --------------------------------------------------
# Columna participant del full Gràfics
col_participant_ranking = None
for col in df_ranking.columns:
    if "participant" in col.lower():
        col_participant_ranking = col

if col_participant_ranking is None:
    st.error("No s'ha trobat la columna de participant al full Gràfics.")
    st.write("Columnes detectades:", list(df_ranking.columns))
    st.stop()

# Columna punts del full Gràfics
col_punts = None
for col in df_ranking.columns:
    if "punt" in col.lower():
        col_punts = col

if col_punts is None:
    st.error("No s'ha trobat la columna de punts al full Gràfics.")
    st.write("Columnes detectades:", list(df_ranking.columns))
    st.stop()

# --------------------------------------------------
# NETEJA DADES
# --------------------------------------------------
# Eliminar fila "Total general"
df_ranking = df_ranking[
    ~df_ranking[col_participant_ranking]
    .astype(str)
    .str.contains("Total", case=False, na=False)
]

# Convertir punts a numèric
df_ranking[col_punts] = pd.to_numeric(df_ranking[col_punts], errors="coerce")

# Eliminar files sense punts
df_ranking = df_ranking.dropna(subset=[col_punts])

# Ordenar general
df_ranking = df_ranking.sort_values(col_punts, ascending=False).reset_index(drop=True)

# Posició general
df_ranking["Posició general"] = df_ranking.index + 1

# Arrodonir punts
df_ranking[col_punts] = df_ranking[col_punts].round(1)

# --------------------------------------------------
# ESTILS
# --------------------------------------------------
st.markdown("""
<style>
.title {
    font-size: 44px;
    font-weight: 800;
    margin-bottom: 0px;
}

.subtitle {
    font-size: 18px;
    color: #666;
    margin-top: 0px;
    margin-bottom: 25px;
}

.card {
    padding: 22px;
    border-radius: 18px;
    text-align: center;
    box-shadow: 0px 4px 20px rgba(0,0,0,0.08);
}

.gold {
    background: linear-gradient(135deg, #ffd700, #fff1a8);
    color: #111;
}

.silver {
    background: linear-gradient(135deg, #c0c0c0, #f2f2f2);
    color: #111;
}

.bronze {
    background: linear-gradient(135deg, #cd7f32, #f0b27a);
    color: white;
}

.small-note {
    font-size: 13px;
    color: #777;
}
</style>
""", unsafe_allow_html=True)

st.markdown('<p class="title">🏆 PORRA MUNDIAL</p>', unsafe_allow_html=True)
st.markdown('<p class="subtitle">Classificació en viu i lliguetes personalitzades</p>', unsafe_allow_html=True)

# --------------------------------------------------
# FILTRE PARTICIPANTS
# --------------------------------------------------
st.subheader("🎛️ Filtre de participants")

tots_participants = df_ranking[col_participant_ranking].dropna().astype(str).tolist()

participants_filtrats = st.multiselect(
    "Selecciona participants per crear una lligueta o deixa-ho buit per veure la classificació completa:",
    options=tots_participants,
    default=[]
)

if participants_filtrats:
    df_view = df_ranking[
        df_ranking[col_participant_ranking].astype(str).isin(participants_filtrats)
    ].copy()
else:
    df_view = df_ranking.copy()

# Ordenar vista seleccionada
df_view = df_view.sort_values(col_punts, ascending=False).reset_index(drop=True)

# Posició dins la vista filtrada
df_view["Posició"] = df_view.index + 1

# Diferència respecte al líder de la vista actual
punts_lider = df_view[col_punts].iloc[0]
df_view["Dif líder"] = (df_view[col_punts] - punts_lider).round(1)

# Assegurar 1 decimal
df_view[col_punts] = df_view[col_punts].round(1)

# --------------------------------------------------
# TOP 3
# --------------------------------------------------
st.subheader("🥇 TOP 3")

if len(df_view) >= 3:
    top3 = df_view.head(3)

    c1, c2, c3 = st.columns(3)

    c1.markdown(
        f"""
        <div class='card gold'>
            <h3>🥇 {top3.iloc[0][col_participant_ranking]}</h3>
            <h1>{top3.iloc[0][col_punts]:.1f}</h1>
            <p>punts</p>
        </div>
        """,
        unsafe_allow_html=True
    )

    c2.markdown(
        f"""
