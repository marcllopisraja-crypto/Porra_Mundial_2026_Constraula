import streamlit as st
import pandas as pd

st.set_page_config(layout="wide")

# ------- LOAD DATA -------
df_ranking = pd.read_excel("Porra_Mundial_Final_Definitiva.xlsx", sheet_name="Gràfics")
df_porra = pd.read_excel("Porra_Mundial_Final_Definitiva.xlsx", sheet_name="Porra")

# Netejar columnes
df_ranking.columns = df_ranking.columns.astype(str).str.strip()

# Detectar columna punts
col_punts = [col for col in df_ranking.columns if "punt" in col.lower()][0]

df_ranking = df_ranking.sort_values(col_punts, ascending=False).reset_index(drop=True)

# ------- ESTILS -------
st.markdown("""
<style>
.title {
    font-size: 42px;
    font-weight: bold;
}
.card {
    padding: 20px;
    border-radius: 15px;
    text-align: center;
}
.gold {background-color:#ffd700;}
.silver {background-color:#c0c0c0;}
.bronze {background-color:#cd7f32; color:white;}
</style>
""", unsafe_allow_html=True)

st.markdown('<p class="title">🏆 PORRA MUNDIAL</p>', unsafe_allow_html=True)

# ------- TOP 3 VISUAL -------
st.subheader("🥇 TOP 3")

c1, c2, c3 = st.columns(3)

top3 = df_ranking.head(3)

c1.markdown(f"<div class='card gold'><h3>{top3.iloc[0]['Participant']}</h3><h1>{top3.iloc[0][col_punts]}</h1></div>", unsafe_allow_html=True)
c2.markdown(f"<div class='card silver'><h3>{top3.iloc[1]['Participant']}</h3><h1>{top3.iloc[1][col_punts]}</h1></div>", unsafe_allow_html=True)
c3.markdown(f"<div class='card bronze'><h3>{top3.iloc[2]['Participant']}</h3><h1>{top3.iloc[2][col_punts]}</h1></div>", unsafe_allow_html=True)

# ------- RANKING VISUAL -------
st.subheader("📊 Classificació")

ranking_display = df_ranking.copy()
ranking_display["Posició"] = ranking_display.index + 1

ranking_display["Diferència líder"] = ranking_display[col_punts] - ranking_display[col_punts].iloc[0]

st.dataframe(
    ranking_display[["Posició", "Participant", col_punts, "Diferència líder"]],
    use_container_width=True
)

st.bar_chart(df_ranking.set_index("Participant")[col_punts])

# ------- FITXA PARTICIPANT -------
st.subheader("👤 Anàlisi participant")

jugador = st.selectbox("Selecciona participant", df_porra["Participants"].unique())

df_j = df_porra[df_porra["Participants"] == jugador]

if not df_j.empty:

    total = df_j["Total Punts"].values[0]

    c1, c2 = st.columns(2)

    c1.metric("Total punts", total)

    punts_dict = {
        "1rs": df_j["Punts Grups 1r"].values[0],
        "2ns": df_j["Punts Grups 2n"].values[0],
        "3rs": df_j["Punts Grups 3r"].values[0],
        "Vuitens": df_j["Punts Vuitens"].values[0],
        "Quarts": df_j["Punts Quarts"].values[0],
        "Semis": df_j["Punts Semis"].values[0]
    }

    c1.bar_chart(punts_dict)

    c2.write("### ⚽ Prediccions")
    c2.write(f"🏆 Campió: {df_j['Campió'].values[0]}")
    c2.write(f"⭐ MVP: {df_j['MVP'].values[0]}")
    c2.write(f"⚽ Pichichi: {df_j['Pichichi'].values[0]}")

# ------- COMPARADOR -------
st.subheader("⚔️ Comparador")

j1, j2 = st.columns(2)

jugador1 = j1.selectbox("Jugador 1", df_porra["Participants"].unique(), key="j1")
jugador2 = j2.selectbox("Jugador 2", df_porra["Participants"].unique(), key="j2")

if jugador1 and jugador2:

    df1 = df_porra[df_porra["Participants"] == jugador1]
    df2 = df_porra[df_porra["Participants"] == jugador2]

    if not df1.empty and not df2.empty:

        p1 = df1["Total Punts"].values[0]
        p2 = df2["Total Punts"].values[0]

        c1, c2, c3 = st.columns(3)

        c1.metric(jugador1, p1)
        c2.metric("Diferència", p1 - p2)
        c3.metric(jugador2, p2)

# ------- FOOTER -------
st.markdown("---")
st.write("📡 Actualització automàtica des de Excel")
