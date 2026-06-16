import streamlit as st
import pandas as pd

st.set_page_config(layout="wide")

# ------- LOAD DATA -------
df_ranking = pd.read_excel("Porra_Mundial_Final_Definitiva.xlsx", sheet_name="Gràfics")
df_porra = pd.read_excel("Porra_Mundial_Final_Definitiva.xlsx", sheet_name="Porra")

df_ranking.columns = df_ranking.columns.str.strip()

df_ranking = df_ranking.sort_values("Punts", ascending=False)


# ------- ESTILS -------
st.markdown("""
    <style>
    .title {
        font-size: 40px;
        font-weight: bold;
    }
    .top1 {color: gold; font-size: 22px;}
    .top2 {color: silver; font-size: 20px;}
    .top3 {color: #cd7f32; font-size: 18px;}
    </style>
""", unsafe_allow_html=True)

st.markdown('<p class="title">🏆 PORRA MUNDIAL</p>', unsafe_allow_html=True)

# ------- TOP 3 -------
st.subheader("🥇 Classificació líder")

col1, col2, col3 = st.columns(3)

top3 = df_ranking.head(3)

col1.markdown(f"<p class='top1'>🥇 {top3.iloc[0]['Participant']}<br>{top3.iloc[0]['Punts']} punts</p>", unsafe_allow_html=True)
col2.markdown(f"<p class='top2'>🥈 {top3.iloc[1]['Participant']}<br>{top3.iloc[1]['Punts']} punts</p>", unsafe_allow_html=True)
col3.markdown(f"<p class='top3'>🥉 {top3.iloc[2]['Participant']}<br>{top3.iloc[2]['Punts']} punts</p>", unsafe_allow_html=True)

# ------- TAULA -------
st.subheader("📊 Classificació completa")
st.dataframe(df_ranking, use_container_width=True)

# ------- DETALL -------
st.subheader("👤 Fitxa participant")

jugador = st.selectbox("Selecciona participant", df_porra["Participants"])

df_j = df_porra[df_porra["Participants"] == jugador]

if not df_j.empty:

    total = df_j["Total Punts"].values[0]

    st.metric("Total punts", total)

    col1, col2 = st.columns(2)

    punts_dict = {
        "Grups 1r": df_j["Punts Grups 1r"].values[0],
        "Grups 2n": df_j["Punts Grups 2n"].values[0],
        "Grups 3r": df_j["Punts Grups 3r"].values[0],
        "Vuitens": df_j["Punts Vuitens"].values[0],
        "Quarts": df_j["Punts Quarts"].values[0],
        "Semis": df_j["Punts Semis"].values[0]
    }

    col1.bar_chart(punts_dict)

    st.write("📌 Prediccions destacades:")
    col2.write(f"🏆 Campió: {df_j['Campió'].values[0]}")
    col2.write(f"⭐ MVP: {df_j['MVP'].values[0]}")
    col2.write(f"⚽ Pichichi: {df_j['Pichichi'].values[0]}")

# ------- COMPARADOR -------
st.subheader("⚔️ Comparador")

jugador1 = st.selectbox("Jugador 1", df_porra["Participants"], key="j1")
jugador2 = st.selectbox("Jugador 2", df_porra["Participants"], key="j2")

if jugador1 and jugador2:

    df1 = df_porra[df_porra["Participants"] == jugador1]
    df2 = df_porra[df_porra["Participants"] == jugador2]

    if not df1.empty and not df2.empty:

        p1 = df1["Total Punts"].values[0]
        p2 = df2["Total Punts"].values[0]

        col1, col2, col3 = st.columns(3)

        col1.metric(jugador1, p1)
        col2.metric("Diferència", p1 - p2)
        col3.metric(jugador2, p2)

# ------- FOOTER -------
st.markdown("---")
st.write("📡 Actualització automàtica des de Excel")
