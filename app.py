import streamlit as st
import pandas as pd
import altair as alt
import base64
import os
from datetime import datetime
from zoneinfo import ZoneInfo

st.set_page_config(layout="wide")

EXCEL_FILE = "Porra_Mundial_Final_Definitiva.xlsx"
BACKGROUND_IMAGE = "fifa-Trionda.jpg"
PREU = 5

# --------------------------------------------------
# CACHE
# --------------------------------------------------
@st.cache_data
def load_data(file):
    sheets = pd.read_excel(file, sheet_name=["Porra", "Resultats Reals"])
    return sheets["Porra"], sheets["Resultats Reals"]

def last_update(file):
    t = os.path.getmtime(file)
    dt = datetime.fromtimestamp(t, tz=ZoneInfo("Europe/Madrid"))
    return dt.strftime("%d/%m/%Y %H:%M")

# --------------------------------------------------
# CSS RESPONSIVE (IMPORTANT)
# --------------------------------------------------
st.markdown("""
<style>
.card {
    padding: 18px;
    border-radius: 18px;
    text-align: center;
    min-height: 160px;

    display:flex;
    flex-direction:column;
    justify-content:center;
    align-items:center;

    overflow:hidden;
}

.card h3 {
    font-size: clamp(14px,2vw,22px);
    white-space:nowrap;
    overflow:hidden;
    text-overflow:ellipsis;
}

.card h1 {
    font-size: clamp(24px,4vw,40px);
    margin:0;
}

.card p {
    font-size: clamp(12px,1.5vw,14px);
    white-space:nowrap;
    overflow:hidden;
    text-overflow:ellipsis;
}

.dark {background:#2f4c65;color:white;}
.green {background:linear-gradient(135deg,#2ea568,#79d7a6);color:white;}
.blue {background:linear-gradient(135deg,#2f7cc2,#6aa8d9);color:white;}
.gold {background:#ffd700;}
.silver{background:#c0c0c0;}
.bronze{background:#cd7f32;color:white;}

@media (max-width:768px){
    div[data-testid="column"] {
        width:100%!important;
    }
}
</style>
""", unsafe_allow_html=True)

# --------------------------------------------------
# LOAD
# --------------------------------------------------
df_porra, df_res = load_data(EXCEL_FILE)

df_rank = df_porra[["Participants","Total Punts"]].copy()
df_rank.columns = ["Participant","Punts"]
df_rank["Punts"] = pd.to_numeric(df_rank["Punts"], errors="coerce")
df_rank = df_rank.dropna().sort_values("Punts",ascending=False).reset_index(drop=True)

df_rank["Posició"] = df_rank.index+1
df_rank["Dif líder"] = (df_rank["Punts"] - df_rank["Punts"].iloc[0]).round(1)

participants = len(df_rank)
premi = participants * PREU

# --------------------------------------------------
# HEADER
# --------------------------------------------------
st.title("🏆 PORRA MUNDIAL")

c1,c2,c3 = st.columns(3)

c1.markdown(f"<div class='card dark'><h3>🕒 Dades actualitzades</h3><h1>{last_update(EXCEL_FILE)}</h1></div>",unsafe_allow_html=True)
c2.markdown(f"<div class='card green'><h3>🎁 Premi</h3><h1>{premi} €</h1><p>{participants} x {PREU}€</p></div>",unsafe_allow_html=True)
c3.markdown(f"<div class='card blue'><h3>👥 Participants</h3><h1>{participants}</h1></div>",unsafe_allow_html=True)

# --------------------------------------------------
# TOP 3
# --------------------------------------------------
st.subheader("🥇 TOP 3")
cols = st.columns(3)

for i,c in enumerate(cols):
    try:
        row = df_rank.iloc[i]
        c.markdown(f"<div class='card'>{row['Participant']}<br><h1>{row['Punts']:.1f}</h1></div>",unsafe_allow_html=True)
    except:
        pass

# --------------------------------------------------
# CLASSIFICACIÓ
# --------------------------------------------------
st.subheader("📊 Classificació")
st.dataframe(df_rank[["Posició","Participant","Punts","Dif líder"]],use_container_width=True)

# --------------------------------------------------
# GRÀFIC
# --------------------------------------------------
st.subheader("📈 Gràfic")

chart = alt.Chart(df_rank).mark_bar().encode(
    x="Punts",
    y=alt.Y("Participant", sort='-x')
)

st.altair_chart(chart,use_container_width=True)

# --------------------------------------------------
# PARTICIPANT
# --------------------------------------------------
st.subheader("👤 Participant")

jugador = st.selectbox("Selecciona participant", df_rank["Participant"], index=None)

if jugador:
    df_j = df_porra[df_porra["Participants"]==jugador]

    st.metric("Total punts", float(df_j["Total Punts"]))

    st.write("### Prediccions bàsiques")
    st.write("🏆 Campió:", df_j["Campió"].values[0])
    st.write("⭐ MVP:", df_j["MVP"].values[0])
    st.write("⚽ Pichichi:", df_j["Pichichi"].values[0])

    # FASE ELIMINATORIA
    st.write("### 🧭 Prediccions eliminatòria")

    tabs = st.tabs(["Vuitens","Quarts","Semis","Final","Campió"])

    def build(prefix,n):
        data=[]
        for i in range(1,n+1):
            col=f"{prefix}_{i}"
            if col in df_j:
                val=str(df_j[col].values[0])
            else:
                val="Pendent"
            data.append(val)
        return pd.DataFrame({"Equip":data})

    with tabs[0]:
        st.dataframe(build("Vuitens",16), use_container_width=True)
    with tabs[1]:
        st.dataframe(build("Quarts",8), use_container_width=True)
    with tabs[2]:
        st.dataframe(build("Semis",4), use_container_width=True)
    with tabs[3]:
        final = pd.DataFrame({"Equip":[df_j["Final_1"].values[0], df_j["Final_2"].values[0]]})
        st.dataframe(final)
    with tabs[4]:
        st.dataframe(pd.DataFrame({"Campió":[df_j["Campió"].values[0]]}))

# --------------------------------------------------
# LLIGUETA
# --------------------------------------------------
st.subheader("🏟️ Lligueta")

sel = st.multiselect("Participants", df_rank["Participant"])

if sel:
    df_l = df_rank[df_rank["Participant"].isin(sel)]
    st.dataframe(df_l[["Posició","Participant","Punts","Dif líder"]])

# --------------------------------------------------
# RESULTATS REALS (simplificat)
# --------------------------------------------------
st.subheader("✅ Resultats reals")

df_res = df_res.fillna("")

# PICHICHI correcte
pichichi="Pendent"
gols="Pendent"

if "Gols" in df_res:
    temp = df_res.copy()
    temp["Gols"] = pd.to_numeric(temp["Gols"], errors='coerce')
    temp = temp[temp["Gols"]>=1]

    if not temp.empty:
        pichichi = temp.iloc[0]["Jugador Pichichi"]
        gols = int(temp.iloc[0]["Gols"])

st.write("⚽ Pichichi:", pichichi, "-", gols)

# FASE GRUPS COLUMNES
st.write("### Grups")

grups={}
for _,r in df_res.iterrows():
    g=r.get("Grup","")
    if g=="":
        continue
    if g not in grups:
        grups[g]=["","",""]

    if r["Posició"]=="1r":
        grups[g][0]=r["Equip"]
    elif r["Posició"]=="2n":
        grups[g][1]=r["Equip"]
    elif r["Posició"]=="3r":
        grups[g][2]=r["Equip"]

st.dataframe(pd.DataFrame(grups,index=["1r","2n","3r"]))

# --------------------------------------------------
# FINAL
# --------------------------------------------------
st.markdown("---")
st.write("📡 App optimitzada")
