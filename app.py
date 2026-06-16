import streamlit as st
import pandas as pd
import altair as alt
import base64
import os
import unicodedata

# --------------------------------------------------
# CONFIGURACIÓ GENERAL
# --------------------------------------------------
st.set_page_config(
    page_title="Porra Mundial",
    layout="wide"
)

EXCEL_FILE = "Porra_Mundial_Final_Definitiva.xlsx"
BACKGROUND_IMAGE = "fifa-Trionda.jpg"


# --------------------------------------------------
# FUNCIONS CACHEJADES
# --------------------------------------------------
@st.cache_data(show_spinner=False)
def carregar_dades(excel_file):
    sheets = pd.read_excel(
        excel_file,
        sheet_name=["Porra", "Resultats Reals"],
        engine="openpyxl"
    )

    df_porra = sheets["Porra"]
    df_resultats = sheets["Resultats Reals"]

    df_porra.columns = df_porra.columns.astype(str).str.strip()
    df_resultats.columns = df_resultats.columns.astype(str).str.strip()

    return df_porra, df_resultats


@st.cache_data(show_spinner=False)
def carregar_imatge_base64(image_path):
    if not os.path.exists(image_path):
        return None

    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode()


# --------------------------------------------------
# FUNCIONS AUXILIARS
# --------------------------------------------------
def normalitzar_text(text):
    text = str(text).strip().lower()
    text = unicodedata.normalize("NFD", text)
    text = "".join(char for char in text if unicodedata.category(char) != "Mn")
    return text


def valor_o_pendent(valor):
    if pd.isna(valor):
        return "Pendent"

    valor_text = str(valor).strip()

    if valor_text == "" or valor_text.lower() in ["nan", "nat"]:
        return "Pendent"

    return valor_text


def llista_valors_no_buits(df, columna):
    if columna not in df.columns:
        return []

    valors = (
        df[columna]
        .astype(str)
        .str.strip()
        .replace("nan", "")
        .replace("NaT", "")
    )

    return [v for v in valors if v != ""]


def primer_valor_o_pendent(df, columna):
    valors = llista_valors_no_buits(df, columna)

    if len(valors) == 0:
        return "Pendent"

    return valors[0]


def preparar_taula_buida(df):
    df = df.copy()
    df = df.dropna(how="all")
    df = df.dropna(axis=1, how="all")
    df = df.fillna("")
    return df


def fila_no_buida(df):
    return (
        df.astype(str)
        .apply(lambda fila: "".join(fila), axis=1)
        .str.strip() != ""
    )


def trobar_col_resultat_final_porra(df_porra):
    # Al full Porra hi ha "Resultat final" com a predicció.
    # Més endavant també pot existir "Resultat Final" com a columna de punts.
    # Per això prioritzem exactament "Resultat final".
    for col in df_porra.columns:
        if col.strip() == "Resultat final":
            return col

    for col in df_porra.columns:
        col_norm = normalitzar_text(col)
        if "resultat" in col_norm and "final" in col_norm and "punt" not in col_norm:
            return col

    return None


def crear_ranking_des_de_porra(df_porra):
    if "Participants" not in df_porra.columns:
        st.error("No s'ha trobat la columna 'Participants' al full Porra.")
        st.write("Columnes detectades:", list(df_porra.columns))
        st.stop()

    if "Total Punts" not in df_porra.columns:
        st.error("No s'ha trobat la columna 'Total Punts' al full Porra.")
        st.write("Columnes detectades:", list(df_porra.columns))
        st.stop()

    df = df_porra[["Participants", "Total Punts"]].copy()
    df = df.rename(columns={
        "Participants": "Participant",
        "Total Punts": "Punts"
    })

    df["Participant"] = df["Participant"].astype(str).str.strip()
    df["Punts"] = pd.to_numeric(df["Punts"], errors="coerce")

    df = df.dropna(subset=["Punts"])
    df = df[df["Participant"] != ""]
    df = df[~df["Participant"].str.contains("Total", case=False, na=False)]

    df = df.sort_values("Punts", ascending=False).reset_index(drop=True)
    df["Punts"] = df["Punts"].round(1)

    df["Posició"] = df.index + 1

    punts_lider = float(df["Punts"].iloc[0])
    df["Dif líder"] = (df["Punts"] - punts_lider).round(1)

    return df


def highlight_leader(row):
    if row["Posició"] == 1:
        return ["background-color: #ffe066; font-weight: bold;"] * len(row)
    return [""] * len(row)


def mostrar_taula_ranking(df):
    df_display = df[["Posició", "Participant", "Punts", "Dif líder"]].copy()

    df_display["Punts"] = df_display["Punts"].astype(float).round(1)
    df_display["Dif líder"] = df_display["Dif líder"].astype(float).round(1)

    styled = (
        df_display
        .style
        .apply(highlight_leader, axis=1)
        .format({
            "Punts": "{:.1f}",
            "Dif líder": "{:.1f}"
        })
    )

    st.dataframe(
        styled,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Posició": st.column_config.NumberColumn("Posició", format="%d"),
            "Punts": st.column_config.NumberColumn("Punts", format="%.1f"),
            "Dif líder": st.column_config.NumberColumn("Dif líder", format="%.1f"),
        }
    )


def mostrar_grafic_punts(df, color="#0b70c9", altura_minima=950):
    chart_data = df[["Posició", "Participant", "Punts", "Dif líder"]].copy()
    chart_data = chart_data.sort_values("Punts", ascending=False)

    chart_height = max(altura_minima, len(chart_data) * 36)

    chart = (
        alt.Chart(chart_data)
        .mark_bar(color=color)
        .encode(
            x=alt.X(
                "Punts:Q",
                title="Punts",
                scale=alt.Scale(zero=False)
            ),
            y=alt.Y(
                "Participant:N",
                sort="-x",
                title=None,
                axis=alt.Axis(labelLimit=560, labelFontSize=12)
            ),
            tooltip=[
                alt.Tooltip("Posició:Q", title="Posició"),
                alt.Tooltip("Participant:N", title="Participant"),
                alt.Tooltip("Punts:Q", title="Punts", format=".1f"),
                alt.Tooltip("Dif líder:Q", title="Dif. líder", format=".1f")
            ]
        )
        .properties(height=chart_height)
    )

    st.altair_chart(chart, use_container_width=True)


# --------------------------------------------------
# ESTILS + FONS
# --------------------------------------------------
img_base64 = carregar_imatge_base64(BACKGROUND_IMAGE)

if img_base64:
    background_css = f"""
    background-image:
        linear-gradient(rgba(0,0,0,0.58), rgba(0,0,0,0.74)),
        url("data:image/jpg;base64,{img_base64}");
    background-size: cover;
    background-position: center;
    background-attachment: fixed;
    """
else:
    background_css = "background: #eef2f7;"


st.markdown(
    f"""
    <style>
    .stApp {{
        {background_css}
    }}

    .block-container {{
        padding-top: 2rem;
        padding-bottom: 2rem;
        background: rgba(255,255,255,0.91);
        border-radius: 24px;
        margin-top: 24px;
        margin-bottom: 24px;
        box-shadow: 0px 8px 30px rgba(0,0,0,0.25);
    }}

    .title {{
        font-size: 52px;
        font-weight: 900;
        margin-bottom: 0px;
        color: #102a43;
        letter-spacing: -1px;
    }}

    .subtitle {{
        font-size: 18px;
        color: #334e68;
        margin-top: 0px;
        margin-bottom: 25px;
    }}

    .card {{
        padding: 22px;
        border-radius: 18px;
        text-align: center;
        box-shadow: 0px 4px 20px rgba(0,0,0,0.18);
        min-height: 150px;
    }}

    .gold {{
        background: linear-gradient(135deg, #ffd700, #fff1a8);
        color: #111;
    }}

    .silver {{
        background: linear-gradient(135deg, #c0c0c0, #f2f2f2);
        color: #111;
    }}

    .bronze {{
        background: linear-gradient(135deg, #cd7f32, #f0b27a);
        color: white;
    }}

    .bluecard {{
        background: linear-gradient(135deg, #0b70c9, #7cc5ff);
        color: white;
    }}

    .greencard {{
        background: linear-gradient(135deg, #0f9d58, #8ee6b3);
        color: white;
    }}

    .darkcard {{
        background: linear-gradient(135deg, #102a43, #486581);
        color: white;
    }}

    .card h3 {{
        margin-bottom: 10px;
    }}

    .card h1 {{
        margin: 0px;
        font-size: 42px;
    }}

    .card p {{
        margin: 5px 0px 0px 0px;
    }}
    </style>
    """,
    unsafe_allow_html=True
)


# --------------------------------------------------
# CARREGAR DADES
# --------------------------------------------------
df_porra, df_resultats = carregar_dades(EXCEL_FILE)
df_ranking = crear_ranking_des_de_porra(df_porra)


# --------------------------------------------------
# TÍTOL
# --------------------------------------------------
st.markdown('<p class="title">🏆 PORRA MUNDIAL</p>', unsafe_allow_html=True)
st.markdown(
    '<p class="subtitle">Classificació en viu, detall per participant, lliguetes personalitzades i resultats reals</p>',
    unsafe_allow_html=True
)


# --------------------------------------------------
# TOP 3 GENERAL
# --------------------------------------------------
st.subheader("🥇 TOP 3 General")

top3 = df_ranking.head(3)

c1, c2, c3 = st.columns(3)

c1.markdown(
    f"""
    <div class='card gold'>
        <h3>🥇 {top3.iloc[0]["Participant"]}</h3>
        <h1>{float(top3.iloc[0]["Punts"]):.1f}</h1>
        <p>punts</p>
    </div>
    """,
    unsafe_allow_html=True
)

c2.markdown(
    f"""
    <div class='card silver'>
        <h3>🥈 {top3.iloc[1]["Participant"]}</h3>
        <h1>{float(top3.iloc[1]["Punts"]):.1f}</h1>
        <p>punts</p>
    </div>
    """,
    unsafe_allow_html=True
)

c3.markdown(
    f"""
    <div class='card bronze'>
        <h3>🥉 {top3.iloc[2]["Participant"]}</h3>
        <h1>{float(top3.iloc[2]["Punts"]):.1f}</h1>
        <p>punts</p>
    </div>
    """,
    unsafe_allow_html=True
)


# --------------------------------------------------
# CLASSIFICACIÓ GENERAL
# --------------------------------------------------
st.subheader("📊 Classificació general")
mostrar_taula_ranking(df_ranking)


# --------------------------------------------------
# GRÀFIC GENERAL
# --------------------------------------------------
st.subheader("📈 Gràfic general de punts")
mostrar_grafic_punts(df_ranking, color="#0b70c9", altura_minima=1000)


# --------------------------------------------------
# FITXA PARTICIPANT
# --------------------------------------------------
st.subheader("👤 Fitxa participant")

participants_porra = df_porra["Participants"].dropna().astype(str).unique()

jugador = st.selectbox(
    "Selecciona participant",
    participants_porra,
    index=None,
    placeholder="Selecciona un participant..."
)

if jugador is not None:
    df_j = df_porra[df_porra["Participants"].astype(str) == str(jugador)]

    if not df_j.empty:
        total = pd.to_numeric(df_j["Total Punts"].values[0], errors="coerce")

        c1, c2 = st.columns(2)

        c1.metric("Total punts", f"{total:.1f}")

        punts_dict = {
            "1rs grup": pd.to_numeric(df_j["Punts Grups 1r"].values[0], errors="coerce"),
            "2ns grup": pd.to_numeric(df_j["Punts Grups 2n"].values[0], errors="coerce"),
            "3rs grup": pd.to_numeric(df_j["Punts Grups 3r"].values[0], errors="coerce"),
            "Vuitens": pd.to_numeric(df_j["Punts Vuitens"].values[0], errors="coerce"),
            "Quarts": pd.to_numeric(df_j["Punts Quarts"].values[0], errors="coerce"),
            "Semis": pd.to_numeric(df_j["Punts Semis"].values[0], errors="coerce"),
            "Finalistes": pd.to_numeric(df_j["Punts Finalistes"].values[0], errors="coerce"),
            "Campió": pd.to_numeric(df_j["Punts Campió"].values[0], errors="coerce"),
            "MVP": pd.to_numeric(df_j["Punts MVP"].values[0], errors="coerce"),
            "Pichichi": pd.to_numeric(df_j["Punts Pichichi"].values[0], errors="coerce"),
        }

        punts_categoria = pd.DataFrame({
            "Categoria": list(punts_dict.keys()),
            "Punts": list(punts_dict.values())
        })

        punts_categoria["Punts"] = punts_categoria["Punts"].fillna(0).round(1)

        chart_cat = (
            alt.Chart(punts_categoria)
            .mark_bar(color="#1f77b4")
            .encode(
                x=alt.X("Categoria:N", sort=None, title=None),
                y=alt.Y("Punts:Q", title="Punts"),
                tooltip=[
                    alt.Tooltip("Categoria:N", title="Categoria"),
                    alt.Tooltip("Punts:Q", title="Punts", format=".1f")
                ]
            )
            .properties(height=320)
        )

        c1.altair_chart(chart_cat, use_container_width=True)

        col_resultat_final_porra = trobar_col_resultat_final_porra(df_porra)

        if col_resultat_final_porra is not None:
            resultat_final = valor_o_pendent(df_j[col_resultat_final_porra].values[0])
        else:
            resultat_final = "Pendent"

        c2.write("### ⚽ Prediccions")
        c2.write(f"🏆 Campió: {valor_o_pendent(df_j['Campió'].values[0])}")
        c2.write(f"📌 Resultat final: {resultat_final}")
        c2.write(f"⭐ MVP: {valor_o_pendent(df_j['MVP'].values[0])}")
        c2.write(f"⚽ Pichichi: {valor_o_pendent(df_j['Pichichi'].values[0])}")

else:
    st.info("Selecciona un participant per veure el detall de punts i prediccions.")


# --------------------------------------------------
# LLIGUETES - JUST ABANS DELS RESULTATS
# --------------------------------------------------
st.subheader("🏟️ Lligueta personalitzada")

tots_participants = df_ranking["Participant"].dropna().astype(str).tolist()

participants_filtrats = st.multiselect(
    "Selecciona participants per crear una lligueta:",
    options=tots_participants,
    default=[],
    placeholder="Tria participants..."
)

if participants_filtrats:
    df_lligueta = df_ranking[
        df_ranking["Participant"].astype(str).isin(participants_filtrats)
    ].copy()

    df_lligueta = df_lligueta.sort_values("Punts", ascending=False).reset_index(drop=True)
    df_lligueta["Posició"] = df_lligueta.index + 1

    punts_lider_lligueta = float(df_lligueta["Punts"].iloc[0])
    df_lligueta["Dif líder"] = (df_lligueta["Punts"] - punts_lider_lligueta).round(1)

    st.write(f"Participants seleccionats: **{len(participants_filtrats)}**")

    mostrar_taula_ranking(df_lligueta)

    st.write("#### 📈 Gràfic de la lligueta")
    mostrar_grafic_punts(df_lligueta, color="#0f9d58", altura_minima=350)

else:
    st.write("Selecciona participants per crear una classificació reduïda tipus lligueta.")


# --------------------------------------------------
# RESULTATS REALS - DASHBOARD
# --------------------------------------------------
st.subheader("✅ Resultats reals")

df_resultats_display = preparar_taula_buida(df_resultats)

COL_GRUP = "Grup"
COL_POSICIO = "Posició"
COL_EQUIP = "Equip"

COL_VUITENS = "Vuitens"
COL_QUARTS = "Quarts"
COL_SEMIS = "Semis"
COL_FINALISTES = "Finalistes"
COL_CAMPIO = "Campió"
COL_MVP = "MVP"

COL_RESULTAT_FINAL = "Resultat Final"

COL_PICHICHI = "Jugador Pichichi"
COL_GOLS = "Gols"


# --------------------------------------------------
# CARDS RESUM RESULTATS REALS
# --------------------------------------------------
campio_real = primer_valor_o_pendent(df_resultats_display, COL_CAMPIO)
mvp_real = primer_valor_o_pendent(df_resultats_display, COL_MVP)
resultat_final_real = primer_valor_o_pendent(df_resultats_display, COL_RESULTAT_FINAL)
pichichi_real = primer_valor_o_pendent(df_resultats_display, COL_PICHICHI)

if COL_GOLS in df_resultats_display.columns:
    gols_series = pd.to_numeric(df_resultats_display[COL_GOLS], errors="coerce")
    gols_valids = gols_series.dropna()

    if len(gols_valids) > 0:
        gols_pichichi = str(int(gols_valids.iloc[0]))
    else:
        gols_pichichi = "Pendent"
else:
    gols_pichichi = "Pendent"

st.write("### 🏟️ Resum oficial")

r1, r2, r3, r4 = st.columns(4)

r1.markdown(
    f"""
    <div class='card gold'>
        <h3>🏆 Campió</h3>
        <h1 style='font-size:28px'>{campio_real}</h1>
    </div>
    """,
    unsafe_allow_html=True
)

r2.markdown(
    f"""
    <div class='card silver'>
        <h3>⭐ MVP</h3>
        <h1 style='font-size:28px'>{mvp_real}</h1>
    </div>
    """,
    unsafe_allow_html=True
)

r3.markdown(
    f"""
    <div class='card bronze'>
        <h3>⚽ Pichichi</h3>
        <h1 style='font-size:25px'>{pichichi_real}</h1>
        <p>{gols_pichichi} gols</p>
    </div>
    """,
    unsafe_allow_html=True
)

r4.markdown(
    f"""
    <div class='card bluecard'>
        <h3>🏁 Resultat final</h3>
        <h1 style='font-size:28px'>{resultat_final_real}</h1>
    </div>
    """,
    unsafe_allow_html=True
)


# --------------------------------------------------
# 1. FASE DE GRUPS
# --------------------------------------------------
st.write("### 🧩 Fase de grups")

cols_grups = [COL_GRUP, COL_POSICIO, COL_EQUIP]
cols_grups_existents = [col for col in cols_grups if col in df_resultats_display.columns]

if len(cols_grups_existents) > 0:
    taula_grups = df_resultats_display[cols_grups_existents].copy()
    taula_grups = taula_grups[fila_no_buida(taula_grups)]

    if COL_EQUIP in taula_grups.columns:
        taula_grups = taula_grups[
            taula_grups[COL_EQUIP].astype(str).str.strip() != ""
        ]

    st.dataframe(
        taula_grups,
        use_container_width=True,
        hide_index=True
    )
else:
    st.info("No hi ha dades de fase de grups configurades.")


# --------------------------------------------------
# 2. FASE ELIMINATÒRIA
# --------------------------------------------------
st.write("### ⚔️ Fase eliminatòria")

fases_eliminatoria = [
    COL_VUITENS,
    COL_QUARTS,
    COL_SEMIS,
    COL_FINALISTES,
    COL_CAMPIO,
    COL_MVP
]

files_eliminatoria = []

for fase in fases_eliminatoria:
    if fase in df_resultats_display.columns:
        valors = llista_valors_no_buits(df_resultats_display, fase)

        if len(valors) == 0:
            detall = "Pendent"
        else:
            detall = " · ".join(valors)

        files_eliminatoria.append({
            "Fase": fase,
            "Resultat": detall
        })

if len(files_eliminatoria) > 0:
    taula_eliminatoria = pd.DataFrame(files_eliminatoria)

    st.dataframe(
        taula_eliminatoria,
        use_container_width=True,
        hide_index=True
    )
else:
    st.info("No hi ha dades de fase eliminatòria configurades.")


# --------------------------------------------------
# 3. PICHICHI
# --------------------------------------------------
st.write("### ⚽ Jugador pichichi")

cols_pichichi = []

if COL_PICHICHI in df_resultats_display.columns:
    cols_pichichi.append(COL_PICHICHI)

if COL_GOLS in df_resultats_display.columns:
    cols_pichichi.append(COL_GOLS)

if len(cols_pichichi) > 0:
    taula_pichichi = df_resultats_display[cols_pichichi].copy()
    taula_pichichi = taula_pichichi[fila_no_buida(taula_pichichi)]

    if COL_PICHICHI in taula_pichichi.columns:
        taula_pichichi = taula_pichichi[
            taula_pichichi[COL_PICHICHI].astype(str).str.strip() != ""
        ]

    if COL_GOLS in taula_pichichi.columns:
        taula_pichichi[COL_GOLS] = pd.to_numeric(
            taula_pichichi[COL_GOLS],
            errors="coerce"
        ).astype("Int64")

    st.dataframe(
        taula_pichichi,
        use_container_width=True,
        hide_index=True
    )
else:
    st.info("No hi ha dades de pichichi configurades.")


# --------------------------------------------------
# 4. RESULTAT DE LA FINAL
# --------------------------------------------------
st.write("### 🏁 Resultat de la final")

resultat_final = primer_valor_o_pendent(
    df_resultats_display,
    COL_RESULTAT_FINAL
)

taula_final = pd.DataFrame({
    "Concepte": ["Resultat de la final"],
    "Valor": [resultat_final]
})

st.dataframe(
    taula_final,
    use_container_width=True,
    hide_index=True
)


# --------------------------------------------------
# FOOTER
# --------------------------------------------------
st.markdown("---")
st.write("📡 Actualització automàtica des de Excel")
