import streamlit as st
import pandas as pd
import altair as alt
import base64
import os

st.set_page_config(
    page_title="Porra Mundial",
    layout="wide"
)

# --------------------------------------------------
# CONFIGURACIÓ
# --------------------------------------------------
EXCEL_FILE = "Porra_Mundial_Final_Definitiva.xlsx"
BACKGROUND_IMAGE = "fifa-Trionda.jpg"

# --------------------------------------------------
# FUNCIÓ IMATGE DE FONS
# --------------------------------------------------
def get_base64_image(image_path):
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode()

# --------------------------------------------------
# ESTILS + FONS
# --------------------------------------------------
if os.path.exists(BACKGROUND_IMAGE):
    img_base64 = get_base64_image(BACKGROUND_IMAGE)

    st.markdown(
        f"""
        <style>
        .stApp {{
            background-image:
                linear-gradient(rgba(0,0,0,0.58), rgba(0,0,0,0.70)),
                url("data:image/jpg;base64,{img_base64}");
            background-size: cover;
            background-position: center;
            background-attachment: fixed;
        }}

        .block-container {{
            padding-top: 2rem;
            padding-bottom: 2rem;
            background: rgba(255,255,255,0.88);
            border-radius: 20px;
            margin-top: 20px;
            margin-bottom: 20px;
        }}

        .title {{
            font-size: 48px;
            font-weight: 900;
            margin-bottom: 0px;
            color: #102a43;
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

        .card h3 {{
            margin-bottom: 10px;
        }}

        .card h1 {{
            margin: 0px;
            font-size: 42px;
        }}

        .small-note {{
            font-size: 13px;
            color: #777;
        }}
        </style>
        """,
        unsafe_allow_html=True
    )
else:
    st.markdown(
        """
        <style>
        .title {
            font-size: 48px;
            font-weight: 900;
            margin-bottom: 0px;
            color: #102a43;
        }

        .subtitle {
            font-size: 18px;
            color: #334e68;
            margin-top: 0px;
            margin-bottom: 25px;
        }

        .card {
            padding: 22px;
            border-radius: 18px;
            text-align: center;
            box-shadow: 0px 4px 20px rgba(0,0,0,0.18);
            min-height: 150px;
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

        .card h3 {
            margin-bottom: 10px;
        }

        .card h1 {
            margin: 0px;
            font-size: 42px;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

# --------------------------------------------------
# LOAD DATA
# --------------------------------------------------
df_ranking = pd.read_excel(EXCEL_FILE, sheet_name="Gràfics")
df_porra = pd.read_excel(EXCEL_FILE, sheet_name="Porra")
df_resultats = pd.read_excel(EXCEL_FILE, sheet_name="Resultats Reals")

# Netejar noms de columnes
df_ranking.columns = df_ranking.columns.astype(str).str.strip()
df_porra.columns = df_porra.columns.astype(str).str.strip()
df_resultats.columns = df_resultats.columns.astype(str).str.strip()

# --------------------------------------------------
# DETECTAR COLUMNES DEL RÀNQUING
# --------------------------------------------------
col_participant_ranking = None
for col in df_ranking.columns:
    if "participant" in col.lower():
        col_participant_ranking = col

if col_participant_ranking is None:
    st.error("No s'ha trobat la columna de participant al full Gràfics.")
    st.write("Columnes detectades:", list(df_ranking.columns))
    st.stop()

col_punts = None
for col in df_ranking.columns:
    if "punt" in col.lower():
        col_punts = col

if col_punts is None:
    st.error("No s'ha trobat la columna de punts al full Gràfics.")
    st.write("Columnes detectades:", list(df_ranking.columns))
    st.stop()

# --------------------------------------------------
# NETEJA DEL RÀNQUING
# --------------------------------------------------
# Eliminar fila "Total general"
df_ranking = df_ranking[
    ~df_ranking[col_participant_ranking]
    .astype(str)
    .str.contains("Total", case=False, na=False)
].copy()

# Convertir punts a numèric
df_ranking[col_punts] = pd.to_numeric(df_ranking[col_punts], errors="coerce")

# Eliminar files sense punts
df_ranking = df_ranking.dropna(subset=[col_punts])

# Ordenar general
df_ranking = df_ranking.sort_values(col_punts, ascending=False).reset_index(drop=True)

# Arrodonir punts a 1 decimal
df_ranking[col_punts] = df_ranking[col_punts].round(1)

# --------------------------------------------------
# TÍTOL
# --------------------------------------------------
st.markdown('<p class="title">🏆 PORRA MUNDIAL</p>', unsafe_allow_html=True)
st.markdown(
    '<p class="subtitle">Classificació en viu, lliguetes personalitzades i resultats reals</p>',
    unsafe_allow_html=True
)

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

# Posició dins la vista actual
df_view["Posició"] = df_view.index + 1

# Diferència respecte líder de la vista actual
if len(df_view) > 0:
    punts_lider = float(df_view[col_punts].iloc[0])
    df_view["Dif líder"] = (df_view[col_punts] - punts_lider).round(1)
else:
    st.warning("No hi ha participants seleccionats.")
    st.stop()

# Assegurar 1 decimal
df_view[col_punts] = df_view[col_punts].round(1)

# --------------------------------------------------
# TOP 3
# --------------------------------------------------
st.subheader("🥇 TOP 3")

if len(df_view) >= 3:
    top3 = df_view.head(3)

    c1, c2, c3 = st.columns(3)

    nom1 = top3.iloc[0][col_participant_ranking]
    punts1 = float(top3.iloc[0][col_punts])

    nom2 = top3.iloc[1][col_participant_ranking]
    punts2 = float(top3.iloc[1][col_punts])

    nom3 = top3.iloc[2][col_participant_ranking]
    punts3 = float(top3.iloc[2][col_punts])

    c1.markdown(
        f"""
        <div class='card gold'>
            <h3>🥇 {nom1}</h3>
            <h1>{punts1:.1f}</h1>
            <p>punts</p>
        </div>
        """,
        unsafe_allow_html=True
    )

    c2.markdown(
        f"""
        <div class='card silver'>
            <h3>🥈 {nom2}</h3>
            <h1>{punts2:.1f}</h1>
            <p>punts</p>
        </div>
        """,
        unsafe_allow_html=True
    )

    c3.markdown(
        f"""
        <div class='card bronze'>
            <h3>🥉 {nom3}</h3>
            <h1>{punts3:.1f}</h1>
            <p>punts</p>
        </div>
        """,
        unsafe_allow_html=True
    )

elif len(df_view) > 0:
    st.info("Selecciona com a mínim 3 participants si vols veure el TOP 3 complet.")

# --------------------------------------------------
# CLASSIFICACIÓ
# --------------------------------------------------
st.subheader("📊 Classificació")

ranking_display = df_view[
    ["Posició", col_participant_ranking, col_punts, "Dif líder"]
].copy()

ranking_display = ranking_display.rename(columns={
    col_participant_ranking: "Participant",
    col_punts: "Punts"
})

ranking_display["Punts"] = ranking_display["Punts"].astype(float).round(1)
ranking_display["Dif líder"] = ranking_display["Dif líder"].astype(float).round(1)

def highlight_leader(row):
    if row["Posició"] == 1:
        return ["background-color: #ffe066; font-weight: bold;"] * len(row)
    return [""] * len(row)

styled_ranking = (
    ranking_display
    .style
    .apply(highlight_leader, axis=1)
    .format({
        "Punts": "{:.1f}",
        "Dif líder": "{:.1f}"
    })
)

st.dataframe(
    styled_ranking,
    use_container_width=True,
    hide_index=True,
    column_config={
        "Posició": st.column_config.NumberColumn("Posició", format="%d"),
        "Punts": st.column_config.NumberColumn("Punts", format="%.1f"),
        "Dif líder": st.column_config.NumberColumn("Dif líder", format="%.1f"),
    }
)

# --------------------------------------------------
# GRÀFIC ORDENAT
# --------------------------------------------------
st.subheader("📈 Gràfic de punts")

chart_data = ranking_display.copy()
chart_data = chart_data.sort_values("Punts", ascending=False)

chart_height = max(350, min(900, len(chart_data) * 24))

chart = (
    alt.Chart(chart_data)
    .mark_bar(color="#0b70c9")
    .encode(
        x=alt.X(
            "Punts:Q",
            title="Punts",
            scale=alt.Scale(zero=False)
        ),
        y=alt.Y(
            "Participant:N",
            sort="-x",
            title=None
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
# FITXA PARTICIPANT
# --------------------------------------------------
st.subheader("👤 Fitxa participant")

participants_porra = df_porra["Participants"].dropna().astype(str).unique()

jugador = st.selectbox(
    "Selecciona participant",
    participants_porra
)

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

    c2.write("### ⚽ Prediccions")
    c2.write(f"🏆 Campió: {df_j['Campió'].values[0]}")
    c2.write(f"⭐ MVP: {df_j['MVP'].values[0]}")
    c2.write(f"⚽ Pichichi: {df_j['Pichichi'].values[0]}")

# --------------------------------------------------
# LLIGUETA SELECCIONADA
# --------------------------------------------------
st.subheader("🏟️ Lligueta seleccionada")

if participants_filtrats:
    st.write(f"Participants seleccionats: **{len(participants_filtrats)}**")

    st.dataframe(
        styled_ranking,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Posició": st.column_config.NumberColumn("Posició", format="%d"),
            "Punts": st.column_config.NumberColumn("Punts", format="%.1f"),
            "Dif líder": st.column_config.NumberColumn("Dif líder", format="%.1f"),
        }
    )
else:
    st.write("Selecciona participants al filtre superior per crear una lligueta personalitzada.")

# --------------------------------------------------
# RESULTATS REALS
# --------------------------------------------------
st.subheader("✅ Resultats reals")

df_resultats_display = df_resultats.copy()

# Treure files totalment buides
df_resultats_display = df_resultats_display.dropna(how="all")

# Treure columnes totalment buides
df_resultats_display = df_resultats_display.dropna(axis=1, how="all")

# Omplir buits per visualitzar millor
df_resultats_display = df_resultats_display.fillna("")

st.dataframe(
    df_resultats_display,
    use_container_width=True,
    hide_index=True
)

# --------------------------------------------------
# FOOTER
# --------------------------------------------------
st.markdown("---")
st.write("📡 Actualització automàtica des de Excel")
