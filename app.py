import streamlit as st
import pandas as pd
import altair as alt
import base64
import os
import json
import unicodedata
from datetime import datetime
from zoneinfo import ZoneInfo

# --------------------------------------------------
# CONFIGURACIÓ GENERAL
# --------------------------------------------------
st.set_page_config(
    page_title="Porra Mundial",
    layout="wide"
)

EXCEL_FILE = "Porra_Mundial_Final_Definitiva.xlsx"
BACKGROUND_IMAGE = "fifa-Trionda.jpg"
PREU_PARTICIPACIO = 5

SNAPSHOT_CURRENT_FILE = "ranking_snapshot_current.csv"
SNAPSHOT_DISPLAY_FILE = "ranking_snapshot_display.csv"
SNAPSHOT_META_FILE = "ranking_snapshot_meta.json"


# --------------------------------------------------
# BOTÓ RESET SNAPSHOT
# --------------------------------------------------
with st.sidebar:
    if st.button("🔄 Reiniciar comparativa de moviments"):
        for fitxer in [
            SNAPSHOT_CURRENT_FILE,
            SNAPSHOT_DISPLAY_FILE,
            SNAPSHOT_META_FILE
        ]:
            if os.path.exists(fitxer):
                os.remove(fitxer)
        st.rerun()


# --------------------------------------------------
# BANDERES
# --------------------------------------------------
FLAGS = {
    "mexic": "🇲🇽",
    "corea del sud": "🇰🇷",
    "republica txeca": "🇨🇿",
    "suissa": "🇨🇭",
    "canada": "🇨🇦",
    "qatar": "🇶🇦",
    "escocia": "🏴",
    "marroc": "🇲🇦",
    "brasil": "🇧🇷",
    "estats units": "🇺🇸",
    "ee.uu": "🇺🇸",
    "australia": "🇦🇺",
    "turquia": "🇹🇷",
    "alemanya": "🇩🇪",
    "costa d'ivori": "🇨🇮",
    "cote d'ivoire": "🇨🇮",
    "equador": "🇪🇨",
    "suecia": "🇸🇪",
    "japo": "🇯🇵",
    "paisos baixos": "🇳🇱",
    "nova zelanda": "🇳🇿",
    "iran": "🇮🇷",
    "belgica": "🇧🇪",
    "uruguai": "🇺🇾",
    "arabia saudita": "🇸🇦",
    "espanya": "🇪🇸",
    "franca": "🇫🇷",
    "senegal": "🇸🇳",
    "iraq": "🇮🇶",
    "argentina": "🇦🇷",
    "algeria": "🇩🇿",
    "austria": "🇦🇹",
    "portugal": "🇵🇹",
    "rd congo": "🇨🇩",
    "uzbekistan": "🇺🇿",
    "anglaterra": "🏴",
    "croacia": "🇭🇷",
    "ghana": "🇬🇭",
    "egipte": "🇪🇬",
    "noruega": "🇳🇴",
    "colombia": "🇨🇴",
    "colòmbia": "🇨🇴",
    "bosnia i hercegovina": "🇧🇦",
    "paraguai": "🇵🇾",
    "tunisia": "🇹🇳",
    "tunísia": "🇹🇳",
    "cap verd": "🇨🇻",
    "jordania": "🇯🇴",
    "jordània": "🇯🇴",
    "panama": "🇵🇦",
    "panamà": "🇵🇦",
    "curaçao": "🇨🇼",
    "curacao": "🇨🇼",
    "haiti": "🇭🇹",
    "haití": "🇭🇹",
    "sud-africa": "🇿🇦",
    "sud-àfrica": "🇿🇦"
}


# --------------------------------------------------
# FUNCIONS CACHEJADES
# --------------------------------------------------
@st.cache_data(show_spinner=False)
def carregar_dades(excel_file, file_mtime):
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


def obtenir_columna_departament(df):
    for col in df.columns:
        if normalitzar_text(col) == "departament":
            return col

    for col in df.columns:
        if "depart" in normalitzar_text(col):
            return col

    return None


def afegir_bandera(valor):
    if pd.isna(valor):
        return "Pendent"

    text = str(valor).strip()

    if text == "" or normalitzar_text(text) in ["nan", "nat", "pendent"]:
        return "Pendent"

    text_norm = normalitzar_text(text)

    for pais, bandera in FLAGS.items():
        if pais in text_norm:
            return f"{bandera} {text}"

    return text


def valor_o_pendent(valor):
    if pd.isna(valor):
        return "Pendent"

    valor_text = str(valor).strip()

    if valor_text == "" or valor_text.lower() in ["nan", "nat"]:
        return "Pendent"

    return valor_text


def obtenir_data_actualitzacio_fitxer(path):
    if not os.path.exists(path):
        return "No disponible"

    timestamp = os.path.getmtime(path)
    dt = datetime.fromtimestamp(timestamp, tz=ZoneInfo("Europe/Madrid"))

    return dt.strftime("%d/%m/%Y %H:%M")


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

    valors_nets = []

    for valor in valors:
        if valor == "":
            continue

        if normalitzar_text(valor) in ["nan", "nat"]:
            continue

        if normalitzar_text(valor) == "pendent":
            valor = "Pendent"

        valors_nets.append(valor)

    valors_unics = []
    vistos = set()

    for valor in valors_nets:
        clau = normalitzar_text(valor)
        if clau not in vistos:
            valors_unics.append(valor)
            vistos.add(clau)

    if len(valors_unics) > 0 and all(normalitzar_text(v) == "pendent" for v in valors_unics):
        return ["Pendent"]

    return valors_unics


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


def trobar_col_resultat_final_porra(df_porra):
    for col in df_porra.columns:
        if col.strip() == "Resultat final":
            return col

    for col in df_porra.columns:
        col_norm = normalitzar_text(col)
        if "resultat" in col_norm and "final" in col_norm and "punt" not in col_norm:
            return col

    return None


def recalcular_posicions(df):
    df = df.copy()
    df = df.sort_values("Punts", ascending=False).reset_index(drop=True)
    df["Posició"] = df.index + 1

    if not df.empty:
        punts_lider = float(df["Punts"].iloc[0])
        df["Dif líder"] = (df["Punts"] - punts_lider).round(1)
    else:
        df["Dif líder"] = 0

    return df


def crear_ranking_des_de_porra(df_porra):
    if "Participants" not in df_porra.columns:
        st.error("No s'ha trobat la columna 'Participants' al full Porra.")
        st.write("Columnes detectades:", list(df_porra.columns))
        st.stop()

    if "Total Punts" not in df_porra.columns:
        st.error("No s'ha trobat la columna 'Total Punts' al full Porra.")
        st.write("Columnes detectades:", list(df_porra.columns))
        st.stop()

    col_dep = obtenir_columna_departament(df_porra)

    cols_base = ["Participants", "Total Punts"]

    if col_dep is not None:
        cols_base.append(col_dep)

    df = df_porra[cols_base].copy()

    rename_map = {
        "Participants": "Participant",
        "Total Punts": "Punts"
    }

    if col_dep is not None:
        rename_map[col_dep] = "Departament"

    df = df.rename(columns=rename_map)

    df["Participant"] = df["Participant"].astype(str).str.strip()
    df["Punts"] = pd.to_numeric(df["Punts"], errors="coerce")

    if "Departament" in df.columns:
        df["Departament"] = df["Departament"].fillna("Sense departament").astype(str).str.strip()
        df["Departament"] = df["Departament"].replace("", "Sense departament")

    df = df.dropna(subset=["Punts"])
    df = df[df["Participant"] != ""]
    df = df[~df["Participant"].str.contains("Total", case=False, na=False)]

    df["Punts"] = df["Punts"].round(1)
    df = recalcular_posicions(df)

    return df


def crear_ranking_departaments(df_ranking):
    if "Departament" not in df_ranking.columns:
        return pd.DataFrame()

    df_temp = df_ranking.copy()
    df_temp["Departament"] = df_temp["Departament"].fillna("Sense departament").astype(str).str.strip()

    resum = (
        df_temp
        .groupby("Departament", as_index=False)
        .agg(
            Participants=("Participant", "count"),
            Punts_totals=("Punts", "sum"),
            Mitjana_punts=("Punts", "mean"),
            Millor_puntuacio=("Punts", "max")
        )
    )

    lider_departament = (
        df_temp
        .sort_values("Punts", ascending=False)
        .drop_duplicates("Departament")
        [["Departament", "Participant"]]
        .rename(columns={"Participant": "Líder departament"})
    )

    resum = resum.merge(lider_departament, on="Departament", how="left")

    resum["Punts_totals"] = resum["Punts_totals"].round(1)
    resum["Mitjana_punts"] = resum["Mitjana_punts"].round(1)
    resum["Millor_puntuacio"] = resum["Millor_puntuacio"].round(1)

    resum = resum.sort_values(
        ["Mitjana_punts", "Punts_totals"],
        ascending=[False, False]
    ).reset_index(drop=True)

    resum["Posició"] = resum.index + 1

    if not resum.empty:
        lider = float(resum["Mitjana_punts"].iloc[0])
        resum["Dif líder"] = (resum["Mitjana_punts"] - lider).round(1)
    else:
        resum["Dif líder"] = 0

    return resum[
        [
            "Posició",
            "Departament",
            "Participants",
            "Mitjana_punts",
            "Punts_totals",
            "Millor_puntuacio",
            "Líder departament",
            "Dif líder"
        ]
    ]


# --------------------------------------------------
# SNAPSHOT / MOVIMENT AUTOMÀTIC
# --------------------------------------------------
def aplicar_moviment(df_ranking, excel_mtime):
    df_actual = df_ranking.copy()

    SNAPSHOT_PREV = "ranking_snapshot_previous.csv"
    SNAPSHOT_CURR = "ranking_snapshot_current.csv"

    def carregar(path):
        if os.path.exists(path):
            try:
                return pd.read_csv(path)
            except:
                return pd.DataFrame()
        return pd.DataFrame()

    def guardar(path, df):
        df.to_csv(path, index=False)

    # Carreguem snapshots
    df_prev = carregar(SNAPSHOT_PREV)
    df_curr = carregar(SNAPSHOT_CURR)

    # Si és la primera vegada
    if df_curr.empty:
        df_actual["Indicador"] = "⚪ —"
        df_actual["Mov. posició"] = "—"
        df_actual["Canvi punts"] = 0.0

        guardar(SNAPSHOT_CURR, df_actual[["Participant", "Punts", "Posició"]])
        return df_actual

    # Si Excel NO ha canviat → no recalcular res
    meta = carregar_meta_snapshot()
    meta_mtime = meta.get("excel_mtime", None)

    if meta_mtime == excel_mtime and not df_curr.empty:
        # tornar mateix estat (no recalcular)
        df_actual = df_actual.merge(
            df_curr,
            on="Participant",
            how="left",
            suffixes=("", "_old")
        )

        return df_actual

    # --- Excel ha canviat ---
    # movem current → previous
    if not df_curr.empty:
        guardar(SNAPSHOT_PREV, df_curr)

    # calculem comparant amb anterior
    df_prev = carregar(SNAPSHOT_PREV)

    if df_prev.empty:
        df_actual["Indicador"] = "⚪ —"
        df_actual["Mov. posició"] = "—"
        df_actual["Canvi punts"] = 0.0
    else:
        df_prev["Participant"] = df_prev["Participant"].astype(str).str.strip()

        df_actual = df_actual.merge(
            df_prev,
            on="Participant",
            how="left",
            suffixes=("", "_prev")
        )

        df_actual["Punts_prev"] = pd.to_numeric(df_actual["Punts_prev"], errors="coerce")
        df_actual["Posició_prev"] = pd.to_numeric(df_actual["Posició_prev"], errors="coerce")

        df_actual["Canvi punts"] = (df_actual["Punts"] - df_actual["Punts_prev"]).fillna(0).round(1)
        df_actual["Canvi posició"] = (df_actual["Posició_prev"] - df_actual["Posició"]).fillna(0)

        def indicador(row):
            if pd.isna(row["Punts_prev"]):
                return "⚪ —"

            if row["Canvi punts"] > 0:
                return "🟢 ▲"
            elif row["Canvi punts"] < 0:
                return "🔴 ▼"
            else:
                return "⚪ —"

        def mov(row):
            if pd.isna(row["Posició_prev"]):
                return "—"

            c = int(row["Canvi posició"])

            if c > 0:
                return f"+{c}"
            elif c < 0:
                return str(c)
            else:
                return "—"

        df_actual["Indicador"] = df_actual.apply(indicador, axis=1)
        df_actual["Mov. posició"] = df_actual.apply(mov, axis=1)

    # guardem nou current
    guardar(
        SNAPSHOT_CURR,
        df_actual[["Participant", "Punts", "Posició", "Indicador", "Mov. posició", "Canvi punts"]]
    )

    guardar_meta_snapshot(excel_mtime)

    return df_actual

# --------------------------------------------------
# TAULES I GRÀFICS
# --------------------------------------------------
def highlight_leader(row):
    if row["Posició"] == 1:
        return ["background-color: #ffe066; font-weight: bold;"] * len(row)
    return [""] * len(row)


def mostrar_taula_ranking(df):
    cols = ["Posició"]

    if "Indicador" in df.columns:
        cols.append("Indicador")

    if "Mov. posició" in df.columns:
        cols.append("Mov. posició")

    cols.append("Participant")

    if "Departament" in df.columns:
        cols.append("Departament")

    cols.append("Punts")
    cols.append("Dif líder")

    # Canvi punts al final
    if "Canvi punts" in df.columns:
        cols.append("Canvi punts")

    cols_existents = [c for c in cols if c in df.columns]

    df_display = df[cols_existents].copy()

    df_display["Punts"] = df_display["Punts"].astype(float).round(1)
    df_display["Dif líder"] = df_display["Dif líder"].astype(float).round(1)

    format_dict = {
        "Punts": "{:.1f}",
        "Dif líder": "{:.1f}"
    }

    if "Canvi punts" in df_display.columns:
        df_display["Canvi punts"] = pd.to_numeric(
            df_display["Canvi punts"],
            errors="coerce"
        ).fillna(0.0).round(1)

        format_dict["Canvi punts"] = "{:+.1f}"

    styled = (
        df_display
        .style
        .apply(highlight_leader, axis=1)
        .format(format_dict)
    )

    column_config = {
        "Posició": st.column_config.NumberColumn("Posició", format="%d"),
        "Punts": st.column_config.NumberColumn("Punts", format="%.1f"),
        "Dif líder": st.column_config.NumberColumn("Dif líder", format="%.1f"),
    }

    if "Canvi punts" in df_display.columns:
        column_config["Canvi punts"] = st.column_config.NumberColumn(
            "Canvi punts",
            format="%+.1f"
        )

    st.dataframe(
        styled,
        use_container_width=True,
        hide_index=True,
        column_config=column_config
    )


def mostrar_taula_departaments(df_dep):
    if df_dep.empty:
        st.info("Afegeix una columna 'Departament' al costat de 'Participants' al full Porra per activar aquest mode.")
        return

    styled = (
        df_dep
        .style
        .apply(highlight_leader, axis=1)
        .format({
            "Mitjana_punts": "{:.1f}",
            "Punts_totals": "{:.1f}",
            "Millor_puntuacio": "{:.1f}",
            "Dif líder": "{:.1f}"
        })
    )

    st.dataframe(
        styled,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Posició": st.column_config.NumberColumn("Posició", format="%d"),
            "Participants": st.column_config.NumberColumn("Participants", format="%d"),
            "Mitjana_punts": st.column_config.NumberColumn("Mitjana punts", format="%.1f"),
            "Punts_totals": st.column_config.NumberColumn("Punts totals", format="%.1f"),
            "Millor_puntuacio": st.column_config.NumberColumn("Millor puntuació", format="%.1f"),
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


def mostrar_grafic_departaments(df_dep):
    if df_dep.empty:
        return

    chart_data = df_dep.copy().sort_values("Mitjana_punts", ascending=False)
    chart_height = max(350, len(chart_data) * 46)

    chart = (
        alt.Chart(chart_data)
        .mark_bar(color="#0f9d58")
        .encode(
            x=alt.X(
                "Mitjana_punts:Q",
                title="Mitjana de punts",
                scale=alt.Scale(zero=False)
            ),
            y=alt.Y(
                "Departament:N",
                sort="-x",
                title=None,
                axis=alt.Axis(labelLimit=560, labelFontSize=13)
            ),
            tooltip=[
                alt.Tooltip("Posició:Q", title="Posició"),
                alt.Tooltip("Departament:N", title="Departament"),
                alt.Tooltip("Participants:Q", title="Participants"),
                alt.Tooltip("Mitjana_punts:Q", title="Mitjana punts", format=".1f"),
                alt.Tooltip("Punts_totals:Q", title="Punts totals", format=".1f"),
                alt.Tooltip("Líder departament:N", title="Líder departament")
            ]
        )
        .properties(height=chart_height)
    )

    st.altair_chart(chart, use_container_width=True)


def obtenir_pichichi_real(df_resultats_display, col_pichichi, col_gols):
    if col_pichichi not in df_resultats_display.columns or col_gols not in df_resultats_display.columns:
        return "Pendent", "Pendent"

    taula = df_resultats_display[[col_pichichi, col_gols]].copy()

    taula[col_pichichi] = taula[col_pichichi].astype(str).str.strip()
    taula[col_gols] = pd.to_numeric(taula[col_gols], errors="coerce")

    taula = taula[
        (taula[col_pichichi] != "") &
        (~taula[col_pichichi].str.lower().isin(["nan", "nat", "pendent"])) &
        (taula[col_gols] >= 1)
    ]

    if taula.empty:
        return "Pendent", "Pendent"

    taula = taula.sort_values(col_gols, ascending=False).reset_index(drop=True)

    jugador = taula.iloc[0][col_pichichi]
    gols = int(taula.iloc[0][col_gols])

    return jugador, str(gols)


def obtenir_prediccions_fase(df_j, prefix, quantitat):
    files = []

    for i in range(1, quantitat + 1):
        col = f"{prefix}_{i}"

        if col in df_j.columns:
            valor = valor_o_pendent(df_j[col].values[0])
        else:
            valor = "Pendent"

        files.append({
            "Posició": i,
            "Equip": afegir_bandera(valor)
        })

    return pd.DataFrame(files)


def mostrar_prediccions_eliminatoria_participant(df_j):
    st.write("### 🧭 Prediccions fase eliminatòria")

    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "Vuitens",
        "Quarts",
        "Semis",
        "Final",
        "Campió"
    ])

    with tab1:
        df_vuitens = obtenir_prediccions_fase(df_j, "Vuitens", 16)
        st.dataframe(df_vuitens, use_container_width=True, hide_index=True)

    with tab2:
        df_quarts = obtenir_prediccions_fase(df_j, "Quarts", 8)
        st.dataframe(df_quarts, use_container_width=True, hide_index=True)

    with tab3:
        df_semis = obtenir_prediccions_fase(df_j, "Semis", 4)
        st.dataframe(df_semis, use_container_width=True, hide_index=True)

    with tab4:
        finalistes = []

        for col in ["Final_1", "Final_2"]:
            if col in df_j.columns:
                finalistes.append(afegir_bandera(valor_o_pendent(df_j[col].values[0])))
            else:
                finalistes.append("Pendent")

        df_final = pd.DataFrame({
            "Finalista": ["Finalista 1", "Finalista 2"],
            "Equip": finalistes
        })

        st.dataframe(df_final, use_container_width=True, hide_index=True)

    with tab5:
        campio = valor_o_pendent(df_j["Campió"].values[0]) if "Campió" in df_j.columns else "Pendent"

        df_campio = pd.DataFrame({
            "Concepte": ["Campió previst"],
            "Equip": [afegir_bandera(campio)]
        })

        st.dataframe(df_campio, use_container_width=True, hide_index=True)


# --------------------------------------------------
# ESTILS + FONS
# --------------------------------------------------
img_base64 = carregar_imatge_base64(BACKGROUND_IMAGE)

if img_base64:
    background_css = f"""
    background-image:
        linear-gradient(rgba(0,0,0,0.52), rgba(0,0,0,0.72)),
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
        background: rgba(255,255,255,0.92);
        border-radius: 24px;
        margin-top: 24px;
        margin-bottom: 24px;
        box-shadow: 0px 8px 30px rgba(0,0,0,0.25);
    }}

    .title {{
        font-size: clamp(32px, 5vw, 52px);
        font-weight: 900;
        margin-bottom: 0px;
        color: #102a43;
        letter-spacing: -1px;
    }}

    .subtitle {{
        font-size: clamp(14px, 2vw, 18px);
        color: #334e68;
        margin-top: 0px;
        margin-bottom: 25px;
    }}

    .card {{
        padding: 18px;
        border-radius: 18px;
        text-align: center;
        box-shadow: 0px 4px 20px rgba(0,0,0,0.18);
        height: 178px;
        min-height: 178px;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        box-sizing: border-box;
        overflow: hidden;
        width: 100%;
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

    .purplecard {{
        background: linear-gradient(135deg, #6f42c1, #b982ff);
        color: white;
        margin-top: 18px;
        margin-bottom: 18px;
    }}

    .card h3 {{
        margin: 0px 0px 14px 0px;
        font-size: clamp(15px, 2vw, 24px);
        line-height: 1.15;
        max-width: 100%;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }}

    .card h1 {{
        margin: 0px;
        font-size: clamp(24px, 4vw, 40px);
        line-height: 1.1;
        max-width: 100%;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }}

    .card p {{
        margin: 12px 0px 0px 0px;
        font-size: clamp(11px, 1.5vw, 15px);
        max-width: 100%;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }}

    @media (max-width: 768px) {{
        .block-container {{
            padding-left: 0.8rem;
            padding-right: 0.8rem;
            border-radius: 16px;
        }}

        .card {{
            height: auto;
            min-height: 140px;
            margin-bottom: 12px;
        }}

        .card h3 {{
            white-space: normal;
        }}

        .card h1 {{
            white-space: normal;
        }}

        .card p {{
            white-space: normal;
        }}
    }}
    </style>
    """,
    unsafe_allow_html=True
)


# --------------------------------------------------
# CARREGAR DADES
# --------------------------------------------------
excel_mtime = os.path.getmtime(EXCEL_FILE) if os.path.exists(EXCEL_FILE) else 0
data_actualitzacio = obtenir_data_actualitzacio_fitxer(EXCEL_FILE)

df_porra, df_resultats = carregar_dades(EXCEL_FILE, excel_mtime)
df_ranking = crear_ranking_des_de_porra(df_porra)
df_ranking = aplicar_moviment(df_ranking, excel_mtime)

df_departaments = crear_ranking_departaments(df_ranking)

num_participants = len(df_ranking)
premi_guanyador = num_participants * PREU_PARTICIPACIO

te_departaments = "Departament" in df_ranking.columns and not df_departaments.empty


# --------------------------------------------------
# TÍTOL
# --------------------------------------------------
st.markdown('<p class="title">🏆 PORRA MUNDIAL</p>', unsafe_allow_html=True)
st.markdown(
    '<p class="subtitle">Classificació en viu, moviment respecte l’última actualització, competició per departaments i resultats reals</p>',
    unsafe_allow_html=True
)


# --------------------------------------------------
# INFO PRINCIPAL
# --------------------------------------------------
info1, info2, info3 = st.columns(3, gap="small")

info1.markdown(
    f"""
    <div class='card darkcard'>
        <h3>🕒 Dades actualitzades</h3>
        <h1>{data_actualitzacio}</h1>
    </div>
    """,
    unsafe_allow_html=True
)

info2.markdown(
    f"""
    <div class='card greencard'>
        <h3>🎁 Premi guanyador</h3>
        <h1>{premi_guanyador} €</h1>
        <p>{num_participants} participants x {PREU_PARTICIPACIO} €</p>
    </div>
    """,
    unsafe_allow_html=True
)

info3.markdown(
    f"""
    <div class='card bluecard'>
        <h3>👥 Participants</h3>
        <h1>{num_participants}</h1>
        <p>porres registrades</p>
    </div>
    """,
    unsafe_allow_html=True
)


# --------------------------------------------------
# DEPARTAMENT LÍDER
# --------------------------------------------------
if te_departaments:
    dept_lider = df_departaments.iloc[0]

    st.markdown(
        f"""
        <div class='card purplecard'>
            <h3>🏢 Departament líder</h3>
            <h1>{dept_lider["Departament"]}</h1>
            <p>Mitjana {float(dept_lider["Mitjana_punts"]):.1f} punts · {int(dept_lider["Participants"])} participants</p>
        </div>
        """,
        unsafe_allow_html=True
    )


# --------------------------------------------------
# TOP 3 GENERAL
# --------------------------------------------------
st.subheader("🥇 TOP 3 General")

top3 = df_ranking.head(3)

c1, c2, c3 = st.columns(3, gap="small")

top_cards = [
    ("🥇", "gold", top3.iloc[0]),
    ("🥈", "silver", top3.iloc[1]),
    ("🥉", "bronze", top3.iloc[2]),
]

for col, (medalla, classe, row) in zip([c1, c2, c3], top_cards):
    subtext = "punts"

    if "Departament" in row.index:
        subtext = f"{row['Departament']}"

    indicador = row["Indicador"] if "Indicador" in row.index else ""
    mov_pos = row["Mov. posició"] if "Mov. posició" in row.index else ""

    col.markdown(
        f"""
        <div class='card {classe}'>
            <h3>{medalla} {row["Participant"]}</h3>
            <h1>{float(row["Punts"]):.1f}</h1>
            <p>{subtext} · {indicador} {mov_pos}</p>
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

        col_dep_original = obtenir_columna_departament(df_porra)

        if col_dep_original is not None:
            departament_jugador = valor_o_pendent(df_j[col_dep_original].values[0])
            c1.metric("Departament", departament_jugador)

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

        c2.write("### ⚽ Prediccions principals")
        c2.write(f"🏆 Campió: {afegir_bandera(valor_o_pendent(df_j['Campió'].values[0]))}")
        c2.write(f"📌 Resultat final: {resultat_final}")
        c2.write(f"⭐ MVP: {valor_o_pendent(df_j['MVP'].values[0])}")
        c2.write(f"⚽ Pichichi: {valor_o_pendent(df_j['Pichichi'].values[0])}")

        mostrar_prediccions_eliminatoria_participant(df_j)

else:
    st.info("Selecciona un participant per veure el detall de punts i prediccions.")


# --------------------------------------------------
# COMPETICIÓ PER DEPARTAMENTS
# --------------------------------------------------
st.subheader("🏢 Competició per departaments")

if te_departaments:
    st.write("Rànquing calculat per **mitjana de punts** del departament. També es mostren punts totals, millor puntuació i líder del departament.")

    mostrar_taula_departaments(df_departaments)

    st.write("#### 📈 Gràfic departaments")
    mostrar_grafic_departaments(df_departaments)

    st.write("### 🎯 Classificació interna per departament")

    departaments_opcions = sorted(df_ranking["Departament"].dropna().astype(str).unique().tolist())

    departament_sel = st.selectbox(
        "Selecciona departament",
        departaments_opcions,
        index=None,
        placeholder="Selecciona un departament..."
    )

    if departament_sel:
        df_dep_individual = df_ranking[df_ranking["Departament"] == departament_sel].copy()
        df_dep_individual = recalcular_posicions(df_dep_individual)

        st.write(f"### 🥇 TOP 3 · {departament_sel}")

        dep_top = df_dep_individual.head(3)
        dep_cols = st.columns(3, gap="small")
        medalles = ["🥇", "🥈", "🥉"]
        classes = ["gold", "silver", "bronze"]

        for i in range(min(3, len(dep_top))):
            indicador_dep = dep_top.iloc[i]["Indicador"] if "Indicador" in dep_top.columns else ""
            mov_dep = dep_top.iloc[i]["Mov. posició"] if "Mov. posició" in dep_top.columns else ""

            dep_cols[i].markdown(
                f"""
                <div class='card {classes[i]}'>
                    <h3>{medalles[i]} {dep_top.iloc[i]["Participant"]}</h3>
                    <h1>{float(dep_top.iloc[i]["Punts"]):.1f}</h1>
                    <p>{departament_sel} · {indicador_dep} {mov_dep}</p>
                </div>
                """,
                unsafe_allow_html=True
            )

        st.write(f"### 📊 Classificació interna · {departament_sel}")
        mostrar_taula_ranking(df_dep_individual)

        st.write(f"### 📈 Gràfic · {departament_sel}")
        mostrar_grafic_punts(df_dep_individual, color="#6f42c1", altura_minima=350)

else:
    st.info("Per activar aquest apartat, afegeix una columna 'Departament' al costat de 'Participants' al full Porra.")


# --------------------------------------------------
# LLIGUETES
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

    df_lligueta = recalcular_posicions(df_lligueta)

    st.write(f"Participants seleccionats: **{len(participants_filtrats)}**")

    mostrar_taula_ranking(df_lligueta)

    st.write("#### 📈 Gràfic de la lligueta")
    mostrar_grafic_punts(df_lligueta, color="#0f9d58", altura_minima=350)

else:
    st.write("Selecciona participants per crear una classificació reduïda tipus lligueta.")


# --------------------------------------------------
# RESULTATS REALS
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
campio_real = afegir_bandera(primer_valor_o_pendent(df_resultats_display, COL_CAMPIO))
mvp_real = primer_valor_o_pendent(df_resultats_display, COL_MVP)
resultat_final_real = primer_valor_o_pendent(df_resultats_display, COL_RESULTAT_FINAL)

pichichi_real, gols_pichichi = obtenir_pichichi_real(
    df_resultats_display,
    COL_PICHICHI,
    COL_GOLS
)

st.write("### 🏟️ Resum oficial")

r1, r2, r3, r4 = st.columns(4, gap="small")

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

pichichi_subtext = f"{gols_pichichi} gols" if gols_pichichi != "Pendent" else "Pendent"

r3.markdown(
    f"""
    <div class='card bronze'>
        <h3>⚽ Pichichi</h3>
        <h1 style='font-size:25px'>{pichichi_real}</h1>
        <p>{pichichi_subtext}</p>
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
# FASE DE GRUPS
# --------------------------------------------------
st.write("### 🧩 Fase de grups")

grups = {}

if all(col in df_resultats_display.columns for col in [COL_GRUP, COL_POSICIO, COL_EQUIP]):
    for _, row in df_resultats_display.iterrows():
        grup = str(row.get(COL_GRUP, "")).strip()
        posicio = str(row.get(COL_POSICIO, "")).strip()
        equip = str(row.get(COL_EQUIP, "")).strip()

        if grup == "" or equip == "":
            continue

        if grup not in grups:
            grups[grup] = {
                "1r": "",
                "2n": "",
                "3r": ""
            }

        if posicio in ["1r", "1", "1º"]:
            grups[grup]["1r"] = afegir_bandera(equip)
        elif posicio in ["2n", "2", "2º"]:
            grups[grup]["2n"] = afegir_bandera(equip)
        elif posicio in ["3r", "3", "3º"]:
            grups[grup]["3r"] = afegir_bandera(equip)

    if len(grups) > 0:
        df_grups = pd.DataFrame(grups)
        df_grups = df_grups.reindex(["1r", "2n", "3r"])
        df_grups = df_grups.reset_index().rename(columns={"index": "Posició"})

        st.dataframe(
            df_grups,
            use_container_width=True,
            hide_index=True
        )
    else:
        st.info("No hi ha dades de fase de grups configurades.")
else:
    st.info("No hi ha dades de fase de grups configurades.")


# --------------------------------------------------
# FASE ELIMINATÒRIA
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
        elif len(valors) == 1 and normalitzar_text(valors[0]) == "pendent":
            detall = "Pendent"
        else:
            if fase == COL_MVP:
                detall = " · ".join(valors)
            else:
                detall = " · ".join([afegir_bandera(v) for v in valors])

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
# PICHICHI
# --------------------------------------------------
st.write("### ⚽ Jugador pichichi")

if COL_PICHICHI in df_resultats_display.columns and COL_GOLS in df_resultats_display.columns:
    taula_pichichi = df_resultats_display[[COL_PICHICHI, COL_GOLS]].copy()

    taula_pichichi[COL_PICHICHI] = taula_pichichi[COL_PICHICHI].astype(str).str.strip()
    taula_pichichi[COL_GOLS] = pd.to_numeric(
        taula_pichichi[COL_GOLS],
        errors="coerce"
    )

    taula_pichichi = taula_pichichi[
        (taula_pichichi[COL_PICHICHI] != "") &
        (~taula_pichichi[COL_PICHICHI].str.lower().isin(["nan", "nat", "pendent"])) &
        (taula_pichichi[COL_GOLS] >= 1)
    ]

    if taula_pichichi.empty:
        taula_pichichi = pd.DataFrame({
            COL_PICHICHI: ["Pendent"],
            COL_GOLS: ["Pendent"]
        })
    else:
        taula_pichichi = taula_pichichi.sort_values(COL_GOLS, ascending=False)
        taula_pichichi[COL_GOLS] = taula_pichichi[COL_GOLS].astype("Int64")

    st.dataframe(
        taula_pichichi,
        use_container_width=True,
        hide_index=True
    )
else:
    taula_pichichi = pd.DataFrame({
        "Jugador Pichichi": ["Pendent"],
        "Gols": ["Pendent"]
    })

    st.dataframe(
        taula_pichichi,
        use_container_width=True,
        hide_index=True
    )


# --------------------------------------------------
# RESULTAT FINAL
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
