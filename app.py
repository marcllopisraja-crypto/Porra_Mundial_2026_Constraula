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

# --- SISTEMA DE SEGURETAT ANTI-PANTALLA BLANCA ---
if not os.path.exists(EXCEL_FILE):
    st.error(f"❌ No s'ha trobat l'arxiu de dades: **{EXCEL_FILE}**")
    st.warning("Revisa que l'arxiu estigui pujat a GitHub a la carpeta principal i que les majúscules i minúscules del nom coincideixin exactament.")
    st.stop()


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
        st.stop()

    if "Total Punts" not in df_porra.columns:
        st.error("No s'ha trobat la columna 'Total Punts' al full Porra.")
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
def carregar_meta_snapshot():
    if not os.path.exists(SNAPSHOT_META_FILE):
        return {}
    try:
        with open(SNAPSHOT_META_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def guardar_meta_snapshot(excel_mtime):
    meta = {
        "excel_mtime": float(excel_mtime),
        "updated_at": datetime.now(tz=ZoneInfo("Europe/Madrid")).isoformat()
    }
    with open(SNAPSHOT_META_FILE, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)


def carregar_csv_segura(path):
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        return pd.read_csv(path)
    except Exception:
        return pd.DataFrame()


def guardar_snapshot_actual(df_ranking):
    cols = ["Participant", "Punts", "Posició"]
    df_snapshot = df_ranking[cols].copy()
    df_snapshot = df_snapshot.rename(columns={
        "Punts": "Punts anteriors",
        "Posició": "Posició anterior"
    })
    df_snapshot.to_csv(SNAPSHOT_CURRENT_FILE, index=False)


def guardar_snapshot_display(df_ranking):
    cols = [
        "Participant",
        "Indicador",
        "Mov. posició",
        "Canvi punts",
        "Canvi posició",
        "Punts anteriors",
        "Posició anterior"
    ]
    cols_existents = [c for c in cols if c in df_ranking.columns]
    df_display = df_ranking[cols_existents].copy()
    df_display.to_csv(SNAPSHOT_DISPLAY_FILE, index=False)


def aplicar_moviment(df_ranking, excel_mtime):
    df_actual = df_ranking.copy()

    def posar_neutral(df):
        df = df.copy()
        df["Posició anterior"] = df["Posició"]
        df["Punts anteriors"] = df["Punts"]
        df["Canvi punts"] = 0.0
        df["Canvi posició"] = 0
        df["Indicador"] = "⚪ —"
        df["Mov. posició"] = "—"
        return df

    meta = carregar_meta_snapshot()
    meta_mtime = meta.get("excel_mtime", None)

    # 1. Si l'Excel no ha canviat, recuperem el moviment guardat
    if meta_mtime is not None and float(meta_mtime) == float(excel_mtime):
        df_mov = carregar_csv_segura(SNAPSHOT_DISPLAY_FILE)

        if not df_mov.empty and "Participant" in df_mov.columns:
            df_actual = df_actual.merge(df_mov, on="Participant", how="left")

            if "Indicador" in df_actual.columns:
                indicadors = df_actual["Indicador"].dropna().astype(str)
                tot_nou = len(indicadors) > 0 and indicadors.eq("🆕 Nou").all()
            else:
                tot_nou = True

            if tot_nou:
                df_actual = posar_neutral(df_actual)
                guardar_snapshot_actual(df_actual)
                guardar_snapshot_display(df_actual)
                guardar_meta_snapshot(excel_mtime)
                return df_actual

            df_actual["Indicador"] = df_actual["Indicador"].fillna("⚪ —")
            df_actual["Mov. posició"] = df_actual["Mov. posició"].fillna("—")
            df_actual["Canvi punts"] = pd.to_numeric(df_actual["Canvi punts"], errors="coerce").fillna(0.0).round(1)
            
            if "Canvi posició" not in df_actual.columns:
                df_actual["Canvi posició"] = 0

            return df_actual

        # Si falla alguna cosa, posem neutral
        df_actual = posar_neutral(df_actual)
        guardar_snapshot_actual(df_actual)
        guardar_snapshot_display(df_actual)
        guardar_meta_snapshot(excel_mtime)
        return df_actual

    # 2. Si l'Excel HA CANVIAT, comparem contra l'últim snapshot
    df_prev = carregar_csv_segura(SNAPSHOT_CURRENT_FILE)

    if df_prev.empty or "Participant" not in df_prev.columns:
        df_actual = posar_neutral(df_actual)
        guardar_snapshot_actual(df_actual)
        guardar_snapshot_display(df_actual)
        guardar_meta_snapshot(excel_mtime)
        return df_actual

    df_prev["Participant"] = df_prev["Participant"].astype(str).str.strip()

    df_actual = df_actual.merge(df_prev, on="Participant", how="left")

    # Convertim a numèric i calculem diferències
    df_actual["Punts anteriors"] = pd.to_numeric(df_actual["Punts anteriors"], errors="coerce")
    df_actual["Posició anterior"] = pd.to_numeric(df_actual["Posició anterior"], errors="coerce")

    # Càlcul Punts
    df_actual["Canvi punts"] = (df_actual["Punts"] - df_actual["Punts anteriors"]).round(1)
    df_actual["Canvi punts"] = pd.to_numeric(df_actual["Canvi punts"], errors="coerce").fillna(0.0).round(1)

    # Càlcul Posicions (Positiu = puja al ranking, Negatiu = baixa)
    df_actual["Canvi posició"] = (df_actual["Posició anterior"] - df_actual["Posició"]).fillna(0)

    # --- NOVES FUNCIONS DE LÒGICA VISUAL ---
    def indicador_posicio(row):
        if pd.isna(row.get("Posició anterior")):
            return "🆕"
        
        canvi = row["Canvi posició"]
        if canvi > 0:
            return "🟢 ▲"
        elif canvi < 0:
            return "🔴 ▼"
        else:
            return "⚪ —"

    def moviment_posicio(row):
        if pd.isna(row.get("Posició anterior")):
            return "Nou"
        
        canvi = int(row["Canvi posició"])
        if canvi > 0:
            return f"+{canvi}"
        elif canvi < 0:
            return str(canvi)
        else:
            return "—"

    # Apliquem les noves funcions
    df_actual["Indicador"] = df_actual.apply(indicador_posicio, axis=1)
    df_actual["Mov. posició"] = df_actual.apply(moviment_posicio, axis=1)

    # Guardem estat per a la propera vegada
    guardar_snapshot_actual(df_actual)
    guardar_snapshot_display(df_actual)
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
        .properties