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
LOGO_IMAGE = "Logo RGB fondo transparente letra negra Constraula.png"
PREU_PARTICIPACIO = 5

SNAPSHOT_CURRENT_FILE = "ranking_snapshot_current.csv"
SNAPSHOT_DISPLAY_FILE = "ranking_snapshot_display.csv"
SNAPSHOT_META_FILE = "ranking_snapshot_meta.json"

# --------------------------------------------------
# CONFIGURACIÓ DE PUNTS MÀXIMS
# --------------------------------------------------
MAX_PUNTS_CATEGORIA = {
    "1rs grup": 24,
    "2ns grup": 24,
    "3rs grup": 18,
    "Vuitens": 16,
    "Quarts": 8,
    "Semis": 4,
    "Finalistes": 6, 
    "Campió": 5,
    "Resultat final": 5,
    "MVP": 5,
    "Bota d'Or": 10 # Valor orientatiu
}

# --- SISTEMA DE SEGURETAT ANTI-PANTALLA BLANCA ---
if not os.path.exists(EXCEL_FILE):
    st.error(f"❌ No s'ha trobat l'arxiu de dades: **{EXCEL_FILE}**")
    st.warning("Revisa que l'arxiu estigui pujat a GitHub a la carpeta principal i que les majúscules i minúscules del nom coincideixin exactament.")
    st.stop()


# --------------------------------------------------
# BANDERES
# --------------------------------------------------
FLAGS = {
    "mexic": "🇲🇽", "corea del sud": "🇰🇷", "republica txeca": "🇨🇿",
    "suissa": "🇨🇭", "canada": "🇨🇦", "qatar": "🇶🇦", "escocia": "🏴",
    "marroc": "🇲🇦", "brasil": "🇧🇷", "estats units": "🇺🇸", "ee.uu": "🇺🇸",
    "australia": "🇦🇺", "turquia": "🇹🇷", "alemanya": "🇩🇪", "costa d'ivori": "🇨🇮",
    "cote d'ivoire": "🇨🇮", "equador": "🇪🇨", "suecia": "🇸🇪", "japo": "🇯🇵",
    "paisos baixos": "🇳🇱", "nova zelanda": "🇳🇿", "iran": "🇮🇷", "belgica": "🇧🇪",
    "uruguai": "🇺🇾", "arabia saudita": "🇸🇦", "espanya": "🇪🇸", "franca": "🇫🇷",
    "senegal": "🇸🇳", "iraq": "🇮🇶", "argentina": "🇦🇷", "algeria": "🇩🇿",
    "austria": "🇦🇹", "portugal": "🇵🇹", "rd congo": "🇨🇩", "uzbekistan": "🇺🇿",
    "anglaterra": "🏴", "croacia": "🇭🇷", "ghana": "🇬🇭", "egipte": "🇪🇬",
    "noruega": "🇳🇴", "colombia": "🇨🇴", "colòmbia": "🇨🇴", "bosnia i hercegovina": "🇧🇦",
    "paraguai": "🇵🇾", "tunisia": "🇹🇳", "tunísia": "🇹🇳", "cap verd": "🇨🇻",
    "jordania": "🇯🇴", "jordània": "🇯🇴", "panama": "🇵🇦", "panamà": "🇵🇦",
    "curaçao": "🇨🇼", "curacao": "🇨🇼", "haiti": "🇭🇹", "haití": "🇭🇹",
    "sud-africa": "🇿🇦", "sud-àfrica": "🇿🇦"
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
        if normalitzar_text(col) == "departament": return col
    for col in df.columns:
        if "depart" in normalitzar_text(col): return col
    return None

def afegir_bandera(valor):
    if pd.isna(valor): return "Pendent"
    text = str(valor).strip()
    if text == "" or normalitzar_text(text) in ["nan", "nat", "pendent"]: return "Pendent"
    text_norm = normalitzar_text(text)
    for pais, bandera in FLAGS.items():
        if pais in text_norm: return f"{bandera} {text}"
    return text

def valor_o_pendent(valor):
    if pd.isna(valor): return "Pendent"
    valor_text = str(valor).strip()
    if valor_text == "" or valor_text.lower() in ["nan", "nat"]: return "Pendent"
    return valor_text

def obtenir_data_actualitzacio_fitxer(path):
    if not os.path.exists(path): return "No disponible"
    timestamp = os.path.getmtime(path)
    dt = datetime.fromtimestamp(timestamp, tz=ZoneInfo("Europe/Madrid"))
    return dt.strftime("%d/%m/%Y")

def llista_valors_no_buits(df, columna):
    if columna not in df.columns: return []
    valors = df[columna].astype(str).str.strip().replace("nan", "").replace("NaT", "")
    valors_nets = []
    for valor in valors:
        if valor == "" or normalitzar_text(valor) in ["nan", "nat"]: continue
        if normalitzar_text(valor) == "pendent": valor = "Pendent"
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
    if len(valors) == 0: return "Pendent"
    return valors[0]

def preparar_taula_buida(df):
    df = df.copy()
    df = df.dropna(how="all").dropna(axis=1, how="all").fillna("")
    return df

def trobar_col_resultat_final_porra(df_porra):
    for col in df_porra.columns:
        if col.strip() == "Resultat final": return col
    for col in df_porra.columns:
        col_norm = normalitzar_text(col)
        if "resultat" in col_norm and "final" in col_norm and "punt" not in col_norm: return col
    return None

def recalcular_posicions(df):
    df = df.copy().sort_values("Punts", ascending=False).reset_index(drop=True)
    df["Posició"] = df.index + 1
    if not df.empty:
        df["Dif líder"] = (df["Punts"] - float(df["Punts"].iloc[0])).round(1)
    else:
        df["Dif líder"] = 0
    return df


# --------------------------------------------------
# FUNCIONS CREACIÓ DE RÀNQUINGS
# --------------------------------------------------
def crear_ranking_des_de_porra(df_porra):
    col_dep = obtenir_columna_departament(df_porra)
    cols_base = ["Participants", "Total Punts"]
    if col_dep is not None: cols_base.append(col_dep)
    df = df_porra[cols_base].copy()
    rename_map = {"Participants": "Participant", "Total Punts": "Punts"}
    if col_dep is not None: rename_map[col_dep] = "Departament"
    df = df.rename(columns=rename_map)
    df["Participant"] = df["Participant"].astype(str).str.strip()
    df["Punts"] = pd.to_numeric(df["Punts"], errors="coerce")
    if "Departament" in df.columns:
        df["Departament"] = df["Departament"].fillna("Sense departament").astype(str).str.strip().replace("", "Sense departament")
    df = df.dropna(subset=["Punts"])
    df = df[(df["Participant"] != "") & (~df["Participant"].str.contains("Total", case=False, na=False))]
    df["Punts"] = df["Punts"].round(1)
    return recalcular_posicions(df)

def crear_ranking_departaments(df_ranking):
    if "Departament" not in df_ranking.columns: return pd.DataFrame()
    df_temp = df_ranking.copy()
    df_temp["Departament"] = df_temp["Departament"].fillna("Sense departament").astype(str).str.strip()
    resum = df_temp.groupby("Departament", as_index=False).agg(
        Participants=("Participant", "count"),
        Punts_totals=("Punts", "sum"),
        Mitjana_punts=("Punts", "mean"),
        Millor_puntuacio=("Punts", "max")
    )
    lider_departament = df_temp.sort_values("Punts", ascending=False).drop_duplicates("Departament")[["Departament", "Participant"]].rename(columns={"Participant": "Líder departament"})
    resum = resum.merge(lider_departament, on="Departament", how="left")
    resum["Punts_totals"] = resum["Punts_totals"].round(1)
    resum["Mitjana_punts"] = resum["Mitjana_punts"].round(1)
    resum["Millor_puntuacio"] = resum["Millor_puntuacio"].round(1)
    resum = resum.sort_values(["Mitjana_punts", "Punts_totals"], ascending=[False, False]).reset_index(drop=True)
    resum["Posició"] = resum.index + 1
    if not resum.empty: resum["Dif líder"] = (resum["Mitjana_punts"] - float(resum["Mitjana_punts"].iloc[0])).round(1)
    else: resum["Dif líder"] = 0
    return resum[["Posició", "Departament", "Participants", "Mitjana_punts", "Punts_totals", "Millor_puntuacio", "Líder departament", "Dif líder"]]

# --------------------------------------------------
# SNAPSHOT / MOVIMENT AUTOMÀTIC
# --------------------------------------------------
def carregar_meta_snapshot():
    if not os.path.exists(SNAPSHOT_META_FILE): return {}
    try:
        with open(SNAPSHOT_META_FILE, "r", encoding="utf-8") as f: return json.load(f)
    except: return {}

def guardar_meta_snapshot(excel_mtime):
    meta = {"excel_mtime": float(excel_mtime), "updated_at": datetime.now(tz=ZoneInfo("Europe/Madrid")).isoformat()}
    with open(SNAPSHOT_META_FILE, "w", encoding="utf-8") as f: json.dump(meta, f, ensure_ascii=False, indent=2)

def carregar_csv_segura(path):
    if not os.path.exists(path): return pd.DataFrame()
    try: return pd.read_csv(path)
    except: return pd.DataFrame()

def guardar_snapshot_actual(df_ranking):
    df_snapshot = df_ranking[["Participant", "Punts", "Posició"]].copy().rename(columns={"Punts": "Punts anteriors", "Posició": "Posició anterior"})
    df_snapshot.to_csv(SNAPSHOT_CURRENT_FILE, index=False)

def guardar_snapshot_display(df_ranking):
    cols = ["Participant", "Evolució", "Canvi punts", "Canvi posició", "Punts anteriors", "Posició anterior"]
    df_ranking[[c for c in cols if c in df_ranking.columns]].copy().to_csv(SNAPSHOT_DISPLAY_FILE, index=False)

def aplicar_moviment(df_ranking, excel_mtime):
    df_actual = df_ranking.copy()
    def posar_neutral(df):
        df = df.copy()
        df["Posició anterior"] = df["Posició"]
        df["Punts anteriors"] = df["Punts"]
        df["Canvi punts"] = 0.0
        df["Canvi posició"] = 0
        df["Evolució"] = "⚪ —"
        return df
    meta = carregar_meta_snapshot()
    meta_mtime = meta.get("excel_mtime", None)

    if meta_mtime is not None and float(meta_mtime) == float(excel_mtime):
        df_mov = carregar_csv_segura(SNAPSHOT_DISPLAY_FILE)
        if not df_mov.empty and "Participant" in df_mov.columns:
            df_actual = df_actual.merge(df_mov, on="Participant", how="left")
            tot_nou = len(df_actual["Evolució"].dropna()) > 0 and df_actual["Evolució"].dropna().astype(str).eq("🆕 Nou").all() if "Evolució" in df_actual.columns else True
            if tot_nou:
                df_actual = posar_neutral(df_actual)
            else:
                df_actual["Evolució"] = df_actual["Evolució"].fillna("⚪ —")
                df_actual["Canvi punts"] = pd.to_numeric(df_actual["Canvi punts"], errors="coerce").fillna(0.0).round(1)
                if "Canvi posició" not in df_actual.columns: df_actual["Canvi posició"] = 0
            guardar_snapshot_actual(df_actual)
            guardar_snapshot_display(df_actual)
            guardar_meta_snapshot(excel_mtime)
            return df_actual
        return posar_neutral(df_actual)

    df_prev = carregar_csv_segura(SNAPSHOT_CURRENT_FILE)
    if df_prev.empty or "Participant" not in df_prev.columns:
        df_actual = posar_neutral(df_actual)
        guardar_snapshot_actual(df_actual)
        guardar_snapshot_display(df_actual)
        guardar_meta_snapshot(excel_mtime)
        return df_actual

    df_prev["Participant"] = df_prev["Participant"].astype(str).str.strip()
    df_actual = df_actual.merge(df_prev, on="Participant", how="left")
    df_actual["Canvi punts"] = (df_actual["Punts"] - pd.to_numeric(df_actual["Punts anteriors"], errors="coerce")).round(1)
    df_actual["Canvi posició"] = (pd.to_numeric(df_actual["Posició anterior"], errors="coerce") - df_actual["Posició"]).fillna(0)

    if df_actual["Canvi punts"].fillna(0).eq(0).all() and df_actual["Canvi posició"].fillna(0).eq(0).all() and len(df_actual) == len(df_prev):
        df_old_display = carregar_csv_segura(SNAPSHOT_DISPLAY_FILE)
        if not df_old_display.empty and "Participant" in df_old_display.columns:
            df_actual = df_actual.drop(columns=["Canvi punts", "Canvi posició"], errors="ignore").merge(df_old_display[["Participant", "Evolució", "Canvi punts", "Canvi posició"]], on="Participant", how="left")
            df_actual["Evolució"] = df_actual["Evolució"].fillna("⚪ —")
            df_actual["Canvi punts"] = pd.to_numeric(df_actual["Canvi punts"], errors="coerce").fillna(0.0).round(1)
            guardar_meta_snapshot(excel_mtime)
            return df_actual

    df_actual["Canvi punts"] = pd.to_numeric(df_actual["Canvi punts"], errors="coerce").fillna(0.0).round(1)
    def evolucio_unificada(row):
        if pd.isna(row.get("Posició anterior")): return "🆕 Nou"
        canvi = int(row["Canvi posició"])
        if canvi > 0: return f"🟢 ▲ +{canvi}"
        elif canvi < 0: return f"🔴 ▼ {canvi}"
        else: return "⚪ —"
    df_actual["Evolució"] = df_actual.apply(evolucio_unificada, axis=1)
    guardar_snapshot_actual(df_actual)
    guardar_snapshot_display(df_actual)
    guardar_meta_snapshot(excel_mtime)
    return df_actual


# --------------------------------------------------
# TAULES I GRÀFICS
# --------------------------------------------------
def highlight_leader(row):
    if row["Posició"] == 1: return ["background-color: #ffe066; font-weight: bold;"] * len(row)
    return [""] * len(row)

def mostrar_taula_ranking(df):
    cols = ["Posició"] + (["Evolució"] if "Evolució" in df.columns else []) + ["Participant"] + (["Departament"] if "Departament" in df.columns else []) + ["Punts", "Dif líder"] + (["Canvi punts"] if "Canvi punts" in df.columns else [])
    df_display = df[[c for c in cols if c in df.columns]].copy()
    format_dict = {"Punts": "{:.1f}", "Dif líder": "{:.1f}"}
    if "Canvi punts" in df_display.columns:
        df_display["Canvi punts"] = pd.to_numeric(df_display["Canvi punts"], errors="coerce").fillna(0.0).round(1)
        format_dict["Canvi punts"] = "{:+.1f}"
    
    styled = df_display.style.apply(highlight_leader, axis=1).format(format_dict)
    column_config = {
        "Posició": st.column_config.NumberColumn("Posició", format="%d"),
        "Punts": st.column_config.NumberColumn("Punts", format="%.1f"),
        "Dif líder": st.column_config.NumberColumn("Dif líder", format="%.1f"),
    }
    if "Canvi punts" in df_display.columns: column_config["Canvi punts"] = st.column_config.NumberColumn("Canvi punts", format="%+.1f")
    st.dataframe(styled, use_container_width=True, hide_index=True, column_config=column_config)

def mostrar_taula_departaments(df_dep):
    if df_dep.empty:
        st.info("Afegeix una columna 'Departament' al costat de 'Participants' al full Porra per activar aquest mode.")
        return
    styled = df_dep.style.apply(highlight_leader, axis=1).format({"Mitjana_punts": "{:.1f}", "Punts_totals": "{:.1f}", "Millor_puntuacio": "{:.1f}", "Dif líder": "{:.1f}"})
    st.dataframe(styled, use_container_width=True, hide_index=True, column_config={
        "Posició": st.column_config.NumberColumn("Posició", format="%d"),
        "Participants": st.column_config.NumberColumn("Participants", format="%d"),
        "Mitjana_punts": st.column_config.NumberColumn("Mitjana punts", format="%.1f"),
        "Punts_totals": st.column_config.NumberColumn("Punts totals", format="%.1f"),
        "Millor_puntuacio": st.column_config.NumberColumn("Millor puntuació", format="%.1f"),
        "Dif líder": st.column_config.NumberColumn("Dif líder", format="%.1f"),
    })

def mostrar_grafic_punts(df, color_scheme="blues", altura_minima=950):
    if df.empty: return
    chart_data = df[["Posició", "Participant", "Punts", "Dif líder"]].copy()
    chart_data["Punts"] = pd.to_numeric(chart_data["Punts"], errors="coerce").fillna(0.0)
    chart_data = chart_data.sort_values("Punts", ascending=False)
    chart_height = max(altura_minima, len(chart_data) * 40)
    max_dom = 10.0 if pd.isna(chart_data["Punts"].max()) or chart_data["Punts"].max() <= 0 else float(chart_data["Punts"].max() * 1.15) 
    bars = alt.Chart(chart_data).mark_bar(cornerRadiusEnd=6, height=22).encode(
        x=alt.X("Punts:Q", title="Punts", scale=alt.Scale(domain=[0.0, max_dom]), axis=alt.Axis(grid=True, gridColor="#f0f2f6", domain=False)),
        y=alt.Y("Participant:N", sort="-x", title=None, axis=alt.Axis(labelLimit=560, labelFontSize=12, tickSize=0, domain=False)),
        color=alt.Color("Punts:Q", scale=alt.Scale(scheme=color_scheme), legend=None),
        tooltip=[alt.Tooltip("Posició:Q"), alt.Tooltip("Participant:N"), alt.Tooltip("Punts:Q", format=".1f"), alt.Tooltip("Dif líder:Q", format=".1f")]
    )
    text = bars.mark_text(align='left', baseline='middle', dx=8, fontSize=12, fontWeight='bold', color='#334e68').encode(text=alt.Text('Punts:Q', format='.1f'))
    st.altair_chart((bars + text).properties(height=chart_height).configure_view(strokeWidth=0), use_container_width=True, theme="streamlit")

def mostrar_grafic_departaments(df_dep, color_scheme="purples"):
    if df_dep.empty: return
    chart_data = df_dep.copy()
    chart_data["Mitjana_punts"] = pd.to_numeric(chart_data["Mitjana_punts"], errors="coerce").fillna(0.0).sort_values(ascending=False)
    chart_height = max(350, len(chart_data) * 46)
    max_dom = 10.0 if pd.isna(chart_data["Mitjana_punts"].max()) or chart_data["Mitjana_punts"].max() <= 0 else float(chart_data["Mitjana_punts"].max() * 1.15)
    bars = alt.Chart(chart_data).mark_bar(cornerRadiusEnd=6, height=26).encode(
        x=alt.X("Mitjana_punts:Q", title="Mitjana de punts", scale=alt.Scale(domain=[0.0, max_dom]), axis=alt.Axis(grid=True, gridColor="#f0f2f6", domain=False)),
        y=alt.Y("Departament:N", sort="-x", title=None, axis=alt.Axis(labelLimit=560, labelFontSize=13, tickSize=0, domain=False)),
        color=alt.Color("Mitjana_punts:Q", scale=alt.Scale(scheme=color_scheme), legend=None),
        tooltip=[alt.Tooltip("Posició:Q"), alt.Tooltip("Departament:N"), alt.Tooltip("Participants:Q"), alt.Tooltip("Mitjana_punts:Q", format=".1f"), alt.Tooltip("Punts_totals:Q", format=".1f")]
    )
    text = bars.mark_text(align='left', baseline='middle', dx=8, fontSize=13, fontWeight='bold', color='#334e68').encode(text=alt.Text('Mitjana_punts:Q', format='.1f'))
    st.altair_chart((bars + text).properties(height=chart_height).configure_view(strokeWidth=0), use_container_width=True, theme="streamlit")

def obtenir_pichichi_real(df_resultats_display, col_pichichi, col_gols):
    if col_pichichi not in df_resultats_display.columns or col_gols not in df_resultats_display.columns: return "Pendent", "Pendent"
    taula = df_resultats_display[[col_pichichi, col_gols]].copy()
    taula[col_pichichi] = taula[col_pichichi].astype(str).str.strip()
    taula[col_gols] = pd.to_numeric(taula[col_gols], errors="coerce")
    taula = taula[(taula[col_pichichi] != "") & (~taula[col_pichichi].str.lower().isin(["nan", "nat", "pendent"]))]
    if taula.empty or pd.isna(taula[col_gols].max()) or taula[col_gols].max() <= 0: return "Pendent", "Pendent"
    jugadors_top = taula[taula[col_gols] == taula[col_gols].max()][col_pichichi].tolist()
    return " · ".join(jugadors_top), str(int(taula[col_gols].max()))

def obtenir_prediccions_fase(df_j, prefix, quantitat):
    files = []
    for i in range(1, quantitat + 1):
        col = f"{prefix}_{i}"
        valor = valor_o_pendent(df_j[col].values[0]) if col in df_j.columns else "Pendent"
        files.append({"Posició": i, "Equip": afegir_bandera(valor)})
    return pd.DataFrame(files)

def mostrar_prediccions_grups_participant(df_j, df_resultats):
    st.write("### 🧩 Prediccions fase de grups")
    
    # LLEGENDA
    st.markdown("""
    <div style="background-color: rgba(255, 255, 255, 0.7); padding: 12px 18px; border-radius: 12px; margin-bottom: 16px; border: 1px solid rgba(0,0,0,0.1); font-size: 14px;">
        <p style="margin-bottom: 8px;"><strong>Llegenda de resultats:</strong></p>
        <ul style="margin: 0; padding-left: 20px;">
            <li>🟢 <strong>Verd (2 pts):</strong> Has encertat la posició exacta i l'equip passa.</li>
            <li>🟠 <strong>Taronja (1 pt):</strong> L'equip passa, però en una posició diferent a la que havies apostat.</li>
            <li>🟡 <strong>Groc (0,5 pts):</strong> Vas dir que quedava 3r a la posició exacta, però no aconsegueix classificar-se per setzens.</li>
            <li>🔴 <strong>Vermell (0 pts):</strong> L'equip cau eliminat o no es classifica en la posició que vas marcar (i el grup ja s'ha tancat).</li>
            <li>⚪ <strong>Pendent:</strong> Encara no s'ha jugat o no hi ha resultat oficial.</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)
    
    # Extraiem resultats
    grups_reals = {}
    if all(c in df_resultats.columns for c in ["Grup", "Posició", "Equip"]):
        for _, row in df_resultats.iterrows():
            grup_str = str(row.get("Grup", "")).strip().upper()
            lletra_grup = grup_str.split()[-1] if grup_str else ""
            posicio = str(row.get("Posició", "")).strip()
            equip_real = normalitzar_text(str(row.get("Equip", "")))
            if lletra_grup and equip_real and equip_real not in ["nan", "nat", "pendent", ""]:
                if lletra_grup not in grups_reals: grups_reals[lletra_grup] = {}
                if posicio in ["1r", "1", "1º"]: grups_reals[lletra_grup]["1r"] = equip_real
                elif posicio in ["2n", "2", "2º"]: grups_reals[lletra_grup]["2n"] = equip_real
                elif posicio in ["3r", "3", "3º"]: grups_reals[lletra_grup]["3r"] = equip_real

    classificats_oficials = []
    try:
        if df_resultats.shape[1] > 4:
            classificats_oficials = df_resultats.iloc[:, 4].dropna().astype(str).str.strip().apply(normalitzar_text).tolist()
    except: pass

    grups_dict = {}
    for grup in "ABCDEFGHIJKL":
        grup_trobat = False
        grup_data = {"1r": "", "2n": "", "3r": ""}
        real_grup_actual = grups_reals.get(grup, {})
        for col in df_j.columns:
            col_n = normalitzar_text(col)
            if "punt" in col_n: continue
            es_grup_actual = any(x in col_n for x in [f"grup {grup.lower()}", f"grup_{grup.lower()}", f"grup-{grup.lower()}"]) or col_n.startswith(f"{grup.lower()} ") or col_n.startswith(f"{grup.lower()}1") or col_n.startswith(f"{grup.lower()}2") or col_n.startswith(f"{grup.lower()}3")
            if es_grup_actual:
                pos_clau = "1r" if "1" in col_n else "2n" if "2" in col_n else "3r" if "3" in col_n else None
                if pos_clau:
                    val_orig = valor_o_pendent(df_j[col].values[0])
                    equip_norm = normalitzar_text(val_orig)
                    txt_base = afegir_bandera(val_orig)
                    if equip_norm in ["pendent", "nan", "nat", ""]:
                        grup_data[pos_clau] = txt_base
                    else:
                        equip_real_pos = real_grup_actual.get(pos_clau, "")
                        es_classificat = equip_norm in classificats_oficials
                        if equip_real_pos == "": estat = "⚪ "
                        else:
                            if equip_norm == equip_real_pos:
                                estat = "🟡 " if pos_clau == "3r" and not es_classificat else "🟢 "
                            elif es_classificat: estat = "🟠 "
                            else: estat = "🔴 "
                        grup_data[pos_clau] = f"{estat}{txt_base}"
                    grup_trobat = True
        if grup_trobat: grups_dict[f"Grup {grup}"] = grup_data
            
    if grups_dict:
        df_g = pd.DataFrame(grups_dict).reindex(["1r", "2n", "3r"]).reset_index().rename(columns={"index": "Posició"})
        st.dataframe(df_g, use_container_width=True, hide_index=True)
    else:
        st.info("No s'han detectat dades de la fase de grups.")

def mostrar_prediccions_eliminatoria_participant(df_j):
    st.write("### 🧭 Prediccions fase eliminatòria")
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["Vuitens", "Quarts", "Semis", "Final", "Campió"])
    with tab1: st.dataframe(obtenir_prediccions_fase(df_j, "Vuitens", 16), use_container_width=True, hide_index=True)
    with tab2: st.dataframe(obtenir_prediccions_fase(df_j, "Quarts", 8), use_container_width=True, hide_index=True)
    with tab3: st.dataframe(obtenir_prediccions_fase(df_j, "Semis", 4), use_container_width=True, hide_index=True)
    with tab4:
        finalistes = [afegir_bandera(valor_o_pendent(df_j[c].values[0])) if c in df_j.columns else "Pendent" for c in ["Final_1", "Final_2"]]
        st.dataframe(pd.DataFrame({"Finalista": ["Finalista 1", "Finalista 2"], "Equip": finalistes}), use_container_width=True, hide_index=True)
    with tab5:
        campio = valor_o_pendent(df_j["Campió"].values[0]) if "Campió" in df_j.columns else "Pendent"
        st.dataframe(pd.DataFrame({"Concepte": ["Campió previst"], "Equip": [afegir_bandera(campio)]}), use_container_width=True, hide_index=True)


# --------------------------------------------------
# ESTILS + FONS
# --------------------------------------------------
img_base64 = carregar_imatge_base64(BACKGROUND_IMAGE)
background_css = f'background-image: linear-gradient(rgba(0,0,0,0.15), rgba(0,0,0,0.35)), url("data:image/jpg;base64,{img_base64}"); background-size: cover; background-position: center; background-attachment: fixed;' if img_base64 else "background: #eef2f7;"

st.markdown(
    f"""
    <style>
    .stApp {{ {background_css} }}
    .block-container {{ padding-top: 2rem; padding-bottom: 2rem; background: rgba(255, 255, 255, 0.65); backdrop-filter: blur(12px); border: 1px solid rgba(255, 255, 255, 0.3); border-radius: 24px; margin-top: 24px; margin-bottom: 24px; box-shadow: 0px 8px 30px rgba(0,0,0,0.25); }}
    .title {{ font-size: clamp(32px, 5vw, 52px); font-weight: 900; margin-bottom: 0px; color: #102a43; letter-spacing: -1px; }}
    .subtitle {{ font-size: clamp(14px, 2vw, 18px); color: #334e68; margin-top: 0px; margin-bottom: 25px; }}
    .card-grid-2 {{ display: grid; grid-template-columns: repeat(2, 1fr); gap: 1.5rem; align-items: stretch; margin-bottom: 1rem; }}
    .card-grid-3 {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 1.5rem; align-items: stretch; margin-bottom: 1rem; }}
    .card-grid-4 {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 1.5rem; align-items: stretch; margin-bottom: 1rem; }}
    .card {{ padding: 18px; border-radius: 18px; text-align: center; box-shadow: 0px 4px 20px rgba(0,0,0,0.18); height: 100% !important; min-height: 178px; display: flex; flex-direction: column; justify-content: center; align-items: center; box-sizing: border-box; overflow: hidden; width: 100%; transition: all 0.3s ease-in-out !important; }}
    .card:hover {{ transform: translateY(-8px) scale(1.02) !important; box-shadow: 0px 15px 30px rgba(0,0,0,0.4) !important; cursor: pointer !important; }}
    .gold {{ background: linear-gradient(135deg, #ffd700, #fff1a8); color: #111; }}
    .silver {{ background: linear-gradient(135deg, #c0c0c0, #f2f2f2); color: #111; }}
    .bronze {{ background: linear-gradient(135deg, #cd7f32, #f0b27a); color: white; }}
    .bluecard {{ background: linear-gradient(135deg, #0b70c9, #7cc5ff); color: white; }}
    .greencard {{ background: linear-gradient(135deg, #0f9d58, #8ee6b3); color: white; }}
    .redcard {{ background: linear-gradient(135deg, #dc3545, #f1aeb5); color: white; }}
    .darkcard {{ background: linear-gradient(135deg, #102a43, #486581); color: white; }}
    .purplecard {{ background: linear-gradient(135deg, #6f42c1, #b982ff); color: white; margin-top: 18px; margin-bottom: 18px; }}
    .card h3 {{ margin: 0px 0px 14px 0px; font-size: clamp(15px, 2vw, 24px); line-height: 1.15; max-width: 100%; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
    .card h1 {{ margin: 0px; font-size: clamp(24px, 4vw, 40px); line-height: 1.1; max-width: 100%; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
    .card p {{ margin: 12px 0px 0px 0px; font-size: clamp(11px, 1.5vw, 15px); max-width: 100%; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
    @media (max-width: 768px) {{ .block-container {{ padding-left: 0.8rem; padding-right: 0.8rem; border-radius: 16px; }} .card-grid-2, .card-grid-3, .card-grid-4 {{ grid-template-columns: 1fr; }} .card {{ min-height: 140px; }} .card h3, .card h1, .card p {{ white-space: normal; }} }}
    </style>
    """, unsafe_allow_html=True
)

# --------------------------------------------------
# CARREGAR DADES
# --------------------------------------------------
excel_mtime = os.path.getmtime(EXCEL_FILE) if os.path.exists(EXCEL_FILE) else 0
data_actualitzacio = obtenir_data_actualitzacio_fitxer(EXCEL_FILE)
df_porra, df_resultats = carregar_dades(EXCEL_FILE, excel_mtime)
df_ranking = aplicar_moviment(crear_ranking_des_de_porra(df_porra), excel_mtime)
df_departaments = crear_ranking_departaments(df_ranking)
num_participants = len(df_ranking)
premi_guanyador = num_participants * PREU_PARTICIPACIO
te_departaments = "Departament" in df_ranking.columns and not df_departaments.empty

# --------------------------------------------------
# TÍTOL I LOGO EMPRESA
# --------------------------------------------------
col_titol, col_logo = st.columns([5, 1])
with col_titol:
    st.markdown('<p class="title">🏆 PORRA MUNDIAL</p>', unsafe_allow_html=True)
    st.markdown('<p class="subtitle">Classificació en viu, moviment respecte l’última actualització, competició per departaments i resultats reals</p>', unsafe_allow_html=True)
with col_logo:
    if os.path.exists(LOGO_IMAGE): st.image(LOGO_IMAGE, use_container_width=True)

# --------------------------------------------------
# INFO PRINCIPAL
# --------------------------------------------------
st.markdown(f"<div class='card-grid-3'><div class='card darkcard'><h3>🕒 Dades actualitzades</h3><h1>{data_actualitzacio}</h1></div><div class='card greencard'><h3>🎁 Premi guanyador</h3><h1>{premi_guanyador} €</h1><p>{num_participants} participants x {PREU_PARTICIPACIO} €</p></div><div class='card bluecard'><h3>👥 Participants</h3><h1>{num_participants}</h1><p>porres registrades</p></div></div>", unsafe_allow_html=True)


# --------------------------------------------------
# MOVIMENTS DESTACATS (PUJADES I BAIXADES)
# --------------------------------------------------
if "Canvi posició" in df_ranking.columns:
    max_p = df_ranking["Canvi posició"].max()
    min_p = df_ranking["Canvi posició"].min()
    
    # Només mostrem la secció si hi ha hagut algun moviment real
    if pd.notna(max_p) and pd.notna(min_p) and (max_p > 0 or min_p < 0):
        st.write("### 🎢 La muntanya russa de posicions")
        html_mov = "<div class='card-grid-2'>"
        
        if max_p > 0:
            pujadors = df_ranking[df_ranking["Canvi posició"] == max_p]["Participant"].tolist()
            noms_p = " · ".join(pujadors[:2]) + ("..." if len(pujadors) > 2 else "")
            html_mov += f"<div class='card greencard'><h3>🚀 La gran remuntada</h3><h1>{noms_p}</h1><p>+{int(max_p)} posicions d'una tacada! 🔥</p></div>"
        else:
            html_mov += "<div class='card greencard'><h3>🚀 La gran remuntada</h3><h1>-</h1><p>Ningú ha guanyat posicions encara 🤷‍♂️</p></div>"
            
        if min_p < 0:
            baixadors = df_ranking[df_ranking["Canvi posició"] == min_p]["Participant"].tolist()
            noms_b = " · ".join(baixadors[:2]) + ("..." if len(baixadors) > 2 else "")
            html_mov += f"<div class='card redcard'><h3>📉 Caiguda lliure</h3><h1>{noms_b}</h1><p>{int(min_p)} posicions avall... 🥶🚑</p></div>"
        else:
            html_mov += "<div class='card redcard'><h3>📉 Caiguda lliure</h3><h1>-</h1><p>Tothom manté el tipus 🧘‍♂️</p></div>"
            
        html_mov += "</div>"
        st.markdown(html_mov, unsafe_allow_html=True)


# --------------------------------------------------
# DEPARTAMENT LÍDER
# --------------------------------------------------
if te_departaments:
    st.markdown(f"<div class='card purplecard'><h3>🏢 Departament líder</h3><h1>{df_departaments.iloc[0]['Departament']}</h1><p>Mitjana {float(df_departaments.iloc[0]['Mitjana_punts']):.1f} punts · {int(df_departaments.iloc[0]['Participants'])} participants</p></div>", unsafe_allow_html=True)

# --------------------------------------------------
# TOP 3 GENERAL
# --------------------------------------------------
st.subheader("🥇 TOP 3 General")
top3 = df_ranking.head(3)
html_top3 = "<div class='card-grid-3'>"
for i, (medalla, classe) in enumerate([("🥇", "gold"), ("🥈", "silver"), ("🥉", "bronze")]):
    if len(top3) > i:
        row = top3.iloc[i]
        html_top3 += f"<div class='card {classe}'><h3>{medalla} {row['Participant']}</h3><h1>{float(row['Punts']):.1f}</h1><p>{row.get('Departament', 'punts')} · {row.get('Evolució', '')}</p></div>"
st.markdown(html_top3 + "</div>", unsafe_allow_html=True)

# --------------------------------------------------
# CLASSIFICACIÓ GENERAL
# --------------------------------------------------
st.subheader("📊 Classificació general")
mostrar_taula_ranking(df_ranking)

st.subheader("📈 Gràfic general de punts")
mostrar_grafic_punts(df_ranking, color_scheme="blues", altura_minima=1000)

# --------------------------------------------------
# FITXA PARTICIPANT
# --------------------------------------------------
st.subheader("👤 Fitxa participant")
jugador = st.selectbox("Selecciona participant", df_porra["Participants"].dropna().astype(str).unique(), index=None, placeholder="Selecciona un participant...")

if jugador is not None:
    df_j = df_porra[df_porra["Participants"].astype(str) == str(jugador)]
    if not df_j.empty:
        c1, c2 = st.columns(2)
        c1.metric("Total punts", f"{pd.to_numeric(df_j['Total Punts'].values[0], errors='coerce'):.1f}")
        col_dep_original = obtenir_columna_departament(df_porra)
        if col_dep_original is not None: c1.metric("Departament", valor_o_pendent(df_j[col_dep_original].values[0]))

        punts_dict = {
            "1rs grup": pd.to_numeric(df_j["Punts Grups 1r"].values[0], errors="coerce"),
            "2ns grup": pd.to_numeric(df_j["Punts Grups 2n"].values[0], errors="coerce"),
            "3rs grup": pd.to_numeric(df_j["Punts Grups 3r"].values[0], errors="coerce"),
            "Vuitens": pd.to_numeric(df_j["Punts Vuitens"].values[0], errors="coerce"),
            "Quarts": pd.to_numeric(df_j["Punts Quarts"].values[0], errors="coerce"),
            "Semis": pd.to_numeric(df_j["Punts Semis"].values[0], errors="coerce"),
            "Finalistes": pd.to_numeric(df_j["Punts Finalistes"].values[0], errors="coerce"),
            "Campió": pd.to_numeric(df_j["Punts Campió"].values[0], errors="coerce"),
            "Resultat final": pd.to_numeric(df_j["Punts Resultat Final"].values[0], errors="coerce") if "Punts Resultat Final" in df_j.columns else 0,
            "MVP": pd.to_numeric(df_j["Punts MVP"].values[0], errors="coerce"),
            "Bota d'Or": pd.to_numeric(df_j["Punts Pichichi"].values[0], errors="coerce") if "Punts Pichichi" in df_j.columns else 0,
        }
        
        punts_categoria = pd.DataFrame({"Categoria": list(punts_dict.keys()), "Punts": list(punts_dict.values())})
        punts_categoria["Punts"] = pd.to_numeric(punts_categoria["Punts"], errors="coerce").fillna(0.0).round(1)
        
        max_p_cat = punts_categoria["Punts"].max()
        bars_cat = alt.Chart(punts_categoria).mark_bar(cornerRadiusEnd=6, size=25).encode(
            x=alt.X("Categoria:N", sort=None, title=None, axis=alt.Axis(labelAngle=-45, labelFontSize=12, tickSize=0, domain=False)),
            y=alt.Y("Punts:Q", title="Punts", scale=alt.Scale(domain=[0.0, 10.0 if pd.isna(max_p_cat) or max_p_cat <= 0 else float(max_p_cat * 1.2)]), axis=alt.Axis(grid=True, gridColor="#f0f2f6", domain=False)),
            color=alt.Color("Punts:Q", scale=alt.Scale(scheme="tealblues"), legend=None),
            tooltip=[alt.Tooltip("Categoria:N"), alt.Tooltip("Punts:Q", format=".1f")]
        )
        c1.altair_chart((bars_cat + bars_cat.mark_text(align='center', baseline='bottom', dy=-8, fontSize=12, fontWeight='bold', color='#334e68').encode(text=alt.Text('Punts:Q', format='.1f'))).properties(height=350).configure_view(strokeWidth=0), use_container_width=True, theme="streamlit")

        col_resultat_final_porra = trobar_col_resultat_final_porra(df_porra)
        c2.write("### ⚽ Prediccions principals")
        c2.write(f"🏆 Campió: {afegir_bandera(valor_o_pendent(df_j['Campió'].values[0]))}")
        c2.write(f"📌 Resultat final: {valor_o_pendent(df_j[col_resultat_final_porra].values[0]) if col_resultat_final_porra else 'Pendent'}")
        c2.write(f"⭐ MVP: {valor_o_pendent(df_j['MVP'].values[0])}")
        
        val_bota = valor_o_pendent(df_j['Pichichi'].values[0]) if 'Pichichi' in df_j.columns else "Pendent"
        gols_bota_str = " (0 gols reals)"
        if val_bota != "Pendent" and "Jugador Pichichi" in df_resultats.columns and "Gols" in df_resultats.columns:
            match = df_resultats[df_resultats["Jugador Pichichi"].astype(str).str.strip().str.lower() == val_bota.strip().lower()]
            if not match.empty and pd.notna(match.iloc[0]["Gols"]) and str(match.iloc[0]["Gols"]).strip() != "":
                gols_bota_str = f" ({int(match.iloc[0]['Gols'])} gols reals)"
        c2.write(f"⚽ Bota d'Or: {val_bota}{gols_bota_str}")

        mostrar_prediccions_grups_participant(df_j, df_resultats)
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
    mostrar_grafic_departaments(df_departaments, color_scheme="purples")
    
    departament_sel = st.selectbox("Selecciona departament", sorted(df_ranking["Departament"].dropna().astype(str).unique()), index=None, placeholder="Selecciona un departament...")
    if departament_sel:
        df_dep_individual = recalcular_posicions(df_ranking[df_ranking["Departament"] == departament_sel].copy())
        st.write(f"### 🥇 TOP 3 · {departament_sel}")
        html_top_dep = "<div class='card-grid-3'>"
        for i, (medalla, classe) in enumerate([("🥇", "gold"), ("🥈", "silver"), ("🥉", "bronze")]):
            if len(df_dep_individual) > i:
                row = df_dep_individual.iloc[i]
                html_top_dep += f"<div class='card {classe}'><h3>{medalla} {row['Participant']}</h3><h1>{float(row['Punts']):.1f}</h1><p>{departament_sel} · {row.get('Evolució', '')}</p></div>"
        st.markdown(html_top_dep + "</div>", unsafe_allow_html=True)
        mostrar_taula_ranking(df_dep_individual)
        mostrar_grafic_punts(df_dep_individual, color_scheme="purples", altura_minima=350)
else:
    st.info("Per activar aquest apartat, afegeix una columna 'Departament' al costat de 'Participants' al full Porra.")

# --------------------------------------------------
# LLIGUETES
# --------------------------------------------------
st.subheader("🏟️ Lligueta personalitzada")
participants_filtrats = st.multiselect("Selecciona participants per crear una lligueta:", options=df_ranking["Participant"].dropna().astype(str).tolist(), default=[], placeholder="Tria participants...")
if participants_filtrats:
    df_lligueta = recalcular_posicions(df_ranking[df_ranking["Participant"].astype(str).isin(participants_filtrats)].copy())
    st.write(f"Participants seleccionats: **{len(participants_filtrats)}**")
    mostrar_taula_ranking(df_lligueta)
    mostrar_grafic_punts(df_lligueta, color_scheme="greens", altura_minima=350)
else:
    st.write("Selecciona participants per crear una classificació reduïda tipus lligueta.")

# --------------------------------------------------
# RESULTATS REALS
# --------------------------------------------------
st.subheader("✅ Resultats reals")
df_resultats_display = preparar_taula_buida(df_resultats)

campio_real = afegir_bandera(primer_valor_o_pendent(df_resultats_display, "Campió"))
mvp_real = primer_valor_o_pendent(df_resultats_display, "MVP")
resultat_final_real = primer_valor_o_pendent(df_resultats_display, "Resultat Final")
pichichi_real, gols_pichichi = obtenir_pichichi_real(df_resultats_display, "Jugador Pichichi", "Gols")
st.markdown(f"<div class='card-grid-4'><div class='card gold'><h3>🏆 Campió</h3><h1 style='font-size:28px'>{campio_real}</h1></div><div class='card silver'><h3>⭐ MVP</h3><h1 style='font-size:28px'>{mvp_real}</h1></div><div class='card bronze'><h3>⚽ Bota d'Or</h3><h1 style='font-size: clamp(16px, 2.5vw, 24px); white-space: normal; line-height: 1.2;'>{pichichi_real}</h1><p>{f'{gols_pichichi} gols' if gols_pichichi != 'Pendent' else 'Pendent'}</p></div><div class='card bluecard'><h3>🏁 Resultat final</h3><h1 style='font-size:28px'>{resultat_final_real}</h1></div></div>", unsafe_allow_html=True)

st.write("### 🧩 Fase de grups")
grups = {}
if all(col in df_resultats_display.columns for col in ["Grup", "Posició", "Equip"]):
    for _, row in df_resultats_display.iterrows():
        grup = str(row.get("Grup", "")).strip()
        posicio = str(row.get("Posició", "")).strip()
        equip = str(row.get("Equip", "")).strip()
        if grup and equip:
            if grup not in grups: grups[grup] = {"1r": "", "2n": "", "3r": ""}
            if posicio in ["1r", "1", "1º"]: grups[grup]["1r"] = afegir_bandera(equip)
            elif posicio in ["2n", "2", "2º"]: grups[grup]["2n"] = afegir_bandera(equip)
            elif posicio in ["3r", "3", "3º"]: grups[grup]["3r"] = afegir_bandera(equip)
    if grups: st.dataframe(pd.DataFrame(grups).reindex(["1r", "2n", "3r"]).reset_index().rename(columns={"index": "Posició"}), use_container_width=True, hide_index=True)
    else: st.info("No hi ha dades de fase de grups configurades.")
else: st.info("No hi ha dades de fase de grups configurades.")

st.write("### ⚔️ Fase eliminatòria")
files_eliminatoria = []
for fase in ["Vuitens", "Quarts", "Semis", "Finalistes", "Campió", "MVP"]:
    if fase in df_resultats_display.columns:
        valors = llista_valors_no_buits(df_resultats_display, fase)
        detall = "Pendent" if not valors or (len(valors) == 1 and normalitzar_text(valors[0]) == "pendent") else " · ".join(valors) if fase == "MVP" else " · ".join([afegir_bandera(v) for v in valors])
        files_eliminatoria.append({"Fase": fase, "Resultat": detall})
if files_eliminatoria: st.dataframe(pd.DataFrame(files_eliminatoria), use_container_width=True, hide_index=True)
else: st.info("No hi ha dades de fase eliminatòria configurades.")

st.write("### ⚽ Bota d'Or")
if "Jugador Pichichi" in df_resultats_display.columns and "Gols" in df_resultats_display.columns:
    taula_pichichi = df_resultats_display[["Jugador Pichichi", "Gols"]].copy()
    taula_pichichi["Jugador Pichichi"] = taula_pichichi["Jugador Pichichi"].astype(str).str.strip()
    taula_pichichi["Gols"] = pd.to_numeric(taula_pichichi["Gols"], errors="coerce")
    taula_pichichi = taula_pichichi[(taula_pichichi["Jugador Pichichi"] != "") & (~taula_pichichi["Jugador Pichichi"].str.lower().isin(["nan", "nat", "pendent"]))]
    if taula_pichichi.empty:
        st.dataframe(pd.DataFrame({"Jugador": ["Pendent"], "Gols": ["Pendent"]}), use_container_width=True, hide_index=True)
    else:
        taula_pichichi = taula_pichichi.sort_values("Gols", ascending=False).reset_index(drop=True)
        taula_pichichi["Gols"] = taula_pichichi["Gols"].fillna(0).astype("Int64")
        st.dataframe(taula_pichichi.rename(columns={"Jugador Pichichi": "Jugador"}), use_container_width=True, hide_index=True, height=(len(taula_pichichi) * 35) + 40)
else:
    st.dataframe(pd.DataFrame({"Jugador": ["Pendent"], "Gols": ["Pendent"]}), use_container_width=True, hide_index=True)

st.write("### 🏁 Resultat de la final")
st.dataframe(pd.DataFrame({"Concepte": ["Resultat de la final"], "Valor": [primer_valor_o_pendent(df_resultats_display, "Resultat Final")]}), use_container_width=True, hide_index=True)

st.markdown("---")
st.write("📡 Actualització automàtica des de Excel")