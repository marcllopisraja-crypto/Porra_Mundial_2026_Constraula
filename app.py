
import streamlit as st
import pandas as pd
import altair as alt
import base64
import os
import json
import unicodedata
import hashlib
import requests
from datetime import datetime
from zoneinfo import ZoneInfo

# ==================================================
# CONFIGURACIÓ GENERAL
# ==================================================
st.set_page_config(page_title="Porra Mundial 2026", layout="wide")

EXCEL_FILE = "Porra_Mundial_Final_Definitiva.xlsx"
BACKGROUND_IMAGE = "fifa-Trionda.jpg"
PREU_PARTICIPACIO = 5

SNAPSHOT_CURRENT_FILE = "ranking_snapshot_current.csv"
SNAPSHOT_DISPLAY_FILE = "ranking_snapshot_display.csv"
SNAPSHOT_META_FILE = "ranking_snapshot_meta.json"
HISTORY_FILE = "ranking_history.csv"

THESTATSAPI_BASE_URL = "https://api.thestatsapi.com/api/football/matches"
DEFAULT_COMPETITION_ID = "comp_6107"
DEFAULT_SEASON_ID = "sn_118868"
API_TIMEOUT_SECONDS = 8

# ==================================================
# NORMALITZACIÓ
# ==================================================
def normalitzar_text(text):
    text = str(text).strip().lower().replace("’", "'").replace("`", "'")
    text = unicodedata.normalize("NFD", text)
    text = "".join(char for char in text if unicodedata.category(char) != "Mn")
    return " ".join(text.split())

TEAM_NAME_MAP_RAW = {
    "mexico": "Mèxic", "mexic": "Mèxic", "méxic": "Mèxic",
    "south africa": "Sud-àfrica", "sud africa": "Sud-àfrica", "sud-africa": "Sud-àfrica", "sud-àfrica": "Sud-àfrica",
    "korea republic": "Corea del Sud", "south korea": "Corea del Sud", "corea del sud": "Corea del Sud",
    "czechia": "República Txeca", "republica txeca": "República Txeca", "república txeca": "República Txeca", "republica checa": "República Txeca",
    "canada": "Canadà", "canadà": "Canadà",
    "bosnia and herzegovina": "Bosnia i Hercegovina", "bosnia & herzegovina": "Bosnia i Hercegovina", "bosnia i hercegovina": "Bosnia i Hercegovina",
    "qatar": "Qatar", "switzerland": "Suïssa", "suissa": "Suïssa", "suïssa": "Suïssa",
    "brazil": "Brasil", "brasil": "Brasil", "morocco": "Marroc", "marroc": "Marroc",
    "haiti": "Haití", "haití": "Haití", "scotland": "Escòcia", "escocia": "Escòcia", "escòcia": "Escòcia",
    "united states": "Estats Units", "usa": "Estats Units", "eeuu": "Estats Units", "ee.uu": "Estats Units", "estats units": "Estats Units",
    "paraguay": "Paraguai", "paraguai": "Paraguai", "australia": "Austràlia", "austràlia": "Austràlia",
    "turkiye": "Turquia", "turkey": "Turquia", "turquia": "Turquia",
    "germany": "Alemanya", "alemanya": "Alemanya", "curacao": "Curaçao", "curaçao": "Curaçao",
    "cote d'ivoire": "Costa d'Ivori", "côte d'ivoire": "Costa d'Ivori", "costa d'ivori": "Costa d'Ivori",
    "ecuador": "Equador", "equador": "Equador", "netherlands": "Països Baixos", "paisos baixos": "Països Baixos", "països baixos": "Països Baixos", "pasïsos baixos": "Països Baixos",
    "japan": "Japó", "japo": "Japó", "japó": "Japó", "sweden": "Suècia", "suecia": "Suècia", "suècia": "Suècia", "tunisia": "Tunísia", "tunísia": "Tunísia",
    "belgium": "Bèlgica", "belgica": "Bèlgica", "bèlgica": "Bèlgica", "egypt": "Egipte", "egipte": "Egipte",
    "ir iran": "Iran", "iran": "Iran", "new zealand": "Nova Zelanda", "nova zelanda": "Nova Zelanda",
    "spain": "Espanya", "espanya": "Espanya", "cabo verde": "Cap Verd", "cap verd": "Cap Verd",
    "saudi arabia": "Aràbia Saudita", "arabia saudita": "Aràbia Saudita", "aràbia saudita": "Aràbia Saudita",
    "uruguay": "Uruguai", "uruguai": "Uruguai", "france": "França", "franca": "França", "frança": "França",
    "senegal": "Senegal", "iraq": "Iraq", "norway": "Noruega", "noruega": "Noruega",
    "argentina": "Argentina", "algeria": "Algèria", "algèria": "Algèria", "austria": "Àustria", "àustria": "Àustria",
    "jordan": "Jordània", "jordania": "Jordània", "jordània": "Jordània", "portugal": "Portugal",
    "congo dr": "RD Congo", "dr congo": "RD Congo", "rd congo": "RD Congo", "uzbekistan": "Uzbekistan",
    "colombia": "Colòmbia", "colòmbia": "Colòmbia", "england": "Anglaterra", "anglaterra": "Anglaterra",
    "croatia": "Croàcia", "croacia": "Croàcia", "croàcia": "Croàcia", "ghana": "Ghana", "panama": "Panamà", "panamà": "Panamà",
}
TEAM_NAME_MAP = {normalitzar_text(k): v for k, v in TEAM_NAME_MAP_RAW.items()}

FLAGS = {
    "mexic": "🇲🇽", "corea del sud": "🇰🇷", "republica txeca": "🇨🇿", "suissa": "🇨🇭", "canada": "🇨🇦", "qatar": "🇶🇦",
    "escocia": "🏴", "marroc": "🇲🇦", "brasil": "🇧🇷", "estats units": "🇺🇸", "australia": "🇦🇺", "turquia": "🇹🇷",
    "alemanya": "🇩🇪", "costa d'ivori": "🇨🇮", "equador": "🇪🇨", "suecia": "🇸🇪", "japo": "🇯🇵", "paisos baixos": "🇳🇱",
    "nova zelanda": "🇳🇿", "iran": "🇮🇷", "belgica": "🇧🇪", "uruguai": "🇺🇾", "arabia saudita": "🇸🇦", "espanya": "🇪🇸",
    "franca": "🇫🇷", "senegal": "🇸🇳", "iraq": "🇮🇶", "argentina": "🇦🇷", "algeria": "🇩🇿", "austria": "🇦🇹",
    "portugal": "🇵🇹", "rd congo": "🇨🇩", "uzbekistan": "🇺🇿", "anglaterra": "🏴", "croacia": "🇭🇷", "ghana": "🇬🇭",
    "egipte": "🇪🇬", "noruega": "🇳🇴", "colombia": "🇨🇴", "bosnia i hercegovina": "🇧🇦", "paraguai": "🇵🇾", "tunisia": "🇹🇳",
    "cap verd": "🇨🇻", "jordania": "🇯🇴", "panama": "🇵🇦", "curacao": "🇨🇼", "haiti": "🇭🇹", "sud-africa": "🇿🇦",
}

def normalitzar_equip(valor):
    if pd.isna(valor):
        return ""
    text = str(valor).strip()
    if text == "" or normalitzar_text(text) in ["pendent", "nan", "nat", "none"]:
        return ""
    return TEAM_NAME_MAP.get(normalitzar_text(text), text)

def afegir_bandera(valor):
    if pd.isna(valor):
        return "Pendent"
    text = str(valor).strip()
    if text == "" or normalitzar_text(text) in ["pendent", "nan", "nat", "none"]:
        return "Pendent"
    equip = normalitzar_equip(text)
    norm = normalitzar_text(equip)
    for pais, bandera in FLAGS.items():
        if pais in norm:
            return f"{bandera} {equip}"
    return equip

def valor_o_pendent(valor):
    if pd.isna(valor):
        return "Pendent"
    text = str(valor).strip()
    if text == "" or normalitzar_text(text) in ["nan", "nat", "none", "pendent"]:
        return "Pendent"
    return text

def convertir_a_int(valor):
    if valor is None:
        return None
    try:
        if pd.isna(valor):
            return None
    except Exception:
        pass
    try:
        return int(valor)
    except Exception:
        try:
            return int(float(str(valor).strip()))
        except Exception:
            return None

# ==================================================
# SECRETS / API
# ==================================================
def llegir_secret(nom, defecte=None):
    try:
        return st.secrets[nom]
    except Exception:
        return defecte

def obtenir_config_api():
    return {
        "use_api": str(llegir_secret("USE_API_RESULTS", "false")).lower() == "true",
        "api_key": str(llegir_secret("THESTATSAPI_KEY", "")).strip(),
        "competition_id": str(llegir_secret("THESTATSAPI_COMPETITION_ID", DEFAULT_COMPETITION_ID)).strip(),
        "season_id": str(llegir_secret("THESTATSAPI_SEASON_ID", DEFAULT_SEASON_ID)).strip(),
    }

# ==================================================
# CÀRREGA EXCEL
# ==================================================
@st.cache_data(show_spinner=False)
def carregar_dades(excel_file, file_mtime):
    sheets = pd.read_excel(excel_file, sheet_name=["Porra", "Resultats Reals"], engine="openpyxl")
    df_porra = sheets["Porra"]
    df_resultats = sheets["Resultats Reals"]
    df_porra.columns = df_porra.columns.astype(str).str.strip()
    df_resultats.columns = df_resultats.columns.astype(str).str.strip()
    return df_porra, df_resultats

@st.cache_data(show_spinner=False)
def carregar_imatge_base64(image_path):
    if not os.path.exists(image_path):
        return None
    with open(image_path, "rb") as f:
        return base64.b64encode(f.read()).decode()

def obtenir_columna_departament(df):
    for col in df.columns:
        if normalitzar_text(col) == "departament":
            return col
    for col in df.columns:
        if "depart" in normalitzar_text(col):
            return col
    return None

def trobar_columna_flexible(df, nom):
    nom_norm = normalitzar_text(nom)
    for col in df.columns:
        if normalitzar_text(col) == nom_norm:
            return col
    return None

def obtenir_data_actualitzacio_fitxer(path):
    if not os.path.exists(path):
        return "No disponible"
    timestamp = os.path.getmtime(path)
    dt = datetime.fromtimestamp(timestamp, tz=ZoneInfo("Europe/Madrid"))
    return dt.strftime("%d/%m/%Y")

def preparar_taula_buida(df):
    return df.copy().dropna(how="all").dropna(axis=1, how="all").fillna("")

def llista_valors_no_buits(df, columna):
    col = trobar_columna_flexible(df, columna)
    if col is None:
        return []
    valors = df[col].astype(str).str.strip().tolist()
    nets = []
    vistos = set()
    for valor in valors:
        if valor == "" or normalitzar_text(valor) in ["nan", "nat", "none"]:
            continue
        if normalitzar_text(valor) == "pendent":
            continue
        equip = normalitzar_equip(valor)
        if not equip:
            continue
        key = normalitzar_text(equip)
        if key not in vistos:
            nets.append(equip)
            vistos.add(key)
    return nets

def primer_valor_o_pendent(df, columna):
    valors = llista_valors_no_buits(df, columna)
    return valors[0] if valors else "Pendent"

def trobar_col_resultat_final_porra(df_porra):
    for col in df_porra.columns:
        col_norm = normalitzar_text(col)
        if col_norm == "resultat final" or ("resultat" in col_norm and "final" in col_norm and "punt" not in col_norm):
            return col
    return None

# ==================================================
# API THESTATSAPI
# ==================================================
def obtenir_camp_dict(obj, claus, defecte=None):
    if not isinstance(obj, dict):
        return defecte
    for clau in claus:
        if clau in obj:
            return obj.get(clau)
    return defecte

def extreure_nom_equip(team_obj):
    if isinstance(team_obj, dict):
        return normalitzar_equip(obtenir_camp_dict(team_obj, ["name", "team_name", "display_name", "short_name"]))
    return normalitzar_equip(team_obj)

def extreure_score(match):
    score = match.get("score", {})
    if isinstance(score, dict):
        home = obtenir_camp_dict(score, ["home", "home_score", "home_goals"])
        away = obtenir_camp_dict(score, ["away", "away_score", "away_goals"])
        if home is None and isinstance(score.get("full_time"), dict):
            home = obtenir_camp_dict(score["full_time"], ["home", "home_score"])
            away = obtenir_camp_dict(score["full_time"], ["away", "away_score"])
        return convertir_a_int(home), convertir_a_int(away)
    return convertir_a_int(obtenir_camp_dict(match, ["home_score", "score_home", "home_goals"])), convertir_a_int(obtenir_camp_dict(match, ["away_score", "score_away", "away_goals"]))

def es_partit_finalitzat(match):
    status = str(match.get("status", "")).strip().lower()
    phase = str(match.get("phase", "")).strip().lower()
    finals = ["ft", "full_time", "finished", "final", "after_extra_time", "aet", "ft_pen", "finished_penalties", "penalties"]
    if status in finals or phase in finals:
        return True
    home, away = extreure_score(match)
    return home is not None and away is not None and status not in ["scheduled", "pre", "postponed", "cancelled"]

def obtenir_guanyador_match(match):
    home = extreure_nom_equip(match.get("home_team"))
    away = extreure_nom_equip(match.get("away_team"))
    hs, aw = extreure_score(match)
    if hs is None or aw is None:
        return ""
    if hs > aw:
        return home
    if aw > hs:
        return away
    winner = match.get("winner") or match.get("winning_team") or match.get("winner_team")
    return extreure_nom_equip(winner) if winner else ""

def obtenir_grup_match(match):
    group_label = match.get("group_label") or match.get("group") or match.get("group_name")
    if group_label is None:
        return ""
    text = str(group_label).strip().replace("Group", "").replace("Grup", "").strip()
    return text.upper() if text else ""

def obtenir_stage_match(match):
    return normalitzar_text(match.get("stage_name") or match.get("stage") or match.get("round") or match.get("round_name") or match.get("phase_name") or "")

@st.cache_data(ttl=900, show_spinner=False)
def carregar_matches_thestatsapi(api_key, competition_id, season_id):
    if not api_key:
        return [], "No hi ha THESTATSAPI_KEY configurada."
    headers = {"Authorization": f"Bearer {api_key}"}
    matches = []
    errors = []
    for page in [1, 2]:
        params = {"competition_id": competition_id, "season_id": season_id, "per_page": 100, "page": page}
        try:
            response = requests.get(THESTATSAPI_BASE_URL, headers=headers, params=params, timeout=API_TIMEOUT_SECONDS)
            if response.status_code != 200:
                errors.append(f"HTTP {response.status_code} pàgina {page}")
                continue
            data = response.json()
            if isinstance(data, dict):
                page_matches = data.get("data", data.get("matches", data.get("response", [])))
            elif isinstance(data, list):
                page_matches = data
            else:
                page_matches = []
            if isinstance(page_matches, dict):
                page_matches = list(page_matches.values())
            if page_matches:
                matches.extend(page_matches)
        except Exception as e:
            errors.append(str(e))
    if not matches:
        return [], "No s'han pogut obtenir partits de TheStatsAPI. " + " | ".join(errors)
    return matches, ""

def construir_resultats_des_api(matches):
    group_matches, knockout_matches = [], []
    for match in matches:
        stage = obtenir_stage_match(match)
        grup = obtenir_grup_match(match)
        if grup and grup in list("ABCDEFGHIJKL"):
            group_matches.append(match)
        elif ("group" in stage or "regular" in stage) and grup:
            group_matches.append(match)
        else:
            knockout_matches.append(match)

    standings = {}
    for match in group_matches:
        if not es_partit_finalitzat(match):
            continue
        grup = obtenir_grup_match(match)
        home = extreure_nom_equip(match.get("home_team"))
        away = extreure_nom_equip(match.get("away_team"))
        hs, aw = extreure_score(match)
        if not grup or not home or not away or hs is None or aw is None:
            continue
        standings.setdefault(grup, {})
        for equip in [home, away]:
            standings[grup].setdefault(equip, {"Equip": equip, "PJ": 0, "PG": 0, "PE": 0, "PP": 0, "GF": 0, "GC": 0, "DG": 0, "Pts": 0})
        standings[grup][home]["PJ"] += 1
        standings[grup][away]["PJ"] += 1
        standings[grup][home]["GF"] += hs
        standings[grup][home]["GC"] += aw
        standings[grup][away]["GF"] += aw
        standings[grup][away]["GC"] += hs
        if hs > aw:
            standings[grup][home]["PG"] += 1
            standings[grup][away]["PP"] += 1
            standings[grup][home]["Pts"] += 3
        elif aw > hs:
            standings[grup][away]["PG"] += 1
            standings[grup][home]["PP"] += 1
            standings[grup][away]["Pts"] += 3
        else:
            standings[grup][home]["PE"] += 1
            standings[grup][away]["PE"] += 1
            standings[grup][home]["Pts"] += 1
            standings[grup][away]["Pts"] += 1
        for equip in [home, away]:
            standings[grup][equip]["DG"] = standings[grup][equip]["GF"] - standings[grup][equip]["GC"]

    group_positions, tercers = {}, []
    for grup, equips in standings.items():
        taula = pd.DataFrame(list(equips.values()))
        if taula.empty:
            continue
        taula = taula.sort_values(["Pts", "DG", "GF"], ascending=[False, False, False]).reset_index(drop=True)
        group_positions[grup] = {}
        if len(taula) >= 1:
            group_positions[grup]["1r"] = taula.iloc[0]["Equip"]
        if len(taula) >= 2:
            group_positions[grup]["2n"] = taula.iloc[1]["Equip"]
        if len(taula) >= 3:
            group_positions[grup]["3r"] = taula.iloc[2]["Equip"]
            tercer = taula.iloc[2].to_dict()
            tercer["Grup"] = grup
            tercers.append(tercer)

    setzens = []
    for grup in sorted(group_positions.keys()):
        if "1r" in group_positions[grup]:
            setzens.append(group_positions[grup]["1r"])
        if "2n" in group_positions[grup]:
            setzens.append(group_positions[grup]["2n"])
    if tercers:
        df_tercers = pd.DataFrame(tercers).sort_values(["Pts", "DG", "GF"], ascending=[False, False, False]).head(8)
        setzens.extend(df_tercers["Equip"].tolist())

    fases = {"Vuitens": [], "Quarts": [], "Semis": [], "Finalistes": [], "Campió": []}
    for match in knockout_matches:
        if not es_partit_finalitzat(match):
            continue
        winner = obtenir_guanyador_match(match)
        if not winner:
            continue
        stage = obtenir_stage_match(match)
        if "round of 32" in stage or "r32" in stage or "setzens" in stage or "32" in stage:
            fases["Vuitens"].append(winner)
        elif "round of 16" in stage or "r16" in stage or "vuitens" in stage or "16" in stage:
            fases["Quarts"].append(winner)
        elif "quarter" in stage or "quart" in stage:
            fases["Semis"].append(winner)
        elif "semi" in stage:
            fases["Finalistes"].append(winner)
        elif "final" in stage and "third" not in stage and "3rd" not in stage:
            fases["Campió"] = [winner]

    return {
        "source": "API", "group_positions": group_positions,
        "Setzens": list(dict.fromkeys(setzens)), "Vuitens": list(dict.fromkeys(fases["Vuitens"])),
        "Quarts": list(dict.fromkeys(fases["Quarts"])), "Semis": list(dict.fromkeys(fases["Semis"])),
        "Finalistes": list(dict.fromkeys(fases["Finalistes"])), "Campió": list(dict.fromkeys(fases["Campió"])),
        "MVP": [], "Pichichi": [], "Gols Pichichi": None, "api_error": ""
    }

def construir_resultats_des_excel(df_resultats):
    df = preparar_taula_buida(df_resultats)
    resultats = {"source": "Excel", "group_positions": {}, "Setzens": [], "Vuitens": [], "Quarts": [], "Semis": [], "Finalistes": [], "Campió": [], "MVP": [], "Pichichi": [], "Gols Pichichi": None, "api_error": ""}

    col_grup = trobar_columna_flexible(df, "Grup")
    col_pos = trobar_columna_flexible(df, "Posició")
    col_equip = trobar_columna_flexible(df, "Equip")
    if col_grup and col_pos and col_equip:
        for _, row in df.iterrows():
            grup = str(row.get(col_grup, "")).strip().replace("Grup ", "").strip()
            pos = str(row.get(col_pos, "")).strip()
            equip = normalitzar_equip(row.get(col_equip, ""))
            if not grup or not pos or not equip:
                continue
            resultats["group_positions"].setdefault(grup, {})
            if pos in ["1r", "1", "1º"]:
                resultats["group_positions"][grup]["1r"] = equip
            elif pos in ["2n", "2", "2º"]:
                resultats["group_positions"][grup]["2n"] = equip
            elif pos in ["3r", "3", "3º"]:
                resultats["group_positions"][grup]["3r"] = equip

    for fase in ["Setzens", "Vuitens", "Quarts", "Semis", "Finalistes", "Campió", "MVP"]:
        resultats[fase] = llista_valors_no_buits(df, fase)

    col_pichichi = trobar_columna_flexible(df, "Jugador Pichichi")
    col_gols = trobar_columna_flexible(df, "Gols")
    if col_pichichi and col_gols:
        taula = df[[col_pichichi, col_gols]].copy()
        taula[col_pichichi] = taula[col_pichichi].astype(str).str.strip()
        taula[col_gols] = pd.to_numeric(taula[col_gols], errors="coerce")
        taula = taula[(taula[col_pichichi] != "") & (~taula[col_pichichi].str.lower().isin(["nan", "nat", "pendent"])) & (taula[col_gols] >= 1)]
        if not taula.empty:
            taula = taula.sort_values(col_gols, ascending=False).reset_index(drop=True)
            resultats["Pichichi"] = [taula.iloc[0][col_pichichi]]
            resultats["Gols Pichichi"] = int(taula.iloc[0][col_gols])
    return resultats

def fusionar_resultats(excel_resultats, api_resultats):
    merged = excel_resultats.copy()
    merged["source"] = "API"
    merged["api_error"] = api_resultats.get("api_error", "")
    if api_resultats.get("group_positions"):
        merged["group_positions"] = api_resultats["group_positions"]
    for fase in ["Setzens", "Vuitens", "Quarts", "Semis", "Finalistes", "Campió", "MVP", "Pichichi"]:
        if api_resultats.get(fase):
            merged[fase] = api_resultats[fase]
    if api_resultats.get("Gols Pichichi") is not None:
        merged["Gols Pichichi"] = api_resultats.get("Gols Pichichi")
    return merged

def obtenir_resultats_actuals(df_resultats):
    excel_resultats = construir_resultats_des_excel(df_resultats)
    config = obtenir_config_api()
    if config["use_api"] and config["api_key"]:
        matches, error = carregar_matches_thestatsapi(config["api_key"], config["competition_id"], config["season_id"])
        if matches:
            return fusionar_resultats(excel_resultats, construir_resultats_des_api(matches))
        excel_resultats["api_error"] = error
    return excel_resultats

# ==================================================
# CÀLCUL PUNTS API
# ==================================================
def punts_grup(pred, real_exacte, reals_altres, equips_classificats, es_tercer=False):
    pred = normalitzar_equip(pred)
    if not pred:
        return 0.0
    pred_norm = normalitzar_text(pred)
    real_exacte_norm = normalitzar_text(real_exacte)
    classificats = {normalitzar_text(x) for x in equips_classificats if x}
    altres = {normalitzar_text(x) for x in reals_altres if x}
    if pred_norm == real_exacte_norm:
        return 2.0 if (not es_tercer or pred_norm in classificats) else 0.5
    if pred_norm in altres or pred_norm in classificats:
        return 1.0
    return 0.0

def calcular_punts_participant_api(row, resultats):
    p1 = p2 = p3 = 0.0
    group_positions = resultats.get("group_positions", {})
    setzens = resultats.get("Setzens", [])
    for grup in list("ABCDEFGHIJKL"):
        real = group_positions.get(grup, {})
        r1, r2, r3 = real.get("1r", ""), real.get("2n", ""), real.get("3r", "")
        p1 += punts_grup(row.get(f"Grup {grup} 1r", ""), r1, [r2, r3], setzens, False)
        p2 += punts_grup(row.get(f"Grup {grup} 2n", ""), r2, [r1, r3], setzens, False)
        p3 += punts_grup(row.get(f"Grup {grup} 3r", ""), r3, [r1, r2], setzens, True)

    def punts_fase(prefix, quantitat, real_list):
        real_norm = {normalitzar_text(x) for x in real_list if x}
        total = 0.0
        for i in range(1, quantitat + 1):
            pred = normalitzar_equip(row.get(f"{prefix}_{i}", ""))
            if pred and normalitzar_text(pred) in real_norm:
                total += 1.0
        return total

    ps = punts_fase("Setzens", 32, resultats.get("Setzens", []))
    pv = punts_fase("Vuitens", 16, resultats.get("Vuitens", []))
    pq = punts_fase("Quarts", 8, resultats.get("Quarts", []))
    psem = punts_fase("Semis", 4, resultats.get("Semis", []))

    pf = 0.0
    finalistes = {normalitzar_text(x) for x in resultats.get("Finalistes", []) if x}
    for col in ["Final_1", "Final_2"]:
        pred = normalitzar_equip(row.get(col, ""))
        if pred and normalitzar_text(pred) in finalistes:
            pf += 1.0

    campio = normalitzar_equip(row.get("Campió", ""))
    campio_real = {normalitzar_text(x) for x in resultats.get("Campió", []) if x}
    pc = 2.0 if campio and normalitzar_text(campio) in campio_real else 0.0

    pmvp = pd.to_numeric(row.get("Punts MVP", 0), errors="coerce")
    ppichichi = pd.to_numeric(row.get("Punts Pichichi", 0), errors="coerce")
    presfinal = pd.to_numeric(row.get("Resultat Final", 0), errors="coerce")
    pmvp = 0.0 if pd.isna(pmvp) else float(pmvp)
    ppichichi = 0.0 if pd.isna(ppichichi) else float(ppichichi)
    presfinal = 0.0 if pd.isna(presfinal) else float(presfinal)
    total = p1 + p2 + p3 + ps + pv + pq + psem + pf + pc + pmvp + ppichichi + presfinal
    return {
        "Punts Grups 1r": round(p1, 1), "Punts Grups 2n": round(p2, 1), "Punts Grups 3r": round(p3, 1),
        "Punts Setzens": round(ps, 1), "Punts Vuitens": round(pv, 1), "Punts Quarts": round(pq, 1), "Punts Semis": round(psem, 1),
        "Punts Finalistes": round(pf, 1), "Punts Campió": round(pc, 1), "Punts MVP": round(pmvp, 1), "Punts Pichichi": round(ppichichi, 1),
        "Resultat Final": round(presfinal, 1), "Total Punts": round(total, 1)
    }

def aplicar_recompte_api_a_porra(df_porra, resultats):
    if resultats.get("source") != "API":
        return df_porra
    df = df_porra.copy()
    df_punts = pd.DataFrame([calcular_punts_participant_api(row, resultats) for _, row in df.iterrows()])
    for col in df_punts.columns:
        df[col] = df_punts[col]
    return df

# ==================================================
# RÀNQUINGS / MOVIMENTS / HISTÒRIC
# ==================================================
def recalcular_posicions(df):
    df = df.copy().sort_values("Punts", ascending=False).reset_index(drop=True)
    df["Posició"] = df.index + 1
    df["Punts"] = pd.to_numeric(df["Punts"], errors="coerce").fillna(0).round(1)
    df["Dif líder"] = (df["Punts"] - float(df["Punts"].iloc[0])).round(1) if not df.empty else 0.0
    return df

def crear_ranking_des_de_porra(df_porra):
    if "Participants" not in df_porra.columns or "Total Punts" not in df_porra.columns:
        st.error("Falten columnes obligatòries al full Porra: Participants i/o Total Punts")
        st.write("Columnes detectades:", list(df_porra.columns))
        st.stop()
    col_dep = obtenir_columna_departament(df_porra)
    cols = ["Participants", "Total Punts"] + ([col_dep] if col_dep else [])
    df = df_porra[cols].copy()
    rename = {"Participants": "Participant", "Total Punts": "Punts"}
    if col_dep:
        rename[col_dep] = "Departament"
    df = df.rename(columns=rename)
    df["Participant"] = df["Participant"].astype(str).str.strip()
    df["Punts"] = pd.to_numeric(df["Punts"], errors="coerce")
    if "Departament" in df.columns:
        df["Departament"] = df["Departament"].fillna("Sense departament").astype(str).str.strip().replace("", "Sense departament")
    df = df.dropna(subset=["Punts"])
    df = df[df["Participant"] != ""]
    df = df[~df["Participant"].str.contains("Total", case=False, na=False)]
    return recalcular_posicions(df)

def crear_ranking_departaments(df_ranking):
    if "Departament" not in df_ranking.columns:
        return pd.DataFrame()
    df = df_ranking.copy()
    resum = df.groupby("Departament", as_index=False).agg(
        Participants=("Participant", "count"), Punts_totals=("Punts", "sum"), Mitjana_punts=("Punts", "mean"), Millor_puntuacio=("Punts", "max")
    )
    lider = df.sort_values("Punts", ascending=False).drop_duplicates("Departament")[["Departament", "Participant"]].rename(columns={"Participant": "Líder departament"})
    resum = resum.merge(lider, on="Departament", how="left")
    for col in ["Punts_totals", "Mitjana_punts", "Millor_puntuacio"]:
        resum[col] = pd.to_numeric(resum[col], errors="coerce").fillna(0).round(1)
    resum = resum.sort_values(["Mitjana_punts", "Punts_totals"], ascending=[False, False]).reset_index(drop=True)
    resum["Posició"] = resum.index + 1
    resum["Dif líder"] = (resum["Mitjana_punts"] - float(resum["Mitjana_punts"].iloc[0])).round(1) if not resum.empty else 0.0
    return resum[["Posició", "Departament", "Participants", "Mitjana_punts", "Punts_totals", "Millor_puntuacio", "Líder departament", "Dif líder"]]

def ranking_signature(df_ranking, resultats_actuals):
    cols = ["Participant", "Punts", "Posició"]
    if "Departament" in df_ranking.columns:
        cols.append("Departament")
    payload = df_ranking[cols].sort_values("Participant").to_json(orient="records", force_ascii=False)
    payload += json.dumps({k: resultats_actuals.get(k) for k in ["source", "Setzens", "Vuitens", "Quarts", "Semis", "Finalistes", "Campió"]}, ensure_ascii=False, sort_keys=True)
    return hashlib.md5(payload.encode("utf-8")).hexdigest()

def carregar_csv_segura(path):
    try:
        return pd.read_csv(path) if os.path.exists(path) else pd.DataFrame()
    except Exception:
        return pd.DataFrame()

def carregar_meta_snapshot():
    if not os.path.exists(SNAPSHOT_META_FILE):
        return {}
    try:
        with open(SNAPSHOT_META_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def guardar_meta_snapshot(snapshot_key):
    with open(SNAPSHOT_META_FILE, "w", encoding="utf-8") as f:
        json.dump({"snapshot_key": snapshot_key, "updated_at": datetime.now(tz=ZoneInfo("Europe/Madrid")).isoformat()}, f, ensure_ascii=False, indent=2)

def guardar_snapshot_actual(df_ranking):
    cols = ["Participant", "Punts", "Posició"] + (["Departament"] if "Departament" in df_ranking.columns else [])
    df_ranking[cols].rename(columns={"Punts": "Punts anteriors", "Posició": "Posició anterior"}).to_csv(SNAPSHOT_CURRENT_FILE, index=False)

def guardar_snapshot_display(df_ranking):
    cols = ["Participant", "Canvi posició", "Canvi punts", "Canvi num posició", "Punts anteriors", "Posició anterior"]
    df_ranking[[c for c in cols if c in df_ranking.columns]].to_csv(SNAPSHOT_DISPLAY_FILE, index=False)

def aplicar_moviment(df_ranking, snapshot_key):
    df = df_ranking.copy()
    meta = carregar_meta_snapshot()
    if meta.get("snapshot_key") == snapshot_key:
        df_mov = carregar_csv_segura(SNAPSHOT_DISPLAY_FILE)
        if not df_mov.empty and "Participant" in df_mov.columns:
            df = df.merge(df_mov, on="Participant", how="left")
            df["Canvi posició"] = df["Canvi posició"].fillna("⚪ —")
            df["Canvi punts"] = pd.to_numeric(df["Canvi punts"], errors="coerce").fillna(0.0).round(1)
            return df
    prev = carregar_csv_segura(SNAPSHOT_CURRENT_FILE)
    if prev.empty or "Participant" not in prev.columns:
        df["Posició anterior"] = df["Posició"]
        df["Punts anteriors"] = df["Punts"]
        df["Canvi punts"] = 0.0
        df["Canvi num posició"] = 0
        df["Canvi posició"] = "⚪ —"
    else:
        prev["Participant"] = prev["Participant"].astype(str).str.strip()
        df = df.merge(prev, on="Participant", how="left")
        df["Punts anteriors"] = pd.to_numeric(df["Punts anteriors"], errors="coerce")
        df["Posició anterior"] = pd.to_numeric(df["Posició anterior"], errors="coerce")
        df["Canvi punts"] = (df["Punts"] - df["Punts anteriors"]).fillna(0).round(1)
        df["Canvi num posició"] = (df["Posició anterior"] - df["Posició"]).fillna(0).astype(int)
        df["Canvi posició"] = df["Canvi num posició"].apply(lambda x: f"🟢 ▲ +{x}" if x > 0 else (f"🔴 ▼ {x}" if x < 0 else "⚪ —"))
    guardar_snapshot_actual(df)
    guardar_snapshot_display(df)
    guardar_meta_snapshot(snapshot_key)
    return df

def aplicar_moviment_departament(df_dep_actual, df_ranking_global, departament):
    df = df_dep_actual.copy()
    for col in ["Punts anteriors", "Posició anterior", "Punts anteriors dep", "Posició anterior dep", "Canvi punts", "Canvi num posició", "Canvi posició"]:
        if col in df.columns:
            df = df.drop(columns=[col])
    if "Punts anteriors" not in df_ranking_global.columns or "Departament" not in df_ranking_global.columns:
        df["Canvi posició"] = "⚪ —"
        df["Canvi punts"] = 0.0
        return df
    prev = df_ranking_global[df_ranking_global["Departament"].astype(str).str.strip() == str(departament).strip()].copy()
    prev["Punts anteriors"] = pd.to_numeric(prev.get("Punts anteriors"), errors="coerce")
    prev = prev.dropna(subset=["Punts anteriors"])
    if prev.empty:
        df["Canvi posició"] = "⚪ —"
        df["Canvi punts"] = 0.0
        return df
    prev = prev.sort_values("Punts anteriors", ascending=False).reset_index(drop=True)
    prev["Posició anterior dep"] = prev.index + 1
    prev = prev[["Participant", "Punts anteriors", "Posició anterior dep"]].rename(columns={"Punts anteriors": "Punts anteriors dep"})
    df = df.merge(prev, on="Participant", how="left")
    df["Punts anteriors dep"] = pd.to_numeric(df["Punts anteriors dep"], errors="coerce")
    df["Posició anterior dep"] = pd.to_numeric(df["Posició anterior dep"], errors="coerce")
    df["Canvi punts"] = (df["Punts"] - df["Punts anteriors dep"]).fillna(0).round(1)
    df["Canvi num posició"] = (df["Posició anterior dep"] - df["Posició"]).fillna(0).astype(int)
    df["Canvi posició"] = df["Canvi num posició"].apply(lambda x: f"🟢 ▲ +{x}" if x > 0 else (f"🔴 ▼ {x}" if x < 0 else "⚪ —"))
    return df

def carregar_historic():
    return carregar_csv_segura(HISTORY_FILE)

def registrar_historic(df_ranking, snapshot_key):
    hist = carregar_historic()
    if not hist.empty and "snapshot_key" in hist.columns and snapshot_key in hist["snapshot_key"].astype(str).values:
        return hist
    cols = ["Participant", "Punts", "Posició"] + (["Departament"] if "Departament" in df_ranking.columns else [])
    nou = df_ranking[cols].copy()
    nou["snapshot_key"] = snapshot_key
    nou["Actualització"] = datetime.now(tz=ZoneInfo("Europe/Madrid")).strftime("%d/%m/%Y %H:%M")
    hist = nou if hist.empty else pd.concat([hist, nou], ignore_index=True)
    hist.to_csv(HISTORY_FILE, index=False)
    return hist

# ==================================================
# VISUALS
# ==================================================
def highlight_leader(row):
    return ["background-color: #ffe066; font-weight: bold;" if row.get("Posició") == 1 else "" for _ in row]

def mostrar_taula_ranking(df):
    cols = ["Posició", "Canvi posició", "Participant"]
    if "Departament" in df.columns:
        cols.append("Departament")
    cols += ["Punts", "Dif líder", "Canvi punts"]
    cols = [c for c in cols if c in df.columns]
    d = df[cols].copy()
    for c in ["Punts", "Dif líder", "Canvi punts"]:
        if c in d.columns:
            d[c] = pd.to_numeric(d[c], errors="coerce").fillna(0).round(1)
    fmt = {c: "{:.1f}" for c in ["Punts", "Dif líder"] if c in d.columns}
    if "Canvi punts" in d.columns:
        fmt["Canvi punts"] = "{:+.1f}"
    st.dataframe(d.style.apply(highlight_leader, axis=1).format(fmt), use_container_width=True, hide_index=True)

def mostrar_taula_departaments(df_dep):
    if df_dep.empty:
        st.info("Afegeix una columna 'Departament' al full Porra per activar aquest mode.")
        return
    d = df_dep.copy()
    for c in ["Mitjana_punts", "Punts_totals", "Millor_puntuacio", "Dif líder"]:
        if c in d.columns:
            d[c] = pd.to_numeric(d[c], errors="coerce").fillna(0).round(1)
    st.dataframe(d.style.apply(highlight_leader, axis=1).format({c: "{:.1f}" for c in ["Mitjana_punts", "Punts_totals", "Millor_puntuacio", "Dif líder"] if c in d.columns}), use_container_width=True, hide_index=True)

def mostrar_grafic_punts(df, color="#0b70c9", altura_minima=950):
    if df.empty:
        return
    data = df[["Participant", "Punts", "Posició", "Dif líder"]].copy().sort_values("Punts", ascending=False)
    data["Punts"] = pd.to_numeric(data["Punts"], errors="coerce").fillna(0).round(1)
    chart = alt.Chart(data).mark_bar(color=color).encode(
        x=alt.X("Punts:Q", title="Punts", scale=alt.Scale(zero=False)),
        y=alt.Y("Participant:N", sort="-x", title=None),
        tooltip=["Posició", "Participant", alt.Tooltip("Punts:Q", format=".1f"), alt.Tooltip("Dif líder:Q", format=".1f")]
    ).properties(height=max(altura_minima, len(data) * 36))
    st.altair_chart(chart, use_container_width=True)

def mostrar_grafic_departaments(df_dep):
    if df_dep.empty:
        return
    data = df_dep.copy().sort_values("Mitjana_punts", ascending=False)
    chart = alt.Chart(data).mark_bar(color="#0f9d58").encode(
        x=alt.X("Mitjana_punts:Q", title="Mitjana de punts", scale=alt.Scale(zero=False)),
        y=alt.Y("Departament:N", sort="-x", title=None),
        tooltip=list(data.columns)
    ).properties(height=max(350, len(data) * 46))
    st.altair_chart(chart, use_container_width=True)

def mostrar_evolucio_temporal(df_hist, df_ranking):
    st.subheader("📉 Evolució temporal")
    if df_hist.empty or "Actualització" not in df_hist.columns or df_hist["Actualització"].nunique() < 2:
        st.info("Encara no hi ha històric suficient. Es registrarà quan canviïn els punts o les posicions.")
        return
    participants_sel = st.multiselect("Selecciona participants", sorted(df_hist["Participant"].astype(str).unique()), default=df_ranking.head(5)["Participant"].tolist())
    if not participants_sel:
        return
    data = df_hist[df_hist["Participant"].isin(participants_sel)].copy()
    data["Punts"] = pd.to_numeric(data["Punts"], errors="coerce")
    data["Posició"] = pd.to_numeric(data["Posició"], errors="coerce")
    ordre = list(data["Actualització"].drop_duplicates())
    st.write("### 📈 Evolució de punts")
    chart = alt.Chart(data).mark_line(point=True).encode(
        x=alt.X("Actualització:N", sort=ordre, title="Actualització"), y=alt.Y("Punts:Q", title="Punts"), color="Participant:N",
        tooltip=["Actualització", "Participant", alt.Tooltip("Punts:Q", format=".1f"), "Posició"]
    ).properties(height=420)
    st.altair_chart(chart, use_container_width=True)
    st.write("### 📉 Evolució de posició")
    chart2 = alt.Chart(data).mark_line(point=True).encode(
        x=alt.X("Actualització:N", sort=ordre, title="Actualització"), y=alt.Y("Posició:Q", title="Posició", scale=alt.Scale(reverse=True)), color="Participant:N",
        tooltip=["Actualització", "Participant", "Posició", alt.Tooltip("Punts:Q", format=".1f")]
    ).properties(height=420)
    st.altair_chart(chart2, use_container_width=True)

def mostrar_animacio_evolucio(df_hist):
    st.subheader("📽️ Animació evolució del rànquing")
    if df_hist.empty or "Actualització" not in df_hist.columns or df_hist["Actualització"].nunique() < 2:
        st.info("Calen almenys dues versions per reproduir l’animació.")
        return
    max_pos = st.slider("Nombre de posicions a mostrar", 5, 20, 10, 1)
    if st.button("▶️ Reproduir evolució"):
        placeholder = st.empty()
        for act in df_hist["Actualització"].drop_duplicates():
            frame = df_hist[df_hist["Actualització"] == act].copy()
            frame["Posició"] = pd.to_numeric(frame["Posició"], errors="coerce")
            frame["Punts"] = pd.to_numeric(frame["Punts"], errors="coerce").round(1)
            frame = frame.sort_values("Posició").head(max_pos)
            with placeholder.container():
                st.write(f"### 🕒 {act}")
                st.dataframe(frame[["Posició", "Participant", "Punts"]], use_container_width=True, hide_index=True)

def obtenir_pichichi_real(df_resultats_display, col_pichichi, col_gols):
    cp = trobar_columna_flexible(df_resultats_display, col_pichichi)
    cg = trobar_columna_flexible(df_resultats_display, col_gols)
    if not cp or not cg:
        return "Pendent", "Pendent"
    taula = df_resultats_display[[cp, cg]].copy()
    taula[cp] = taula[cp].astype(str).str.strip()
    taula[cg] = pd.to_numeric(taula[cg], errors="coerce")
    taula = taula[(taula[cp] != "") & (~taula[cp].str.lower().isin(["nan", "nat", "pendent"])) & (taula[cg] >= 1)]
    if taula.empty:
        return "Pendent", "Pendent"
    taula = taula.sort_values(cg, ascending=False).reset_index(drop=True)
    return taula.iloc[0][cp], str(int(taula.iloc[0][cg]))

def obtenir_prediccions_fase(df_j, prefix, quantitat):
    files = []
    for i in range(1, quantitat + 1):
        col = f"{prefix}_{i}"
        valor = valor_o_pendent(df_j[col].values[0]) if col in df_j.columns else "Pendent"
        files.append({"Posició": i, "Equip": afegir_bandera(valor)})
    return pd.DataFrame(files)

def mostrar_prediccions_eliminatoria_participant(df_j):
    st.write("### 🧭 Prediccions fase eliminatòria")
    tabs = st.tabs(["Setzens", "Vuitens", "Quarts", "Semis", "Final", "Campió"])
    for idx, (prefix, qty) in enumerate([("Setzens", 32), ("Vuitens", 16), ("Quarts", 8), ("Semis", 4)]):
        with tabs[idx]:
            st.dataframe(obtenir_prediccions_fase(df_j, prefix, qty), use_container_width=True, hide_index=True)
    with tabs[4]:
        finalistes = [afegir_bandera(valor_o_pendent(df_j[col].values[0])) if col in df_j.columns else "Pendent" for col in ["Final_1", "Final_2"]]
        st.dataframe(pd.DataFrame({"Finalista": ["Finalista 1", "Finalista 2"], "Equip": finalistes}), use_container_width=True, hide_index=True)
    with tabs[5]:
        campio = valor_o_pendent(df_j["Campió"].values[0]) if "Campió" in df_j.columns else "Pendent"
        st.dataframe(pd.DataFrame({"Concepte": ["Campió previst"], "Equip": [afegir_bandera(campio)]}), use_container_width=True, hide_index=True)

def mostrar_fase_eliminatoria_responsive(files_eliminatoria):
    if not files_eliminatoria:
        st.info("No hi ha dades de fase eliminatòria configurades.")
        return
    html_files = ""
    for fila in files_eliminatoria:
        fase = fila["Fase"]
        resultat = str(fila["Resultat"]).strip()
        if resultat == "" or normalitzar_text(resultat) == "pendent":
            equips_html = "<span class='elim-pending'>Pendent</span>"
        else:
            equips = [x.strip() for x in resultat.split(" · ") if x.strip()]
            equips_html = "".join([f"<span class='elim-badge'>{equip}</span>" for equip in equips])
        html_files += f"<div class='elim-row'><div class='elim-phase'>{fase}</div><div class='elim-teams'>{equips_html}</div></div>"
    st.markdown(f"<div class='elim-wrapper'>{html_files}</div>", unsafe_allow_html=True)

# ==================================================
# ESTILS
# ==================================================
img_base64 = carregar_imatge_base64(BACKGROUND_IMAGE)
if img_base64:
    background_css = f"""
    background-image: linear-gradient(rgba(0,0,0,0.35), rgba(0,0,0,0.55)), url('data:image/jpg;base64,{img_base64}');
    background-size: cover; background-position: center; background-attachment: fixed;
    """
else:
    background_css = "background: #eef2f7;"

st.markdown(f"""
<style>
.stApp {{ {background_css} }}
.block-container {{ padding-top: 2rem; padding-bottom: 2rem; background: rgba(255,255,255,0.88); backdrop-filter: blur(6px); border-radius: 26px; margin-top: 24px; margin-bottom: 24px; box-shadow: 0px 10px 35px rgba(0,0,0,0.32); }}
h1,h2,h3,h4,h5,h6 {{ color: #102a43; text-shadow: 0px 1px 2px rgba(255,255,255,0.6); }}
.title {{ font-size: clamp(34px, 5vw, 58px); font-weight: 900; margin-bottom: 0px; color: #102a43; letter-spacing: -1px; text-shadow: 0px 2px 8px rgba(255,255,255,0.8); }}
.subtitle {{ font-size: clamp(14px, 2vw, 19px); color: #334e68; margin-top: 0px; margin-bottom: 25px; text-shadow: 0px 1px 4px rgba(255,255,255,0.7); }}
.card {{ padding: 20px; border-radius: 20px; text-align: center; box-shadow: 0px 6px 24px rgba(0,0,0,0.24); min-height: 182px; display: flex; flex-direction: column; justify-content: center; align-items: center; box-sizing: border-box; overflow: hidden; width: 100%; transition: all 0.25s ease; }}
.card:hover {{ transform: translateY(-4px); box-shadow: 0px 10px 30px rgba(0,0,0,0.32); }}
.card h3 {{ margin: 0 0 14px 0; font-size: clamp(15px,2vw,24px); line-height:1.15; max-width:100%; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; text-shadow:0px 2px 6px rgba(0,0,0,0.35); }}
.card h1 {{ margin:0; font-size:clamp(26px,4vw,42px); line-height:1.1; max-width:100%; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; text-shadow:0px 2px 6px rgba(0,0,0,0.35); }}
.card p {{ margin:12px 0 0 0; font-size:clamp(11px,1.5vw,15px); max-width:100%; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; text-shadow:0px 2px 6px rgba(0,0,0,0.35); }}
.gold {{ background: linear-gradient(135deg,#ffd700,#fff1a8); color:#111; }} .silver {{ background: linear-gradient(135deg,#c0c0c0,#f2f2f2); color:#111; }} .bronze {{ background: linear-gradient(135deg,#cd7f32,#f0b27a); color:white; }}
.bluecard {{ background: linear-gradient(135deg,#0b70c9,#7cc5ff); color:white; }} .greencard {{ background: linear-gradient(135deg,#0f9d58,#8ee6b3); color:white; }} .darkcard {{ background: linear-gradient(135deg,#102a43,#486581); color:white; }} .purplecard {{ background: linear-gradient(135deg,#6f42c1,#b982ff); color:white; margin-top:18px; margin-bottom:18px; }}
.elim-wrapper {{ width:100%; display:flex; flex-direction:column; gap:12px; margin-top:12px; margin-bottom:18px; }} .elim-row {{ display:grid; grid-template-columns:150px minmax(0,1fr); gap:12px; align-items:flex-start; background:rgba(255,255,255,0.78); border:1px solid rgba(0,0,0,0.08); border-radius:16px; padding:14px; box-shadow:0px 4px 16px rgba(0,0,0,0.10); }} .elim-phase {{ font-weight:800; color:#102a43; font-size:15px; padding-top:6px; }} .elim-teams {{ display:flex; flex-wrap:wrap; gap:7px; line-height:1.35; min-width:0; }} .elim-badge {{ display:inline-block; background:#edf2f7; color:#102a43; border:1px solid #d9e2ec; border-radius:999px; padding:6px 10px; font-size:13px; font-weight:600; white-space:normal; max-width:100%; box-shadow:0px 2px 6px rgba(0,0,0,0.06); }} .elim-pending {{ display:inline-block; background:#fff3cd; color:#665200; border:1px solid #ffe08a; border-radius:999px; padding:6px 10px; font-size:13px; font-weight:700; }}
@media (max-width:768px) {{ .block-container {{ padding-left:0.8rem; padding-right:0.8rem; border-radius:16px; }} .card {{ height:auto; min-height:140px; margin-bottom:12px; }} .card h3,.card h1,.card p {{ white-space:normal; }} .elim-row {{ grid-template-columns:1fr; padding:12px; }} .elim-phase {{ font-size:16px; padding-top:0; border-bottom:1px solid rgba(0,0,0,0.08); padding-bottom:8px; }} .elim-badge {{ font-size:12px; padding:5px 8px; }} }}
</style>
""", unsafe_allow_html=True)

# ==================================================
# SIDEBAR
# ==================================================
with st.sidebar:
    st.write("### ⚙️ Gestió")
    if st.button("🔄 Reiniciar comparativa de moviments"):
        for fitxer in [SNAPSHOT_CURRENT_FILE, SNAPSHOT_DISPLAY_FILE, SNAPSHOT_META_FILE]:
            if os.path.exists(fitxer):
                os.remove(fitxer)
        st.rerun()
    if st.button("🧹 Reiniciar històric d’evolució"):
        if os.path.exists(HISTORY_FILE):
            os.remove(HISTORY_FILE)
        st.rerun()
    st.markdown("---")
    st.caption("Mode híbrid: Excel manual + API TheStatsAPI opcional.")

# ==================================================
# CARREGAR DADES
# ==================================================
excel_mtime = os.path.getmtime(EXCEL_FILE) if os.path.exists(EXCEL_FILE) else 0
data_actualitzacio = obtenir_data_actualitzacio_fitxer(EXCEL_FILE)
df_porra_original, df_resultats = carregar_dades(EXCEL_FILE, excel_mtime)
resultats_actuals = obtenir_resultats_actuals(df_resultats)
df_porra = aplicar_recompte_api_a_porra(df_porra_original, resultats_actuals)
df_ranking_base = crear_ranking_des_de_porra(df_porra)
snapshot_key = ranking_signature(df_ranking_base, resultats_actuals)
df_ranking = aplicar_moviment(df_ranking_base, snapshot_key)
df_historic = registrar_historic(df_ranking_base, snapshot_key)
df_departaments = crear_ranking_departaments(df_ranking)
num_participants = len(df_ranking)
premi_guanyador = num_participants * PREU_PARTICIPACIO
te_departaments = "Departament" in df_ranking.columns and not df_departaments.empty

# ==================================================
# CAPÇALERA
# ==================================================
st.markdown('<p class="title">🏆 PORRA MUNDIAL</p>', unsafe_allow_html=True)
st.markdown('<p class="subtitle">Classificació en viu, resultats Excel/API, competició per departaments i resultats reals</p>', unsafe_allow_html=True)

info1, info2, info3 = st.columns(3, gap="small")
info1.markdown(f"<div class='card darkcard'><h3>🕒 Dades actualitzades</h3><h1>{data_actualitzacio}</h1></div>", unsafe_allow_html=True)
info2.markdown(f"<div class='card greencard'><h3>🎁 Premi guanyador</h3><h1>{premi_guanyador} €</h1><p>{num_participants} participants x {PREU_PARTICIPACIO} €</p></div>", unsafe_allow_html=True)
info3.markdown(f"<div class='card bluecard'><h3>👥 Participants</h3><h1>{num_participants}</h1><p>porres registrades</p></div>", unsafe_allow_html=True)

if te_departaments:
    dept_lider = df_departaments.iloc[0]
    st.markdown(f"<div class='card purplecard'><h3>🏢 Departament líder</h3><h1>{dept_lider['Departament']}</h1><p>Mitjana {float(dept_lider['Mitjana_punts']):.1f} punts · {int(dept_lider['Participants'])} participants</p></div>", unsafe_allow_html=True)

# ==================================================
# RÀNQUING
# ==================================================
st.subheader("🥇 TOP 3 General")
top3 = df_ranking.head(3)
top_cols = st.columns(3, gap="small")
for i, (klass, medal) in enumerate(zip(["gold", "silver", "bronze"], ["🥇", "🥈", "🥉"])):
    if i < len(top3):
        row = top3.iloc[i]
        subtext = row["Departament"] if "Departament" in row.index else "punts"
        canvi_pos = row["Canvi posició"] if "Canvi posició" in row.index else ""
        top_cols[i].markdown(f"<div class='card {klass}'><h3>{medal} {row['Participant']}</h3><h1>{float(row['Punts']):.1f}</h1><p>{subtext} · {canvi_pos}</p></div>", unsafe_allow_html=True)

st.subheader("📊 Classificació general")
mostrar_taula_ranking(df_ranking)

st.subheader("📈 Gràfic general de punts")
mostrar_grafic_punts(df_ranking, color="#0b70c9", altura_minima=1000)

mostrar_evolucio_temporal(df_historic, df_ranking_base)
mostrar_animacio_evolucio(df_historic)

# ==================================================
# FITXA PARTICIPANT
# ==================================================
st.subheader("👤 Fitxa participant")
participants_porra = df_porra["Participants"].dropna().astype(str).unique()
jugador = st.selectbox("Selecciona participant", participants_porra, index=None, placeholder="Selecciona un participant...")
if jugador is not None:
    df_j = df_porra[df_porra["Participants"].astype(str) == str(jugador)]
    if not df_j.empty:
        total = pd.to_numeric(df_j["Total Punts"].values[0], errors="coerce")
        c1, c2 = st.columns(2)
        c1.metric("Total punts", f"{total:.1f}")
        col_dep_original = obtenir_columna_departament(df_porra)
        if col_dep_original is not None:
            c1.metric("Departament", valor_o_pendent(df_j[col_dep_original].values[0]))
        punts_dict = {}
        for label, col in [("1rs grup", "Punts Grups 1r"), ("2ns grup", "Punts Grups 2n"), ("3rs grup", "Punts Grups 3r"), ("Setzens", "Punts Setzens"), ("Vuitens", "Punts Vuitens"), ("Quarts", "Punts Quarts"), ("Semis", "Punts Semis"), ("Finalistes", "Punts Finalistes"), ("Campió", "Punts Campió"), ("MVP", "Punts MVP"), ("Pichichi", "Punts Pichichi")]:
            if col in df_j.columns:
                punts_dict[label] = pd.to_numeric(df_j[col].values[0], errors="coerce")
        if punts_dict:
            punts_categoria = pd.DataFrame({"Categoria": list(punts_dict.keys()), "Punts": list(punts_dict.values())})
            punts_categoria["Punts"] = punts_categoria["Punts"].fillna(0).round(1)
            chart_cat = alt.Chart(punts_categoria).mark_bar(color="#1f77b4").encode(x=alt.X("Categoria:N", sort=None, title=None), y=alt.Y("Punts:Q", title="Punts"), tooltip=["Categoria", alt.Tooltip("Punts:Q", format=".1f")]).properties(height=320)
            c1.altair_chart(chart_cat, use_container_width=True)
        c2.write("### ⚽ Prediccions principals")
        campio = valor_o_pendent(df_j["Campió"].values[0]) if "Campió" in df_j.columns else "Pendent"
        mvp = valor_o_pendent(df_j["MVP"].values[0]) if "MVP" in df_j.columns else "Pendent"
        pichichi = valor_o_pendent(df_j["Pichichi"].values[0]) if "Pichichi" in df_j.columns else "Pendent"
        col_resultat_final_porra = trobar_col_resultat_final_porra(df_porra)
        resultat_final = valor_o_pendent(df_j[col_resultat_final_porra].values[0]) if col_resultat_final_porra is not None else "Pendent"
        c2.write(f"🏆 Campió: {afegir_bandera(campio)}")
        c2.write(f"📌 Resultat final: {resultat_final}")
        c2.write(f"⭐ MVP: {mvp}")
        c2.write(f"⚽ Pichichi: {pichichi}")
        mostrar_prediccions_eliminatoria_participant(df_j)
else:
    st.info("Selecciona un participant per veure el detall de punts i prediccions.")

# ==================================================
# DEPARTAMENTS
# ==================================================
st.subheader("🏢 Competició per departaments")
if te_departaments:
    st.write("Rànquing calculat per **mitjana de punts** del departament.")
    mostrar_taula_departaments(df_departaments)
    st.write("#### 📈 Gràfic departaments")
    mostrar_grafic_departaments(df_departaments)
    departaments_opcions = sorted(df_ranking["Departament"].dropna().astype(str).unique().tolist())
    departament_sel = st.selectbox("Selecciona departament", departaments_opcions, index=None, placeholder="Selecciona un departament...")
    if departament_sel:
        df_dep_individual = recalcular_posicions(df_ranking[df_ranking["Departament"] == departament_sel].copy())
        df_dep_individual = aplicar_moviment_departament(df_dep_individual, df_ranking, departament_sel)
        st.write(f"### 📊 Classificació interna · {departament_sel}")
        mostrar_taula_ranking(df_dep_individual)
        st.write(f"### 📈 Gràfic · {departament_sel}")
        mostrar_grafic_punts(df_dep_individual, color="#6f42c1", altura_minima=350)
else:
    st.info("Per activar aquest apartat, afegeix una columna 'Departament' al full Porra.")

# ==================================================
# LLIGUETA
# ==================================================
st.subheader("🏟️ Lligueta personalitzada")
tots_participants = df_ranking["Participant"].dropna().astype(str).tolist()
participants_filtrats = st.multiselect("Selecciona participants per crear una lligueta:", options=tots_participants, default=[], placeholder="Tria participants...")
if participants_filtrats:
    df_lligueta = recalcular_posicions(df_ranking[df_ranking["Participant"].astype(str).isin(participants_filtrats)].copy())
    st.write(f"Participants seleccionats: **{len(participants_filtrats)}**")
    mostrar_taula_ranking(df_lligueta)
    st.write("#### 📈 Gràfic de la lligueta")
    mostrar_grafic_punts(df_lligueta, color="#0f9d58", altura_minima=350)
else:
    st.write("Selecciona participants per crear una classificació reduïda tipus lligueta.")

# ==================================================
# RESULTATS REALS
# ==================================================
st.subheader("✅ Resultats reals")
df_resultats_display = preparar_taula_buida(df_resultats)
COL_GRUP = "Grup"; COL_POSICIO = "Posició"; COL_EQUIP = "Equip"; COL_SETZENS = "Setzens"; COL_VUITENS = "Vuitens"; COL_QUARTS = "Quarts"; COL_SEMIS = "Semis"; COL_FINALISTES = "Finalistes"; COL_CAMPIO = "Campió"; COL_MVP = "MVP"; COL_RESULTAT_FINAL = "Resultat Final"; COL_PICHICHI = "Jugador Pichichi"; COL_GOLS = "Gols"

campio_real = afegir_bandera(resultats_actuals.get("Campió")[0]) if resultats_actuals.get("Campió") else afegir_bandera(primer_valor_o_pendent(df_resultats_display, COL_CAMPIO))
mvp_real = primer_valor_o_pendent(df_resultats_display, COL_MVP)
resultat_final_real = primer_valor_o_pendent(df_resultats_display, COL_RESULTAT_FINAL)
if resultats_actuals.get("Pichichi"):
    pichichi_real = resultats_actuals.get("Pichichi")[0]
    gols_pichichi = resultats_actuals.get("Gols Pichichi", "Pendent")
else:
    pichichi_real, gols_pichichi = obtenir_pichichi_real(df_resultats_display, COL_PICHICHI, COL_GOLS)

st.write("### 🏟️ Resum oficial")
r1, r2, r3, r4 = st.columns(4, gap="small")
r1.markdown(f"<div class='card gold'><h3>🏆 Campió</h3><h1 style='font-size:28px'>{campio_real}</h1></div>", unsafe_allow_html=True)
r2.markdown(f"<div class='card silver'><h3>⭐ MVP</h3><h1 style='font-size:28px'>{mvp_real}</h1></div>", unsafe_allow_html=True)
pichichi_subtext = f"{gols_pichichi} gols" if gols_pichichi != "Pendent" else "Pendent"
r3.markdown(f"<div class='card bronze'><h3>⚽ Pichichi</h3><h1 style='font-size:25px'>{pichichi_real}</h1><p>{pichichi_subtext}</p></div>", unsafe_allow_html=True)
r4.markdown(f"<div class='card bluecard'><h3>🏁 Resultat final</h3><h1 style='font-size:28px'>{resultat_final_real}</h1></div>", unsafe_allow_html=True)

st.write("### 🧩 Fase de grups")
grups = {}
group_positions = resultats_actuals.get("group_positions", {})
if group_positions:
    for grup, posicions in group_positions.items():
        grups[f"Grup {grup}"] = {"1r": afegir_bandera(posicions.get("1r", "")) if posicions.get("1r") else "", "2n": afegir_bandera(posicions.get("2n", "")) if posicions.get("2n") else "", "3r": afegir_bandera(posicions.get("3r", "")) if posicions.get("3r") else ""}
elif all(trobar_columna_flexible(df_resultats_display, c) for c in [COL_GRUP, COL_POSICIO, COL_EQUIP]):
    cg, cp, ce = trobar_columna_flexible(df_resultats_display, COL_GRUP), trobar_columna_flexible(df_resultats_display, COL_POSICIO), trobar_columna_flexible(df_resultats_display, COL_EQUIP)
    for _, row in df_resultats_display.iterrows():
        grup = str(row.get(cg, "")).strip(); posicio = str(row.get(cp, "")).strip(); equip = str(row.get(ce, "")).strip()
        if not grup or not equip:
            continue
        grups.setdefault(grup, {"1r": "", "2n": "", "3r": ""})
        if posicio in ["1r", "1", "1º"]:
            grups[grup]["1r"] = afegir_bandera(equip)
        elif posicio in ["2n", "2", "2º"]:
            grups[grup]["2n"] = afegir_bandera(equip)
        elif posicio in ["3r", "3", "3º"]:
            grups[grup]["3r"] = afegir_bandera(equip)
if grups:
    st.dataframe(pd.DataFrame(grups).reindex(["1r", "2n", "3r"]).reset_index().rename(columns={"index": "Posició"}), use_container_width=True, hide_index=True)
else:
    st.info("No hi ha dades de fase de grups configurades.")

st.write("### ⚔️ Fase eliminatòria")
fases_eliminatoria = [COL_SETZENS, COL_VUITENS, COL_QUARTS, COL_SEMIS, COL_FINALISTES, COL_CAMPIO, COL_MVP]
files_eliminatoria = []
for fase in fases_eliminatoria:
    if resultats_actuals.get(fase):
        valors = resultats_actuals.get(fase, [])
        detall = " · ".join([afegir_bandera(v) for v in valors]) if fase != COL_MVP else " · ".join(valors)
    else:
        detall = "Pendent"
    files_eliminatoria.append({"Fase": fase, "Resultat": detall})
mostrar_fase_eliminatoria_responsive(files_eliminatoria)

st.write("### ⚽ Jugador pichichi")
col_pichichi = trobar_columna_flexible(df_resultats_display, COL_PICHICHI)
col_gols = trobar_columna_flexible(df_resultats_display, COL_GOLS)
if col_pichichi and col_gols:
    taula_pichichi = df_resultats_display[[col_pichichi, col_gols]].copy()
    taula_pichichi[col_pichichi] = taula_pichichi[col_pichichi].astype(str).str.strip()
    taula_pichichi[col_gols] = pd.to_numeric(taula_pichichi[col_gols], errors="coerce")
    taula_pichichi = taula_pichichi[(taula_pichichi[col_pichichi] != "") & (~taula_pichichi[col_pichichi].str.lower().isin(["nan", "nat", "pendent"])) & (taula_pichichi[col_gols] >= 1)]
    if taula_pichichi.empty:
        taula_pichichi = pd.DataFrame({COL_PICHICHI: ["Pendent"], COL_GOLS: ["Pendent"]})
    else:
        taula_pichichi = taula_pichichi.sort_values(col_gols, ascending=False)
        taula_pichichi[col_gols] = taula_pichichi[col_gols].astype("Int64")
        taula_pichichi = taula_pichichi.rename(columns={col_pichichi: COL_PICHICHI, col_gols: COL_GOLS})
    st.dataframe(taula_pichichi, use_container_width=True, hide_index=True)
else:
    st.dataframe(pd.DataFrame({COL_PICHICHI: ["Pendent"], COL_GOLS: ["Pendent"]}), use_container_width=True, hide_index=True)

st.write("### 🏁 Resultat de la final")
st.dataframe(pd.DataFrame({"Concepte": ["Resultat de la final"], "Valor": [primer_valor_o_pendent(df_resultats_display, COL_RESULTAT_FINAL)]}), use_container_width=True, hide_index=True)

# ==================================================
# ESTAT API AL FINAL
# ==================================================
st.markdown("---")
if resultats_actuals.get("source") == "API":
    st.success("Mode resultats: API TheStatsAPI activada. Si alguna fase no té dades API, es completa amb el full Excel.")
else:
    st.info("Mode resultats: Excel. L’app utilitza el full 'Resultats Reals' del fitxer Excel.")
if resultats_actuals.get("api_error"):
    st.warning(f"Error API: {resultats_actuals.get('api_error')}. L’app ha continuat amb Excel.")
st.write("📡 Actualització automàtica des de Excel / API TheStatsAPI")
