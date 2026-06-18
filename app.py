
import streamlit as st
import pandas as pd
import altair as alt
import base64
import os
import json
import unicodedata
import time
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


def llegir_secret(nom, defecte=None):
    try:
        return st.secrets[nom]
    except Exception:
        return defecte


def obtenir_config_api():
    return {
        "use_api": str(llegir_secret("USE_API_RESULTS", "false")).lower() == "true",
        "api_key": llegir_secret("THESTATSAPI_KEY", ""),
        "competition_id": llegir_secret("THESTATSAPI_COMPETITION_ID", DEFAULT_COMPETITION_ID),
        "season_id": llegir_secret("THESTATSAPI_SEASON_ID", DEFAULT_SEASON_ID),
    }

FLAGS = {
    "mexic": "🇲🇽", "mexico": "🇲🇽", "corea del sud": "🇰🇷", "south korea": "🇰🇷", "korea republic": "🇰🇷",
    "republica txeca": "🇨🇿", "czechia": "🇨🇿", "suissa": "🇨🇭", "switzerland": "🇨🇭", "canada": "🇨🇦",
    "qatar": "🇶🇦", "escocia": "🏴", "scotland": "🏴", "marroc": "🇲🇦", "morocco": "🇲🇦",
    "brasil": "🇧🇷", "brazil": "🇧🇷", "estats units": "🇺🇸", "united states": "🇺🇸", "usa": "🇺🇸", "eeuu": "🇺🇸", "ee.uu": "🇺🇸",
    "australia": "🇦🇺", "turquia": "🇹🇷", "turkiye": "🇹🇷", "turkey": "🇹🇷", "alemanya": "🇩🇪", "germany": "🇩🇪",
    "costa d'ivori": "🇨🇮", "cote d'ivoire": "🇨🇮", "côte d'ivoire": "🇨🇮", "equador": "🇪🇨", "ecuador": "🇪🇨",
    "suecia": "🇸🇪", "sweden": "🇸🇪", "japo": "🇯🇵", "japan": "🇯🇵", "paisos baixos": "🇳🇱", "netherlands": "🇳🇱",
    "nova zelanda": "🇳🇿", "new zealand": "🇳🇿", "iran": "🇮🇷", "ir iran": "🇮🇷", "belgica": "🇧🇪", "belgium": "🇧🇪",
    "uruguai": "🇺🇾", "uruguay": "🇺🇾", "arabia saudita": "🇸🇦", "saudi arabia": "🇸🇦", "espanya": "🇪🇸", "spain": "🇪🇸",
    "franca": "🇫🇷", "france": "🇫🇷", "senegal": "🇸🇳", "argentina": "🇦🇷", "algeria": "🇩🇿", "austria": "🇦🇹",
    "portugal": "🇵🇹", "rd congo": "🇨🇩", "dr congo": "🇨🇩", "congo dr": "🇨🇩", "uzbekistan": "🇺🇿",
    "anglaterra": "🏴", "england": "🏴", "croacia": "🇭🇷", "croatia": "🇭🇷", "ghana": "🇬🇭", "egipte": "🇪🇬", "egypt": "🇪🇬",
    "noruega": "🇳🇴", "norway": "🇳🇴", "colombia": "🇨🇴", "colòmbia": "🇨🇴", "bosnia i hercegovina": "🇧🇦", "bosnia and herzegovina": "🇧🇦", "bosnia & herzegovina": "🇧🇦",
    "paraguai": "🇵🇾", "paraguay": "🇵🇾", "tunisia": "🇹🇳", "tunísia": "🇹🇳", "cap verd": "🇨🇻", "cabo verde": "🇨🇻",
    "jordania": "🇯🇴", "jordània": "🇯🇴", "jordan": "🇯🇴", "panama": "🇵🇦", "panamà": "🇵🇦", "curacao": "🇨🇼", "curaçao": "🇨🇼",
    "haiti": "🇭🇹", "haití": "🇭🇹", "sud-africa": "🇿🇦", "sud-àfrica": "🇿🇦", "south africa": "🇿🇦",
}

TEAM_NAME_MAP = {k: v for k, v in {
    "mexico": "Mèxic", "mexic": "Mèxic", "south africa": "Sud-àfrica", "sud africa": "Sud-àfrica", "sud-africa": "Sud-àfrica", "sud-àfrica": "Sud-àfrica",
    "korea republic": "Corea del Sud", "south korea": "Corea del Sud", "corea del sud": "Corea del Sud", "czechia": "República Txeca", "republica txeca": "República Txeca", "republica checa": "República Txeca",
    "canada": "Canadà", "bosnia and herzegovina": "Bosnia i Hercegovina", "bosnia & herzegovina": "Bosnia i Hercegovina", "bosnia i hercegovina": "Bosnia i Hercegovina", "qatar": "Qatar",
    "switzerland": "Suïssa", "suissa": "Suïssa", "brazil": "Brasil", "brasil": "Brasil", "morocco": "Marroc", "marroc": "Marroc", "haiti": "Haití", "haití": "Haití", "scotland": "Escòcia", "escocia": "Escòcia",
    "united states": "Estats Units", "usa": "Estats Units", "eeuu": "Estats Units", "ee.uu": "Estats Units", "estats units": "Estats Units", "paraguay": "Paraguai", "paraguai": "Paraguai",
    "australia": "Austràlia", "turkiye": "Turquia", "turkey": "Turquia", "turquia": "Turquia", "germany": "Alemanya", "alemanya": "Alemanya", "curacao": "Curaçao", "curaçao": "Curaçao",
    "cote d'ivoire": "Costa d'Ivori", "côte d'ivoire": "Costa d'Ivori", "costa d'ivori": "Costa d'Ivori", "ecuador": "Equador", "equador": "Equador", "netherlands": "Països Baixos", "paisos baixos": "Països Baixos", "pasïsos baixos": "Països Baixos",
    "japan": "Japó", "japo": "Japó", "sweden": "Suècia", "suecia": "Suècia", "tunisia": "Tunísia", "belgium": "Bèlgica", "belgica": "Bèlgica", "egypt": "Egipte", "egipte": "Egipte", "ir iran": "Iran", "iran": "Iran",
    "new zealand": "Nova Zelanda", "nova zelanda": "Nova Zelanda", "spain": "Espanya", "espanya": "Espanya", "cabo verde": "Cap Verd", "cap verd": "Cap Verd", "saudi arabia": "Aràbia Saudita", "arabia saudita": "Aràbia Saudita",
    "uruguay": "Uruguai", "uruguai": "Uruguai", "france": "França", "franca": "França", "senegal": "Senegal", "iraq": "Iraq", "norway": "Noruega", "noruega": "Noruega", "argentina": "Argentina", "algeria": "Algèria", "austria": "Àustria",
    "jordan": "Jordània", "jordania": "Jordània", "portugal": "Portugal", "congo dr": "RD Congo", "dr congo": "RD Congo", "rd congo": "RD Congo", "uzbekistan": "Uzbekistan", "colombia": "Colòmbia", "colòmbia": "Colòmbia",
    "england": "Anglaterra", "anglaterra": "Anglaterra", "croatia": "Croàcia", "croacia": "Croàcia", "ghana": "Ghana", "panama": "Panamà", "panamà": "Panamà",
}.items()}


def normalitzar_text(text):
    text = str(text).strip().lower().replace("’", "'").replace("`", "'")
    text = unicodedata.normalize("NFD", text)
    text = "".join(char for char in text if unicodedata.category(char) != "Mn")
    return " ".join(text.split())


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
    if text == "" or normalitzar_text(text) in ["nan", "nat", "pendent", "none"]:
        return "Pendent"
    text_norm = normalitzar_text(text)
    for pais, bandera in FLAGS.items():
        if pais in text_norm:
            return f"{bandera} {text}"
    return text


def valor_o_pendent(valor):
    if pd.isna(valor):
        return "Pendent"
    text = str(valor).strip()
    if text == "" or normalitzar_text(text) in ["nan", "nat", "none"]:
        return "Pendent"
    return text


def obtenir_columna_departament(df):
    for col in df.columns:
        if normalitzar_text(col) == "departament":
            return col
    for col in df.columns:
        if "depart" in normalitzar_text(col):
            return col
    return None


def obtenir_data_actualitzacio_fitxer(path):
    if not os.path.exists(path):
        return "No disponible"
    timestamp = os.path.getmtime(path)
    dt = datetime.fromtimestamp(timestamp, tz=ZoneInfo("Europe/Madrid"))
    return dt.strftime("%d/%m/%Y")


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


def preparar_taula_buida(df):
    df = df.copy().dropna(how="all").dropna(axis=1, how="all").fillna("")
    return df


def llista_valors_no_buits(df, columna):
    if columna not in df.columns:
        return []
    valors = df[columna].astype(str).str.strip().replace("nan", "").replace("NaT", "")
    valors_nets = []
    for valor in valors:
        if valor == "" or normalitzar_text(valor) in ["nan", "nat"]:
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
    return valors[0] if valors else "Pendent"


def trobar_col_resultat_final_porra(df_porra):
    for col in df_porra.columns:
        if col.strip() == "Resultat final":
            return col
    for col in df_porra.columns:
        col_norm = normalitzar_text(col)
        if "resultat" in col_norm and "final" in col_norm and "punt" not in col_norm:
            return col
    return None

# API helpers
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
    if status in ["ft", "full_time", "finished", "final", "after_extra_time", "aet", "ft_pen", "finished_penalties", "penalties"] or phase in ["ft", "full_time", "finished", "final", "after_extra_time", "aet", "ft_pen", "finished_penalties", "penalties"]:
        return True
    home, away = extreure_score(match)
    return home is not None and away is not None and status not in ["scheduled", "pre", "postponed", "cancelled"]


def obtenir_guanyador_match(match):
    home_team = extreure_nom_equip(match.get("home_team"))
    away_team = extreure_nom_equip(match.get("away_team"))
    home_score, away_score = extreure_score(match)
    if home_score is None or away_score is None:
        return ""
    if home_score > away_score:
        return home_team
    if away_score > home_score:
        return away_team
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
            response = requests.get(THESTATSAPI_BASE_URL, headers=headers, params=params, timeout=20)
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
    if len(matches) == 0:
        return [], "No s'han pogut obtenir partits de TheStatsAPI. " + " | ".join(errors)
    return matches, ""


def construir_resultats_des_api(matches):
    group_matches = []
    knockout_matches = []
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
            if equip not in standings[grup]:
                standings[grup][equip] = {"Equip": equip, "PJ": 0, "PG": 0, "PE": 0, "PP": 0, "GF": 0, "GC": 0, "DG": 0, "Pts": 0}
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
        standings[grup][home]["DG"] = standings[grup][home]["GF"] - standings[grup][home]["GC"]
        standings[grup][away]["DG"] = standings[grup][away]["GF"] - standings[grup][away]["GC"]
    group_positions = {}
    tercers = []
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
    if len(tercers) > 0:
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
    return {"source": "API", "group_positions": group_positions, "Setzens": list(dict.fromkeys(setzens)), "Vuitens": list(dict.fromkeys(fases["Vuitens"])), "Quarts": list(dict.fromkeys(fases["Quarts"])), "Semis": list(dict.fromkeys(fases["Semis"])), "Finalistes": list(dict.fromkeys(fases["Finalistes"])), "Campió": list(dict.fromkeys(fases["Campió"])), "MVP": [], "Pichichi": [], "Gols Pichichi": None, "api_error": ""}


def construir_resultats_des_excel(df_resultats):
    df = preparar_taula_buida(df_resultats)
    resultats = {"source": "Excel", "group_positions": {}, "Setzens": [], "Vuitens": [], "Quarts": [], "Semis": [], "Finalistes": [], "Campió": [], "MVP": [], "Pichichi": [], "Gols Pichichi": None, "api_error": ""}
    if all(c in df.columns for c in ["Grup", "Posició", "Equip"]):
        for _, row in df.iterrows():
            grup = str(row.get("Grup", "")).strip()
            pos = str(row.get("Posició", "")).strip()
            equip = normalitzar_equip(row.get("Equip", ""))
            if not grup or not pos or not equip:
                continue
            grup_key = grup.replace("Grup ", "").strip()
            resultats["group_positions"].setdefault(grup_key, {})
            if pos in ["1r", "1", "1º"]:
                resultats["group_positions"][grup_key]["1r"] = equip
            elif pos in ["2n", "2", "2º"]:
                resultats["group_positions"][grup_key]["2n"] = equip
            elif pos in ["3r", "3", "3º"]:
                resultats["group_positions"][grup_key]["3r"] = equip
    for fase in ["Setzens", "Vuitens", "Quarts", "Semis", "Finalistes", "Campió", "MVP"]:
        if fase in df.columns:
            valors = [normalitzar_equip(valor) for valor in df[fase].tolist()]
            resultats[fase] = list(dict.fromkeys([v for v in valors if v]))
    if "Jugador Pichichi" in df.columns and "Gols" in df.columns:
        taula = df[["Jugador Pichichi", "Gols"]].copy()
        taula["Jugador Pichichi"] = taula["Jugador Pichichi"].astype(str).str.strip()
        taula["Gols"] = pd.to_numeric(taula["Gols"], errors="coerce")
        taula = taula[(taula["Jugador Pichichi"] != "") & (~taula["Jugador Pichichi"].str.lower().isin(["nan", "nat", "pendent"])) & (taula["Gols"] >= 1)]
        if not taula.empty:
            taula = taula.sort_values("Gols", ascending=False).reset_index(drop=True)
            resultats["Pichichi"] = [taula.iloc[0]["Jugador Pichichi"]]
            resultats["Gols Pichichi"] = int(taula.iloc[0]["Gols"])
    return resultats


def obtenir_resultats_actuals(df_resultats):
    config_api = obtenir_config_api()
    if config_api["use_api"] and config_api["api_key"]:
        matches, error = carregar_matches_thestatsapi(config_api["api_key"], config_api["competition_id"], config_api["season_id"])
        if matches:
            return construir_resultats_des_api(matches)
        resultats = construir_resultats_des_excel(df_resultats)
        resultats["api_error"] = error
        return resultats
    return construir_resultats_des_excel(df_resultats)

# Punt calc
def punts_grup(pred, real_exacte, reals_altres, equips_classificats, es_tercer=False):
    pred = normalitzar_equip(pred)
    if not pred:
        return 0.0
    pred_norm = normalitzar_text(pred)
    real_exacte_norm = normalitzar_text(real_exacte)
    equips_classificats_norm = {normalitzar_text(x) for x in equips_classificats if x}
    altres_norm = {normalitzar_text(x) for x in reals_altres if x}
    if pred_norm == real_exacte_norm:
        if es_tercer:
            return 2.0 if pred_norm in equips_classificats_norm else 0.5
        return 2.0
    if pred_norm in altres_norm or pred_norm in equips_classificats_norm:
        return 1.0
    return 0.0


def calcular_punts_participant_api(row, resultats):
    punts_1r = punts_2n = punts_3r = 0.0
    group_positions = resultats.get("group_positions", {})
    setzens = resultats.get("Setzens", [])
    for grup in list("ABCDEFGHIJKL"):
        real = group_positions.get(grup, {})
        real_1r, real_2n, real_3r = real.get("1r", ""), real.get("2n", ""), real.get("3r", "")
        punts_1r += punts_grup(row.get(f"Grup {grup} 1r", ""), real_1r, [real_2n, real_3r], setzens, False)
        punts_2n += punts_grup(row.get(f"Grup {grup} 2n", ""), real_2n, [real_1r, real_3r], setzens, False)
        punts_3r += punts_grup(row.get(f"Grup {grup} 3r", ""), real_3r, [real_1r, real_2n], setzens, True)
    def punts_fase(prefix, quantitat, real_list, punts_per_encert=1.0):
        total = 0.0
        real_norm = {normalitzar_text(x) for x in real_list if x}
        for i in range(1, quantitat + 1):
            pred = normalitzar_equip(row.get(f"{prefix}_{i}", ""))
            if pred and normalitzar_text(pred) in real_norm:
                total += punts_per_encert
        return total
    punts_setzens = punts_fase("Setzens", 32, resultats.get("Setzens", []), 1.0)
    punts_vuitens = punts_fase("Vuitens", 16, resultats.get("Vuitens", []), 1.0)
    punts_quarts = punts_fase("Quarts", 8, resultats.get("Quarts", []), 1.0)
    punts_semis = punts_fase("Semis", 4, resultats.get("Semis", []), 1.0)
    punts_finalistes = 0.0
    finalistes_norm = {normalitzar_text(x) for x in resultats.get("Finalistes", []) if x}
    for col in ["Final_1", "Final_2"]:
        pred = normalitzar_equip(row.get(col, ""))
        if pred and normalitzar_text(pred) in finalistes_norm:
            punts_finalistes += 1.0
    campio = normalitzar_equip(row.get("Campió", ""))
    campio_real_norm = {normalitzar_text(x) for x in resultats.get("Campió", []) if x}
    punts_campio = 2.0 if campio and normalitzar_text(campio) in campio_real_norm else 0.0
    punts_mvp = pd.to_numeric(row.get("Punts MVP", 0), errors="coerce")
    punts_pichichi = pd.to_numeric(row.get("Punts Pichichi", 0), errors="coerce")
    punts_resultat_final = pd.to_numeric(row.get("Resultat Final", 0), errors="coerce")
    punts_mvp = 0.0 if pd.isna(punts_mvp) else float(punts_mvp)
    punts_pichichi = 0.0 if pd.isna(punts_pichichi) else float(punts_pichichi)
    punts_resultat_final = 0.0 if pd.isna(punts_resultat_final) else float(punts_resultat_final)
    total = punts_1r + punts_2n + punts_3r + punts_setzens + punts_vuitens + punts_quarts + punts_semis + punts_finalistes + punts_campio + punts_mvp + punts_pichichi + punts_resultat_final
    return {"Punts Grups 1r": round(punts_1r, 1), "Punts Grups 2n": round(punts_2n, 1), "Punts Grups 3r": round(punts_3r, 1), "Punts Setzens": round(punts_setzens, 1), "Punts Vuitens": round(punts_vuitens, 1), "Punts Quarts": round(punts_quarts, 1), "Punts Semis": round(punts_semis, 1), "Punts Finalistes": round(punts_finalistes, 1), "Punts Campió": round(punts_campio, 1), "Punts MVP": round(punts_mvp, 1), "Punts Pichichi": round(punts_pichichi, 1), "Resultat Final": round(punts_resultat_final, 1), "Total Punts": round(total, 1)}


def aplicar_recompte_api_a_porra(df_porra, resultats):
    if resultats.get("source") != "API":
        return df_porra
    df = df_porra.copy()
    df_punts = pd.DataFrame([calcular_punts_participant_api(row, resultats) for _, row in df.iterrows()])
    for col in df_punts.columns:
        df[col] = df_punts[col]
    return df

# Ranking, snapshots and views are appended from a compact, validated block


def recalcular_posicions(df):
    df = df.copy().sort_values("Punts", ascending=False).reset_index(drop=True)
    df["Posició"] = df.index + 1
    df["Dif líder"] = (df["Punts"] - float(df["Punts"].iloc[0])).round(1) if not df.empty else 0
    return df


def crear_ranking_des_de_porra(df_porra):
    if "Participants" not in df_porra.columns or "Total Punts" not in df_porra.columns:
        st.error("Falten columnes obligatòries al full Porra: Participants i/o Total Punts")
        st.write("Columnes detectades:", list(df_porra.columns))
        st.stop()
    col_dep = obtenir_columna_departament(df_porra)
    cols = ["Participants", "Total Punts"] + ([col_dep] if col_dep else [])
    df = df_porra[cols].copy()
    rename_map = {"Participants": "Participant", "Total Punts": "Punts"}
    if col_dep:
        rename_map[col_dep] = "Departament"
    df = df.rename(columns=rename_map)
    df["Participant"] = df["Participant"].astype(str).str.strip()
    df["Punts"] = pd.to_numeric(df["Punts"], errors="coerce")
    if "Departament" in df.columns:
        df["Departament"] = df["Departament"].fillna("Sense departament").astype(str).str.strip().replace("", "Sense departament")
    df = df.dropna(subset=["Punts"])
    df = df[df["Participant"] != ""]
    df = df[~df["Participant"].str.contains("Total", case=False, na=False)]
    df["Punts"] = df["Punts"].round(1)
    return recalcular_posicions(df)


def crear_ranking_departaments(df_ranking):
    if "Departament" not in df_ranking.columns:
        return pd.DataFrame()
    df = df_ranking.copy()
    resum = df.groupby("Departament", as_index=False).agg(Participants=("Participant", "count"), Punts_totals=("Punts", "sum"), Mitjana_punts=("Punts", "mean"), Millor_puntuacio=("Punts", "max"))
    lider = df.sort_values("Punts", ascending=False).drop_duplicates("Departament")[["Departament", "Participant"]].rename(columns={"Participant": "Líder departament"})
    resum = resum.merge(lider, on="Departament", how="left")
    for c in ["Punts_totals", "Mitjana_punts", "Millor_puntuacio"]:
        resum[c] = resum[c].round(1)
    resum = resum.sort_values(["Mitjana_punts", "Punts_totals"], ascending=[False, False]).reset_index(drop=True)
    resum["Posició"] = resum.index + 1
    resum["Dif líder"] = (resum["Mitjana_punts"] - float(resum["Mitjana_punts"].iloc[0])).round(1) if not resum.empty else 0
    return resum[["Posició", "Departament", "Participants", "Mitjana_punts", "Punts_totals", "Millor_puntuacio", "Líder departament", "Dif líder"]]


def carregar_csv_segura(path):
    try:
        return pd.read_csv(path) if os.path.exists(path) else pd.DataFrame()
    except Exception:
        return pd.DataFrame()


def aplicar_moviment(df_ranking, excel_mtime):
    df = df_ranking.copy()
    df["Posició anterior"] = df["Posició"]
    df["Punts anteriors"] = df["Punts"]
    df["Canvi punts"] = 0.0
    df["Canvi num posició"] = 0
    df["Canvi posició"] = "⚪ —"
    return df


def aplicar_moviment_departament(df_dep_actual, df_ranking_global, departament):
    df = df_dep_actual.copy()
    df["Canvi posició"] = "⚪ —"
    df["Canvi punts"] = 0.0
    return df


def carregar_historic():
    return carregar_csv_segura(HISTORY_FILE)


def registrar_historic(df_ranking, excel_mtime, data_actualitzacio):
    df_hist = carregar_historic()
    # Minimal safe history: do not duplicate same mtime
    if not df_hist.empty and "excel_mtime" in df_hist.columns:
        mtimes = pd.to_numeric(df_hist["excel_mtime"], errors="coerce").dropna().astype(float).unique()
        if float(excel_mtime) in mtimes:
            return df_hist
    cols = ["Participant", "Punts", "Posició"] + (["Departament"] if "Departament" in df_ranking.columns else [])
    df_nou = df_ranking[cols].copy()
    df_nou["excel_mtime"] = float(excel_mtime)
    df_nou["Actualització"] = str(data_actualitzacio)
    df_hist = df_nou if df_hist.empty else pd.concat([df_hist, df_nou], ignore_index=True)
    df_hist.to_csv(HISTORY_FILE, index=False)
    return df_hist


def highlight_leader(row):
    return ["background-color: #ffe066; font-weight: bold;" if row.get("Posició") == 1 else "" for _ in row]


def mostrar_taula_ranking(df):
    cols = ["Posició", "Canvi posició", "Participant"]
    if "Departament" in df.columns:
        cols.append("Departament")
    cols += ["Punts", "Dif líder", "Canvi punts"]
    cols = [c for c in cols if c in df.columns]
    df_display = df[cols].copy()
    for c in ["Punts", "Dif líder", "Canvi punts"]:
        if c in df_display.columns:
            df_display[c] = pd.to_numeric(df_display[c], errors="coerce").fillna(0).round(1)
    st.dataframe(df_display.style.apply(highlight_leader, axis=1), use_container_width=True, hide_index=True)


def mostrar_taula_departaments(df_dep):
    if df_dep.empty:
        st.info("Afegeix una columna 'Departament' al full Porra per activar aquest mode.")
        return
    st.dataframe(df_dep, use_container_width=True, hide_index=True)


def mostrar_grafic_punts(df, color="#0b70c9", altura_minima=950):
    if df.empty:
        return
    data = df[["Participant", "Punts", "Posició", "Dif líder"]].copy().sort_values("Punts", ascending=False)
    chart = alt.Chart(data).mark_bar(color=color).encode(x=alt.X("Punts:Q", title="Punts", scale=alt.Scale(zero=False)), y=alt.Y("Participant:N", sort="-x", title=None), tooltip=["Posició", "Participant", alt.Tooltip("Punts:Q", format=".1f"), alt.Tooltip("Dif líder:Q", format=".1f")]).properties(height=max(altura_minima, len(data) * 36))
    st.altair_chart(chart, use_container_width=True)


def mostrar_grafic_departaments(df_dep):
    if df_dep.empty:
        return
    data = df_dep.copy().sort_values("Mitjana_punts", ascending=False)
    chart = alt.Chart(data).mark_bar(color="#0f9d58").encode(x=alt.X("Mitjana_punts:Q", title="Mitjana de punts", scale=alt.Scale(zero=False)), y=alt.Y("Departament:N", sort="-x", title=None), tooltip=list(data.columns)).properties(height=max(350, len(data) * 46))
    st.altair_chart(chart, use_container_width=True)


def mostrar_evolucio_temporal(df_hist, df_ranking):
    st.subheader("📉 Evolució temporal")
    if df_hist.empty or "Actualització" not in df_hist.columns or df_hist["Actualització"].nunique() < 2:
        st.info("Encara no hi ha històric suficient.")
        return
    participants_sel = st.multiselect("Selecciona participants", sorted(df_hist["Participant"].astype(str).unique()), default=df_ranking.head(5)["Participant"].tolist())
    if not participants_sel:
        return
    data = df_hist[df_hist["Participant"].isin(participants_sel)].copy()
    data["Punts"] = pd.to_numeric(data["Punts"], errors="coerce")
    chart = alt.Chart(data).mark_line(point=True).encode(x="Actualització:N", y="Punts:Q", color="Participant:N", tooltip=["Actualització", "Participant", "Punts", "Posició"]).properties(height=420)
    st.altair_chart(chart, use_container_width=True)


def mostrar_animacio_evolucio(df_hist):
    st.subheader("📽️ Animació evolució del rànquing")
    if df_hist.empty or "Actualització" not in df_hist.columns or df_hist["Actualització"].nunique() < 2:
        st.info("Calen almenys dues versions per reproduir l’animació.")
        return


def obtenir_pichichi_real(df_resultats_display, col_pichichi, col_gols):
    if col_pichichi not in df_resultats_display.columns or col_gols not in df_resultats_display.columns:
        return "Pendent", "Pendent"
    taula = df_resultats_display[[col_pichichi, col_gols]].copy()
    taula[col_pichichi] = taula[col_pichichi].astype(str).str.strip()
    taula[col_gols] = pd.to_numeric(taula[col_gols], errors="coerce")
    taula = taula[(taula[col_pichichi] != "") & (~taula[col_pichichi].str.lower().isin(["nan", "nat", "pendent"])) & (taula[col_gols] >= 1)]
    if taula.empty:
        return "Pendent", "Pendent"
    taula = taula.sort_values(col_gols, ascending=False).reset_index(drop=True)
    return taula.iloc[0][col_pichichi], str(int(taula.iloc[0][col_gols]))


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
    phases = [("Setzens", 32), ("Vuitens", 16), ("Quarts", 8), ("Semis", 4)]
    for idx, (prefix, qty) in enumerate(phases):
        with tabs[idx]:
            st.dataframe(obtenir_prediccions_fase(df_j, prefix, qty), use_container_width=True, hide_index=True)
    with tabs[4]:
        finalistes = [afegir_bandera(valor_o_pendent(df_j[col].values[0])) if col in df_j.columns else "Pendent" for col in ["Final_1", "Final_2"]]
        st.dataframe(pd.DataFrame({"Finalista": ["Finalista 1", "Finalista 2"], "Equip": finalistes}), use_container_width=True, hide_index=True)
    with tabs[5]:
        campio = valor_o_pendent(df_j["Campió"].values[0]) if "Campió" in df_j.columns else "Pendent"
        st.dataframe(pd.DataFrame({"Concepte": ["Campió previst"], "Equip": [afegir_bandera(campio)]}), use_container_width=True, hide_index=True)


def mostrar_fase_eliminatoria_responsive(files_eliminatoria):
    if len(files_eliminatoria) == 0:
        st.info("No hi ha dades de fase eliminatòria configurades.")
        return
    html_files = ""
    for fila in files_eliminatoria:
        fase = fila["Fase"]
        resultat = str(fila["Resultat"]).strip()
        if resultat == "" or normalitzar_text(resultat) == "pendent":
            equips_html = "<span class='elim-pending'>Pendent</span>"
        else:
            equips = [x.strip() for x in resultat.split(" · ") if x.strip() != ""]
            equips_html = "".join([f"<span class='elim-badge'>{equip}</span>" for equip in equips])
        html_files += f"<div class='elim-row'><div class='elim-phase'>{fase}</div><div class='elim-teams'>{equips_html}</div></div>"
    st.markdown(f"<div class='elim-wrapper'>{html_files}</div>", unsafe_allow_html=True)

# Styles
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

# Data load
excel_mtime = os.path.getmtime(EXCEL_FILE) if os.path.exists(EXCEL_FILE) else 0
data_actualitzacio = obtenir_data_actualitzacio_fitxer(EXCEL_FILE)
df_porra_original, df_resultats = carregar_dades(EXCEL_FILE, excel_mtime)
resultats_actuals = obtenir_resultats_actuals(df_resultats)
df_porra = aplicar_recompte_api_a_porra(df_porra_original, resultats_actuals)
df_ranking = crear_ranking_des_de_porra(df_porra)
df_ranking = aplicar_moviment(df_ranking, excel_mtime)
df_historic = registrar_historic(df_ranking, excel_mtime, data_actualitzacio)
df_departaments = crear_ranking_departaments(df_ranking)
num_participants = len(df_ranking)
premi_guanyador = num_participants * PREU_PARTICIPACIO
te_departaments = "Departament" in df_ranking.columns and not df_departaments.empty

# Header
st.markdown('<p class="title">🏆 PORRA MUNDIAL</p>', unsafe_allow_html=True)
st.markdown('<p class="subtitle">Classificació en viu, resultats Excel/API, competició per departaments i resultats reals</p>', unsafe_allow_html=True)
if resultats_actuals.get("source") == "API":
    st.success("Mode resultats: API TheStatsAPI activada.")
else:
    st.info("Mode resultats: Excel. L’app utilitza el full 'Resultats Reals' del fitxer Excel.")
if resultats_actuals.get("api_error"):
    st.warning(f"L'API no ha retornat dades vàlides i s'ha fet servir Excel. Detall: {resultats_actuals.get('api_error')}")

# Cards
info1, info2, info3 = st.columns(3, gap="small")
info1.markdown(f"<div class='card darkcard'><h3>🕒 Dades actualitzades</h3><h1>{data_actualitzacio}</h1></div>", unsafe_allow_html=True)
info2.markdown(f"<div class='card greencard'><h3>🎁 Premi guanyador</h3><h1>{premi_guanyador} €</h1><p>{num_participants} participants x {PREU_PARTICIPACIO} €</p></div>", unsafe_allow_html=True)
info3.markdown(f"<div class='card bluecard'><h3>👥 Participants</h3><h1>{num_participants}</h1><p>porres registrades</p></div>", unsafe_allow_html=True)

if te_departaments:
    dept_lider = df_departaments.iloc[0]
    st.markdown(f"<div class='card purplecard'><h3>🏢 Departament líder</h3><h1>{dept_lider['Departament']}</h1><p>Mitjana {float(dept_lider['Mitjana_punts']):.1f} punts · {int(dept_lider['Participants'])} participants</p></div>", unsafe_allow_html=True)

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
mostrar_evolucio_temporal(df_historic, df_ranking)
mostrar_animacio_evolucio(df_historic)

# Participant detail
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

# Department
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

# Custom league
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

# Results
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
elif all(col in df_resultats_display.columns for col in [COL_GRUP, COL_POSICIO, COL_EQUIP]):
    for _, row in df_resultats_display.iterrows():
        grup = str(row.get(COL_GRUP, "")).strip(); posicio = str(row.get(COL_POSICIO, "")).strip(); equip = str(row.get(COL_EQUIP, "")).strip()
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
    if fase in ["Setzens", "Vuitens", "Quarts", "Semis", "Finalistes", "Campió"] and resultats_actuals.get(fase):
        valors = resultats_actuals.get(fase, [])
        detall = "Pendent" if len(valors) == 0 else " · ".join([afegir_bandera(v) for v in valors])
    elif fase in df_resultats_display.columns:
        valors = llista_valors_no_buits(df_resultats_display, fase)
        if len(valors) == 0 or (len(valors) == 1 and normalitzar_text(valors[0]) == "pendent"):
            detall = "Pendent"
        else:
            detall = " · ".join(valors) if fase == COL_MVP else " · ".join([afegir_bandera(v) for v in valors])
    else:
        continue
    files_eliminatoria.append({"Fase": fase, "Resultat": detall})
mostrar_fase_eliminatoria_responsive(files_eliminatoria)

st.write("### ⚽ Jugador pichichi")
if COL_PICHICHI in df_resultats_display.columns and COL_GOLS in df_resultats_display.columns:
    taula_pichichi = df_resultats_display[[COL_PICHICHI, COL_GOLS]].copy()
    taula_pichichi[COL_PICHICHI] = taula_pichichi[COL_PICHICHI].astype(str).str.strip()
    taula_pichichi[COL_GOLS] = pd.to_numeric(taula_pichichi[COL_GOLS], errors="coerce")
    taula_pichichi = taula_pichichi[(taula_pichichi[COL_PICHICHI] != "") & (~taula_pichichi[COL_PICHICHI].str.lower().isin(["nan", "nat", "pendent"])) & (taula_pichichi[COL_GOLS] >= 1)]
    if taula_pichichi.empty:
        taula_pichichi = pd.DataFrame({COL_PICHICHI: ["Pendent"], COL_GOLS: ["Pendent"]})
    else:
        taula_pichichi = taula_pichichi.sort_values(COL_GOLS, ascending=False)
        taula_pichichi[COL_GOLS] = taula_pichichi[COL_GOLS].astype("Int64")
    st.dataframe(taula_pichichi, use_container_width=True, hide_index=True)
else:
    st.dataframe(pd.DataFrame({"Jugador Pichichi": ["Pendent"], "Gols": ["Pendent"]}), use_container_width=True, hide_index=True)

st.write("### 🏁 Resultat de la final")
st.dataframe(pd.DataFrame({"Concepte": ["Resultat de la final"], "Valor": [primer_valor_o_pendent(df_resultats_display, COL_RESULTAT_FINAL)]}), use_container_width=True, hide_index=True)

st.markdown("---")
st.write("📡 Actualització automàtica des de Excel / API TheStatsAPI")
