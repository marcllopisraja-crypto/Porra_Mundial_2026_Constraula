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
st.set_page_config(
    page_title="Porra Mundial 2026",
    layout="wide"
)

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


# ==================================================
# SECRETS / CONFIG API
# ==================================================
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


# ==================================================
# SIDEBAR
# ==================================================
with st.sidebar:
    st.write("### ⚙️ Gestió")

    if st.button("🔄 Reiniciar comparativa de moviments"):
        for fitxer in [
            SNAPSHOT_CURRENT_FILE,
            SNAPSHOT_DISPLAY_FILE,
            SNAPSHOT_META_FILE
        ]:
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
# NORMALITZACIÓ I BANDERES
# ==================================================
FLAGS = {
    "mexic": "🇲🇽",
    "mexico": "🇲🇽",
    "corea del sud": "🇰🇷",
    "south korea": "🇰🇷",
    "korea republic": "🇰🇷",
    "republica txeca": "🇨🇿",
    "czechia": "🇨🇿",
    "suissa": "🇨🇭",
    "switzerland": "🇨🇭",
    "canada": "🇨🇦",
    "qatar": "🇶🇦",
    "escocia": "🏴",
    "scotland": "🏴",
    "marroc": "🇲🇦",
    "morocco": "🇲🇦",
    "brasil": "🇧🇷",
    "brazil": "🇧🇷",
    "estats units": "🇺🇸",
    "united states": "🇺🇸",
    "usa": "🇺🇸",
    "eeuu": "🇺🇸",
    "ee.uu": "🇺🇸",
    "australia": "🇦🇺",
    "turquia": "🇹🇷",
    "turkiye": "🇹🇷",
    "turkey": "🇹🇷",
    "alemanya": "🇩🇪",
    "germany": "🇩🇪",
    "costa d'ivori": "🇨🇮",
    "cote d'ivoire": "🇨🇮",
    "côte d'ivoire": "🇨🇮",
    "equador": "🇪🇨",
    "ecuador": "🇪🇨",
    "suecia": "🇸🇪",
    "sweden": "🇸🇪",
    "japo": "🇯🇵",
    "japan": "🇯🇵",
    "paisos baixos": "🇳🇱",
    "netherlands": "🇳🇱",
    "nova zelanda": "🇳🇿",
    "new zealand": "🇳🇿",
    "iran": "🇮🇷",
    "ir iran": "🇮🇷",
    "belgica": "🇧🇪",
    "belgium": "🇧🇪",
    "uruguai": "🇺🇾",
    "uruguay": "🇺🇾",
    "arabia saudita": "🇸🇦",
    "saudi arabia": "🇸🇦",
    "espanya": "🇪🇸",
    "spain": "🇪🇸",
    "franca": "🇫🇷",
    "france": "🇫🇷",
    "senegal": "🇸🇳",
    "argentina": "🇦🇷",
    "algeria": "🇩🇿",
    "austria": "🇦🇹",
    "portugal": "🇵🇹",
    "rd congo": "🇨🇩",
    "dr congo": "🇨🇩",
    "congo dr": "🇨🇩",
    "uzbekistan": "🇺🇿",
    "anglaterra": "🏴",
    "england": "🏴",
    "croacia": "🇭🇷",
    "croatia": "🇭🇷",
    "ghana": "🇬🇭",
    "egipte": "🇪🇬",
    "egypt": "🇪🇬",
    "noruega": "🇳🇴",
    "norway": "🇳🇴",
    "colombia": "🇨🇴",
    "colòmbia": "🇨🇴",
    "bosnia i hercegovina": "🇧🇦",
    "bosnia and herzegovina": "🇧🇦",
    "bosnia & herzegovina": "🇧🇦",
    "paraguai": "🇵🇾",
    "paraguay": "🇵🇾",
    "tunisia": "🇹🇳",
    "tunísia": "🇹🇳",
    "cap verd": "🇨🇻",
    "cabo verde": "🇨🇻",
    "jordania": "🇯🇴",
    "jordània": "🇯🇴",
    "jordan": "🇯🇴",
    "panama": "🇵🇦",
    "panamà": "🇵🇦",
    "curacao": "🇨🇼",
    "curaçao": "🇨🇼",
    "haiti": "🇭🇹",
    "haití": "🇭🇹",
    "sud-africa": "🇿🇦",
    "sud-àfrica": "🇿🇦",
    "south africa": "🇿🇦"
}

TEAM_NAME_MAP = {
    "mexico": "Mèxic",
    "mexic": "Mèxic",
    "south africa": "Sud-àfrica",
    "sud africa": "Sud-àfrica",
    "sud-africa": "Sud-àfrica",
    "korea republic": "Corea del Sud",
    "south korea": "Corea del Sud",
    "corea del sud": "Corea del Sud",
    "czechia": "República Txeca",
    "republica txeca": "República Txeca",
    "republica checa": "República Txeca",
    "canada": "Canadà",
    "bosnia and herzegovina": "Bosnia i Hercegovina",
    "bosnia & herzegovina": "Bosnia i Hercegovina",
    "bosnia i hercegovina": "Bosnia i Hercegovina",
    "qatar": "Qatar",
    "switzerland": "Suïssa",
    "suissa": "Suïssa",
    "brazil": "Brasil",
    "brasil": "Brasil",
    "morocco": "Marroc",
    "marroc": "Marroc",
    "haiti": "Haití",
    "scotland": "Escòcia",
    "escocia": "Escòcia",
    "united states": "Estats Units",
    "usa": "Estats Units",
    "eeuu": "Estats Units",
    "ee.uu": "Estats Units",
    "estats units": "Estats Units",
    "paraguay": "Paraguai",
    "paraguai": "Paraguai",
    "australia": "Austràlia",
    "turkiye": "Turquia",
    "turkey": "Turquia",
    "turquia": "Turquia",
    "germany": "Alemanya",
    "alemanya": "Alemanya",
    "curacao": "Curaçao",
    "curaçao": "Curaçao",
    "cote d'ivoire": "Costa d'Ivori",
    "côte d'ivoire": "Costa d'Ivori",
    "costa d'ivori": "Costa d'Ivori",
    "ecuador": "Equador",
    "equador": "Equador",
    "netherlands": "Països Baixos",
    "paisos baixos": "Països Baixos",
    "japan": "Japó",
    "japo": "Japó",
    "sweden": "Suècia",
    "suecia": "Suècia",
    "tunisia": "Tunísia",
    "tunisia": "Tunísia",
    "belgium": "Bèlgica",
    "belgica": "Bèlgica",
    "egypt": "Egipte",
    "egipte": "Egipte",
    "ir iran": "Iran",
    "iran": "Iran",
    "new zealand": "Nova Zelanda",
    "nova zelanda": "Nova Zelanda",
    "spain": "Espanya",
    "espanya": "Espanya",
    "cabo verde": "Cap Verd",
    "cap verd": "Cap Verd",
    "saudi arabia": "Aràbia Saudita",
    "arabia saudita": "Aràbia Saudita",
    "uruguay": "Uruguai",
    "uruguai": "Uruguai",
    "france": "França",
    "franca": "França",
    "senegal": "Senegal",
    "iraq": "Iraq",
    "norway": "Noruega",
    "noruega": "Noruega",
    "argentina": "Argentina",
    "algeria": "Algèria",
    "austria": "Àustria",
    "jordan": "Jordània",
    "jordania": "Jordània",
    "portugal": "Portugal",
    "congo dr": "RD Congo",
    "dr congo": "RD Congo",
    "rd congo": "RD Congo",
    "uzbekistan": "Uzbekistan",
    "colombia": "Colòmbia",
    "colòmbia": "Colòmbia",
    "england": "Anglaterra",
    "anglaterra": "Anglaterra",
    "croatia": "Croàcia",
    "croacia": "Croàcia",
    "ghana": "Ghana",
    "panama": "Panamà",
    "panamà": "Panamà"
}


def normalitzar_text(text):
    text = str(text).strip().lower()
    text = text.replace("’", "'").replace("`", "'")
    text = unicodedata.normalize("NFD", text)
    text = "".join(char for char in text if unicodedata.category(char) != "Mn")
    text = " ".join(text.split())
    return text


def normalitzar_equip(valor):
    if pd.isna(valor):
        return ""

    text = str(valor).strip()

    if text == "" or normalitzar_text(text) in ["pendent", "nan", "nat", "none"]:
        return ""

    clau = normalitzar_text(text)

    if clau in TEAM_NAME_MAP:
        return TEAM_NAME_MAP[clau]

    return text


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

    valor_text = str(valor).strip()

    if valor_text == "" or normalitzar_text(valor_text) in ["nan", "nat", "none"]:
        return "Pendent"

    return valor_text


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


# ==================================================
# CÀRREGA EXCEL
# ==================================================
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

    with open(image_path, "rb") as f:
        return base64.b64encode(f.read()).decode()


def preparar_taula_buida(df):
    df = df.copy()
    df = df.dropna(how="all")
    df = df.dropna(axis=1, how="all")
    df = df.fillna("")
    return df


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


def trobar_col_resultat_final_porra(df_porra):
    for col in df_porra.columns:
        if col.strip() == "Resultat final":
            return col

    for col in df_porra.columns:
        col_norm = normalitzar_text(col)
        if "resultat" in col_norm and "final" in col_norm and "punt" not in col_norm:
            return col

    return None


# ==================================================
# API THESTATSAPI
# ==================================================
def obtenir_config_api():
    return {
        "use_api": str(llegir_secret("USE_API_RESULTS", "false")).lower() == "true",
        "api_key": llegir_secret("THESTATSAPI_KEY", ""),
        "competition_id": llegir_secret("THESTATSAPI_COMPETITION_ID", DEFAULT_COMPETITION_ID),
        "season_id": llegir_secret("THESTATSAPI_SEASON_ID", DEFAULT_SEASON_ID),
    }


def llegir_secret(nom, defecte=None):
    try:
        return st.secrets[nom]
    except Exception:
        return defecte


def obtenir_camp_dict(obj, claus, defecte=None):
    if not isinstance(obj, dict):
        return defecte

    for clau in claus:
        if clau in obj:
            return obj.get(clau)

    return defecte


def extreure_nom_equip(team_obj):
    if isinstance(team_obj, dict):
        nom = obtenir_camp_dict(team_obj, ["name", "team_name", "display_name", "short_name"])
        return normalitzar_equip(nom)

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

    home = obtenir_camp_dict(match, ["home_score", "score_home", "home_goals"])
    away = obtenir_camp_dict(match, ["away_score", "score_away", "away_goals"])

    return convertir_a_int(home), convertir_a_int(away)


def es_partit_finalitzat(match):
    status = str(match.get("status", "")).strip().lower()
    phase = str(match.get("phase", "")).strip().lower()

    estats_final = [
        "ft",
        "full_time",
        "finished",
        "final",
        "after_extra_time",
        "aet",
        "ft_pen",
        "finished_penalties",
        "penalties"
    ]

    if status in estats_final or phase in estats_final:
        return True

    home, away = extreure_score(match)

    if home is not None and away is not None and status not in ["scheduled", "pre", "postponed", "cancelled"]:
        return True

    return False


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

    if winner:
        return extreure_nom_equip(winner)

    return ""


def obtenir_grup_match(match):
    group_label = match.get("group_label") or match.get("group") or match.get("group_name")

    if group_label is None:
        return ""

    text = str(group_label).strip()

    if text == "":
        return ""

    text = text.replace("Group", "").replace("Grup", "").strip()
    return text.upper()


def obtenir_stage_match(match):
    stage = (
        match.get("stage_name")
        or match.get("stage")
        or match.get("round")
        or match.get("round_name")
        or match.get("phase_name")
        or ""
    )

    return normalitzar_text(stage)


@st.cache_data(ttl=900, show_spinner=False)
def carregar_matches_thestatsapi(api_key, competition_id, season_id):
    if not api_key:
        return [], "No hi ha THESTATSAPI_KEY configurada."

    headers = {
        "Authorization": f"Bearer {api_key}"
    }

    matches = []
    errors = []

    for page in [1, 2]:
        params = {
            "competition_id": competition_id,
            "season_id": season_id,
            "per_page": 100,
            "page": page
        }

        try:
            response = requests.get(
                THESTATSAPI_BASE_URL,
                headers=headers,
                params=params,
                timeout=20
            )

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
        elif "group" in stage or "regular" in stage:
            if grup:
                group_matches.append(match)
        else:
            knockout_matches.append(match)

    standings = {}

    for match in group_matches:
        if not es_partit_finalitzat(match):
            continue

        grup = obtenir_grup_match(match)

        if not grup:
            continue

        home = extreure_nom_equip(match.get("home_team"))
        away = extreure_nom_equip(match.get("away_team"))
        hs, aw = extreure_score(match)

        if not home or not away or hs is None or aw is None:
            continue

        if grup not in standings:
            standings[grup] = {}

        for equip in [home, away]:
            if equip not in standings[grup]:
                standings[grup][equip] = {
                    "Equip": equip,
                    "PJ": 0,
                    "PG": 0,
                    "PE": 0,
                    "PP": 0,
                    "GF": 0,
                    "GC": 0,
                    "DG": 0,
                    "Pts": 0
                }

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

        taula = taula.sort_values(
            ["Pts", "DG", "GF"],
            ascending=[False, False, False]
        ).reset_index(drop=True)

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
        df_tercers = pd.DataFrame(tercers)
        df_tercers = df_tercers.sort_values(
            ["Pts", "DG", "GF"],
            ascending=[False, False, False]
        ).head(8)

        setzens.extend(df_tercers["Equip"].tolist())

    fases = {
        "Vuitens": [],
        "Quarts": [],
        "Semis": [],
        "Finalistes": [],
        "Campió": []
    }

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
        "source": "API",
        "group_positions": group_positions,
        "Setzens": list(dict.fromkeys(setzens)),
        "Vuitens": list(dict.fromkeys(fases["Vuitens"])),
        "Quarts": list(dict.fromkeys(fases["Quarts"])),
        "Semis": list(dict.fromkeys(fases["Semis"])),
        "Finalistes": list(dict.fromkeys(fases["Finalistes"])),
        "Campió": list(dict.fromkeys(fases["Campió"])),
        "MVP": [],
        "Pichichi": [],
        "Gols Pichichi": None,
        "api_error": ""
    }


def construir_resultats_des_excel(df_resultats):
    df = preparar_taula_buida(df_resultats)

    resultats = {
        "source": "Excel",
        "group_positions": {},
        "Setzens": [],
        "Vuitens": [],
        "Quarts": [],
        "Semis": [],
        "Finalistes": [],
        "Campió": [],
        "MVP": [],
        "Pichichi": [],
        "Gols Pichichi": None,
        "api_error": ""
    }

    if all(c in df.columns for c in ["Grup", "Posició", "Equip"]):
        for _, row in df.iterrows():
            grup = str(row.get("Grup", "")).strip()
            pos = str(row.get("Posició", "")).strip()
            equip = normalitzar_equip(row.get("Equip", ""))

            if not grup or not pos or not equip:
                continue

            grup_key = grup.replace("Grup ", "").strip()

            if grup_key not in resultats["group_positions"]:
                resultats["group_positions"][grup_key] = {}

            if pos in ["1r", "1", "1º"]:
                resultats["group_positions"][grup_key]["1r"] = equip
            elif pos in ["2n", "2", "2º"]:
                resultats["group_positions"][grup_key]["2n"] = equip
            elif pos in ["3r", "3", "3º"]:
                resultats["group_positions"][grup_key]["3r"] = equip

    for fase in ["Setzens", "Vuitens", "Quarts", "Semis", "Finalistes", "Campió", "MVP"]:
        if fase in df.columns:
            valors = []
            for valor in df[fase].tolist():
                equip = normalitzar_equip(valor)
                if equip:
                    valors.append(equip)
            resultats[fase] = list(dict.fromkeys(valors))

    if "Jugador Pichichi" in df.columns and "Gols" in df.columns:
        taula = df[["Jugador Pichichi", "Gols"]].copy()
        taula["Jugador Pichichi"] = taula["Jugador Pichichi"].astype(str).str.strip()
        taula["Gols"] = pd.to_numeric(taula["Gols"], errors="coerce")

        taula = taula[
            (taula["Jugador Pichichi"] != "") &
            (~taula["Jugador Pichichi"].str.lower().isin(["nan", "nat", "pendent"])) &
            (taula["Gols"] >= 1)
        ]

        if not taula.empty:
            taula = taula.sort_values("Gols", ascending=False).reset_index(drop=True)
            resultats["Pichichi"] = [taula.iloc[0]["Jugador Pichichi"]]
            resultats["Gols Pichichi"] = int(taula.iloc[0]["Gols"])

    return resultats


def obtenir_resultats_actuals(df_resultats):
    config_api = obtenir_config_api()

    if config_api["use_api"] and config_api["api_key"]:
        matches, error = carregar_matches_thestatsapi(
            config_api["api_key"],
            config_api["competition_id"],
            config_api["season_id"]
        )

        if matches:
            return construir_resultats_des_api(matches)

        resultats = construir_resultats_des_excel(df_resultats)
        resultats["api_error"] = error
        return resultats

    return construir_resultats_des_excel(df_resultats)


# ==================================================
# CÀLCUL DE PUNTS
# ==================================================
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
            if pred_norm in equips_classificats_norm:
                return 2.0
            return 0.5
        return 2.0

    if pred_norm in altres_norm or pred_norm in equips_classificats_norm:
        return 1.0

    return 0.0


def calcular_punts_participant_api(row, resultats):
    punts_1r = 0.0
    punts_2n = 0.0
    punts_3r = 0.0

    group_positions = resultats.get("group_positions", {})
    setzens = resultats.get("Setzens", [])

    for grup in list("ABCDEFGHIJKL"):
        real = group_positions.get(grup, {})

        real_1r = real.get("1r", "")
        real_2n = real.get("2n", "")
        real_3r = real.get("3r", "")

        pred_1r = row.get(f"Grup {grup} 1r", "")
        pred_2n = row.get(f"Grup {grup} 2n", "")
        pred_3r = row.get(f"Grup {grup} 3r", "")

        punts_1r += punts_grup(pred_1r, real_1r, [real_2n, real_3r], setzens, es_tercer=False)
        punts_2n += punts_grup(pred_2n, real_2n, [real_1r, real_3r], setzens, es_tercer=False)
        punts_3r += punts_grup(pred_3r, real_3r, [real_1r, real_2n], setzens, es_tercer=True)

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

    total = (
        punts_1r +
        punts_2n +
        punts_3r +
        punts_setzens +
        punts_vuitens +
        punts_quarts +
        punts_semis +
        punts_finalistes +
        punts_campio +
        punts_mvp +
        punts_pichichi +
        punts_resultat_final
    )

    return {
        "Punts Grups 1r": round(punts_1r, 1),
        "Punts Grups 2n": round(punts_2n, 1),
        "Punts Grups 3r": round(punts_3r, 1),
        "Punts Setzens": round(punts_setzens, 1),
        "Punts Vuitens": round(punts_vuitens, 1),
        "Punts Quarts": round(punts_quarts, 1),
        "Punts Semis": round(punts_semis, 1),
        "Punts Finalistes": round(punts_finalistes, 1),
        "Punts Campió": round(punts_campio, 1),
        "Punts MVP": round(punts_mvp, 1),
        "Punts Pichichi": round(punts_pichichi, 1),
        "Resultat Final": round(punts_resultat_final, 1),
        "Total Punts": round(total, 1)
    }


def aplicar_recompte_api_a_porra(df_porra, resultats):
    if resultats.get("source") != "API":
        return df_porra

    df = df_porra.copy()
    files_punts = []

    for _, row in df.iterrows():
        files_punts.append(calcular_punts_participant_api(row, resultats))

    df_punts = pd.DataFrame(files_punts)

    for col in df_punts.columns:
        df[col] = df_punts[col]

    return df


# ==================================================
# RÀNQUINGS
# ==================================================
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


# ==================================================
# SNAPSHOTS I MOVIMENTS
# ==================================================
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

    if "Departament" in df_ranking.columns:
        cols.append("Departament")

    df_snapshot = df_ranking[cols].copy()
    df_snapshot = df_snapshot.rename(columns={
        "Punts": "Punts anteriors",
        "Posició": "Posició anterior"
    })

    df_snapshot.to_csv(SNAPSHOT_CURRENT_FILE, index=False)


def guardar_snapshot_display(df_ranking):
    cols = [
        "Participant",
        "Canvi posició",
        "Canvi punts",
        "Canvi num posició",
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
        df["Canvi num posició"] = 0
        df["Canvi posició"] = "⚪ —"
        return df

    meta = carregar_meta_snapshot()
    meta_mtime = meta.get("excel_mtime", None)

    if meta_mtime is not None and float(meta_mtime) == float(excel_mtime):
        df_mov = carregar_csv_segura(SNAPSHOT_DISPLAY_FILE)

        if not df_mov.empty and "Participant" in df_mov.columns:
            df_actual = df_actual.merge(df_mov, on="Participant", how="left")

            df_actual["Canvi posició"] = df_actual["Canvi posició"].fillna("⚪ —")
            df_actual["Canvi punts"] = pd.to_numeric(
                df_actual["Canvi punts"],
                errors="coerce"
            ).fillna(0.0).round(1)

            if "Canvi num posició" not in df_actual.columns:
                df_actual["Canvi num posició"] = 0

            return df_actual

        df_actual = posar_neutral(df_actual)
        guardar_snapshot_actual(df_actual)
        guardar_snapshot_display(df_actual)
        guardar_meta_snapshot(excel_mtime)
        return df_actual

    df_prev = carregar_csv_segura(SNAPSHOT_CURRENT_FILE)

    if df_prev.empty or "Participant" not in df_prev.columns:
        df_actual = posar_neutral(df_actual)
        guardar_snapshot_actual(df_actual)
        guardar_snapshot_display(df_actual)
        guardar_meta_snapshot(excel_mtime)
        return df_actual

    df_prev["Participant"] = df_prev["Participant"].astype(str).str.strip()
    df_actual = df_actual.merge(df_prev, on="Participant", how="left")

    df_actual["Punts anteriors"] = pd.to_numeric(df_actual["Punts anteriors"], errors="coerce")
    df_actual["Posició anterior"] = pd.to_numeric(df_actual["Posició anterior"], errors="coerce")

    df_actual["Canvi punts"] = (
        df_actual["Punts"] - df_actual["Punts anteriors"]
    ).round(1)

    df_actual["Canvi punts"] = pd.to_numeric(
        df_actual["Canvi punts"],
        errors="coerce"
    ).fillna(0.0).round(1)

    df_actual["Canvi num posició"] = (
        df_actual["Posició anterior"] - df_actual["Posició"]
    ).fillna(0).astype(int)

    def indicador_final(row):
        canvi_pos = int(row["Canvi num posició"])

        if canvi_pos > 0:
            return f"🟢 ▲ +{canvi_pos}"
        elif canvi_pos < 0:
            return f"🔴 ▼ {canvi_pos}"
        else:
            return "⚪ —"

    df_actual["Canvi posició"] = df_actual.apply(indicador_final, axis=1)

    guardar_snapshot_actual(df_actual)
    guardar_snapshot_display(df_actual)
    guardar_meta_snapshot(excel_mtime)

    return df_actual


def aplicar_moviment_departament(df_dep_actual, df_ranking_global, departament):
    df_dep = df_dep_actual.copy()

    columnes_a_netejar = [
        "Punts anteriors",
        "Posició anterior",
        "Punts anteriors dep",
        "Posició anterior dep",
        "Canvi punts",
        "Canvi num posició",
        "Canvi posició"
    ]

    for col in columnes_a_netejar:
        if col in df_dep.columns:
            df_dep = df_dep.drop(columns=[col])

    if "Punts anteriors" not in df_ranking_global.columns or "Departament" not in df_ranking_global.columns:
        df_dep["Canvi posició"] = "⚪ —"
        df_dep["Canvi punts"] = 0.0
        return df_dep

    df_prev_dep = df_ranking_global[
        df_ranking_global["Departament"].astype(str).str.strip() == str(departament).strip()
    ].copy()

    if df_prev_dep.empty:
        df_dep["Canvi posició"] = "⚪ —"
        df_dep["Canvi punts"] = 0.0
        return df_dep

    df_prev_dep["Punts anteriors"] = pd.to_numeric(
        df_prev_dep["Punts anteriors"],
        errors="coerce"
    )

    df_prev_dep = df_prev_dep.dropna(subset=["Punts anteriors"])

    if df_prev_dep.empty:
        df_dep["Canvi posició"] = "⚪ —"
        df_dep["Canvi punts"] = 0.0
        return df_dep

    df_prev_dep = df_prev_dep.sort_values("Punts anteriors", ascending=False).reset_index(drop=True)
    df_prev_dep["Posició anterior dep"] = df_prev_dep.index + 1

    df_prev_dep = df_prev_dep[
        ["Participant", "Punts anteriors", "Posició anterior dep"]
    ].copy()

    df_prev_dep = df_prev_dep.rename(columns={
        "Punts anteriors": "Punts anteriors dep"
    })

    df_dep = df_dep.merge(df_prev_dep, on="Participant", how="left")

    df_dep["Punts anteriors dep"] = pd.to_numeric(df_dep["Punts anteriors dep"], errors="coerce")
    df_dep["Posició anterior dep"] = pd.to_numeric(df_dep["Posició anterior dep"], errors="coerce")

    df_dep["Canvi punts"] = (
        df_dep["Punts"] - df_dep["Punts anteriors dep"]
    ).fillna(0.0).round(1)

    df_dep["Canvi num posició"] = (
        df_dep["Posició anterior dep"] - df_dep["Posició"]
    ).fillna(0).astype(int)

    def indicador_dep(row):
        canvi_pos = int(row["Canvi num posició"])

        if canvi_pos > 0:
            return f"🟢 ▲ +{canvi_pos}"
        elif canvi_pos < 0:
            return f"🔴 ▼ {canvi_pos}"
        else:
            return "⚪ —"

    df_dep["Canvi posició"] = df_dep.apply(indicador_dep, axis=1)

    return df_dep


# ==================================================
# HISTÒRIC
# ==================================================
def carregar_historic():
    if not os.path.exists(HISTORY_FILE):
        return pd.DataFrame()

    try:
        return pd.read_csv(HISTORY_FILE)
    except Exception:
        return pd.DataFrame()


def guardar_historic(df_hist):
    df_hist.to_csv(HISTORY_FILE, index=False)


def registrar_historic(df_ranking, excel_mtime, data_actualitzacio):
    df_hist = carregar_historic()

    if not df_hist.empty and "excel_mtime" in df_hist.columns:
        mtimes = pd.to_numeric(
            df_hist["excel_mtime"],
            errors="coerce"
        ).dropna().astype(float).unique()

        if float(excel_mtime) in mtimes:
            return df_hist

    cols = ["Participant", "Punts", "Posició"]

    if "Departament" in df_ranking.columns:
        cols.append("Departament")

    df_nou = df_ranking[cols].copy()
    df_nou["excel_mtime"] = float(excel_mtime)
    df_nou["Actualització"] = str(data_actualitzacio)

    if df_hist.empty:
        df_hist = df_nou
    else:
        df_hist = pd.concat([df_hist, df_nou], ignore_index=True)

    guardar_historic(df_hist)

    return df_hist


# ==================================================
# VISUALITZACIONS
# ==================================================
def highlight_leader(row):
    if row["Posició"] == 1:
        return ["background-color: #ffe066; font-weight: bold;"] * len(row)
    return [""] * len(row)


def mostrar_taula_ranking(df):
    cols = ["Posició"]

    if "Canvi posició" in df.columns:
        cols.append("Canvi posició")

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

    styled = df_display.style.apply(highlight_leader, axis=1).format(format_dict)

    column_config = {
        "Posició": st.column_config.NumberColumn("Posició", format="%d"),
        "Punts": st.column_config.NumberColumn("Punts", format="%.1f"),
        "Dif líder": st.column_config.NumberColumn("Dif líder", format="%.1f"),
    }

    if "Canvi posició" in df_display.columns:
        column_config["Canvi posició"] = st.column_config.TextColumn("Canvi posició")

    if "Canvi punts" in df_display.columns:
        column_config["Canvi punts"] = st.column_config.NumberColumn("Canvi punts", format="%+.1f")

    st.dataframe(
        styled,
        use_container_width=True,
        hide_index=True,
        column_config=column_config
    )


def mostrar_taula_departaments(df_dep):
    if df_dep.empty:
        st.info("Afegeix una columna 'Departament' al full Porra per activar aquest mode.")
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

    st.dataframe(styled, use_container_width=True, hide_index=True)


def mostrar_grafic_punts(df, color="#0b70c9", altura_minima=950):
    chart_data = df[["Posició", "Participant", "Punts", "Dif líder"]].copy()
    chart_data = chart_data.sort_values("Punts", ascending=False)
    chart_height = max(altura_minima, len(chart_data) * 36)

    chart = (
        alt.Chart(chart_data)
        .mark_bar(color=color)
        .encode(
            x=alt.X("Punts:Q", title="Punts", scale=alt.Scale(zero=False)),
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
            x=alt.X("Mitjana_punts:Q", title="Mitjana de punts", scale=alt.Scale(zero=False)),
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


def mostrar_evolucio_temporal(df_hist, df_ranking):
    st.subheader("📉 Evolució temporal")

    if df_hist.empty or "Actualització" not in df_hist.columns:
        st.info("Encara no hi ha històric suficient.")
        return

    if df_hist["Actualització"].nunique() < 2:
        st.info("Encara només hi ha una versió registrada.")
        return

    top_default = df_ranking.head(5)["Participant"].tolist()

    participants_sel = st.multiselect(
        "Selecciona participants per veure l’evolució",
        options=sorted(df_hist["Participant"].dropna().astype(str).unique().tolist()),
        default=top_default
    )

    if not participants_sel:
        return

    df_evo = df_hist[df_hist["Participant"].isin(participants_sel)].copy()
    df_evo["Punts"] = pd.to_numeric(df_evo["Punts"], errors="coerce")
    df_evo["Posició"] = pd.to_numeric(df_evo["Posició"], errors="coerce")
    df_evo["Ordre versió"] = pd.to_numeric(df_evo["excel_mtime"], errors="coerce")
    df_evo = df_evo.sort_values(["Ordre versió", "Participant"])

    ordre_actualitzacions = list(df_evo["Actualització"].drop_duplicates())

    st.write("### 📈 Evolució de mark_line(point=True)
        .encode(
            x=alt.X("Actualització:N", title="Versió", sort=ordre_actualitzacions),
            y=alt.Y("Punts:Q", title="Punts"),
            color=alt.Color("Participant:N", title="Participant"),
            tooltip=[
                alt.Tooltip("Actualització:N", title="Actualització"),
                alt.Tooltip("Participant:N", title="Participant"),
                alt.Tooltip("Punts:Q", title="Punts", format=".1f"),
                alt.Tooltip("Posició:Q", title="Posició")
            ]
        )
        .properties(height=420)
    )

    st.altair_chart(chart_punts, use_container_width=True)

    st.write("### 📉 Evolució de posició")

    chart_pos = (
        alt.Chart(df_evo)
        .mark_line(point=True)
        .encode(
            x=alt.X("Actualització:N", title="Versió", sort=ordre_actualitzacions),
            y=alt.Y("Posició:Q", title="Posició", scale=alt.Scale(reverse=True)),
            color=alt.Color("Participant:N", title="Participant"),
            tooltip=[
                alt.Tooltip("Actualització:N", title="Actualització"),
                alt.Tooltip("Participant:N", title="Participant"),
                alt.Tooltip("Posició:Q", title="Posició"),
                alt.Tooltip("Punts:Q", title="Punts", format=".1f")
            ]
        )
        .properties(height=420)
    )

    st.altair_chart(chart_pos, use_container_width=True)


def mostrar_animacio_evolucio(df_hist):
    st.subheader("📽️ Animació evolució del rànquing")

    if df_hist.empty or "Actualització" not in df_hist.columns:
        st.info("Encara no hi ha històric suficient per animar.")
        return

    if df_hist["Actualització"].nunique() < 2:
        st.info("Calen almenys dues versions per reproduir l’animació.")
        return

    max_pos = st.slider("Nombre de posicions a mostrar", 5, 20, 10, 1)
    velocitat = st.slider("Velocitat de reproducció", 0.3, 2.0, 0.9, 0.1)

    if st.button("▶️ Reproduir evolució"):
        df_anim = df_hist.copy()
        df_anim["excel_mtime"] = pd.to_numeric(df_anim["excel_mtime"], errors="coerce")
        df_anim["Posició"] = pd.to_numeric(df_anim["Posició"], errors="coerce")
        df_anim["Punts"] = pd.to_numeric(df_anim["Punts"], errors="coerce")

        versions = (
            df_anim[["excel_mtime", "Actualització"]]
            .drop_duplicates()
            .sort_values("excel_mtime")
        )

        placeholder = st.empty()

        for _, versio in versions.iterrows():
            data_versio = versio["Actualització"]
            mtime_versio = versio["excel_mtime"]

            frame = (
                df_anim[df_anim["excel_mtime"] == mtime_versio]
                .sort_values("Posició")
                .head(max_pos)
                [["Posició", "Participant", "Punts"]]
                .copy()
            )

            frame["Punts"] = frame["Punts"].round(1)

            with placeholder.container():
                st.write(f"### 🕒 Versió: {data_versio}")
                st.dataframe(frame, use_container_width=True, hide_index=True)

                chart = (
                    alt.Chart(frame)
                    .mark_bar(color="#0b70c9")
                    .encode(
                        x=alt.X("Punts:Q", title="Punts"),
                        y=alt.Y("Participant:N", sort="-x", title=None),
                        tooltip=[
                            alt.Tooltip("Posició:Q", title="Posició"),
                            alt.Tooltip("Participant:N", title="Participant"),
                            alt.Tooltip("Punts:Q", title="Punts", format=".1f")
                        ]
                    )
                    .properties(height=max(300, len(frame) * 32))
                )

                st.altair_chart(chart, use_container_width=True)

            time.sleep(velocitat)


# ==================================================
# RESULTATS REALS / FASES
# ==================================================
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

    tabs = st.tabs(["Setzens", "Vuitens", "Quarts", "Semis", "Final", "Campió"])

    with tabsdf_setzens = obtenir_prediccions_fase(df_j, "Setzens", 32)
        st.dataframe(df_setzens, use_container_width=True, hide_index=True)

    with tabsdf_vuitens = obtenir_prediccions_fase(df_j, "Vuitens", 16)
        st.dataframe(df_vuitens, use_container_width=True, hide_index=True)

    with tabsdf_quarts = obtenir_prediccions_fase(df_j, "Quarts", 8)
        st.dataframe(df_quarts, use_container_width=True, hide_index=True)

    with tabsdf_semis = obtenir_prediccions_fase(df_j, "Semis", 4)
        st.dataframe(df_semis, use_container_width=True, hide_index=True)

    with tabsfinalistes = []

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

    with tabs[5]:
        campio = valor_o_pendent(df_j["Campió"].values[0]) if "Campió" in df_j.columns else "Pendent"

        df_campio = pd.DataFrame({
            "Concepte": ["Campió previst"],
            "Equip": [afegir_bandera(campio)]
        })

        st.dataframe(df_campio, use_container_width=True, hide_index=True)


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
            equips_html = "".join([
                f"<span class='elim-badge'>{equip}</span>"
                for equip in equips
            ])

        html_files += f"""
        <div class="elim-row">
            <div class="elim-phase">{fase}</div>
            <div class="elim-teams">{equips_html}</div>
        </div>
        """

    st.markdown(
        f"""
        <div class="elim-wrapper">
            {html_files}
        </div>
        """,
        unsafe_allow_html=True
    )


# ==================================================
# ESTILS
# ==================================================
img_base64 = carregar_imatge_base64(BACKGROUND_IMAGE)

if img_base64:
    background_css = f"""
    background-image:
        linear-gradient(rgba(0,0,0,0.35), rgba(0,0,0,0.55)),
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
        background: rgba(255,255,255,0.88);
        backdrop-filter: blur(6px);
        border-radius: 26px;
        margin-top: 24px;
        margin-bottom: 24px;
        box-shadow: 0px 10px 35px rgba(0,0,0,0.32);
    }}

    h1, h2, h3, h4, h5, h6 {{
        color: #102a43;
        text-shadow: 0px 1px 2px rgba(255,255,255,0.6);
    }}

    .title {{
        font-size: clamp(34px, 5vw, 58px);
        font-weight: 900;
        margin-bottom: 0px;
        color: #102a43;
        letter-spacing: -1px;
        text-shadow: 0px 2px 8px rgba(255,255,255,0.8);
    }}

    .subtitle {{
        font-size: clamp(14px, 2vw, 19px);
        color: #334e68;
        margin-top: 0px;
        margin-bottom: 25px;
        text-shadow: 0px 1px 4px rgba(255,255,255,0.7);
    }}

    .card {{
        padding: 20px;
        border-radius: 20px;
        text-align: center;
        box-shadow: 0px 6px 24px rgba(0,0,0,0.24);
        height: 182px;
        min-height: 182px;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        box-sizing: border-box;
        overflow: hidden;
        width: 100%;
        transform: translateY(0);
        transition: all 0.25s ease;
    }}

    .card:hover {{
        transform: translateY(-4px);
        box-shadow: 0px 10px 30px rgba(0,0,0,0.32);
    }}

    .card h3 {{
        margin: 0px 0px 14px 0px;
        font-size: clamp(15px, 2vw, 24px);
        line-height: 1.15;
        max-width: 100%;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        text-shadow: 0px 2px 6px rgba(0,0,0,0.35);
    }}

    .card h1 {{
        margin: 0px;
        font-size: clamp(26px, 4vw, 42px);
        line-height: 1.1;
        max-width: 100%;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        text-shadow: 0px 2px 6px rgba(0,0,0,0.35);
    }}

    .card p {{
        margin: 12px 0px 0px 0px;
        font-size: clamp(11px, 1.5vw, 15px);
        max-width: 100%;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        text-shadow: 0px 2px 6px rgba(0,0,0,0.35);
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

    .stMetric {{
        background: rgba(255,255,255,0.72);
        padding: 12px;
        border-radius: 14px;
        box-shadow: 0px 4px 16px rgba(0,0,0,0.12);
    }}

    .elim-wrapper {{
        width: 100%;
        display: flex;
        flex-direction: column;
        gap: 12px;
        margin-top: 12px;
        margin-bottom: 18px;
    }}

    .elim-row {{
        display: grid;
        grid-template-columns: 150px minmax(0, 1fr);
        gap: 12px;
        align-items: flex-start;
        background: rgba(255,255,255,0.78);
        border: 1px solid rgba(0,0,0,0.08);
        border-radius: 16px;
        padding: 14px;
        box-shadow: 0px 4px 16px rgba(0,0,0,0.10);
    }}

    .elim-phase {{
        font-weight: 800;
        color: #102a43;
        font-size: 15px;
        padding-top: 6px;
    }}

    .elim-teams {{
        display: flex;
        flex-wrap: wrap;
        gap: 7px;
        line-height: 1.35;
        min-width: 0;
    }}

    .elim-badge {{
        display: inline-block;
        background: #edf2f7;
        color: #102a43;
        border: 1px solid #d9e2ec;
        border-radius: 999px;
        padding: 6px 10px;
        font-size: 13px;
        font-weight: 600;
        white-space: normal;
        max-width: 100%;
        box-shadow: 0px 2px 6px rgba(0,0,0,0.06);
    }}

    .elim-pending {{
        display: inline-block;
        background: #fff3cd;
        color: #665200;
        border: 1px solid #ffe08a;
        border-radius: 999px;
        padding: 6px 10px;
        font-size: 13px;
        font-weight: 700;
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

        .card h3,
        .card h1,
        .card p {{
            white-space: normal;
        }}

        .elim-row {{
            grid-template-columns: 1fr;
            padding: 12px;
        }}

        .elim-phase {{
            font-size: 16px;
            padding-top: 0;
            border-bottom: 1px solid rgba(0,0,0,0.08);
            padding-bottom: 8px;
        }}

        .elim-badge {{
            font-size: 12px;
            padding: 5px 8px;
        }}
    }}
    </style>
    """,
    unsafe_allow_html=True
)


# ==================================================
# CARREGAR DADES
# ==================================================
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


# ==================================================
# TÍTOL
# ==================================================
st.markdown('<p class="title">🏆 PORRA MUNDIAL</p>', unsafe_allow_html=True)
st.markdown(
    '<p class="subtitle">Classificació en viu, resultats Excel/API, moviment entre versions, competició per departaments i resultats reals</p>',
    unsafe_allow_html=True
)

if resultats_actuals.get("source") == "API":
    st.success("Mode resultats: API TheStatsAPI activada.")
else:
    st.info("Mode resultats: Excel. L’app utilitza el full 'Resultats Reals' del fitxer Excel.")

if resultats_actuals.get("api_error"):
    st.warning(f"L'API no ha retornat dades vàlides i s'ha fet servir Excel. Detall: {resultats_actuals.get('api_error')}")


# ==================================================
# INFO PRINCIPAL
# ==================================================
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


# ==================================================
# DEPARTAMENT LÍDER
# ==================================================
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


# ==================================================
# TOP 3 GENERAL
# ==================================================
st.subheader("🥇 TOP 3 General")

top3 = df_ranking.head(3)
top_cols = st.columns(3, gap="small")
top_classes = ["gold", "silver", "bronze"]
top_medals = ["🥇", "🥈", "🥉"]

for i in range(min(3, len(top3))):
    row = top3.iloc[i]
    subtext = row["Departament"] if "Departament" in row.index else "punts"
    canvi_pos = row["Canvi posició"] if "Canvi posició" in row.index else ""

    top_cols[i].markdown(
        f"""
        <div class='card {top_classes[i]}'>
            <h3>{top_medals[i]} {row["Participant"]}</h3>
            <h1>{float(row["Punts"]):.1f}</h1>
            <p>{subtext} · {canvi_pos}</p>
        </div>
        """,
        unsafe_allow_html=True
    )


# ==================================================
# CLASSIFICACIÓ GENERAL
# ==================================================
st.subheader("📊 Classificació general")
mostrar_taula_ranking(df_ranking)


# ==================================================
# GRÀFIC GENERAL
# ==================================================
st.subheader("📈 Gràfic general de punts")
mostrar_grafic_punts(df_ranking, color="#0b70c9", altura_minima=1000)


# ==================================================
# EVOLUCIÓ
# ==================================================
mostrar_evolucio_temporal(df_historic, df_ranking)
mostrar_animacio_evolucio(df_historic)


# ==================================================
# FITXA PARTICIPANT
# ==================================================
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

        punts_dict = {}

        for label, col in [
            ("1rs grup", "Punts Grups 1r"),
            ("2ns grup", "Punts Grups 2n"),
            ("3rs grup", "Punts Grups 3r"),
            ("Setzens", "Punts Setzens"),
            ("Vuitens", "Punts Vuitens"),
            ("Quarts", "Punts Quarts"),
            ("Semis", "Punts Semis"),
            ("Finalistes", "Punts Finalistes"),
            ("Campió", "Punts Campió"),
            ("MVP", "Punts MVP"),
            ("Pichichi", "Punts Pichichi"),
        ]:
            if col in df_j.columns:
                punts_dict[label] = pd.to_numeric(df_j[col].values[0], errors="coerce")

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

        campio = valor_o_pendent(df_j["Campió"].values[0]) if "Campió" in df_j.columns else "Pendent"
        mvp = valor_o_pendent(df_j["MVP"].values[0]) if "MVP" in df_j.columns else "Pendent"
        pichichi = valor_o_pendent(df_j["Pichichi"].values[0]) if "Pichichi" in df_j.columns else "Pendent"

        c2.write(f"🏆 Campió: {afegir_bandera(campio)}")
        c2.write(f"📌 Resultat final: {resultat_final}")
        c2.write(f"⭐ MVP: {mvp}")
        c2.write(f"⚽ Pichichi: {pichichi}")

        mostrar_prediccions_eliminatoria_participant(df_j)

else:
    st.info("Selecciona un participant per veure el detall de punts i prediccions.")


# ==================================================
# COMPETICIÓ PER DEPARTAMENTS
# ==================================================
st.subheader("🏢 Competició per departaments")

if te_departaments:
    st.write("Rànquing calculat per **mitjana de punts** del departament.")

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

        df_dep_individual = aplicar_moviment_departament(
            df_dep_individual,
            df_ranking,
            departament_sel
        )

        st.write(f"### 🥇 TOP 3 · {departament_sel}")

        dep_top = df_dep_individual.head(3)
        dep_cols = st.columns(3, gap="small")
        medalles = ["🥇", "🥈", "🥉"]
        classes = ["gold", "silver", "bronze"]

        for i in range(min(3, len(dep_top))):
            canvi_dep = dep_top.iloc[i]["Canvi posició"] if "Canvi posició" in dep_top.columns else ""

            dep_cols[i].markdown(
                f"""
                <div class='card {classes[i]}'>
                    <h3>{medalles[i]} {dep_top.iloc[i]["Participant"]}</h3>
                    <h1>{float(dep_top.iloc[i]["Punts"]):.1f}</h1>
                    <p>{departament_sel} · {canvi_dep}</p>
                </div>
                """,
                unsafe_allow_html=True
            )

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


# ==================================================
# RESULTATS REALS
# ==================================================
st.subheader("✅ Resultats reals")

df_resultats_display = preparar_taula_buida(df_resultats)

COL_GRUP = "Grup"
COL_POSICIO = "Posició"
COL_EQUIP = "Equip"

COL_SETZENS = "Setzens"
COL_VUITENS = "Vuitens"
COL_QUARTS = "Quarts"
COL_SEMIS = "Semis"
COL_FINALISTES = "Finalistes"
COL_CAMPIO = "Campió"
COL_MVP = "MVP"
COL_RESULTAT_FINAL = "Resultat Final"
COL_PICHICHI = "Jugador Pichichi"
COL_GOLS = "Gols"

if resultats_actuals.get("Campió"):
    campio_real = afegir_bandera(resultats_actuals.get("Campió")[0])
else:
    campio_real = afegir_bandera(primer_valor_o_pendent(df_resultats_display, COL_CAMPIO))

mvp_real = primer_valor_o_pendent(df_resultats_display, COL_MVP)
resultat_final_real = primer_valor_o_pendent(df_resultats_display, COL_RESULTAT_FINAL)

if resultats_actuals.get("Pichichi"):
    pichichi_real = resultats_actuals.get("Pichichi")[0]
    gols_pichichi = resultats_actuals.get("Gols Pichichi", "Pendent")
else:
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


# ==================================================
# FASE DE GRUPS
# ==================================================
st.write("### 🧩 Fase de grups")

grups = {}
group_positions = resultats_actuals.get("group_positions", {})

if group_positions:
    for grup, posicions in group_positions.items():
        grup_label = f"Grup {grup}"
        grups[grup_label] = {
            "1r": afegir_bandera(posicions.get("1r", "")) if posicions.get("1r") else "",
            "2n": afegir_bandera(posicions.get("2n", "")) if posicions.get("2n") else "",
            "3r": afegir_bandera(posicions.get("3r", "")) if posicions.get("3r") else ""
        }

elif all(col in df_resultats_display.columns for col in [COL_GRUP, COL_POSICIO, COL_EQUIP]):
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

    st.dataframe(df_grups, use_container_width=True, hide_index=True)
else:
    st.info("No hi ha dades de fase de grups configurades.")


# ==================================================
# FASE ELIMINATÒRIA RESPONSIVE
# ==================================================
st.write("### ⚔️ Fase eliminatòria")

fases_eliminatoria = [
    COL_SETZENS,
    COL_VUITENS,
    COL_QUARTS,
    COL_SEMIS,
    COL_FINALISTES,
    COL_CAMPIO,
    COL_MVP
]

files_eliminatoria = []

for fase in fases_eliminatoria:
    if fase in ["Setzens", "Vuitens", "Quarts", "Semis", "Finalistes", "Campió"] and resultats_actuals.get(fase):
        valors = resultats_actuals.get(fase, [])

        if len(valors) == 0:
            detall = "Pendent"
        else:
            detall = " · ".join([afegir_bandera(v) for v in valors])

        files_eliminatoria.append({
            "Fase": fase,
            "Resultat": detall
        })

    elif fase in df_resultats_display.columns:
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

mostrar_fase_eliminatoria_responsive(files_eliminatoria)


# ==================================================
# PICHICHI
# ==================================================
st.write("### ⚽ Jugador pichichi")

if COL_PICHICHI in df_resultats_display.columns and COL_GOLS in df_resultats_display.columns:
    taula_pichichi = df_resultats_display[[COL_PICHICHI, COL_GOLS]].copy()

    taula_pichichi[COL_PICHICHI] = taula_pichichi[COL_PICHICHI].astype(str).str.strip()
    taula_pichichi[COL_GOLS] = pd.to_numeric(taula_pichichi[COL_GOLS], errors="coerce")

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

    st.dataframe(taula_pichichi, use_container_width=True, hide_index=True)
else:
    taula_pichichi = pd.DataFrame({
        "Jugador Pichichi": ["Pendent"],
        "Gols": ["Pendent"]
    })

    st.dataframe(taula_pichichi, use_container_width=True, hide_index=True)


# ==================================================
# RESULTAT FINAL
# ==================================================
st.write("### 🏁 Resultat de la final")

resultat_final = primer_valor_o_pendent(
    df_resultats_display,
    COL_RESULTAT_FINAL
)

taula_final = pd.DataFrame({
    "Concepte": ["Resultat de la final"],
    "Valor": [resultat_final]
})

st.dataframe(taula_final, use_container_width=True, hide_index=True)


# ==================================================
# FOOTER
# ==================================================
st.markdown("---")
st.write("📡 Actualització automàtica des de Excel / API TheStatsAPI")
