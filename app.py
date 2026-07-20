import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import altair as alt
import base64
import os
import json
import unicodedata
import re
import html as html_lib
from datetime import datetime
from zoneinfo import ZoneInfo

# --------------------------------------------------
# CONFIGURACIÓ GENERAL
# --------------------------------------------------
st.set_page_config(
    page_title="Porra Mundial",
    layout="wide"
)
# V11.4.12 FANALET VERMELL BUILD

EXCEL_FILE = "Porra_Mundial_Final_Definitiva.xlsx"
BACKGROUND_IMAGE = "fifa-Trionda.jpg"
LOGO_IMAGE = "Logo RGB fondo transparente letra negra Constraula.png"
PREU_PARTICIPACIO = 5

# --------------------------------------------------
# V11.4.6 · ANIMACIÓ INICIAL DE CELEBRACIÓ
# --------------------------------------------------
def mostrar_animacio_celebracio_inicial():
    """Mostra confeti i focs artificials una sola vegada per sessió."""
    if st.session_state.get("celebracio_inicial_v1146_mostrada", False):
        return
    st.session_state["celebracio_inicial_v1146_mostrada"] = True

    # Efecte nadiu de Streamlit
    st.balloons()

    confeti = "".join([
        f"<span class='confetti-piece c{i}'></span>" for i in range(1, 25)
    ])
    fireworks = "".join([
        f"<span class='firework f{i}'></span>" for i in range(1, 8)
    ])

    st.markdown(f"""
    <style>
    .celebration-overlay {{
        position: fixed;
        inset: 0;
        pointer-events: none;
        z-index: 999999;
        overflow: hidden;
        animation: celebrationFadeOut 8.5s ease forwards;
    }}
    .celebration-title {{
        position: absolute;
        top: 7%;
        left: 50%;
        transform: translateX(-50%);
        padding: 14px 24px;
        border-radius: 999px;
        background: rgba(8, 52, 92, 0.78);
        border: 1px solid rgba(255,255,255,0.30);
        color: white;
        font-weight: 1000;
        font-size: clamp(24px, 4vw, 48px);
        text-shadow: 0 4px 18px rgba(0,0,0,0.45);
        box-shadow: 0 16px 40px rgba(0,0,0,0.30);
        backdrop-filter: blur(8px);
        animation: titlePop 1.2s ease both;
        white-space: nowrap;
    }}
    .confetti-piece {{
        position: absolute;
        top: -12vh;
        width: 12px;
        height: 22px;
        border-radius: 4px;
        opacity: 0.95;
        animation-name: confettiFall;
        animation-timing-function: linear;
        animation-fill-mode: forwards;
    }}
    .firework {{
        position: absolute;
        width: 9px;
        height: 9px;
        border-radius: 50%;
        background: #fff1a8;
        box-shadow:
            0 0 0 0 rgba(255, 241, 168, 0.90),
            25px 0 0 0 #ffd166,
            -25px 0 0 0 #06d6a0,
            0 25px 0 0 #4dabf7,
            0 -25px 0 0 #ff6b6b,
            18px 18px 0 0 #f72585,
            -18px 18px 0 0 #b8f2e6,
            18px -18px 0 0 #ffd60a,
            -18px -18px 0 0 #90dbf4;
        opacity: 0;
        animation: fireworkExplode 2.1s ease-out infinite;
    }}
    @keyframes confettiFall {{
        0% {{ transform: translateY(-12vh) rotate(0deg); opacity: 1; }}
        82% {{ opacity: 1; }}
        100% {{ transform: translateY(112vh) rotate(720deg); opacity: 0; }}
    }}
    @keyframes fireworkExplode {{
        0% {{ transform: scale(0.2); opacity: 0; box-shadow: 0 0 0 0 rgba(255,255,255,0.90); }}
        18% {{ opacity: 1; }}
        55% {{ transform: scale(1.25); opacity: 1; }}
        100% {{ transform: scale(1.95); opacity: 0; }}
    }}
    @keyframes titlePop {{
        0% {{ opacity: 0; transform: translateX(-50%) scale(0.84); }}
        100% {{ opacity: 1; transform: translateX(-50%) scale(1); }}
    }}
    @keyframes celebrationFadeOut {{
        0%, 78% {{ opacity: 1; }}
        100% {{ opacity: 0; visibility: hidden; }}
    }}
    .c1 {{ left: 4%; background:#ffd166; animation-duration: 6.8s; animation-delay: .1s; }}
    .c2 {{ left: 8%; background:#ef476f; animation-duration: 7.6s; animation-delay: .6s; }}
    .c3 {{ left: 13%; background:#06d6a0; animation-duration: 6.9s; animation-delay: .2s; }}
    .c4 {{ left: 18%; background:#118ab2; animation-duration: 8.1s; animation-delay: .9s; }}
    .c5 {{ left: 23%; background:#ffd60a; animation-duration: 7.2s; animation-delay: .4s; }}
    .c6 {{ left: 27%; background:#f72585; animation-duration: 7.9s; animation-delay: .0s; }}
    .c7 {{ left: 31%; background:#4dabf7; animation-duration: 6.7s; animation-delay: .7s; }}
    .c8 {{ left: 36%; background:#b8f2e6; animation-duration: 7.5s; animation-delay: .3s; }}
    .c9 {{ left: 41%; background:#ff9f1c; animation-duration: 8.4s; animation-delay: .5s; }}
    .c10 {{ left: 46%; background:#e9ff70; animation-duration: 6.6s; animation-delay: .8s; }}
    .c11 {{ left: 51%; background:#ff006e; animation-duration: 8.0s; animation-delay: .2s; }}
    .c12 {{ left: 56%; background:#8338ec; animation-duration: 7.0s; animation-delay: .6s; }}
    .c13 {{ left: 61%; background:#3a86ff; animation-duration: 8.2s; animation-delay: .1s; }}
    .c14 {{ left: 66%; background:#fb5607; animation-duration: 7.3s; animation-delay: .9s; }}
    .c15 {{ left: 71%; background:#80ed99; animation-duration: 6.8s; animation-delay: .3s; }}
    .c16 {{ left: 76%; background:#ffd166; animation-duration: 7.8s; animation-delay: .7s; }}
    .c17 {{ left: 81%; background:#ef476f; animation-duration: 6.9s; animation-delay: .4s; }}
    .c18 {{ left: 85%; background:#06d6a0; animation-duration: 8.3s; animation-delay: .1s; }}
    .c19 {{ left: 89%; background:#4dabf7; animation-duration: 7.4s; animation-delay: .8s; }}
    .c20 {{ left: 93%; background:#ffd60a; animation-duration: 8.0s; animation-delay: .5s; }}
    .c21 {{ left: 16%; background:#ff595e; animation-duration: 6.5s; animation-delay: 1.1s; }}
    .c22 {{ left: 44%; background:#8ac926; animation-duration: 7.1s; animation-delay: 1.0s; }}
    .c23 {{ left: 68%; background:#1982c4; animation-duration: 6.7s; animation-delay: 1.2s; }}
    .c24 {{ left: 96%; background:#ffca3a; animation-duration: 7.6s; animation-delay: 1.1s; }}
    .f1 {{ left: 12%; top: 18%; animation-delay: .1s; }}
    .f2 {{ left: 78%; top: 20%; animation-delay: .45s; }}
    .f3 {{ left: 31%; top: 33%; animation-delay: .9s; }}
    .f4 {{ left: 61%; top: 39%; animation-delay: 1.25s; }}
    .f5 {{ left: 49%; top: 18%; animation-delay: 1.65s; }}
    .f6 {{ left: 88%; top: 48%; animation-delay: 2.05s; }}
    .f7 {{ left: 18%; top: 52%; animation-delay: 2.45s; }}
    </style>
    <div class="celebration-overlay">
        <div class="celebration-title">🎆 Celebració final de la Porra Mundial 🎊</div>
        {confeti}
        {fireworks}
    </div>
    """, unsafe_allow_html=True)


SNAPSHOT_CURRENT_FILE = "ranking_snapshot_current.csv"
SNAPSHOT_DISPLAY_FILE = "ranking_snapshot_display.csv"
SNAPSHOT_META_FILE = "ranking_snapshot_meta.json"


# --------------------------------------------------
# V11.4.7 · CERIMÒNIA CELEBRATION DELUXE
# --------------------------------------------------
def mostrar_ceremonia_celebration_deluxe(df_ranking, num_participants):
    """V11.4.8: cerimònia visible dins un component HTML propi, amb confeti i focs artificials reals."""
    if df_ranking.empty:
        return

    if not st.session_state.get("celebracio_deluxe_v1148_mostrada", False):
        st.session_state["celebracio_deluxe_v1148_mostrada"] = True
        st.balloons()

    df_final = df_ranking.copy().sort_values("Punts", ascending=False).reset_index(drop=True)
    guanyador = df_final.iloc[0]
    segon = df_final.iloc[1] if len(df_final) > 1 else None
    tercer = df_final.iloc[2] if len(df_final) > 2 else None

    def esc(valor):
        return html_lib.escape(str(valor))

    top_cards = []
    for grup in obtenir_grups_podi(df_final, nom_col="Participant", punts_col="Punts", pos_col="Posició", max_pos=3):
        medalla = medalla_per_posicio(grup["posicio"])
        noms = esc(format_noms_podi(grup["noms"]))
        top_cards.append(
            "<div class='podium-card'>"
            f"<div class='medal'>{medalla}</div>"
            f"<div class='podium-name'>{noms}</div>"
            f"<div class='podium-points'>{float(grup['punts']):.1f} punts</div>"
            "</div>"
        )
    top_html = "".join(top_cards)

    winner = esc(guanyador["Participant"])
    points = f"{float(guanyador['Punts']):.1f}"
    participants = str(num_participants)

    html_template = """
<!DOCTYPE html>
<html>
<head>
<meta charset='utf-8'>
<style>
html, body { margin: 0; padding: 0; background: transparent; font-family: Inter, Montserrat, Arial, sans-serif; overflow: hidden; }
.stage {
  position: relative;
  width: 100%;
  height: 700px;
  overflow: hidden;
  border-radius: 32px;
  color: #fff;
  background:
    radial-gradient(circle at 12% 12%, rgba(255, 214, 10, .45), transparent 24%),
    radial-gradient(circle at 86% 16%, rgba(247, 37, 133, .35), transparent 26%),
    radial-gradient(circle at 50% 92%, rgba(77, 171, 247, .32), transparent 30%),
    linear-gradient(135deg, rgba(4, 18, 34, .99), rgba(7, 56, 102, .97) 48%, rgba(15, 113, 201, .92));
  border: 1px solid rgba(255,255,255,.24);
  box-shadow: 0 22px 60px rgba(0,0,0,.38);
}
.stage:before {
  content: '';
  position: absolute;
  inset: 0;
  background: linear-gradient(120deg, transparent, rgba(255,255,255,.12), transparent);
  animation: shine 4.6s ease-in-out infinite;
}
canvas { position: absolute; inset: 0; width: 100%; height: 100%; z-index: 3; pointer-events: none; }
.content { position: relative; z-index: 4; height: 100%; display: flex; flex-direction: column; align-items: center; justify-content: center; text-align: center; padding: 28px; box-sizing: border-box; }
.kicker { color: #fff1a8; text-transform: uppercase; font-size: clamp(15px, 2vw, 22px); font-weight: 1000; letter-spacing: 5px; }
.title { margin-top: 10px; font-weight: 1000; letter-spacing: -2px; line-height: .92; font-size: clamp(42px, 7vw, 96px); text-shadow: 0 8px 30px rgba(0,0,0,.50); animation: titlePop 1.1s cubic-bezier(.2,.9,.2,1) both; }
.winner { margin-top: 28px; color: #fff1a8; font-weight: 1000; line-height: .98; font-size: clamp(36px, 6vw, 80px); text-shadow: 0 8px 30px rgba(0,0,0,.45); animation: pulse 2.1s ease-in-out infinite; }
.points { margin-top: 12px; font-size: clamp(26px, 3.4vw, 44px); font-weight: 1000; }
.ribbon { margin: 22px auto 12px auto; color: #fff1a8; font-size: clamp(20px, 3vw, 34px); font-weight: 1000; letter-spacing: 4px; padding: 12px 20px; border-top: 2px solid rgba(255,241,168,.65); border-bottom: 2px solid rgba(255,241,168,.65); }
.meta { padding: 11px 18px; border-radius: 999px; background: rgba(255,255,255,.13); border: 1px solid rgba(255,255,255,.22); font-weight: 850; color: rgba(255,255,255,.9); }
.podium { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 14px; width: min(920px, 100%); margin-top: 22px; }
.podium-card { border-radius: 22px; padding: 16px; background: rgba(255,255,255,.14); border: 1px solid rgba(255,255,255,.24); backdrop-filter: blur(8px); }
.medal { font-size: 36px; }
.podium-name { margin-top: 6px; font-size: 20px; font-weight: 950; line-height: 1.1; }
.podium-points { margin-top: 5px; color: #fff1a8; font-weight: 900; }
.firework { position: absolute; width: 8px; height: 8px; border-radius: 50%; background: #fff1a8; opacity: 0; z-index: 2; box-shadow: 0 -34px #ffd60a, 0 34px #4dabf7, 34px 0 #ff6b6b, -34px 0 #06d6a0, 24px 24px #f72585, -24px 24px #b8f2e6, 24px -24px #ff9f1c, -24px -24px #90dbf4; animation: explode 2.4s ease-out infinite; }
.f1 { left: 10%; top: 16%; animation-delay: .1s; } .f2 { left: 80%; top: 18%; animation-delay: .45s; } .f3 { left: 25%; top: 42%; animation-delay: .9s; } .f4 { left: 65%; top: 44%; animation-delay: 1.25s; }
.f5 { left: 46%; top: 13%; animation-delay: 1.65s; } .f6 { left: 88%; top: 56%; animation-delay: 2.05s; } .f7 { left: 16%; top: 58%; animation-delay: 2.45s; } .f8 { left: 54%; top: 60%; animation-delay: 2.85s; }
@keyframes explode { 0% { transform: scale(.25); opacity: 0; } 18% { opacity: 1; } 55% { transform: scale(1.3); opacity: 1; } 100% { transform: scale(2.1); opacity: 0; } }
@keyframes pulse { 0%,100% { transform: scale(1); } 50% { transform: scale(1.025); } }
@keyframes titlePop { 0% { opacity:0; transform: translateY(20px) scale(.94); } 100% { opacity:1; transform: translateY(0) scale(1); } }
@keyframes shine { 0% { transform: translateX(-120%); } 100% { transform: translateX(120%); } }
@media (max-width: 900px) { .stage { height: 840px; } .podium { grid-template-columns: 1fr; } }
</style>
</head>
<body>
<div class='stage'>
  <canvas id='confetti'></canvas>
  <span class='firework f1'></span><span class='firework f2'></span><span class='firework f3'></span><span class='firework f4'></span>
  <span class='firework f5'></span><span class='firework f6'></span><span class='firework f7'></span><span class='firework f8'></span>
  <div class='content'>
    <div class='kicker'>Cerimònia de clausura</div>
    <div class='title'>PORRA MUNDIAL 2026</div>
    <div class='winner'>__WINNER__</div>
    <div class='points'>__POINTS__ punts</div>
    <div class='ribbon'>════════ WINNER 2026 ════════</div>
    <div class='meta'>🥇 1a posició · __PARTICIPANTS__ participants · Celebració final</div>
    <div class='podium'>__TOPHTML__</div>
  </div>
</div>
<script>
const canvas = document.getElementById('confetti');
const ctx = canvas.getContext('2d');
let W, H;
function resize(){ W = canvas.width = canvas.offsetWidth; H = canvas.height = canvas.offsetHeight; }
window.addEventListener('resize', resize); resize();
const colors = ['#ffd166','#ef476f','#06d6a0','#118ab2','#f72585','#4dabf7','#ff9f1c','#e9ff70','#8338ec','#ff006e'];
const pieces = Array.from({length: 170}, () => ({
  x: Math.random() * W,
  y: -Math.random() * H,
  w: 6 + Math.random() * 9,
  h: 10 + Math.random() * 16,
  c: colors[Math.floor(Math.random() * colors.length)],
  s: 1.4 + Math.random() * 3.8,
  r: Math.random() * Math.PI,
  vr: (Math.random() - .5) * .22,
  drift: (Math.random() - .5) * 1.2
}));
function draw(){
  ctx.clearRect(0,0,W,H);
  for(const p of pieces){
    p.y += p.s; p.x += p.drift; p.r += p.vr;
    if(p.y > H + 30){ p.y = -30; p.x = Math.random() * W; }
    if(p.x < -30) p.x = W + 30; if(p.x > W + 30) p.x = -30;
    ctx.save(); ctx.translate(p.x,p.y); ctx.rotate(p.r); ctx.fillStyle=p.c; ctx.fillRect(-p.w/2,-p.h/2,p.w,p.h); ctx.restore();
  }
  requestAnimationFrame(draw);
}
draw();
</script>
</body>
</html>
"""
    html_doc = (html_template
        .replace("__WINNER__", winner)
        .replace("__POINTS__", points)
        .replace("__PARTICIPANTS__", participants)
        .replace("__TOPHTML__", top_html)
    )
    components.html(html_doc, height=730, scrolling=False)
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
    try:
        sheets = pd.read_excel(
            excel_file,
            sheet_name=["Porra", "Resultats Reals", "Calendari"],
            engine="openpyxl"
        )
        df_porra = sheets["Porra"]
        df_resultats = sheets["Resultats Reals"]
        df_calendari = sheets["Calendari"]
    except ValueError:
        sheets = pd.read_excel(
            excel_file,
            sheet_name=["Porra", "Resultats Reals"],
            engine="openpyxl"
        )
        df_porra = sheets["Porra"]
        df_resultats = sheets["Resultats Reals"]
        df_calendari = pd.DataFrame(columns=["Fase", "Partit", "Data", "Hora", "Resultat"])

    df_porra.columns = df_porra.columns.astype(str).str.strip()
    df_resultats.columns = df_resultats.columns.astype(str).str.strip()
    if not df_calendari.empty:
        df_calendari.columns = df_calendari.columns.astype(str).str.strip()

    return df_porra, df_resultats, df_calendari


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

def reduir_nom(nom_complet):
    parts = str(nom_complet).strip().split()
    if len(parts) > 1:
        return f"{parts[0]} {parts[1][0]}."
    return parts[0] if parts else ""

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
    
    if not df.empty:
        df["Posició"] = df["Punts"].rank(method="min", ascending=False).astype(int)
        df["Dif líder"] = (df["Punts"] - float(df["Punts"].iloc[0])).round(1)
    else:
        df["Posició"] = 1
        df["Dif líder"] = 0
        
    return df



# --------------------------------------------------
# V11.4.11 · PODIS AMB EMPATS INTEL·LIGENTS
# --------------------------------------------------
def obtenir_grups_podi(df, nom_col="Participant", punts_col="Punts", pos_col="Posició", max_pos=3):
    """Retorna grups de podi respectant empats reals per posició.
    Exemple: 1r, 2n, 3r+3r retorna tres grups i el tercer pot tenir diversos noms.
    """
    if df is None or df.empty or punts_col not in df.columns:
        return []
    tmp = df.copy()
    tmp[punts_col] = pd.to_numeric(tmp[punts_col], errors="coerce")
    tmp = tmp.dropna(subset=[punts_col]).sort_values(punts_col, ascending=False).reset_index(drop=True)
    if tmp.empty:
        return []
    if pos_col not in tmp.columns:
        tmp[pos_col] = tmp[punts_col].rank(method="min", ascending=False).astype(int)
    else:
        tmp[pos_col] = pd.to_numeric(tmp[pos_col], errors="coerce").fillna(999).astype(int)
    grups = []
    for posicio in sorted([p for p in tmp[pos_col].unique() if p <= max_pos]):
        subset = tmp[tmp[pos_col] == posicio].copy()
        if subset.empty:
            continue
        noms = subset[nom_col].dropna().astype(str).tolist() if nom_col in subset.columns else []
        grups.append({
            "posicio": int(posicio),
            "noms": noms,
            "punts": float(subset[punts_col].iloc[0]),
            "files": subset,
        })
    return grups

def medalla_per_posicio(posicio):
    return {1: "🥇", 2: "🥈", 3: "🥉"}.get(int(posicio), "🏅")

def classe_per_posicio(posicio):
    return {1: "gold", 2: "silver", 3: "bronze"}.get(int(posicio), "bluecard")

def format_noms_podi(noms):
    return " · ".join([str(n) for n in noms if str(n).strip()]) if noms else "Pendent"

def html_card_podi_participants(grup, extra=""):
    medalla = medalla_per_posicio(grup["posicio"])
    classe = classe_per_posicio(grup["posicio"])
    noms = format_noms_podi(grup["noms"])
    punts = float(grup["punts"])
    return f"<div class='card {classe}'><h3>{medalla} {noms}</h3><h1>{punts:.1f}</h1><p>{extra}</p></div>"

def html_card_podi_departaments(grup):
    medalla = medalla_per_posicio(grup["posicio"])
    classe = classe_per_posicio(grup["posicio"])
    noms = format_noms_podi(grup["noms"])
    punts = float(grup["punts"])
    files = grup["files"]
    participants_total = int(files["Participants"].sum()) if "Participants" in files.columns else 0
    lider = ""
    if "Líder departament" in files.columns:
        lider = " · Líder: " + " · ".join(files["Líder departament"].dropna().astype(str).tolist())
    return f"<div class='card {classe}'><h3>{medalla} {noms}</h3><h1>{punts:.1f}</h1><p>Mitjana · {participants_total} part.{lider}</p></div>"

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
    resum["Posició"] = resum["Mitjana_punts"].rank(method="min", ascending=False).astype(int)
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


# --------------------------------------------------
# V11.1 PRE-FINAL: EVOLUCIÓ PER RONDA I PREDICCIONS FINAL
# --------------------------------------------------
def valor_numeric_fila(df_j, columna):
    if columna not in df_j.columns:
        return 0.0
    valor = pd.to_numeric(df_j[columna].values[0], errors="coerce")
    return 0.0 if pd.isna(valor) else float(valor)

def obtenir_evolucio_punts_ronda(df_j):
    def num(*candidats):
        if df_j.empty:
            return 0.0
        mapa = {normalitzar_text(c): c for c in df_j.columns}
        for cand in candidats:
            key = normalitzar_text(cand)
            if key in mapa:
                valor = pd.to_numeric(df_j[mapa[key]].values[0], errors="coerce")
                return 0.0 if pd.isna(valor) else float(valor)
        for cand in candidats:
            tokens = [t for t in normalitzar_text(cand).split() if t]
            for col in df_j.columns:
                col_norm = normalitzar_text(col)
                if tokens and all(t in col_norm for t in tokens):
                    valor = pd.to_numeric(df_j[col].values[0], errors="coerce")
                    return 0.0 if pd.isna(valor) else float(valor)
        return 0.0

    grups = num("Punts Grups 1r") + num("Punts Grups 2n") + num("Punts Grups 3r")
    punts_rondes = [
        ("Grups", grups),
        ("Vuitens", num("Punts Vuitens")),
        ("Quarts", num("Punts Quarts")),
        ("Semis", num("Punts Semis")),
        ("Finalistes", num("Punts Finalistes")),
        ("Campió", num("Punts Campió", "Punts Campio")),
        ("Resultat final", num("Punts Resultat Final", "Punts Resultat final", "Resultat final punts")),
        ("MVP", num("Punts MVP")),
        ("Bota d'Or", num("Punts Pichichi", "Punts Bota d'Or", "Punts Bota dOr", "Punts Bota")),
    ]
    acumulat = 0.0
    files = []
    for ronda, punts in punts_rondes:
        punts = 0.0 if pd.isna(punts) else float(punts)
        acumulat += punts
        files.append({"Ronda": ronda, "Punts ronda": round(punts, 1), "Acumulat": round(acumulat, 1)})

    total_excel = num("Total Punts", "Punts totals", "Total")
    if total_excel:
        if abs(total_excel - acumulat) > 0.05:
            diff = total_excel - acumulat
            files.append({"Ronda": "Total final", "Punts ronda": round(diff, 1), "Acumulat": round(total_excel, 1)})
        else:
            files.append({"Ronda": "Total final", "Punts ronda": 0.0, "Acumulat": round(total_excel, 1)})
    return pd.DataFrame(files)
def mostrar_evolucio_punts_ronda(df_j):
    st.write("### 📈 Evolució de punts per ronda")
    df_evo = obtenir_evolucio_punts_ronda(df_j)
    if df_evo.empty:
        st.info("No hi ha dades de punts per ronda.")
        return
    ordre = df_evo["Ronda"].tolist()
    line = alt.Chart(df_evo).mark_line(point=True, strokeWidth=4, color="#0b70c9").encode(
        x=alt.X("Ronda:N", sort=ordre, title=None, axis=alt.Axis(labelAngle=-35)),
        y=alt.Y("Acumulat:Q", title="Punts acumulats", scale=alt.Scale(zero=True)),
        tooltip=["Ronda", alt.Tooltip("Punts ronda:Q", format=".1f"), alt.Tooltip("Acumulat:Q", format=".1f")]
    )
    text = alt.Chart(df_evo).mark_text(dy=-12, fontSize=12, fontWeight="bold", color="#102a43").encode(
        x=alt.X("Ronda:N", sort=ordre),
        y="Acumulat:Q",
        text=alt.Text("Acumulat:Q", format=".1f")
    )
    st.altair_chart((line + text).properties(height=330).configure_view(strokeWidth=0), use_container_width=True, theme="streamlit")
    st.dataframe(df_evo, use_container_width=True, hide_index=True, column_config={
        "Punts ronda": st.column_config.NumberColumn("Punts ronda", format="%.1f"),
        "Acumulat": st.column_config.NumberColumn("Acumulat", format="%.1f"),
    })


def obtenir_evolucio_tots_participants(df_porra):
    files = []
    for _, row in df_porra.iterrows():
        participant = str(row.get("Participants", "")).strip()
        if participant == "" or participant.lower() in ["nan", "total"] or "total" in participant.lower():
            continue
        df_j = pd.DataFrame([row])
        df_evo = obtenir_evolucio_punts_ronda(df_j)
        if df_evo.empty:
            continue
        for _, evo in df_evo.iterrows():
            ronda = str(evo.get("Ronda", ""))
            files.append({"Participant": participant, "Participant curt": reduir_nom(participant), "Ronda": ronda, "Punts acumulats": float(evo.get("Acumulat", 0))})
    return pd.DataFrame(files)

def mostrar_evolucio_tots_participants(df_porra):
    st.subheader("📈 Evolució de punts per ronda")
    df_evo = obtenir_evolucio_tots_participants(df_porra)
    if df_evo.empty:
        st.info("No hi ha dades suficients per mostrar l’evolució de punts per ronda.")
        return
    ordre = ["Grups", "Vuitens", "Quarts", "Semis", "Finalistes", "Campió", "Resultat final", "MVP", "Bota d'Or", "Total final"]
    rondes_existents = [r for r in ordre if r in df_evo["Ronda"].unique()]
    ordre_llegenda = df_evo.groupby("Participant curt", as_index=False)["Punts acumulats"].max().sort_values("Punts acumulats", ascending=False)["Participant curt"].tolist()
    max_punts = float(df_evo["Punts acumulats"].max()) if not df_evo.empty else 40.0
    domini_y = [35, max(40.0, max_punts + 2)]
    ultima_ronda = rondes_existents[-1] if rondes_existents else None
    df_labels = df_evo[df_evo["Ronda"] == ultima_ronda].copy() if ultima_ronda else pd.DataFrame()
    st.caption("Llegenda ordenada de més a menys punts. Passant el cursor pels punts veuràs el participant i el detall; a l’última boleta també surt el nom abreujat.")
    base = alt.Chart(df_evo).encode(
        x=alt.X("Ronda:N", sort=rondes_existents, title=None, axis=alt.Axis(labelAngle=-35, labelFontSize=12)),
        y=alt.Y("Punts acumulats:Q", title="Punts acumulats", scale=alt.Scale(domain=domini_y, zero=False)),
        color=alt.Color("Participant curt:N", sort=ordre_llegenda, title="Participants", legend=alt.Legend(orient="right", columns=1, labelLimit=190, labelFontSize=11, titleFontSize=13, symbolSize=90)),
        tooltip=["Participant", "Participant curt", "Ronda", alt.Tooltip("Punts acumulats:Q", format=".1f")]
    )
    lines = base.mark_line(point=False, strokeWidth=2.25)
    points = base.mark_circle(size=58, opacity=0.9)
    labels = alt.Chart(df_labels).mark_text(align="left", dx=8, fontSize=11, fontWeight="bold").encode(
        x=alt.X("Ronda:N", sort=rondes_existents),
        y=alt.Y("Punts acumulats:Q", scale=alt.Scale(domain=domini_y, zero=False)),
        text="Participant curt:N",
        color=alt.Color("Participant curt:N", sort=ordre_llegenda, legend=None),
        tooltip=["Participant", "Ronda", alt.Tooltip("Punts acumulats:Q", format=".1f")]
    ) if not df_labels.empty else alt.Chart(pd.DataFrame({"x": []})).mark_text()
    st.altair_chart((lines + points + labels).properties(height=max(560, min(1040, 26 * max(10, len(ordre_llegenda))))).configure_view(strokeWidth=0), use_container_width=True, theme="streamlit")

def obtenir_evolucio_departaments(df_porra):
    col_dep = obtenir_columna_departament(df_porra)
    if col_dep is None:
        return pd.DataFrame()
    df_base = df_porra.copy()
    df_base[col_dep] = df_base[col_dep].fillna("Sense departament").astype(str).str.strip().replace("", "Sense departament")
    files = []
    for _, row in df_base.iterrows():
        participant = str(row.get("Participants", "")).strip()
        if participant == "" or participant.lower() in ["nan", "total"] or "total" in participant.lower():
            continue
        departament = str(row.get(col_dep, "Sense departament")).strip() or "Sense departament"
        df_j = pd.DataFrame([row])
        df_evo = obtenir_evolucio_punts_ronda(df_j)
        for _, evo in df_evo.iterrows():
            ronda = str(evo.get("Ronda", ""))
            files.append({"Departament": departament, "Ronda": ronda, "Punts acumulats": float(evo.get("Acumulat", 0))})
    if not files:
        return pd.DataFrame()
    df = pd.DataFrame(files)
    return df.groupby(["Departament", "Ronda"], as_index=False)["Punts acumulats"].mean().round({"Punts acumulats": 1})

def mostrar_evolucio_departaments(df_porra):
    st.write("#### 📈 Evolució de punts per departament")
    df_dep_evo = obtenir_evolucio_departaments(df_porra)
    if df_dep_evo.empty:
        st.info("No hi ha dades suficients per mostrar l’evolució per departaments.")
        return
    ordre = ["Grups", "Vuitens", "Quarts", "Semis", "Finalistes", "Campió", "Resultat final", "MVP", "Bota d'Or", "Total final"]
    rondes_existents = [r for r in ordre if r in df_dep_evo["Ronda"].unique()]
    deps = df_dep_evo.groupby("Departament", as_index=False)["Punts acumulats"].max().sort_values("Punts acumulats", ascending=False)["Departament"].tolist()
    max_punts = float(df_dep_evo["Punts acumulats"].max()) if not df_dep_evo.empty else 40.0
    domini_y = [35, max(40.0, max_punts + 2)]
    ultima_ronda = rondes_existents[-1] if rondes_existents else None
    df_labels = df_dep_evo[df_dep_evo["Ronda"] == ultima_ronda].copy() if ultima_ronda else pd.DataFrame()
    base = alt.Chart(df_dep_evo).encode(
        x=alt.X("Ronda:N", sort=rondes_existents, title=None, axis=alt.Axis(labelAngle=-35, labelFontSize=12)),
        y=alt.Y("Punts acumulats:Q", title="Mitjana punts acumulats", scale=alt.Scale(domain=domini_y, zero=False)),
        color=alt.Color("Departament:N", sort=deps, title="Departaments", legend=alt.Legend(orient="right", columns=1, labelLimit=240, labelFontSize=12, titleFontSize=13, symbolSize=120)),
        tooltip=["Departament", "Ronda", alt.Tooltip("Punts acumulats:Q", format=".1f")]
    )
    lines = base.mark_line(point=False, strokeWidth=3)
    points = base.mark_circle(size=70, opacity=0.9)
    labels = alt.Chart(df_labels).mark_text(align="left", dx=8, fontSize=12, fontWeight="bold").encode(
        x=alt.X("Ronda:N", sort=rondes_existents),
        y=alt.Y("Punts acumulats:Q", scale=alt.Scale(domain=domini_y, zero=False)),
        text="Departament:N",
        color=alt.Color("Departament:N", sort=deps, legend=None),
        tooltip=["Departament", "Ronda", alt.Tooltip("Punts acumulats:Q", format=".1f")]
    ) if not df_labels.empty else alt.Chart(pd.DataFrame({"x": []})).mark_text()
    st.altair_chart((lines + points + labels).properties(height=max(460, 42 * max(6, len(deps)))).configure_view(strokeWidth=0), use_container_width=True, theme="streamlit")

def preparar_prediccions_resultat_final(df_porra):
    col_res = trobar_col_resultat_final_porra(df_porra)
    if col_res is None or "Participants" not in df_porra.columns:
        return pd.DataFrame(), pd.DataFrame()
    df = df_porra[["Participants", col_res]].copy()
    df.columns = ["Participant", "Resultat apostat"]
    df["Participant"] = df["Participant"].astype(str).str.strip()
    df["Resultat apostat"] = df["Resultat apostat"].apply(valor_o_pendent)
    df = df[(df["Participant"] != "") & (df["Resultat apostat"] != "Pendent")]
    df["Participant curt"] = df["Participant"].apply(reduir_nom)
    detall = df[["Participant", "Resultat apostat"]].sort_values(["Resultat apostat", "Participant"]).reset_index(drop=True)
    agrupat = df.groupby("Resultat apostat", as_index=False).agg(
        Participants=("Participant curt", lambda noms: ", ".join(sorted([str(n) for n in noms if str(n).strip()]))),
        Total=("Participant curt", "count")
    ).sort_values(["Total", "Resultat apostat"], ascending=[False, True]).reset_index(drop=True)
    agrupat = agrupat[["Resultat apostat", "Participants", "Total"]]
    return detall, agrupat

def opcio_mes_apostada(df_porra, columna):
    if columna not in df_porra.columns or "Participants" not in df_porra.columns:
        return "Pendent", ""
    df = df_porra[["Participants", columna]].copy()
    df[columna] = df[columna].apply(valor_o_pendent)
    df = df[df[columna] != "Pendent"]
    if df.empty:
        return "Pendent", ""
    recompte = df.groupby(columna)["Participants"].apply(lambda s: ", ".join(sorted(s.astype(str).apply(reduir_nom)))).reset_index(name="Participants")
    recompte["Total"] = recompte["Participants"].apply(lambda x: len([n for n in x.split(", ") if n]))
    recompte = recompte.sort_values(["Total", columna], ascending=[False, True]).reset_index(drop=True)
    valor = recompte.iloc[0][columna]
    participants = recompte.iloc[0]["Participants"]
    return valor, participants


# --------------------------------------------------
# V11.3.2 · TENDENCIA DEL RESULTAT NUMERIC DE LA FINAL
# --------------------------------------------------
def extreure_marcador_final(valor):
    """Extreu gols Espanya-local i Argentina-visitant des de textos tipus 2-1, 2 - 1, Espanya 2-1 Argentina."""
    if pd.isna(valor):
        return None
    text = str(valor).strip()
    if text == "" or normalitzar_text(text) in ["pendent", "nan", "nat", "none"]:
        return None
    nums = re.findall(r"\d+", text)
    if len(nums) < 2:
        return None
    try:
        return int(nums[0]), int(nums[1])
    except Exception:
        return None

def preparar_tendencia_resultat_final(df_porra):
    col_res = trobar_col_resultat_final_porra(df_porra)
    if col_res is None or "Participants" not in df_porra.columns:
        return pd.DataFrame(), {}
    files = []
    for _, row in df_porra.iterrows():
        participant = str(row.get("Participants", "")).strip()
        if participant == "" or participant.lower() == "nan" or "total" in participant.lower():
            continue
        resultat = valor_o_pendent(row.get(col_res, ""))
        marcador = extreure_marcador_final(resultat)
        if marcador is None:
            continue
        gols_esp, gols_arg = marcador
        if gols_esp > gols_arg:
            tendencia = "Espanya"
        elif gols_arg > gols_esp:
            tendencia = "Argentina"
        else:
            tendencia = "Empat"
        files.append({
            "Participant": participant,
            "Participant curt": reduir_nom(participant),
            "Resultat apostat": resultat,
            "Gols Espanya": gols_esp,
            "Gols Argentina": gols_arg,
            "Tendència": tendencia,
        })
    df = pd.DataFrame(files)
    if df.empty:
        return df, {
            "espanya": 0, "argentina": 0, "empat": 0,
            "avg_esp": 0.0, "avg_arg": 0.0, "resultat_top": "Pendent", "resultat_top_total": 0,
            "participants_esp": "", "participants_arg": "", "participants_emp": ""
        }
    resultats_counts = df.groupby("Resultat apostat", as_index=False).agg(Total=("Participant", "count")).sort_values(["Total", "Resultat apostat"], ascending=[False, True]).reset_index(drop=True)
    resultat_top = str(resultats_counts.iloc[0]["Resultat apostat"]) if not resultats_counts.empty else "Pendent"
    resultat_top_total = int(resultats_counts.iloc[0]["Total"]) if not resultats_counts.empty else 0
    def participants_tendencia(nom):
        noms = df[df["Tendència"] == nom]["Participant curt"].dropna().astype(str).tolist()
        return ", ".join(sorted(noms))
    resum = {
        "espanya": int((df["Tendència"] == "Espanya").sum()),
        "argentina": int((df["Tendència"] == "Argentina").sum()),
        "empat": int((df["Tendència"] == "Empat").sum()),
        "avg_esp": round(float(df["Gols Espanya"].mean()), 2),
        "avg_arg": round(float(df["Gols Argentina"].mean()), 2),
        "resultat_top": resultat_top,
        "resultat_top_total": resultat_top_total,
        "participants_esp": participants_tendencia("Espanya"),
        "participants_arg": participants_tendencia("Argentina"),
        "participants_emp": participants_tendencia("Empat"),
    }
    return df, resum

def mostrar_resum_visual_resultats_apostats(df_porra):
    df_tend, resum = preparar_tendencia_resultat_final(df_porra)
    if df_tend.empty:
        st.info("Encara no hi ha resultats finals apostats amb marcador numèric.")
        return
    st.write("### ⚽ Resum visual dels resultats apostats")
    guanya = "Espanya" if resum["espanya"] > resum["argentina"] else ("Argentina" if resum["argentina"] > resum["espanya"] else "Empat")
    st.markdown(
        f"""
        <div class='final-summary-box'>
            <h3>📊 Lectura ràpida</h3>
            <p><strong>Tendència principal:</strong> {guanya}</p>
            <p><strong>Mitjana de gols prevista:</strong> Espanya {resum['avg_esp']:.2f} · Argentina {resum['avg_arg']:.2f}</p>
            <p><strong>Resultat més repetit:</strong> {resum['resultat_top']} ({resum['resultat_top_total']} participants)</p>
        </div>
        """,
        unsafe_allow_html=True
    )
    resultats = df_tend.groupby("Resultat apostat", as_index=False).agg(
        Total=("Participant", "count"),
        Participants=("Participant curt", lambda noms: ", ".join(sorted([str(n) for n in noms if str(n).strip()])))
    ).sort_values(["Total", "Resultat apostat"], ascending=[False, True]).reset_index(drop=True)
    chart = alt.Chart(resultats.head(12)).mark_bar(cornerRadiusEnd=6).encode(
        x=alt.X("Total:Q", title="Participants"),
        y=alt.Y("Resultat apostat:N", sort="-x", title="Resultat apostat"),
        color=alt.Color("Total:Q", scale=alt.Scale(scheme="blues"), legend=None),
        tooltip=["Resultat apostat", "Total", "Participants"]
    ).properties(height=max(260, min(520, 34 * len(resultats.head(12)))))
    text = chart.mark_text(align="left", baseline="middle", dx=6, fontWeight="bold", color="#334e68").encode(text="Total:Q")
    st.altair_chart((chart + text).configure_view(strokeWidth=0), use_container_width=True, theme="streamlit")
    with st.expander("👥 Veure participants per tendència", expanded=False):
        st.markdown(f"**🇪🇸 Victòria Espanya ({resum['espanya']})**")
        st.write(resum["participants_esp"] if resum["participants_esp"] else "Cap")
        st.markdown(f"**🇦🇷 Victòria Argentina ({resum['argentina']})**")
        st.write(resum["participants_arg"] if resum["participants_arg"] else "Cap")
        st.markdown(f"**🤝 Empat ({resum['empat']})**")
        st.write(resum["participants_emp"] if resum["participants_emp"] else "Cap")

def mostrar_prediccions_resultat_final(df_porra):
    st.subheader("🏁 Prediccions del resultat final")
    detall, agrupat = preparar_prediccions_resultat_final(df_porra)
    if detall.empty and agrupat.empty:
        st.info("Encara no hi ha prediccions del resultat final o no s’ha trobat la columna corresponent.")
        return
    tab_resum, tab_detall = st.tabs(["📊 Resultats apostats", "👥 Prediccions per participant"])
    with tab_resum:
        st.dataframe(agrupat, use_container_width=True, hide_index=True, column_config={
            "Resultat apostat": st.column_config.TextColumn("Resultat apostat"),
            "Participants": st.column_config.TextColumn("Participants"),
            "Total": st.column_config.NumberColumn("Total", format="%d"),
        })
    with tab_detall:
        st.dataframe(detall, use_container_width=True, hide_index=True)
    mostrar_resum_visual_resultats_apostats(df_porra)

def mostrar_estadistiques_prefinal(df_porra):
    st.subheader("📊 Estadístiques pre-final")
    _, resultats_agrupats = preparar_prediccions_resultat_final(df_porra)
    df_tend, resum = preparar_tendencia_resultat_final(df_porra)
    resultat_top = resum.get("resultat_top", "Pendent")
    resultat_top_total = resum.get("resultat_top_total", 0)

    dades = [
        ("🇪🇸 Victòria Espanya", resum.get("espanya", 0), f"Mitjana gols: {resum.get('avg_esp', 0):.2f}", "bluecard"),
        ("🇦🇷 Victòria Argentina", resum.get("argentina", 0), f"Mitjana gols: {resum.get('avg_arg', 0):.2f}", "greencard"),
        ("🤝 Empat", resum.get("empat", 0), "Marcador igualat", "purplecard"),
        ("🏁 Resultat més apostat", resultat_top, f"{resultat_top_total} participants", "darkcard"),
    ]
    html = "<div class='final-trend-grid'>"
    for titol, valor, sub, classe in dades:
        if isinstance(valor, int):
            principal = f"{valor}"
            peu = "participants"
        else:
            principal = valor
            peu = sub
            sub = ""
        html += f"<div class='card final-trend-card {classe}'><h3>{titol}</h3><h1>{principal}</h1><div class='trend-sub'>{peu}</div>{f'<div class=\"trend-sub\">{sub}</div>' if sub else ''}</div>"
    html += "</div>"
    st.markdown(html, unsafe_allow_html=True)

    with st.expander("👥 Veure qui aposta per cada signe de la final", expanded=False):
        st.markdown(f"**🇪🇸 Victòria Espanya ({resum.get('espanya', 0)})**")
        st.write(resum.get("participants_esp", "") or "Cap")
        st.markdown(f"**🇦🇷 Victòria Argentina ({resum.get('argentina', 0)})**")
        st.write(resum.get("participants_arg", "") or "Cap")
        st.markdown(f"**🤝 Empat ({resum.get('empat', 0)})**")
        st.write(resum.get("participants_emp", "") or "Cap")

def obtenir_prediccions_fase(df_j, prefix, quantitat, team_max_phase, dead_teams, vuitens_complet, fase_idx):
    files = []
    for i in range(1, quantitat + 1):
        if prefix == "Campió": col = "Campió"
        else: col = f"{prefix}_{i}"
        
        valor = valor_o_pendent(df_j[col].values[0]) if col in df_j.columns else "Pendent"
        txt_base = afegir_bandera(valor)
        equip_norm = normalitzar_text(valor)
        
        if equip_norm in ["pendent", "nan", "nat", ""]:
            estat = "⚪ "
        else:
            max_phase_reached = team_max_phase.get(equip_norm, 0)
            
            if max_phase_reached >= fase_idx:
                estat = "🟢 "
            elif equip_norm in dead_teams:
                estat = "🔴 "
            elif vuitens_complet and max_phase_reached == 0:
                estat = "🔴 "
            else:
                estat = "⚪ "
        
        if prefix == "Campió":
            files.append({"Concepte": "Campió previst", "Equip": f"{estat}{txt_base}"})
        elif prefix == "Final":
            files.append({"Finalista": f"Finalista {i}", "Equip": f"{estat}{txt_base}"})
        else:
            files.append({"Posició": i, "Equip": f"{estat}{txt_base}"})
    return pd.DataFrame(files)

def mostrar_prediccions_grups_participant(df_j, df_resultats):
    st.write("### 🧩 Prediccions fase de grups")
    
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

def mostrar_prediccions_eliminatoria_participant(df_j, df_resultats, df_calendari):
    st.write("### 🧭 Prediccions fase eliminatòria")
    
    st.markdown("""
    <div style="background-color: rgba(255, 255, 255, 0.7); padding: 12px 18px; border-radius: 12px; margin-bottom: 16px; border: 1px solid rgba(0,0,0,0.1); font-size: 14px;">
        <p style="margin-bottom: 8px;"><strong>Llegenda de resultats:</strong></p>
        <ul style="margin: 0; padding-left: 20px;">
            <li>🟢 <strong>Verd:</strong> L'equip ha arribat oficialment a aquesta fase (encertat).</li>
            <li>🔴 <strong>Vermell:</strong> L'equip ha estat eliminat abans o en aquesta mateixa ronda (fallat).</li>
            <li>⚪ <strong>Pendent:</strong> Encara no ha arribat, però l'equip segueix viu al torneig.</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)
    
    team_max_phase = {}
    for col, idx in [("Vuitens", 1), ("Quarts", 2), ("Semis", 3), ("Finalistes", 4), ("Campió", 5)]:
        if col in df_resultats.columns:
            for val in llista_valors_no_buits(df_resultats, col):
                t = normalitzar_text(val)
                if t != "pendent":
                    team_max_phase[t] = max(team_max_phase.get(t, 0), idx)
    
    dead_teams = set()
    if not df_calendari.empty:
        for _, row in df_calendari.iterrows():
            res = str(row.get("Resultat", "")).strip().lower()
            if res not in ["pendent", "nan", ""]:
                fase_nom = str(row.get("Fase", "")).strip().lower()
                partit = str(row.get("Partit", ""))
                
                target_idx = 0
                if "setzen" in fase_nom or "1/16" in fase_nom or "dieciseis" in fase_nom: target_idx = 1
                elif "vuit" in fase_nom or "octav" in fase_nom or "1/8" in fase_nom: target_idx = 2
                elif "quart" in fase_nom or "cuart" in fase_nom or "1/4" in fase_nom: target_idx = 3
                elif "semi" in fase_nom or "1/2" in fase_nom: target_idx = 4
                elif "final" in fase_nom and "consol" not in fase_nom and "3r" not in fase_nom: target_idx = 5
                
                if target_idx > 0:
                    parts = [t.strip() for t in re.split(r' vs\. | vs | - ', partit, maxsplit=1)]
                    if len(parts) == 2:
                        t1, t2 = normalitzar_text(parts[0]), normalitzar_text(parts[1])
                        
                        matches = re.findall(r'\d+', res)
                        winner = None
                        loser = None
                        
                        if len(matches) >= 2:
                            g1, g2 = int(matches[0]), int(matches[1])
                            if g1 > g2: 
                                winner, loser = t1, t2
                            elif g2 > g1: 
                                winner, loser = t2, t1
                            elif len(matches) >= 4:
                                p1, p2 = int(matches[2]), int(matches[3])
                                if p1 > p2: 
                                    winner, loser = t1, t2
                                elif p2 > p1: 
                                    winner, loser = t2, t1
                                    
                        if winner:
                            team_max_phase[winner] = max(team_max_phase.get(winner, 0), target_idx)
                        if loser:
                            dead_teams.add(loser)
                            team_max_phase[loser] = max(team_max_phase.get(loser, 0), target_idx - 1)
    
    vuitens_complet = False
    if "Vuitens" in df_resultats.columns:
        v_teams = [t for t in llista_valors_no_buits(df_resultats, "Vuitens") if normalitzar_text(t) != "pendent"]
        if len(v_teams) >= 16:
            vuitens_complet = True
            
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["Vuitens", "Quarts", "Semis", "Final", "Campió"])
    with tab1: st.dataframe(obtenir_prediccions_fase(df_j, "Vuitens", 16, team_max_phase, dead_teams, vuitens_complet, 1), use_container_width=True, hide_index=True)
    with tab2: st.dataframe(obtenir_prediccions_fase(df_j, "Quarts", 8, team_max_phase, dead_teams, vuitens_complet, 2), use_container_width=True, hide_index=True)
    with tab3: st.dataframe(obtenir_prediccions_fase(df_j, "Semis", 4, team_max_phase, dead_teams, vuitens_complet, 3), use_container_width=True, hide_index=True)
    with tab4: st.dataframe(obtenir_prediccions_fase(df_j, "Final", 2, team_max_phase, dead_teams, vuitens_complet, 4), use_container_width=True, hide_index=True)
    with tab5: st.dataframe(obtenir_prediccions_fase(df_j, "Campió", 1, team_max_phase, dead_teams, vuitens_complet, 5), use_container_width=True, hide_index=True)



# --------------------------------------------------
# V11.3 · DASHBOARD FINAL DE GUANYADORS I CLAUS
# --------------------------------------------------
def mostrar_dashboard_final_professional(df_ranking, df_departaments, df_porra, df_resultats_display):
    if df_ranking.empty:
        return
    guanyador = df_ranking.iloc[0]
    segon = df_ranking.iloc[1] if len(df_ranking) > 1 else None
    tercer = df_ranking.iloc[2] if len(df_ranking) > 2 else None
    lider_dep_txt = "Pendent"
    if df_departaments is not None and not df_departaments.empty:
        dep = df_departaments.iloc[0]
        lider_dep_txt = f"{dep['Departament']} · {float(dep['Mitjana_punts']):.1f} pts mitjana"
    campio_real = afegir_bandera(primer_valor_o_pendent(df_resultats_display, "Campió")) if df_resultats_display is not None else "Pendent"
    resultat_final_real = primer_valor_o_pendent(df_resultats_display, "Resultat Final") if df_resultats_display is not None else "Pendent"
    pichichi_real, gols_pichichi = obtenir_pichichi_real(df_resultats_display, "Jugador Pichichi", "Gols") if df_resultats_display is not None else ("Pendent", "Pendent")
    bota_text = f"{pichichi_real} ({gols_pichichi})" if gols_pichichi != "Pendent" else pichichi_real
    col_res = trobar_col_resultat_final_porra(df_porra)
    total_resultats = len(df_porra[df_porra[col_res].apply(valor_o_pendent) != "Pendent"]) if col_res is not None else 0
    podi_html = ""
    for medalla, row in [("🥇", guanyador), ("🥈", segon), ("🥉", tercer)]:
        if row is not None:
            podi_html += f"<span class='final-pill'>{medalla} {row['Participant']} · {float(row['Punts']):.1f}</span>"
    st.markdown(f"""
    <div class='final-dashboard'>
        <h2>🏆 Dashboard final · Porra Mundial</h2>
        <p>Resum visual dels guanyadors, el podi i les claus principals del torneig.</p>
        <div class='final-grid'>
            <div class='final-hero'>
                <h3>Guanyador de la porra</h3>
                <div class='winner'>{guanyador['Participant']}</div>
                <div class='points'>{float(guanyador['Punts']):.1f} punts</div>
                <div class='final-pills'>{podi_html}</div>
            </div>
            <div class='final-mini'>
                <h3>🏢 Departament destacat</h3>
                <div class='big'>{lider_dep_txt}</div>
                <p>Classificació calculada per mitjana de punts.</p>
            </div>
            <div class='final-mini'>
                <h3>🔑 Claus del torneig</h3>
                <div class='final-pills'>
                    <span class='final-pill'>Campió: {campio_real}</span>
                    <span class='final-pill'>Final: {resultat_final_real}</span>
                    <span class='final-pill'>Bota d'Or: {bota_text}</span>
                    <span class='final-pill'>{total_resultats} prediccions de resultat</span>
                </div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)


# --------------------------------------------------
# V11.4 · DASHBOARD PREMIUM DEL CAMPIÓ
# --------------------------------------------------

def es_nom_femeni_porres(nom):
    primer = normalitzar_text(str(nom).split()[0]) if str(nom).strip() else ""
    noms_femenins = {
        "agustina", "brenda", "patri", "patricia", "jessica", "sonia", "lena", "veronica", "vero",
        "maite", "montse", "sandra", "carmen", "elena", "nuria", "nouria", "maria", "laura",
        "anna", "ana", "marta", "cristina", "tamara", "silvia", "sílvia", "meritxell", "beatriz"
    }
    return primer in noms_femenins

def textos_genere_guanyador(nom):
    # V11.4.5: textos neutres, sense detectar gènere per nom.
    return {
        "celebracio": "Número 1 de la Porra Mundial 2026",
        "kicker": "1a posició absoluta",
        "guanyador": "Lideratge final",
        "campio": "en primera posició",
        "segon": "2a posició",
    }
def trobar_imatge_ia(base_name):
    candidats = []
    for ext in ["", ".png", ".jpg", ".jpeg", ".webp"]:
        candidats.append(base_name + ext)
    # Compatibilitat si el fitxer s'ha pujat sense accents.
    if "últims" in base_name:
        for ext in ["", ".png", ".jpg", ".jpeg", ".webp"]:
            candidats.append(base_name.replace("últims", "ultims") + ext)
    for path in candidats:
        if os.path.exists(path):
            return path
    return None

def mostrar_bloc_imatge_ia(base_name, titol, subtitol="", placeholder=""):
    path = trobar_imatge_ia(base_name)
    if not path:
        return
    st.markdown(f"<div class='ai-poster-block'><div class='ai-poster-title'>{titol}</div></div>", unsafe_allow_html=True)
    st.image(path, use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

def trobar_imatge_resum_final():
    for base_name in ["imatge_final", "imatge_resum_final", "imatge_partit_final", "imatge_resum_partit"]:
        path = trobar_imatge_ia(base_name)
        if path:
            return path
    return None

def mostrar_bloc_resum_visual_final():
    path = trobar_imatge_resum_final()
    st.markdown("<div class='ai-poster-block'><div class='ai-poster-title'>🎨 Resum visual de la final</div></div>", unsafe_allow_html=True)
    if path:
        st.image(path, use_container_width=True)
    else:
        st.markdown("<div class='ai-placeholder'>Espai reservat per al resum visual de la final</div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

def obtenir_moviments_jornada(df_ranking):
    if df_ranking.empty or "Canvi posició" not in df_ranking.columns:
        return None, None
    tmp = df_ranking.copy()
    tmp["Canvi posició"] = pd.to_numeric(tmp["Canvi posició"], errors="coerce").fillna(0)
    guanya = tmp.sort_values("Canvi posició", ascending=False).iloc[0]
    perd = tmp.sort_values("Canvi posició", ascending=True).iloc[0]
    return guanya, perd

def obtenir_premis_especials(df_ranking, df_departaments, df_porra, df_resultats_display):
    premis = []
    encerts = preparar_qui_la_va_clavar(df_porra, df_resultats_display) if 'preparar_qui_la_va_clavar' in globals() else pd.DataFrame()
    if not encerts.empty and "Distància" in encerts.columns:
        exactes = encerts[pd.to_numeric(encerts["Distància"], errors="coerce").fillna(999) == 0]
        if not exactes.empty:
            noms = " · ".join(exactes["Participant"].astype(str).tolist())
            resultat = exactes.iloc[0].get("Resultat real", "Resultat exacte")
            premis.append(("🎯 Resultat final encertat", noms, f"Resultat exacte: {resultat}"))
        else:
            premis.append(("🎯 Resultat final encertat", "Cap encertant", "Cap participant ha clavat el marcador exacte"))
    else:
        premis.append(("🎯 Resultat final encertat", "Pendent", "Falta resultat real o prediccions numèriques"))

    if "Canvi posició" in df_ranking.columns:
        tmp = df_ranking.copy()
        tmp["Canvi posició"] = pd.to_numeric(tmp["Canvi posició"], errors="coerce").fillna(0)
        max_pujada = tmp["Canvi posició"].max()
        if max_pujada > 0:
            noms_pujada = " · ".join(tmp[tmp["Canvi posició"] == max_pujada]["Participant"].astype(str).tolist())
            premis.append(("🔥 Sprint final", noms_pujada, f"+{int(max_pujada)} posicions"))
        else:
            premis.append(("🔥 Sprint final", "Sense pujades", "Cap participant ha guanyat posicions"))

        max_baixada = tmp["Canvi posició"].min()
        if max_baixada < 0:
            noms_baixada = " · ".join(tmp[tmp["Canvi posició"] == max_baixada]["Participant"].astype(str).tolist())
            premis.append(("💥 Punxada final", noms_baixada, f"{int(max_baixada)} posicions"))
        else:
            premis.append(("💥 Punxada final", "Sense baixades", "Cap participant ha perdut posicions"))
    else:
        premis.append(("🔥 Sprint final", "Pendent", "No hi ha columna de canvi de posició"))
        premis.append(("💥 Punxada final", "Pendent", "No hi ha columna de canvi de posició"))

    if df_departaments is not None and not df_departaments.empty and "Departament" in df_ranking.columns:
        dep_lider = df_departaments.iloc[0]["Departament"]
        candidats = df_ranking[df_ranking["Departament"].astype(str) == str(dep_lider)]
        if not candidats.empty:
            millor_dep = candidats.sort_values("Punts", ascending=False).iloc[0]
            premis.append(("🏢 Orgull del Departament", millor_dep["Participant"], f"Millor del departament {dep_lider}"))
        else:
            premis.append(("🏢 Orgull del Departament", str(dep_lider), "Departament amb millor puntuació"))
    else:
        premis.append(("🏢 Orgull del Departament", "Pendent", "Sense departaments configurats"))
    return premis
def obtenir_resultat_real_final(df_resultats_display):
    valor = primer_valor_o_pendent(df_resultats_display, "Resultat Final") if df_resultats_display is not None else "Pendent"
    marcador = extreure_marcador_final(valor) if 'extreure_marcador_final' in globals() else None
    return valor, marcador

def preparar_qui_la_va_clavar(df_porra, df_resultats_display):
    col_res = trobar_col_resultat_final_porra(df_porra)
    resultat_real, marcador_real = obtenir_resultat_real_final(df_resultats_display)
    if col_res is None or marcador_real is None:
        return pd.DataFrame()
    files = []
    real_esp, real_arg = marcador_real
    for _, row in df_porra.iterrows():
        participant = str(row.get("Participants", "")).strip()
        if participant == "" or participant.lower() == "nan" or "total" in participant.lower():
            continue
        resultat = valor_o_pendent(row.get(col_res, ""))
        marcador = extreure_marcador_final(resultat) if 'extreure_marcador_final' in globals() else None
        if marcador is None:
            continue
        esp, arg = marcador
        distancia = abs(esp - real_esp) + abs(arg - real_arg)
        if distancia == 0:
            estat = "🥇 Clavat"
        elif distancia == 1:
            estat = "🥈 Molt a prop"
        elif distancia == 2:
            estat = "🥉 A prop"
        else:
            estat = "⚪ Lluny"
        files.append({
            "Participant": participant,
            "Participant curt": reduir_nom(participant),
            "Resultat apostat": resultat,
            "Resultat real": resultat_real,
            "Distància": distancia,
            "Estat": estat,
        })
    if not files:
        return pd.DataFrame()
    return pd.DataFrame(files).sort_values(["Distància", "Participant"]).reset_index(drop=True)

def obtenir_remuntada_destacada(df_ranking):
    if df_ranking.empty or "Canvi posició" not in df_ranking.columns:
        return None
    tmp = df_ranking.copy()
    tmp["Canvi posició"] = pd.to_numeric(tmp["Canvi posició"], errors="coerce").fillna(0)
    tmp = tmp.sort_values("Canvi posició", ascending=False)
    if tmp.empty or float(tmp.iloc[0]["Canvi posició"]) <= 0:
        return None
    return tmp.iloc[0]

def generar_frase_final(df_ranking, df_departaments, df_porra, df_resultats_display):
    if df_ranking.empty:
        return "La porra encara no té dades suficients per generar el resum final."
    guanyador = df_ranking.iloc[0]
    punts = float(guanyador.get("Punts", 0))
    resultat_real = primer_valor_o_pendent(df_resultats_display, "Resultat Final") if df_resultats_display is not None else "Pendent"
    df_tend, resum = preparar_tendencia_resultat_final(df_porra) if 'preparar_tendencia_resultat_final' in globals() else (pd.DataFrame(), {})
    resultat_top = resum.get("resultat_top", "Pendent") if isinstance(resum, dict) else "Pendent"
    avg_esp = resum.get("avg_esp", 0) if isinstance(resum, dict) else 0
    avg_arg = resum.get("avg_arg", 0) if isinstance(resum, dict) else 0
    dep_txt = ""
    if df_departaments is not None and not df_departaments.empty:
        dep_txt = f" El departament amb millor puntuació ha estat {df_departaments.iloc[0]['Departament']} amb {float(df_departaments.iloc[0]['Mitjana_punts']):.1f} punts de mitjana."
    return f"{guanyador['Participant']} queda en primera posició de la Porra Mundial amb {punts:.1f} punts. El resultat real de la final és {resultat_real}; el marcador més apostat era {resultat_top}. La porra projectava una mitjana de gols de {avg_esp:.2f} per Espanya i {avg_arg:.2f} per Argentina.{dep_txt}"
def mostrar_dashboard_campio_premium(df_ranking, df_departaments, df_porra, df_resultats_display):
    if df_ranking.empty:
        return
    df_final = df_ranking.copy().sort_values("Punts", ascending=False).reset_index(drop=True)
    guanyador = df_final.iloc[0]
    segon = df_final.iloc[1] if len(df_final) > 1 else None
    tercer = df_final.iloc[2] if len(df_final) > 2 else None
    diff_seg = float(guanyador["Punts"] - segon["Punts"]) if segon is not None else 0.0
    txt = textos_genere_guanyador(guanyador["Participant"])

    podi_cards = []
    for grup in obtenir_grups_podi(df_final, nom_col="Participant", punts_col="Punts", pos_col="Posició", max_pos=3):
        medalla = medalla_per_posicio(grup["posicio"])
        noms = format_noms_podi(grup["noms"])
        files = grup["files"]
        deps = []
        if "Departament" in files.columns:
            deps = [str(d) for d in files["Departament"].dropna().astype(str).unique() if str(d).strip()]
        dep = " · " + " · ".join(deps) if deps else ""
        podi_cards.append(
            "<div class='podium-card'>"
            f"<div class='podium-medal'>{medalla}</div>"
            f"<div><div class='podium-name'>{noms}</div><div class='podium-meta'>{dep}</div></div>"
            f"<div class='podium-score'>{float(grup['punts']):.1f}</div>"
            "</div>"
        )
    podi_html = "".join(podi_cards)

    resultat_real = primer_valor_o_pendent(df_resultats_display, "Resultat Final") if df_resultats_display is not None else "Pendent"
    campio_real = afegir_bandera(primer_valor_o_pendent(df_resultats_display, "Campió")) if df_resultats_display is not None else "Pendent"
    mvp_real = primer_valor_o_pendent(df_resultats_display, "MVP") if df_resultats_display is not None else "Pendent"
    pichichi_real, gols_pichichi = obtenir_pichichi_real(df_resultats_display, "Jugador Pichichi", "Gols") if df_resultats_display is not None else ("Pendent", "Pendent")
    bota_txt = f"{pichichi_real} ({gols_pichichi})" if gols_pichichi != "Pendent" else pichichi_real
    frase = generar_frase_final(df_final, df_departaments, df_porra, df_resultats_display)

    html = (
        "<div class='champion-wrap'>"
        "<div class='champion-title'>"
        f"<div><h2>🏆 {txt['celebracio']}</h2><p>{txt['guanyador']} de la Porra Mundial 2026</p></div>"
        "</div>"
        "<div class='champion-grid'>"
        "<div class='champion-hero'>"
        f"<div class='champion-kicker'>{txt['kicker']}</div>"
        f"<div class='champion-name'>{guanyador['Participant']}</div>"
        f"<div class='champion-points'>{float(guanyador['Punts']):.1f} punts</div>"
        f"<div class='champion-diff'>+{diff_seg:.1f} respecte a la {txt['segon']}</div>"
        "</div>"
        f"<div class='podium-grid'>{podi_html}</div>"
        "</div>"
        "<div class='premium-kpi-grid'>"
        f"<div class='premium-kpi'><h3>🏆 Selecció guanyadora</h3><div class='big'>{campio_real}</div><div class='small'>Guanyadora del Mundial</div></div>"
        f"<div class='premium-kpi'><h3>⭐ MVP</h3><div class='big'>{mvp_real}</div><div class='small'>Millor jugador</div></div>"
        f"<div class='premium-kpi'><h3>⚽ Bota d'Or</h3><div class='big'>{bota_txt}</div><div class='small'>Màxim golejador</div></div>"
        f"<div class='premium-kpi'><h3>🏁 Resultat final</h3><div class='big'>{resultat_real}</div><div class='small'>Marcador oficial</div></div>"
        "</div>"
        "</div>"
    )
    st.markdown(html, unsafe_allow_html=True)

    mostrar_bloc_imatge_ia(
        "imatge_campions",
        "🥇🥈🥉 Els tres primers classificats"
    )


    # --------------------------------------------------
    # V11.4.12 · ESPAI FANALET VERMELL
    # --------------------------------------------------
    if not df_final.empty and "Punts" in df_final.columns:
        min_punts = pd.to_numeric(df_final["Punts"], errors="coerce").min()
        fanalets = df_final[pd.to_numeric(df_final["Punts"], errors="coerce") == min_punts].copy()
        noms_fanalet = " · ".join(fanalets["Participant"].dropna().astype(str).tolist()) if "Participant" in fanalets.columns else "Pendent"
        deps_fanalet = ""
        if "Departament" in fanalets.columns:
            deps = [str(d) for d in fanalets["Departament"].dropna().astype(str).unique() if str(d).strip()]
            deps_fanalet = " · " + " · ".join(deps) if deps else ""

        fanalet_html = (
            "<div class='champion-wrap'>"
            "<div class='champion-title'><div><h2>🔦 Fanalet vermell</h2><p>Espai de tancament de la classificació final</p></div></div>"
            "<div style='border-radius:28px;padding:32px 22px;text-align:center;background:linear-gradient(135deg,rgba(255,104,91,.26),rgba(255,255,255,.10));border:1px solid rgba(255,255,255,.24);box-shadow:inset 0 1px 0 rgba(255,255,255,.15);'>"
            "<div style='font-size:clamp(44px,6vw,82px);line-height:1;'>🔦</div>"
            f"<div style='font-size:clamp(30px,4.4vw,60px);font-weight:1000;color:white;line-height:1.05;margin-top:10px;text-shadow:0 7px 26px rgba(0,0,0,.36);'>{noms_fanalet}</div>"
            f"<div style='font-size:clamp(20px,2.5vw,34px);font-weight:950;color:#fff1a8;margin-top:12px;'>{float(min_punts):.1f} punts</div>"
            f"<div style='font-size:16px;font-weight:800;color:rgba(255,255,255,.86);margin-top:8px;'>{deps_fanalet}</div>"
            "</div>"
            "</div>"
        )
        st.markdown(fanalet_html, unsafe_allow_html=True)
        mostrar_bloc_imatge_ia(
            "imatge_fanalet",
            "🔦 Fanalet vermell"
        )

    dept_cards = []
    if df_departaments is not None and not df_departaments.empty:
        for grup in obtenir_grups_podi(df_departaments, nom_col="Departament", punts_col="Mitjana_punts", pos_col="Posició", max_pos=3):
            medalla = medalla_per_posicio(grup["posicio"])
            noms_dep = format_noms_podi(grup["noms"])
            files = grup["files"]
            participants_total = int(files["Participants"].sum()) if "Participants" in files.columns else 0
            dept_cards.append(
                "<div class='story-card'>"
                f"<h3>{medalla} Departament TOP {grup['posicio']}</h3>"
                f"<div class='hero-text'>{noms_dep}</div>"
                f"<div class='detail'>{float(grup['punts']):.1f} punts de mitjana · {participants_total} participants</div>"
                "</div>"
            )
    if not dept_cards:
        dept_cards = [
            "<div class='story-card'><h3>🏢 Departament TOP</h3><div class='hero-text'>Pendent</div><div class='detail'>Sense departaments configurats</div></div>"
        ]

    html2 = (
        "<div class='champion-wrap'>"
        "<div class='champion-title'><div><h2>🏢 Podi de departaments</h2><p>Els tres departaments amb millor puntuació mitjana</p></div></div>"
        f"<div class='story-grid'>{''.join(dept_cards)}</div>"
        "</div>"
    )
    st.markdown(html2, unsafe_allow_html=True)

    final_html = (
        "<div class='champion-wrap'>"
        "<div class='champion-title'><div><h2>🏁 Final del Mundial</h2><p>Resultat oficial de la final</p></div></div>"
        "<div style='border-radius:28px;padding:34px 22px;text-align:center;background:linear-gradient(135deg,rgba(255,215,0,.25),rgba(255,255,255,.12));border:1px solid rgba(255,255,255,.24);'>"
        "<div style='font-size:clamp(22px,3vw,36px);font-weight:950;color:#fff1a8;text-transform:uppercase;letter-spacing:2px;'>Marcador final</div>"
        f"<div style='font-size:clamp(54px,9vw,122px);font-weight:1000;line-height:.95;color:white;text-shadow:0 8px 30px rgba(0,0,0,.42);margin-top:12px;'>{resultat_real}</div>"
        "</div>"
        "</div>"
    )
    st.markdown(final_html, unsafe_allow_html=True)
    mostrar_bloc_resum_visual_final()

    premis = obtenir_premis_especials(df_final, df_departaments, df_porra, df_resultats_display)
    premis_html = "<div class='champion-wrap'><div class='champion-title'><div><h2>🎖 Premis especials</h2></div></div><div class='special-grid'>"
    for titol, nom, detall in premis:
        premis_html += f"<div class='special-card'><h3>{titol}</h3><div class='special-name'>{nom}</div><div class='special-detail'>{detall}</div></div>"
    premis_html += "</div></div>"
    st.markdown(premis_html, unsafe_allow_html=True)

    encerts = preparar_qui_la_va_clavar(df_porra, df_resultats_display)
    if not encerts.empty:
        exactes = encerts[pd.to_numeric(encerts["Distància"], errors="coerce").fillna(999) == 0]
        if not exactes.empty:
            st.subheader("🎯 Resultat de la final encertat")
            st.dataframe(exactes[["Participant", "Resultat apostat", "Estat"]], use_container_width=True, hide_index=True)
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
    .purplecard {{ background: linear-gradient(135deg, #6f42c1, #b982ff); color: white; }}
    .card h3 {{ margin: 0px 0px 14px 0px; font-size: clamp(15px, 2vw, 24px); line-height: 1.15; max-width: 100%; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
    .card h1 {{ margin: 0px; font-size: clamp(24px, 4vw, 40px); line-height: 1.1; max-width: 100%; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
    .card p {{ margin: 12px 0px 0px 0px; font-size: clamp(11px, 1.5vw, 15px); max-width: 100%; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
    
.prefinal-grid {{ display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 1.25rem; align-items: stretch; margin-bottom: 0.75rem; }}
.prefinal-card {{ min-height: 218px !important; height: 218px !important; margin: 0 !important; padding: 22px 18px !important; justify-content: space-between !important; }}
.prefinal-card h3 {{ min-height: 34px; display: flex; align-items: center; justify-content: center; white-space: normal !important; }}
.prefinal-card h1 {{ white-space: normal !important; line-height: 1.15 !important; min-height: 58px; display: flex; align-items: center; justify-content: center; }}
.prefinal-card p {{ white-space: normal !important; font-weight: 700; opacity: 0.95; }}
.final-dashboard {{ margin-top: 22px; margin-bottom: 28px; padding: 24px; border-radius: 28px; background: linear-gradient(135deg, rgba(16,42,67,0.96), rgba(11,112,201,0.84)); color: white; box-shadow: 0px 14px 40px rgba(0,0,0,0.28); }}
.final-dashboard h2 {{ color: white !important; margin-top: 0; font-size: clamp(28px, 4vw, 44px); text-shadow: none; }}
.final-dashboard p {{ color: rgba(255,255,255,0.9); }}
.final-grid {{ display: grid; grid-template-columns: 1.25fr 1fr 1fr; gap: 1rem; align-items: stretch; }}
.final-hero, .final-mini {{ background: rgba(255,255,255,0.14); border: 1px solid rgba(255,255,255,0.22); border-radius: 22px; padding: 20px; backdrop-filter: blur(8px); }}
.final-hero h3, .final-mini h3 {{ color: white !important; margin: 0 0 10px 0; text-shadow: none; }}
.final-hero .winner {{ font-size: clamp(28px, 4vw, 46px); font-weight: 900; line-height: 1.05; margin: 8px 0; }}
.final-hero .points {{ font-size: 20px; font-weight: 800; color: #fff1a8; }}
.final-mini .big {{ font-size: 24px; font-weight: 900; line-height: 1.15; }}
.final-pills {{ display: flex; flex-wrap: wrap; gap: 8px; margin-top: 14px; }}
.final-pill {{ background: rgba(255,255,255,0.18); border: 1px solid rgba(255,255,255,0.22); border-radius: 999px; padding: 7px 11px; color: white; font-weight: 700; }}
@media (max-width: 1100px) {{ .prefinal-grid {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }} .final-grid {{ grid-template-columns: 1fr; }} }}

.final-trend-grid {{ display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 1.25rem; align-items: stretch; margin-bottom: 1rem; }}
.final-trend-card {{ min-height: 280px !important; height: 280px !important; margin: 0 !important; padding: 28px 22px !important; justify-content: center !important; gap: 18px !important; overflow: visible !important; }}
.final-trend-card h3 {{ margin: 0 !important; min-height: 38px; white-space: normal !important; display: flex; justify-content: center; align-items: center; text-align: center; }}
.final-trend-card h1 {{ margin: 0 !important; line-height: 1.05 !important; white-space: normal !important; display: flex; justify-content: center; align-items: center; text-align: center; }}
.final-trend-card .trend-sub {{ font-size: 16px; font-weight: 800; opacity: 0.98; }}
.final-summary-box {{ padding: 18px 20px; border-radius: 18px; background: rgba(255,255,255,0.78); border: 1px solid rgba(0,0,0,0.08); box-shadow: 0px 4px 18px rgba(0,0,0,0.10); margin-top: 10px; margin-bottom: 16px; }}
.final-summary-box h3 {{ margin-top: 0 !important; }}
@media (max-width: 1100px) {{ .final-trend-grid {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }} }}

/* ========================= V11.4 PREMIUM WINNER DASHBOARD ========================= */
.champion-wrap {{
    margin: 22px 0 30px 0;
    padding: 26px;
    border-radius: 32px;
    background:
        radial-gradient(circle at 15% 10%, rgba(255,215,0,0.34), transparent 28%),
        radial-gradient(circle at 90% 0%, rgba(124,197,255,0.28), transparent 26%),
        linear-gradient(135deg, rgba(5,20,36,0.97), rgba(8,52,92,0.94) 55%, rgba(15,113,201,0.88));
    box-shadow: 0 20px 55px rgba(0,0,0,0.36);
    color: white;
    overflow: hidden;
    position: relative;
}}
.champion-wrap::before {{
    content: "";
    position: absolute;
    inset: 0;
    background-image: linear-gradient(120deg, rgba(255,255,255,0.08), transparent 36%, rgba(255,255,255,0.04));
    pointer-events: none;
}}
.champion-title {{
    position: relative;
    display: flex;
    justify-content: space-between;
    gap: 16px;
    align-items: flex-start;
    margin-bottom: 18px;
}}
.champion-title h2 {{ color: white !important; font-size: clamp(30px, 4.2vw, 56px); margin: 0; letter-spacing: -1.2px; text-shadow: 0 4px 18px rgba(0,0,0,0.33); }}
.champion-title p {{ margin: 4px 0 0 0; color: rgba(255,255,255,0.86); font-weight: 650; }}
.champion-badge {{ border: 1px solid rgba(255,255,255,0.28); background: rgba(255,255,255,0.13); backdrop-filter: blur(8px); border-radius: 999px; padding: 10px 15px; font-weight: 900; white-space: nowrap; }}
.champion-grid {{ display: grid; grid-template-columns: 1.45fr 1fr; gap: 18px; position: relative; }}
.champion-hero {{
    min-height: 330px;
    border-radius: 28px;
    padding: 28px;
    background: linear-gradient(135deg, rgba(255,215,0,0.28), rgba(255,255,255,0.12));
    border: 1px solid rgba(255,255,255,0.24);
    box-shadow: inset 0 1px 0 rgba(255,255,255,0.18);
    display: flex;
    flex-direction: column;
    justify-content: center;
}}
.champion-kicker {{ font-size: 15px; text-transform: uppercase; letter-spacing: 2px; font-weight: 900; color: #fff1a8; }}
.champion-name {{ font-size: clamp(40px, 6vw, 82px); line-height: 0.95; font-weight: 1000; margin: 12px 0; letter-spacing: -2px; text-shadow: 0 4px 22px rgba(0,0,0,0.32); }}
.champion-points {{ font-size: clamp(24px, 3vw, 38px); font-weight: 1000; color: #fff1a8; }}
.champion-diff {{ margin-top: 10px; font-size: 18px; font-weight: 800; color: rgba(255,255,255,0.9); }}
.podium-grid {{ display: grid; grid-template-columns: 1fr; gap: 12px; }}
.podium-card {{ border-radius: 22px; padding: 18px; background: rgba(255,255,255,0.13); border: 1px solid rgba(255,255,255,0.22); backdrop-filter: blur(8px); display: grid; grid-template-columns: 54px 1fr auto; gap: 12px; align-items: center; }}
.podium-medal {{ font-size: 34px; }}
.podium-name {{ font-size: 20px; font-weight: 950; line-height: 1.05; }}
.podium-meta {{ font-size: 13px; color: rgba(255,255,255,0.78); margin-top: 4px; }}
.podium-score {{ font-size: 27px; font-weight: 1000; color: #fff1a8; }}
.premium-kpi-grid {{ display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 14px; margin-top: 18px; position: relative; }}
.premium-kpi {{ border-radius: 22px; padding: 18px; background: rgba(255,255,255,0.12); border: 1px solid rgba(255,255,255,0.20); min-height: 145px; display: flex; flex-direction: column; justify-content: center; align-items: center; text-align: center; }}
.premium-kpi h3 {{ color: white !important; margin: 0 0 8px 0; font-size: 16px; white-space: normal !important; }}
.premium-kpi .big {{ font-size: clamp(24px, 3.2vw, 42px); font-weight: 1000; color: #fff1a8; line-height: 1; }}
.premium-kpi .small {{ font-weight: 750; color: rgba(255,255,255,0.82); margin-top: 8px; }}
.story-grid {{ display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 16px; margin-top: 18px; position: relative; }}
.story-card {{ border-radius: 24px; padding: 20px; background: rgba(255,255,255,0.12); border: 1px solid rgba(255,255,255,0.20); min-height: 170px; }}
.story-card h3 {{ color: white !important; margin: 0 0 10px 0; }}
.story-card .hero-text {{ font-size: 25px; font-weight: 1000; line-height: 1.1; color: #fff1a8; }}
.story-card .detail {{ color: rgba(255,255,255,0.84); font-weight: 700; margin-top: 8px; }}
.final-phrase {{ margin-top: 18px; border-radius: 22px; padding: 18px 20px; background: rgba(255,255,255,0.16); border: 1px solid rgba(255,255,255,0.24); color: white; font-size: 18px; font-weight: 800; line-height: 1.35; }}
@media (max-width: 1100px) {{ .champion-grid, .story-grid {{ grid-template-columns: 1fr; }} .premium-kpi-grid {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }} .champion-title {{ flex-direction: column; }} }}
@media (max-width: 720px) {{ .premium-kpi-grid {{ grid-template-columns: 1fr; }} .podium-card {{ grid-template-columns: 44px 1fr; }} .podium-score {{ grid-column: 2; }} }}

/* ========================= V11.4.3 AI + FINAL CELEBRATION ========================= */
.ai-poster-block {{
    margin: 22px 0 24px 0;
    padding: 18px;
    border-radius: 26px;
    background: rgba(255,255,255,0.12);
    border: 1px solid rgba(255,255,255,0.24);
    box-shadow: 0 14px 34px rgba(0,0,0,0.24);
}}
.ai-poster-title {{
    color: white;
    font-size: clamp(22px, 2.8vw, 34px);
    font-weight: 950;
    margin: 0 0 8px 0;
}}
.ai-poster-subtitle {{
    color: rgba(255,255,255,0.82);
    font-weight: 750;
    margin-bottom: 14px;
}}
.ai-placeholder {{
    min-height: 260px;
    border-radius: 22px;
    border: 2px dashed rgba(255,255,255,0.35);
    background: rgba(255,255,255,0.09);
    display: flex;
    align-items: center;
    justify-content: center;
    text-align: center;
    color: rgba(255,255,255,0.82);
    font-weight: 900;
    padding: 24px;
}}
.special-grid {{
    display: grid;
    grid-template-columns: repeat(4, minmax(0, 1fr));
    gap: 14px;
    margin-top: 18px;
}}
.special-card {{
    border-radius: 22px;
    padding: 18px;
    background: rgba(255,255,255,0.12);
    border: 1px solid rgba(255,255,255,0.22);
    min-height: 150px;
}}
.special-card h3 {{ color: white !important; margin: 0 0 10px 0; font-size: 16px; }}
.special-card .special-name {{ color: #fff1a8; font-size: 22px; font-weight: 1000; line-height: 1.05; }}
.special-card .special-detail {{ color: rgba(255,255,255,0.84); font-weight: 750; margin-top: 8px; }}
@media (max-width: 1100px) {{ .special-grid {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }} }}
@media (max-width: 720px) {{ .special-grid {{ grid-template-columns: 1fr; }} }}
@media (max-width: 768px) {{ .block-container {{ padding-left: 0.8rem; padding-right: 0.8rem; border-radius: 16px; }} .card-grid-2, .card-grid-3, .card-grid-4 {{ grid-template-columns: 1fr; }} .card {{ min-height: 140px; }} .card h3, .card h1, .card p {{ white-space: normal; }} }}
    </style>
    """, unsafe_allow_html=True
)

# --------------------------------------------------
# CARREGAR DADES
# --------------------------------------------------
excel_mtime = os.path.getmtime(EXCEL_FILE) if os.path.exists(EXCEL_FILE) else 0
data_actualitzacio = obtenir_data_actualitzacio_fitxer(EXCEL_FILE)

df_porra, df_resultats, df_calendari = carregar_dades(EXCEL_FILE, excel_mtime)

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
# V11.2 · DASHBOARD PRE-FINAL AL PRINCIPI
# --------------------------------------------------
# --------------------------------------------------
# V11.4 · DASHBOARD PREMIUM DEL CAMPIÓ
# --------------------------------------------------

# --------------------------------------------------
# V11.4.2 · DASHBOARD PREMIUM NET
# --------------------------------------------------
mostrar_ceremonia_celebration_deluxe(df_ranking, num_participants)

mostrar_dashboard_campio_premium(df_ranking, df_departaments, df_porra, preparar_taula_buida(df_resultats))

mostrar_estadistiques_prefinal(df_porra)
mostrar_prediccions_resultat_final(df_porra)

# --------------------------------------------------
# MOVIMENTS DESTACATS RETIRATS A LA V11.4.9
# --------------------------------------------------

# --------------------------------------------------
# PODI DE DEPARTAMENTS LÍDERS
# --------------------------------------------------
if te_departaments:
    st.subheader("🏢 TOP 3 Departaments")
    html_top_dep = "<div class='card-grid-3'>"
    for grup in obtenir_grups_podi(df_departaments, nom_col="Departament", punts_col="Mitjana_punts", pos_col="Posició", max_pos=3):
        html_top_dep += html_card_podi_departaments(grup)
    st.markdown(html_top_dep + "</div>", unsafe_allow_html=True)


# --------------------------------------------------
# TOP 3 GENERAL
# --------------------------------------------------
st.subheader("🥇 TOP 3 General")
html_top3 = "<div class='card-grid-3'>"
for grup in obtenir_grups_podi(df_ranking, nom_col="Participant", punts_col="Punts", pos_col="Posició", max_pos=3):
    files = grup["files"]
    deps = []
    if "Departament" in files.columns:
        deps = [str(d) for d in files["Departament"].dropna().astype(str).unique() if str(d).strip()]
    extra = " · ".join(deps) if deps else "punts"
    html_top3 += html_card_podi_participants(grup, extra=extra)
st.markdown(html_top3 + "</div>", unsafe_allow_html=True)


# --------------------------------------------------
# CLASSIFICACIÓ GENERAL
# --------------------------------------------------
st.subheader("📊 Classificació general")
mostrar_taula_ranking(df_ranking)

mostrar_evolucio_tots_participants(df_porra)

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
        mostrar_prediccions_eliminatoria_participant(df_j, df_resultats, df_calendari)
else:
    st.info("Selecciona un participant per veure el detall de punts i prediccions.")

# --------------------------------------------------
# COMPETICIÓ PER DEPARTAMENTS
# --------------------------------------------------
st.subheader("🏢 Competició per departaments")
if te_departaments:
    st.write("Rànquing calculat per **mitjana de punts** del departament. També es mostren punts totals, millor puntuació i líder del departament.")
    mostrar_taula_departaments(df_departaments)
    mostrar_evolucio_departaments(df_porra)
    
    departament_sel = st.selectbox("Selecciona departament", sorted(df_ranking["Departament"].dropna().astype(str).unique()), index=None, placeholder="Selecciona un departament...")
    if departament_sel:
        df_dep_individual = recalcular_posicions(df_ranking[df_ranking["Departament"] == departament_sel].copy())
        st.write(f"### 🥇 TOP 3 · {departament_sel}")
        html_top_dep_ind = "<div class='card-grid-3'>"
        for grup in obtenir_grups_podi(df_dep_individual, nom_col="Participant", punts_col="Punts", pos_col="Posició", max_pos=3):
            html_top_dep_ind += html_card_podi_participants(grup, extra=departament_sel)
        st.markdown(html_top_dep_ind + "</div>", unsafe_allow_html=True)
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
# CALENDARI I RESULTATS DELS PARTITS
# --------------------------------------------------
st.subheader("📅 Calendari i Resultats dels Partits")

if not df_calendari.empty:
    if "Data" in df_calendari.columns:
        df_calendari["Data"] = pd.to_datetime(df_calendari["Data"], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")
        
    df_calendari["Resultat_Net"] = df_calendari["Resultat"].astype(str).str.strip().str.lower()
    
    df_finalitzats = df_calendari[~df_calendari["Resultat_Net"].isin(["pendent", "nan", ""])]
    df_pendents = df_calendari[df_calendari["Resultat_Net"].isin(["pendent", "nan", ""])]
    
    tab_recents, tab_proxims = st.tabs(["✅ Resultats Registrats", "🔜 Pròxims Partits"])
    
    with tab_recents:
        if not df_finalitzats.empty:
            cols_fin = [c for c in ["Fase", "Partit", "Data", "Resultat"] if c in df_finalitzats.columns]
            st.dataframe(df_finalitzats[cols_fin], use_container_width=True, hide_index=True)
        else:
            st.info("Encara no hi ha resultats registrats.")
            
    with tab_proxims:
        if not df_pendents.empty:
            cols_pend = [c for c in ["Fase", "Partit", "Data", "Hora"] if c in df_pendents.columns]
            st.dataframe(df_pendents[cols_pend], use_container_width=True, hide_index=True)
        else:
            st.info("No hi ha partits pendents.")
else:
    st.info("Afegeix una pestanya 'Calendari' al teu Excel per veure els propers partits i resultats.")


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
    if grups: 
        st.dataframe(pd.DataFrame(grups).reindex(["1r", "2n", "3r"]).reset_index().rename(columns={"index": "Posició"}), use_container_width=True, hide_index=True)
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
