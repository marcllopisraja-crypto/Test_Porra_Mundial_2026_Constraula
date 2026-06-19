import streamlit as st
import pandas as pd
import altair as alt
import base64
import os
import json
import unicodedata
import hashlib
import time
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

# ==================================================
# NORMALITZACIÓ I BANDERES
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
    "bosnia and herzegovina": "Bosnia i Hercegovina", "bosnia & herzegovina": "Bosnia i Hercegovina", "bosnia &amp; herzegovina": "Bosnia i Hercegovina", "bosnia i hercegovina": "Bosnia i Hercegovina",
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

# ==================================================
# CÀRREGA EXCEL I UTILITATS
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
    nets, vistos = [], set()
    for valor in valors:
        if valor == "" or normalitzar_text(valor) in ["nan", "nat", "none", "pendent"]:
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
# RESULTATS DES D'EXCEL, SENSE API
# ==================================================
def construir_resultats_des_excel(df_resultats):
    df = preparar_taula_buida(df_resultats)
    resultats = {
        "source": "Excel", "group_positions": {}, "Setzens": [], "Vuitens": [], "Quarts": [], "Semis": [],
        "Finalistes": [], "Campió": [], "MVP": [], "Pichichi": [], "Gols Pichichi": None, "api_error": ""
    }
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
        Participants=("Participant", "count"),
        Punts_totals=("Punts", "sum"),
        Mitjana_punts=("Punts", "mean"),
        Millor_puntuacio=("Punts", "max")
    )
    lider = df.sort_values("Punts", ascending=False).drop_duplicates("Departament")[["Departament", "Participant"]].rename(columns={"Participant": "Líder departament"})
    resum = resum.merge(lider, on="Departament", how="left")
    for col in ["Punts_totals", "Mitjana_punts", "Millor_puntuacio"]:
        resum[col] = pd.to_numeric(resum[col], errors="coerce").fillna(0).round(1)
    resum = resum.sort_values(["Mitjana_punts", "Punts_totals"], ascending=[False, False]).reset_index(drop=True)
    resum["Posició"] = resum.index + 1
    resum["Dif líder"] = (resum["Mitjana_punts"] - float(resum["Mitjana_punts"].iloc[0])).round(1) if not resum.empty else 0.0
    return resum[["Posició", "Departament", "Participants", "Mitjana_punts", "Punts_totals", "Millor_puntuacio", "Líder departament", "Dif líder"]]

def crear_tauler_medalles_departaments(df_ranking):
    if "Departament" not in df_ranking.columns:
        return pd.DataFrame()
    df = df_ranking.copy().sort_values("Posició")
    
    # Comptar presències a diferents trams del rànquing
    top5_counts = df[df["Posició"] <= 5].groupby("Departament")["Participant"].count()
    top10_counts = df[df["Posició"] <= 10].groupby("Departament")["Participant"].count()
    top20_counts = df[df["Posició"] <= 20].groupby("Departament")["Participant"].count()
    
    deps = df["Departament"].unique()
    medalles = []
    for d in deps:
        t5 = int(top5_counts.get(d, 0))
        t10 = int(top10_counts.get(d, 0))
        t20 = int(top20_counts.get(d, 0))
        medalles.append({
            "Departament": d,
            "🥇 En Top 5": t5,
            "🥈 En Top 10": t10,
            "🥉 En Top 20": t20,
            "Total Pòdis": t5 + t10 + t20
        })
    df_medalles = pd.DataFrame(medalles)
    df_medalles = df_medalles.sort_values(by=["🥇 En Top 5", "🥈 En Top 10", "🥉 En Top 20", "Total Pòdis"], ascending=False).reset_index(drop=True)
    df_medalles.insert(0, "Posició Olímpica", df_medalles.index + 1)
    return df_medalles

def ranking_signature(df_ranking, resultats_actuals):
    cols = ["Participant", "Punts", "Posició"]
    if "Departament" in df_ranking.columns:
        cols.append("Departament")
    payload = df_ranking[cols].sort_values("Participant").to_json(orient="records", force_ascii=False)
    payload += json.dumps({
        k: resultats_actuals.get(k) for k in ["source", "Setzens", "Vuitens", "Quarts", "Semis", "Finalistes", "Campió", "Pichichi", "Gols Pichichi"]
    }, ensure_ascii=False, sort_keys=True)
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
    df["Canvi posició"] = df["Canvi num posició"].apply(lambda x: f"🟢 ▲ +{x}" if x > 0 e