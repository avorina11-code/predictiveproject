# -*- coding: utf-8 -*-
"""
=============================================================================
 CONTROL ROOM WFM — PILOTAGE PRODUCTION TEMPS RÉEL & PRÉDICTIF
=============================================================================
Application Streamlit (fichier unique autonome) :
  - Moteur métier WFM (Cumuls, Projections, Point de Non-Retour, Diagnostic IA)
  - Composants Control Room (Bannière animée, Jauge dynamique, Projections)
  - Modèles téléchargeables & Tableau de bord interactif
=============================================================================
"""

import re
import io
import json
import string
import datetime as dt

import numpy as np
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# =============================================================================
# 1. CONFIGURATION GÉNÉRALE & STYLES CSS RESPONSIVES
# =============================================================================
st.set_page_config(
    page_title="Control Room WFM — Pilotage Temps Réel",
    page_icon="📡",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .stApp { background-color: #070b12; }
    section[data-testid="stSidebar"] { background-color: #0c111c; }
    h1, h2, h3, h4 { color: #e6f1ff !important; font-family: 'Segoe UI', sans-serif; }
    .stMetric { background-color: #0f1626; border-radius: 12px; padding: 10px; }
    div[data-testid="stMetricValue"] { color: #00e5ff; }
    hr { border-color: #1c2536; }
    .bloc-titre {
        color:#8fa3c7; font-size:13px; letter-spacing:2px; text-transform:uppercase;
        font-weight:700; margin-top:18px; margin-bottom:6px;
    }
    /* Correctif clé pour la réactivité des composants HTML/JS (pas de zoom nécessaire) */
    iframe {
        width: 100% !important;
        border: none !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# =============================================================================
# 2. UTILITAIRES DE NORMALISATION DES DONNÉES
# =============================================================================
def normaliser_nom_colonne(col: str) -> str:
    col = str(col).strip().upper()
    remplacements = {
        "É": "E", "È": "E", "Ê": "E", "Ë": "E",
        "À": "A", "Â": "A", "Ô": "O", "Î": "I", "Ï": "I", "Ù": "U", "Û": "U", "Ç": "C",
    }
    for a, b in remplacements.items():
        col = col.replace(a, b)
    col = col.replace("%", "PCT")
    col = re.sub(r"[^A-Z0-9]+", "_", col)
    col = re.sub(r"_+", "_", col).strip("_")
    return col


MAP_COLONNES_PROD = {
    "TRANCHE": "tranche",
    "RECUS": "recus",
    "TRAITES": "traites",
    "PREVISION": "prevision",
    "TRP_PCT": "trp",
    "QS_PCT": "qs",
    "SL_PCT": "sl",
    "DMC_S": "dmc",
    "ACW_S": "acw",
    "DMT_S": "dmt",
    "CONNECTES": "connectes",
    "EN_TRAIT": "en_traitement",
    "DISPO_PCT": "dispo",
    "ABAND_MOY": "aband_moy",
    "BESOIN": "besoin",
    "PLANNING": "planning",
    "ECART": "ecart",
}

MAP_COLONNES_PLANNING = {
    "LOGIN_VOCALCOM": "login",
    "CODE_RH": "code_rh",
    "NOM_PRENOM": "nom_prenom",
    "HEURE_DEBUT": "heure_debut",
    "HEURE_FIN": "heure_fin",
    "PAUSE_DEBUT": "pause_debut",
    "PAUSE_FIN": "pause_fin",
}


def parse_heure(val):
    if val is None:
        return None
    if isinstance(val, float) and np.isnan(val):
        return None
    if isinstance(val, pd.Timestamp):
        return val.time()
    if isinstance(val, dt.datetime):
        return val.time()
    if isinstance(val, dt.time):
        return val
    s = str(val).strip()
    if s == "" or s.lower() in ("nan", "none", "nat"):
        return None
    m = re.search(r"(\d{1,2})[:hH](\d{2})", s)
    if m:
        h, mi = int(m.group(1)), int(m.group(2))
        if 0 <= h <= 23 and 0 <= mi <= 59:
            return dt.time(h, mi)
    return None


def normaliser_pourcentage(serie: pd.Series) -> pd.Series:
    serie = pd.to_numeric(serie, errors="coerce")
    if serie.dropna().empty:
        return serie.fillna(0.0)
    if serie.dropna().max() <= 1.5:
        serie = serie * 100.0
    return serie.fillna(0.0)


def charger_fichier(uploaded_file, map_colonnes: dict) -> pd.DataFrame:
    if uploaded_file.name.lower().endswith(".csv"):
        df = pd.read_csv(uploaded_file, sep=None, engine="python")
    else:
        df = pd.read_excel(uploaded_file)
    df.columns = [normaliser_nom_colonne(c) for c in df.columns]
    df = df.rename(columns={k: v for k, v in map_colonnes.items() if k in df.columns})
    return df


# =============================================================================
# 3. GÉNÉRATEURS DE MODÈLES EXCEL (TEMPLATES)
# =============================================================================
_HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=10)
_HEADER_FILL = PatternFill("solid", fgColor="1F2937")
_EXEMPLE_FONT = Font(name="Arial", italic=True, color="6B7280", size=10)
_LEGENDE_TITRE_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=10)
_LEGENDE_TITRE_FILL = PatternFill("solid", fgColor="374151")
_BORDURE = Border(*(Side(style="thin", color="D1D5DB"),) * 4)


def _entete_feuille(ws, colonnes: list):
    for j, col in enumerate(colonnes, start=1):
        cell = ws.cell(row=1, column=j, value=col)
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = _BORDURE
        ws.column_dimensions[get_column_letter(j)].width = max(13, len(col) + 4)
    ws.row_dimensions[1].height = 28


def _ligne_exemple(ws, valeurs: list, ligne: int = 2):
    for j, val in enumerate(valeurs, start=1):
        cell = ws.cell(row=ligne, column=j, value=val)
        cell.font = _EXEMPLE_FONT
        cell.border = _BORDURE


def _feuille_legende(wb, lignes: list, titre_feuille: str = "Légende"):
    ws = wb.create_sheet(titre_feuille)
    entetes = ["Colonne", "Description"]
    for j, e in enumerate(entetes, start=1):
        cell = ws.cell(row=1, column=j, value=e)
        cell.font = _LEGENDE_TITRE_FONT
        cell.fill = _LEGENDE_TITRE_FILL
        cell.border = _BORDURE
    for i, (col, desc) in enumerate(lignes, start=2):
        c1 = ws.cell(row=i, column=1, value=col)
        c2 = ws.cell(row=i, column=2, value=desc)
        c1.font = Font(name="Arial", bold=True, size=10)
        c2.font = Font(name="Arial", size=10)
        c1.alignment = Alignment(vertical="top")
        c2.alignment = Alignment(vertical="top", wrap_text=True)
        c1.border = _BORDURE
        c2.border = _BORDURE
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 85
    ws.row_dimensions[1].height = 22


def generer_modele_production() -> bytes:
    colonnes = [
        "TRANCHE", "REÇUS", "TRAITÉS", "PRÉVISION", "TRP %", "QS %", "SL %",
        "DMC (S)", "ACW (S)", "DMT (S)", "CONNECTÉS", "EN TRAIT.", "DISPO %",
        "ABAND. MOY", "BESOIN", "PLANNING", "ÉCART",
    ]
    exemple = ["08:00", 62, 60, 58, 96.5, 92.0, 88.0, 250, 14, 265, 20, 18, 78.0, 3.2, 21, 21, -1]

    wb = Workbook()
    ws = wb.active
    ws.title = "Production"
    _entete_feuille(ws, colonnes)
    _ligne_exemple(ws, exemple)
    ws.freeze_panes = "A2"

    legende = [
        ("TRANCHE", "Heure de début de la tranche de 30 min, format HH:MM."),
        ("REÇUS", "Nombre d'appels reçus."),
        ("TRAITÉS", "Nombre d'appels traités."),
        ("PRÉVISION", "Volume prévisionnel d'appels."),
        ("TRP %", "Taux de Réponse (%)."),
        ("QS %", "Qualité de Service (%)."),
        ("SL %", "Service Level (%)."),
        ("DMC (S)", "Durée Moyenne de Conversation en secondes."),
        ("ACW (S)", "After Call Work en secondes."),
        ("DMT (S)", "Durée Moyenne de Traitement en secondes."),
        ("CONNECTÉS", "Agents connectés."),
        ("EN TRAIT.", "Agents en communication."),
        ("DISPO %", "Taux de disponibilité."),
        ("ABAND. MOY", "Durée moyenne avant abandon (s)."),
        ("BESOIN", "Effectif requis Erlang."),
        ("PLANNING", "Effectif planifié."),
        ("ÉCART", "Écart Connectés - Planning."),
    ]
    _feuille_legende(wb, legende)

    tampon = io.BytesIO()
    wb.save(tampon)
    return tampon.getvalue()


def generer_modele_planning() -> bytes:
    colonnes = ["Login_Vocalcom", "Code_RH", "Nom_Prenom", "Heure_Debut", "Heure_Fin", "Pause_Debut", "Pause_Fin"]
    exemples = [
        ["AG001", "RH1000", "Dupont Marie", "08:00", "17:00", "12:00", "12:45"],
        ["AG002", "RH1001", "Martin Julien", "07:30", "16:30", "11:30", "12:15"],
    ]

    wb = Workbook()
    ws = wb.active
    ws.title = "Planning"
    _entete_feuille(ws, colonnes)
    for i, ex in enumerate(exemples, start=2):
        _ligne_exemple(ws, ex, ligne=i)
    ws.freeze_panes = "A2"

    legende = [
        ("Login_Vocalcom", "Identifiant téléphonie de l'agent."),
        ("Code_RH", "Matricule RH."),
        ("Nom_Prenom", "Nom et prénom."),
        ("Heure_Debut", "Heure de début de poste."),
        ("Heure_Fin", "Heure de fin de poste."),
        ("Pause_Debut", "Début pause déjeuner."),
        ("Pause_Fin", "Fin pause déjeuner."),
    ]
    _feuille_legende(wb, legende)

    tampon = io.BytesIO()
    wb.save(tampon)
    return tampon.getvalue()


# =============================================================================
# 4. JEU DE DONNÉES DE DÉMONSTRATION
# =============================================================================
def generer_demo_production() -> pd.DataFrame:
    heures = pd.date_range("08:00", "17:30", freq="30min").time
    rng = np.random.default_rng(42)
    lignes = []
    dmt_cible_demo = 280
    for i, h in enumerate(heures):
        prevision = int(60 + 25 * np.sin(i / 3) + rng.integers(-5, 5))
        prevision = max(prevision, 20)
        derive_dmt = 1.0 if i < 6 else 1.0 + min(0.35, (i - 5) * 0.05)
        dmt = dmt_cible_demo * derive_dmt + rng.integers(-8, 8)
        recus = int(prevision * rng.uniform(0.92, 1.12))
        planning = int(prevision / 3.2) + rng.integers(-1, 2)
        connectes = max(planning - (2 if i >= 6 else 0) - rng.integers(0, 2), 1)
        traites = int(recus * rng.uniform(0.9, 1.0))
        sl = max(30, min(98, 95 - (dmt - dmt_cible_demo) * 0.22 + rng.integers(-3, 3)))
        qs = max(50, min(99, sl + rng.integers(-4, 4)))
        trp = max(50, min(100, qs + rng.integers(-2, 2)))
        dispo = max(40, min(95, 80 - (dmt - dmt_cible_demo) * 0.1))
        lignes.append(
            {
                "TRANCHE": h.strftime("%H:%M"),
                "REÇUS": recus,
                "TRAITÉS": traites,
                "PRÉVISION": prevision,
                "TRP %": round(trp, 1),
                "QS %": round(qs, 1),
                "SL %": round(sl, 1),
                "DMC (S)": round(dmt * 0.9, 0),
                "ACW (S)": round(rng.uniform(10, 22), 0),
                "DMT (S)": round(dmt, 0),
                "CONNECTÉS": connectes,
                "EN TRAIT.": max(connectes - rng.integers(0, 2), 0),
                "DISPO %": round(dispo, 1),
                "ABAND. MOY": round(rng.uniform(2, 8), 1),
                "BESOIN": planning + rng.integers(0, 2),
                "PLANNING": planning,
                "ÉCART": connectes - planning,
            }
        )
    return pd.DataFrame(lignes)


def generer_demo_planning(n_agents: int = 28) -> pd.DataFrame:
    rng = np.random.default_rng(7)
    lignes = []
    debuts_possibles = ["07:30", "08:00", "08:30", "09:00", "10:00"]
    fins_possibles = ["16:30", "17:00", "17:30", "18:00", "19:00"]
    pauses_possibles = ["11:30", "12:00", "12:30", "13:00", "13:30"]
    for i in range(n_agents):
        hd = rng.choice(debuts_possibles)
        hf = rng.choice(fins_possibles)
        pd_ = rng.choice(pauses_possibles)
        h, m = map(int, pd_.split(":"))
        fin_pause = dt.time(h, m + 45) if m + 45 < 60 else dt.time(h + 1, m + 45 - 60)
        lignes.append(
            {
                "Login_Vocalcom": f"AG{i+1:03d}",
                "Code_RH": f"RH{1000+i}",
                "Nom_Prenom": f"Agent {i+1:02d}",
                "Heure_Debut": hd,
                "Heure_Fin": hf,
                "Pause_Debut": pd_,
                "Pause_Fin": fin_pause.strftime("%H:%M"),
            }
        )
    return pd.DataFrame(lignes)


# =============================================================================
# 5. BARRE LATÉRALE — CONFIGURATION & PARAMÈTRES
# =============================================================================
with st.sidebar:
    st.markdown("## ⚙️ Configuration des objectifs")

    obj_sl = st.slider("Objectif SL — Service Level (%)", 50, 100, 85)
    seuil_reponse = st.number_input("Seuil de réponse SL (secondes)", min_value=5, value=20, step=5)
    obj_qs = st.slider("Objectif QS / TRP (%)", 50, 100, 95)
    dmt_cible = st.number_input("DMT Cible / Prévue (secondes)", min_value=30, value=280, step=10)
    acw_cible = st.number_input("ACW Cible (secondes)", min_value=0, value=15, step=1)
    capacite_rattrapage = st.slider(
        "Capacité max de rattrapage SL sur le reste de la journée (%)", 50, 100, 95
    ) / 100.0

    st.markdown("---")
    st.markdown("## 📁 Données de production")
    fichier_prod = st.file_uploader(
        "Fichier 1 — Métriques de production (Excel/CSV)",
        type=["xlsx", "xls", "csv"],
        key="fichier_prod",
    )
    st.download_button(
        "📄 Télécharger le modèle — Production",
        data=generer_modele_production(),
        file_name="modele_production_tranches.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_modele_prod",
        use_container_width=True,
    )

    fichier_planning = st.file_uploader(
        "Fichier 2 — Planning agents & pauses (Excel/CSV)",
        type=["xlsx", "xls", "csv"],
        key="fichier_planning",
    )
    st.download_button(
        "📄 Télécharger le modèle — Planning",
        data=generer_modele_planning(),
        file_name="modele_planning_agents.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_modele_planning",
        use_container_width=True,
    )

    mode_demo = fichier_prod is None or fichier_planning is None
    if mode_demo:
        st.info(
            "Mode Démo actif : chargement automatique de données simulées.",
            icon="ℹ️",
        )

    st.markdown("---")
    seuil_alerte_marge = st.slider(
        "Marge (points de SL) déclenchant l'état 'EN DANGER'", 1, 15, 4
    )
    seuil_derive_dmt_pct = st.slider("Seuil de dérive DMT significatif (%)", 5, 30, 10) / 100.0
    seuil_derive_flux_pct = st.slider("Seuil de surflux significatif (%)", 5, 30, 10) / 100.0
    seuil_sous_effectif = st.slider("Seuil sous-effectif (nb agents)", 1, 5, 2)


# =============================================================================
# 6. NETTOYAGE DES FICHIERS
# =============================================================================
if fichier_prod is not None:
    df_prod = charger_fichier(fichier_prod, MAP_COLONNES_PROD)
else:
    df_prod = generer_demo_production()
    df_prod.columns = [normaliser_nom_colonne(c) for c in df_prod.columns]
    df_prod = df_prod.rename(columns={k: v for k, v in MAP_COLONNES_PROD.items() if k in df_prod.columns})

if fichier_planning is not None:
    df_planning = charger_fichier(fichier_planning, MAP_COLONNES_PLANNING)
else:
    df_planning = generer_demo_planning()
    df_planning.columns = [normaliser_nom_colonne(c) for c in df_planning.columns]
    df_planning = df_planning.rename(
        columns={k: v for k, v in MAP_COLONNES_PLANNING.items() if k in df_planning.columns}
    )

for col_num in ["recus", "traites", "prevision", "dmc", "acw", "dmt", "connectes",
                 "en_traitement", "aband_moy", "besoin", "planning", "ecart"]:
    if col_num in df_prod.columns:
        df_prod[col_num] = pd.to_numeric(df_prod[col_num], errors="coerce").fillna(0.0)

for col_pct in ["trp", "qs", "sl", "dispo"]:
    if col_pct in df_prod.columns:
        df_prod[col_pct] = normaliser_pourcentage(df_prod[col_pct])

df_prod["heure_tranche"] = df_prod["tranche"].apply(parse_heure)
df_prod = df_prod.dropna(subset=["heure_tranche"]).sort_values("heure_tranche").reset_index(drop=True)
df_prod["tranche"] = df_prod["heure_tranche"].apply(lambda h: h.strftime("%H:%M"))

for col_h in ["heure_debut", "heure_fin", "pause_debut", "pause_fin"]:
    if col_h in df_planning.columns:
        df_planning[col_h] = df_planning[col_h].apply(parse_heure)


# =============================================================================
# 7. MOTEUR PREDICTIF & DIAGNOSTIC
# =============================================================================
def compter_effectif_planning(df_planning: pd.DataFrame, debut: dt.time, fin: dt.time):
    n_planifies, n_pause = 0, 0
    for _, agent in df_planning.iterrows():
        hd, hf = agent.get("heure_debut"), agent.get("heure_fin")
        if hd is None or hf is None:
            continue
        if hd <= debut and hf >= fin:
            n_planifies += 1
            p_deb, p_fin = agent.get("pause_debut"), agent.get("pause_fin")
            if p_deb is not None and p_fin is not None:
                if not (p_fin <= debut or p_deb >= fin):
                    n_pause += 1
    return n_planifies, n_pause


def diagnostiquer_cause_racine(row: pd.Series, dmt_cible, seuil_derive_dmt_pct,
                                seuil_derive_flux_pct, seuil_sous_effectif):
    ecart_dmt_pct = (row["dmt"] - dmt_cible) / dmt_cible if dmt_cible else 0.0
    effectif_dispo = max(row["planning"] - row["agents_en_pause"], 0)
    ecart_effectif = row["connectes"] - effectif_dispo
    ecart_flux_pct = (row["recus"] - row["prevision"]) / row["prevision"] if row["prevision"] else 0.0

    flag_dmt = ecart_dmt_pct > seuil_derive_dmt_pct
    flag_effectif = ecart_effectif < -seuil_sous_effectif
    flag_flux = ecart_flux_pct > seuil_derive_flux_pct

    causes = []
    if flag_dmt:
        causes.append(f"DMT +{ecart_dmt_pct*100:.0f}% vs cible")
    if flag_effectif:
        causes.append(f"Sous-effectif {ecart_effectif:.0f} agent(s) vs planifié")
    if flag_flux:
        causes.append(f"Surflux +{ecart_flux_pct*100:.0f}% vs prévision")

    if sum([flag_dmt, flag_effectif, flag_flux]) >= 2:
        code = "COMBINAISON"
        libelle = "Combinaison de facteurs (" + " + ".join(causes) + ")"
    elif flag_dmt:
        code = "DMT"
        libelle = f"Dérive opérationnelle sur le temps de traitement (DMT +{ecart_dmt_pct*100:.0f}%)"
    elif flag_effectif:
        code = "EFFECTIF"
        libelle = f"Sous-effectif / inadhérence ({ecart_effectif:.0f} agent(s) manquant(s))"
    elif flag_flux:
        code = "FLUX"
        libelle = f"Surflux volumétrique imprévu (+{ecart_flux_pct*100:.0f}%)"
    else:
        code = "NOMINAL"
        libelle = "Situation nominale — aucune dérive significative détectée"

    flags = {"dmt": flag_dmt, "effectif": flag_effectif, "flux": flag_flux}
    return code, libelle, flags


ACTIONS_PRESCRIPTIVES = {
    "DMT": [
        "Activer les consignes d'écourtage DMT sur les motifs longs",
        "Rappel flash aux superviseurs sur la maîtrise du temps de traitement",
        "Basculer les appels complexes vers les agents seniors / experts",
    ],
    "EFFECTIF": [
        "Décaler ou raccourcir les pauses de 15 min sur les 2 prochaines tranches",
        "Rappeler les agents en formation / coaching disponibles",
        "Solliciter du renfort en heures supplémentaires ou back-up",
    ],
    "FLUX": [
        "Activer un message d'attente / SVI de débordement",
        "Basculer une partie du flux vers un canal différé (email, callback)",
        "Solliciter un plateau de débordement si accord multi-site",
    ],
    "COMBINAISON": [
        "Déclencher la cellule de crise pilotage (WFM + Superviseurs)",
        "Prioriser : stabiliser l'effectif AVANT d'agir sur la DMT",
        "Réévaluer l'objectif SL horaire restant en concertation client",
    ],
    "NOMINAL": [
        "Maintenir la vigilance sur les 2 prochaines tranches",
        "Aucune action corrective nécessaire à ce stade",
    ],
}


def moteur_wfm(df_prod: pd.DataFrame, df_planning: pd.DataFrame, obj_sl_pct: float,
               capacite_rattrapage: float, dmt_cible, seuil_alerte_marge,
               seuil_derive_dmt_pct, seuil_derive_flux_pct, seuil_sous_effectif):
    df = df_prod.copy().reset_index(drop=True)

    agents_en_pause = []
    for _, row in df.iterrows():
        debut = row["heure_tranche"]
        fin_dt = (dt.datetime.combine(dt.date.today(), debut) + dt.timedelta(minutes=30)).time()
        _, n_pause = compter_effectif_planning(df_planning, debut, fin_dt)
        agents_en_pause.append(n_pause)
    df["agents_en_pause"] = agents_en_pause

    df["cum_recus"] = df["recus"].cumsum()
    df["conformes_tranche"] = df["recus"] * (df["sl"] / 100.0)
    df["cum_conformes"] = df["conformes_tranche"].cumsum()
    df["sl_cumule"] = np.where(df["cum_recus"] > 0, df["cum_conformes"] / df["cum_recus"] * 100.0, 0.0)

    total_prevision = df["prevision"].sum()
    df["prevision_cumulee"] = df["prevision"].cumsum()
    df["prevision_future"] = total_prevision - df["prevision_cumulee"]

    denominateur = df["cum_recus"] + df["prevision_future"]
    numerateur = df["cum_conformes"] + df["prevision_future"] * capacite_rattrapage
    df["sl_max_projete"] = np.where(denominateur > 0, numerateur / denominateur * 100.0, df["sl_cumule"])

    df["point_non_retour"] = df["sl_max_projete"] < obj_sl_pct

    def determiner_statut(sl_projete):
        if sl_projete < obj_sl_pct:
            return "POINT DE NON RETOUR ATTEINT"
        elif sl_projete < obj_sl_pct + seuil_alerte_marge:
            return "EN DANGER"
        else:
            return "SOUS CONTROLE"

    df["statut"] = df["sl_max_projete"].apply(determiner_statut)

    codes, libelles = [], []
    for _, row in df.iterrows():
        code, libelle, _ = diagnostiquer_cause_racine(
            row, dmt_cible, seuil_derive_dmt_pct, seuil_derive_flux_pct, seuil_sous_effectif
        )
        codes.append(code)
        libelles.append(libelle)
    df["diagnostic_code"] = codes
    df["diagnostic_libelle"] = libelles

    heure_bascule = None
    tranches_bascule = df.loc[df["point_non_retour"], "tranche"]
    if not tranches_bascule.empty:
        heure_bascule = tranches_bascule.iloc[0]

    # FIX IA : Sélection de la dernière tranche ayant du flux réel
    df_realise = df[df["recus"] > 0]
    derniere_ligne = df_realise.iloc[-1] if not df_realise.empty else df.iloc[-1]

    synthese = {
        "heure_bascule": heure_bascule,
        "statut_actuel": derniere_ligne["statut"],
        "sl_cumule_actuel": derniere_ligne["sl_cumule"],
        "sl_max_projete_actuel": derniere_ligne["sl_max_projete"],
        "diagnostic_code_actuel": derniere_ligne["diagnostic_code"],
        "diagnostic_libelle_actuel": derniere_ligne["diagnostic_libelle"],
        "derniere_tranche": derniere_ligne["tranche"],
    }
    return df, synthese


df_result, synthese = moteur_wfm(
    df_prod, df_planning, obj_sl, capacite_rattrapage, dmt_cible, seuil_alerte_marge,
    seuil_derive_dmt_pct, seuil_derive_flux_pct, seuil_sous_effectif,
)

actions_actuelles = ACTIONS_PRESCRIPTIVES.get(synthese["diagnostic_code_actuel"], [])


# =============================================================================
# 8. COMPOSANTS D'INTERFACE HTML/JS (RESPONSIVES)
# =============================================================================
def render_banniere_alerte(statut: str, heure_bascule, sl_actuel: float, sl_projete: float, obj_sl: float):
    if statut == "POINT DE NON RETOUR ATTEINT":
        couleur, icone, anim = "#ff2d55", "🚨", "blink-neon 1.1s infinite"
        message = f"POINT DE NON-RETOUR ATTEINT À {heure_bascule} — OBJECTIF SL {obj_sl:.0f}% INATTEIGNABLE"
    elif statut == "EN DANGER":
        couleur, icone, anim = "#ffae00", "⚠️", "pulse-orange 1.6s infinite"
        message = "SL EN DANGER — MARGE RÉSIDUELLE FAIBLE"
    else:
        couleur, icone, anim = "#00ff88", "✅", "none"
        message = "SITUATION SOUS CONTRÔLE — OBJECTIF SL ATTEIGNABLE"

    tpl = string.Template(
        """
        <!DOCTYPE html>
        <html>
        <head>
        <style>
        body { margin:0; padding:0; background:transparent; }
        @keyframes blink-neon {
            0%, 100% { opacity:1; box-shadow:0 0 15px 4px $couleur; }
            50% { opacity:0.6; box-shadow:0 0 35px 10px $couleur; }
        }
        @keyframes pulse-orange {
            0%, 100% { box-shadow:0 0 8px 2px $couleur; }
            50% { box-shadow:0 0 20px 6px $couleur; }
        }
        .banniere {
            background: linear-gradient(90deg, #0a0e17 0%, ${couleur}22 100%);
            border: 2px solid $couleur;
            border-radius: 12px;
            padding: 14px 20px;
            display: flex; align-items: center; justify-content: space-between;
            font-family: 'Segoe UI', sans-serif; color: #f2f6ff;
            animation: $anim;
            box-sizing: border-box;
        }
        .banniere-txt { font-size: 16px; font-weight: 800; }
        .banniere-sub { font-size: 12px; opacity: 0.85; margin-top: 4px; color:#b9c6de; }
        .banniere-metric { font-size: 24px; font-weight: 900; color: $couleur; text-align:right; }
        .banniere-metric-label { font-size: 10px; color:#8fa3c7; text-align:right; letter-spacing:1px; }
        </style>
        </head>
        <body>
        <div class="banniere">
          <div>
            <div class="banniere-txt">$icone $message</div>
            <div class="banniere-sub">SL cumulé : $sl_actuel% &nbsp;|&nbsp; SL max projeté : $sl_projete% &nbsp;|&nbsp; Objectif : $obj_sl%</div>
          </div>
          <div>
            <div class="banniere-metric">$sl_projete%</div>
            <div class="banniere-metric-label">SL PROJETÉ</div>
          </div>
        </div>
        </body>
        </html>
        """
    )
    html = tpl.substitute(
        couleur=couleur, icone=icone, anim=anim, message=message,
        sl_actuel=f"{sl_actuel:.1f}", sl_projete=f"{sl_projete:.1f}", obj_sl=f"{obj_sl:.0f}",
    )
    components.html(html, height=100)


def render_jauge_sl(sl_actuel: float, sl_projete: float, obj_sl: float):
    tpl = string.Template(
        """
        <!DOCTYPE html>
        <html>
        <head>
          <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
          <style>
            body { margin: 0; padding: 0; background: transparent; }
            .card {
              background: #0a0e17;
              border: 1px solid #1c2536;
              border-radius: 14px;
              padding: 12px;
              box-sizing: border-box;
            }
            .chart-box { position: relative; width: 100%; height: 160px; }
            .kpi-row { display: flex; justify-content: space-around; margin-top: 6px; font-family: 'Segoe UI', sans-serif; }
            .kpi-val { font-size: 18px; font-weight: 800; }
            .kpi-lbl { color: #8fa3c7; font-size: 10px; letter-spacing: 0.5px; }
          </style>
        </head>
        <body>
          <div class="card">
            <div class="chart-box"><canvas id="gaugeChart"></canvas></div>
            <div class="kpi-row">
              <div style="text-align:center;"><div class="kpi-lbl">SL CUMULÉ</div><div class="kpi-val" style="color:#00e5ff;">$sl_actuel%</div></div>
              <div style="text-align:center;"><div class="kpi-lbl">SL PROJETÉ</div><div class="kpi-val" style="color:#ffffff;">$sl_projete%</div></div>
              <div style="text-align:center;"><div class="kpi-lbl">OBJECTIF</div><div class="kpi-val" style="color:#00ff88;">$obj_sl%</div></div>
            </div>
          </div>
          <script>
            const ctx = document.getElementById('gaugeChart').getContext('2d');
            new Chart(ctx, {
              type: 'doughnut',
              data: {
                datasets: [{
                  data: [$sl_projete, Math.max(0, 100 - $sl_projete)],
                  backgroundColor: ['$sl_projete' >= '$obj_sl' ? '#00ff88' : '#ff2d55', '#1c2536'],
                  borderWidth: 0
                }]
              },
              options: {
                rotation: -90, circumference: 180, cutout: '75%',
                responsive: true, maintainAspectRatio: false,
                plugins: { legend: { display: false }, tooltip: { enabled: false } }
              }
            });
          </script>
        </body>
        </html>
        """
    )
    html = tpl.substitute(
        sl_actuel=f"{sl_actuel:.1f}", sl_projete=f"{sl_projete:.1f}", obj_sl=f"{obj_sl:.0f}"
    )
    components.html(html, height=230)


def render_graphique_projections(df: pd.DataFrame, obj_sl: float):
    tranches_json = json.dumps(df["tranche"].tolist())
    sl_reel_json = json.dumps(df["sl_cumule"].tolist())
    sl_proj_json = json.dumps(df["sl_max_projete"].tolist())

    tpl = string.Template(
        """
        <!DOCTYPE html>
        <html>
        <head>
          <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
          <style>
            body { margin: 0; padding: 0; background: transparent; }
            .card { background: #0a0e17; border: 1px solid #1c2536; border-radius: 14px; padding: 12px; }
            .chart-box { position: relative; width: 100%; height: 210px; }
          </style>
        </head>
        <body>
          <div class="card">
            <div class="chart-box"><canvas id="projChart"></canvas></div>
          </div>
          <script>
            const ctx = document.getElementById('projChart').getContext('2d');
            new Chart(ctx, {
              type: 'line',
              data: {
                labels: $tranches,
                datasets: [
                  { label: 'SL Réel Cumulé (%)', data: $sl_reel, borderColor: '#00e5ff', borderWidth: 2, fill: false },
                  { label: 'SL Max Projeté (%)', data: $sl_proj, borderColor: '#ffae00', borderWidth: 2, borderDash: [4, 4], fill: false },
                  { label: 'Objectif SL', data: Array($tranches.length).fill($obj_sl), borderColor: '#00ff88', borderWidth: 1.5, borderDash: [2, 2], pointRadius: 0 }
                ]
              },
              options: {
                responsive: true, maintainAspectRatio: false,
                scales: {
                  y: { min: 0, max: 100, grid: { color: '#1c2536' }, ticks: { color: '#8fa3c7' } },
                  x: { grid: { color: '#1c2536' }, ticks: { color: '#8fa3c7' } }
                },
                plugins: { legend: { labels: { color: '#e6f1ff' } } }
              }
            });
          </script>
        </body>
        </html>
        """
    )
    html = tpl.substitute(
        tranches=tranches_json, sl_reel=sl_reel_json, sl_proj=sl_proj_json, obj_sl=obj_sl
    )
    components.html(html, height=240)


# =============================================================================
# 9. DISPOSITION DE LA PAGE & RENDU D'AFFICHAGE
# =============================================================================
st.title("📡 Control Room WFM — Pilotage Temps Réel")

# 1. Bandeau supérieur d'alerte
render_banniere_alerte(
    synthese["statut_actuel"],
    synthese["heure_bascule"],
    synthese["sl_cumule_actuel"],
    synthese["sl_max_projete_actuel"],
    obj_sl,
)

st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)

# 2. Section Graphiques & Diagnostic IA
col_g1, col_g2 = st.columns([1, 2])

with col_g1:
    st.markdown("<div class='bloc-titre'>Jauge de Survie SL</div>", unsafe_allow_html=True)
    render_jauge_sl(synthese["sl_cumule_actuel"], synthese["sl_max_projete_actuel"], obj_sl)

with col_g2:
    st.markdown("<div class='bloc-titre'>Projection Trajectoire de Fin de Journée</div>", unsafe_allow_html=True)
    render_graphique_projections(df_result, obj_sl)

st.markdown("---")

# 3. Diagnostic IA & Actions
col_ia1, col_ia2 = st.columns([1, 1])

with col_ia1:
    st.markdown("### 🤖 Diagnostic Vigie IA")
    st.info(f"**Tranche :** {synthese['derniere_tranche']}\n\n**Cause :** {synthese['diagnostic_libelle_actuel']}")

with col_ia2:
    st.markdown("### 📋 Plan d'Action Recommandé")
    for act in actions_actuelles:
        st.markdown(f"* 🔹 {act}")

st.markdown("---")

# 4. Tableau Détaillé par Tranche
st.markdown("### 📊 Tableau de Bord Détaillé")

cols_display = ["tranche", "recus", "prevision", "sl", "sl_cumule", "sl_max_projete", "dmt", "connectes", "planning", "statut"]
df_show = df_result[cols_display].copy()
df_show.columns = ["Tranche", "Reçus", "Prévu", "SL Tranche (%)", "SL Cumulé (%)", "SL Max Projeté (%)", "DMT (s)", "Connectés", "Planning", "Statut"]

def colorer_statut(val):
    if val == "POINT DE NON RETOUR ATTEINT":
        return "background-color: #ff2d55; color: white; font-weight: bold;"
    elif val == "EN DANGER":
        return "background-color: #ffae00; color: black; font-weight: bold;"
    return "background-color: #00ff88; color: black; font-weight: bold;"

st.dataframe(
    df_show.style.applymap(colorer_statut, subset=["Statut"])
    .format({
        "SL Tranche (%)": "{:.1f}%",
        "SL Cumulé (%)": "{:.1f}%",
        "SL Max Projeté (%)": "{:.1f}%",
        "DMT (s)": "{:.0f}s"
    }),
    use_container_width=True,
    height=400
)
