#!/usr/bin/env python3
"""
Collecte de prospects (restaurants Metz) via l'API Pappers v2.

COMPORTEMENT
------------
- Ajoute 5 NOUVEAUX restaurants a chaque execution (NB_PAR_EXECUTION).
- Ne reprend jamais un restaurant deja present dans data/restaurants_metz.csv
  (deduplication par SIREN).
- Colonne "date_ajout" = date du jour d'ajout de chaque restaurant.
- Genere data/restaurants_metz.xlsx : les 5 lignes ajoutees pendant
  l'execution sont en VERT, toutes les lignes precedentes sont sans couleur.
  (Un CSV ne peut pas contenir de couleurs : c'est pour cela qu'un .xlsx
  est genere a cote. Le CSV reste la base de donnees / source de dedup.)

COUT (mesure)
-------------
  /recherche   = 2 credits par appel (20 resultats)
  /entreprise  = 1 credit par societe (pour obtenir les dirigeants)
  -> 5 restaurants = 2 + 5 = 7 credits par execution (9 si une 2e page
     de recherche est necessaire).
Pour supprimer le cout du detail : ENRICHIR_ENTREPRISE = False
(colonne dirigeants alors vide).

SORTIES
-------
data/restaurants_metz.csv   : base de prospection (append)
data/restaurants_metz.xlsx  : meme contenu, nouvelles lignes en vert
data/state.json             : page courante
data/debug_pappers.json     : (si DEBUG) reponses brutes
"""

import csv
import json
import os
import sys
import time
from datetime import date
from pathlib import Path

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ===========================================================================
# CONFIGURATION
# ===========================================================================

# Laisser le placeholder : la cle vient du secret GitHub PAPPERS_API_KEY
CLE_API = "COLLE_TA_CLE_PAPPERS_ICI"

CODE_NAF = "56.10A"          # 56.10A = restauration traditionnelle
CODE_POSTAL = "57000"        # Metz
NB_PAR_EXECUTION = 5         # 5 nouveaux restaurants par run

EXCLURE_RADIEES = True       # ignore les societes radiees / inactives
SIEGE_DANS_LA_ZONE = True    # exige que le SIEGE soit au CODE_POSTAL
ENRICHIR_ENTREPRISE = True   # appel /entreprise pour obtenir les dirigeants
DEBUG = False                # True : ecrit debug_pappers.json

COULEUR_NOUVEAU = "C6EFCE"   # vert clair (style "Bon" d'Excel)

# ===========================================================================

API = "https://api.pappers.fr/v2"
KEY = os.environ.get("PAPPERS_API_KEY") or CLE_API

if not KEY or KEY.startswith("COLLE_TA_CLE"):
    sys.exit(
        "Cle API manquante : definis le secret PAPPERS_API_KEY "
        "ou renseigne CLE_API en haut du fichier."
    )

HEADERS = {"api-key": KEY}
DATA = Path("data")
CSV_PATH = DATA / "restaurants_metz.csv"
XLSX_PATH = DATA / "restaurants_metz.xlsx"
STATE_PATH = DATA / "state.json"
DEBUG_PATH = DATA / "debug_pappers.json"

CHAMPS = [
    "date_ajout", "siren", "siret_siege", "nom", "forme_juridique",
    "code_naf", "libelle_naf", "date_creation", "adresse", "code_postal",
    "ville", "effectif", "statut", "dirigeants",
]

cout = 0        # estimation en credits
debug = {}


def get(endpoint, prix, **params):
    global cout
    r = requests.get(f"{API}/{endpoint}", headers=HEADERS,
                     params=params, timeout=30)

    if r.status_code == 401:
        sys.exit("Cle API refusee (401) : verifie qu'elle est active sur Pappers.")
    if r.status_code == 404:
        print(f"[info] endpoint {endpoint} introuvable (404).")
        return None
    if r.status_code == 429:
        sys.exit("Quota depasse (429) : plus de credits ou trop de requetes.")
    if r.status_code >= 400:
        sys.exit(f"Erreur {r.status_code} sur {endpoint} : {r.text[:300]}")

    cout += prix
    return r.json()


# ---------------------------------------------------------------------------
# Etat et deduplication
# ---------------------------------------------------------------------------

def charger_state():
    if STATE_PATH.exists():
        try:
            return json.loads(STATE_PATH.read_text())
        except json.JSONDecodeError:
            print("[warn] state.json illisible, on repart de la page 1.")
    return {"page": 1}


def lire_csv():
    """Retourne toutes les lignes deja presentes dans le CSV."""
    if not CSV_PATH.exists():
        return []
    try:
        with CSV_PATH.open(encoding="utf-8", newline="") as f:
            return [dict(l) for l in csv.DictReader(f)]
    except Exception as e:
        print(f"[warn] lecture du CSV impossible ({e}), pas de deduplication.")
        return []


# ---------------------------------------------------------------------------
# Detection des dirigeants
# ---------------------------------------------------------------------------

INDICES_PERSONNE = {
    "nom", "prenom", "nom_complet", "qualite", "denomination",
    "prenom_usuel", "nom_usage", "fonction",
}


def ressemble_a_une_personne(d):
    return isinstance(d, dict) and bool(INDICES_PERSONNE & set(d.keys()))


def nom_personne(d):
    for cle in ("nom_complet", "denomination", "nom_usage"):
        if d.get(cle):
            return str(d[cle]).strip()
    morceaux = [d.get("prenom") or d.get("prenom_usuel"), d.get("nom")]
    return " ".join(str(m) for m in morceaux if m).strip()


def formater(liste):
    out = []
    for d in liste:
        if not ressemble_a_une_personne(d):
            continue
        nom = nom_personne(d)
        if not nom:
            continue
        qualite = d.get("qualite") or d.get("fonction") or ""
        entree = " | ".join(x for x in [nom, str(qualite)] if x)
        if entree not in out:
            out.append(entree)
    return " ; ".join(out)


def dirigeants_de(item, tracer=False):
    """
    Cherche uniquement dans les champs 'representants' / 'dirigeants'
    (evite de prendre par erreur les conventions collectives, etc.).
    """
    for cle, valeur in item.items():
        if not isinstance(valeur, list) or not valeur:
            continue
        if "representant" not in cle and "dirigeant" not in cle:
            continue
        texte = formater(valeur)
        if texte:
            if tracer:
                print(f"[diagnostic] dirigeants dans le champ '{cle}'")
            return texte
    if tracer:
        listes = [k for k, v in item.items() if isinstance(v, list)]
        print("[diagnostic] aucun dirigeant trouve.")
        print(f"[diagnostic] champs de type liste : {listes}")
    return ""


# ---------------------------------------------------------------------------
# Filtres et extraction
# ---------------------------------------------------------------------------

def est_active(item):
    statut = (item.get("statut_consolide") or "").lower()
    rcs = (item.get("statut_rcs") or "").lower()
    if "inactif" in statut or "radi" in rcs:
        return False
    return not item.get("entreprise_cessee")


def extraire(item):
    siege = item.get("siege") or {}
    return {
        "date_ajout": date.today().isoformat(),
        "siren": item.get("siren", ""),
        "siret_siege": siege.get("siret", "") or item.get("siret", ""),
        "nom": (item.get("nom_entreprise")
                or item.get("denomination")
                or item.get("nom_complet", "")),
        "forme_juridique": item.get("forme_juridique", ""),
        "code_naf": item.get("code_naf", "") or siege.get("code_naf", ""),
        "libelle_naf": item.get("libelle_code_naf", ""),
        "date_creation": (item.get("date_creation_formate", "")
                          or item.get("date_creation", "")),
        "adresse": (siege.get("adresse_ligne_1", "")
                    or item.get("adresse_ligne_1", "")),
        "code_postal": siege.get("code_postal", "") or item.get("code_postal", ""),
        "ville": siege.get("ville", "") or item.get("ville", ""),
        "effectif": item.get("effectif", ""),
        "statut": item.get("statut_consolide", "") or item.get("statut_rcs", ""),
        "dirigeants": dirigeants_de(item),
    }


# ---------------------------------------------------------------------------
# Export Excel : nouvelles lignes en vert
# ---------------------------------------------------------------------------

def ecrire_xlsx(lignes, sirens_nouveaux):
    """Reecrit tout le fichier .xlsx ; seules les lignes de ce run sont vertes."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Restaurants Metz"

    police = Font(name="Arial", size=10)
    police_titre = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    fond_titre = PatternFill("solid", start_color="1F3864", end_color="1F3864")
    vert = PatternFill("solid", start_color=COULEUR_NOUVEAU,
                       end_color=COULEUR_NOUVEAU)

    ws.append(CHAMPS)
    for c in ws[1]:
        c.font = police_titre
        c.fill = fond_titre
        c.alignment = Alignment(horizontal="center", vertical="center")

    for ligne in lignes:
        ws.append([ligne.get(c, "") for c in CHAMPS])
        nouveau = ligne.get("siren") in sirens_nouveaux
        for c in ws[ws.max_row]:
            c.font = police
            c.alignment = Alignment(vertical="top")
            if nouveau:
                c.fill = vert

    # SIREN / SIRET en texte pour ne pas perdre de zeros ni passer en notation
    # scientifique
    for nom_col in ("siren", "siret_siege"):
        idx = CHAMPS.index(nom_col) + 1
        for row in ws.iter_rows(min_row=2, min_col=idx, max_col=idx):
            row[0].number_format = "@"
            row[0].value = str(row[0].value or "")

    largeurs = {
        "date_ajout": 12, "siren": 12, "siret_siege": 16, "nom": 32,
        "forme_juridique": 30, "code_naf": 9, "libelle_naf": 26,
        "date_creation": 13, "adresse": 36, "code_postal": 11, "ville": 20,
        "effectif": 24, "statut": 9, "dirigeants": 50,
    }
    for i, nom_col in enumerate(CHAMPS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = largeurs.get(nom_col, 15)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(XLSX_PATH)


# ---------------------------------------------------------------------------

def main():
    DATA.mkdir(exist_ok=True)
    state = charger_state()
    existantes = lire_csv()
    vus = {l["siren"] for l in existantes if l.get("siren")}
    page = state.get("page", 1)

    retenues = []
    tours = 0
    trace_faite = False
    page_epuisee = True          # la page courante a-t-elle ete lue en entier ?

    while len(retenues) < NB_PAR_EXECUTION and tours < 10:
        tours += 1
        res = get(
            "recherche", 2,
            code_naf=CODE_NAF,
            code_postal=CODE_POSTAL,
            entreprise_cessee="false",
            par_page=20,
            page=page,
        )

        if not res:
            break
        resultats = res.get("resultats", [])
        if not resultats:
            print(f"Plus de resultats (page {page}), arret.")
            break

        if DEBUG and "recherche" not in debug:
            debug["recherche"] = resultats[0]

        page_epuisee = True
        for item in resultats:
            if len(retenues) >= NB_PAR_EXECUTION:
                # il reste des resultats non traites sur cette page :
                # on ne passera pas a la page suivante (ils seront repris
                # au prochain run, les doublons etant ignores).
                page_epuisee = False
                break

            siren = item.get("siren")
            if not siren or siren in vus:
                continue

            if EXCLURE_RADIEES and not est_active(item):
                print(f"[skip] {item.get('nom_entreprise', siren)} : radiee")
                continue

            ligne = extraire(item)

            if SIEGE_DANS_LA_ZONE and ligne["code_postal"] != CODE_POSTAL:
                print(f"[skip] {ligne['nom'] or siren} : siege hors {CODE_POSTAL} "
                      f"({ligne['code_postal']} {ligne['ville']})")
                continue

            # Detail : 1 credit, seule source fiable des dirigeants
            if ENRICHIR_ENTREPRISE and not ligne["dirigeants"]:
                detail = get("entreprise", 1, siren=siren)
                if detail:
                    if DEBUG and "entreprise" not in debug:
                        debug["entreprise"] = detail
                    ligne["dirigeants"] = dirigeants_de(
                        detail, tracer=not trace_faite
                    )
                    trace_faite = True
                    # completer les champs vides avec le detail
                    for cle, valeur in extraire(detail).items():
                        if cle == "date_ajout":
                            continue
                        if not ligne.get(cle) and valeur:
                            ligne[cle] = valeur
                time.sleep(0.3)

            retenues.append(ligne)
            vus.add(siren)

        if page_epuisee:
            page += 1

    if not retenues:
        print("Aucune nouvelle entreprise retenue - rien n'est ecrit.")
        print(f"Cout estime : {cout} credits")
        return

    # --- CSV : ajout des nouvelles lignes ---------------------------------
    nouveau_fichier = not CSV_PATH.exists()
    with CSV_PATH.open("a", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=CHAMPS)
        if nouveau_fichier:
            w.writeheader()
        w.writerows(retenues)

    # --- XLSX : tout l'historique, seules les nouvelles lignes en vert ----
    sirens_nouveaux = {e["siren"] for e in retenues}
    ecrire_xlsx(existantes + retenues, sirens_nouveaux)

    state["page"] = page
    STATE_PATH.write_text(json.dumps(state, indent=2), encoding="utf-8")

    if DEBUG and debug:
        DEBUG_PATH.write_text(
            json.dumps(debug, indent=2, ensure_ascii=False), encoding="utf-8"
        )

    print(f"\n{len(retenues)} entreprise(s) ajoutee(s) le "
          f"{date.today().isoformat()} :\n")
    for e in retenues:
        print(f"  {e['nom']} ({e['siren']})")
        print(f"    {e['adresse']}, {e['code_postal']} {e['ville']}")
        print(f"    dirigeants : {e['dirigeants'] or 'non renseignes'}")
    sans = sum(1 for e in retenues if not e["dirigeants"])
    if sans:
        print(f"\n[warn] {sans} ligne(s) sans dirigeant.")
    print(f"\nFichiers : {CSV_PATH} et {XLSX_PATH} (nouvelles lignes en vert)")
    print(f"Cout estime : {cout} credits")
    print(f"Prochaine execution : page {page}")


if __name__ == "__main__":
    main()
