#!/usr/bin/env python3
"""
Collecte de prospects (restaurants Metz) via l'API Pappers v2.

COUT REEL (mesure)
------------------
  /recherche                 = 2 credits par appel
  /entreprise                = 1 credit par societe
  /recherche-dirigeants      = 10 credits  -> ABANDONNE, retire du script

Configuration actuelle : 3 prospects par run avec dirigeants
  -> 2 (recherche) + 3 (detail) = 5 credits par execution, ~150 par mois.

Pour reduire : baisser NB_PAR_EXECUTION.
Pour supprimer tout cout de detail : ENRICHIR_ENTREPRISE = False
(2 credits par run, mais colonne dirigeants vide - /recherche ne les fournit pas).

SORTIE
------
data/restaurants_metz.csv  : fichier de prospection (append)
data/state.json            : page courante
data/debug_pappers.json    : (si DEBUG) reponses brutes, pour inspection
"""

import csv
import json
import os
import sys
import time
from datetime import date
from pathlib import Path

import requests

# ===========================================================================
# CONFIGURATION
# ===========================================================================

# Laisser le placeholder : la cle vient du secret GitHub PAPPERS_API_KEY
CLE_API = "COLLE_TA_CLE_PAPPERS_ICI"

CODE_NAF = "56.10A"          # 56.10A = restauration traditionnelle
CODE_POSTAL = "57000"        # Metz
NB_PAR_EXECUTION = 3         # 3 prospects = 5 credits par run

EXCLURE_RADIEES = True       # ignore les societes radiees / inactives
SIEGE_DANS_LA_ZONE = True    # exige que le SIEGE soit au CODE_POSTAL
ENRICHIR_ENTREPRISE = True   # appel /entreprise pour obtenir les dirigeants
DEBUG = True                 # ecrit debug_pappers.json (False ensuite)

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


def sirens_deja_vus():
    if not CSV_PATH.exists():
        return set()
    try:
        with CSV_PATH.open(encoding="utf-8") as f:
            return {l["siren"] for l in csv.DictReader(f) if l.get("siren")}
    except Exception as e:
        print(f"[warn] lecture du CSV impossible ({e}), pas de deduplication.")
        return set()


# ---------------------------------------------------------------------------
# Detection des dirigeants (independante du nom du champ)
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
    """Retient la premiere liste de l'objet qui contient des personnes."""
    for cle, valeur in item.items():
        if isinstance(valeur, list) and valeur:
            if any(ressemble_a_une_personne(x) for x in valeur):
                texte = formater(valeur)
                if texte:
                    if tracer:
                        print(f"[diagnostic] dirigeants dans le champ '{cle}'")
                    return texte
    if tracer:
        listes = [k for k, v in item.items() if isinstance(v, list)]
        print(f"[diagnostic] aucun dirigeant trouve.")
        print(f"[diagnostic] champs disponibles : {sorted(item.keys())}")
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

def main():
    DATA.mkdir(exist_ok=True)
    state = charger_state()
    vus = sirens_deja_vus()
    page = state.get("page", 1)

    retenues = []
    tours = 0
    trace_faite = False

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
        page += 1

        if not res:
            break
        resultats = res.get("resultats", [])
        if not resultats:
            print(f"Plus de resultats (page {page - 1}), arret.")
            break

        if DEBUG and "recherche" not in debug:
            debug["recherche"] = resultats[0]

        for item in resultats:
            if len(retenues) >= NB_PAR_EXECUTION:
                break
            siren = item.get("siren")
            if not siren or siren in vus:
                continue

            if EXCLURE_RADIEES and not est_active(item):
                print(f"[skip] {item.get('nom_entreprise', siren)} : radiee")
                vus.add(siren)
                continue

            ligne = extraire(item)

            if SIEGE_DANS_LA_ZONE and ligne["code_postal"] != CODE_POSTAL:
                print(f"[skip] {ligne['nom'] or siren} : siege hors {CODE_POSTAL} "
                      f"({ligne['code_postal']} {ligne['ville']})")
                vus.add(siren)
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
                        if not ligne.get(cle) and valeur:
                            ligne[cle] = valeur
                time.sleep(0.3)

            retenues.append(ligne)
            vus.add(siren)

    if not retenues:
        print("Aucune nouvelle entreprise retenue - rien n'est ecrit.")
        print(f"Cout estime : {cout} credits")
        return

    nouveau = not CSV_PATH.exists()
    with CSV_PATH.open("a", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=CHAMPS)
        if nouveau:
            w.writeheader()
        w.writerows(retenues)

    state["page"] = page
    STATE_PATH.write_text(json.dumps(state, indent=2), encoding="utf-8")

    if DEBUG and debug:
        DEBUG_PATH.write_text(
            json.dumps(debug, indent=2, ensure_ascii=False), encoding="utf-8"
        )

    print(f"\n{len(retenues)} entreprise(s) ajoutee(s) dans {CSV_PATH} :\n")
    for e in retenues:
        print(f"  {e['nom']} ({e['siren']})")
        print(f"    {e['adresse']}, {e['code_postal']} {e['ville']}")
        print(f"    dirigeants : {e['dirigeants'] or 'non renseignes'}")
    sans = sum(1 for e in retenues if not e["dirigeants"])
    if sans:
        print(f"\n[warn] {sans} ligne(s) sans dirigeant.")
    print(f"\nCout estime : {cout} credits")
    print(f"Prochaine execution : page {page}")


if __name__ == "__main__":
    main()
