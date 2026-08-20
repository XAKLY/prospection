#!/usr/bin/env python3
"""
Collecte de prospects (restaurants Metz) via l'API Pappers v2.

STRATEGIE CREDITS
-----------------
AUCUN appel a /entreprise (qui coute 1 credit PAR societe). On utilise :
  1. /recherche              -> les societes + un maximum de champs
  2. /recherche-dirigeants   -> les dirigeants, en UN seul appel pour tout le lot

Cout typique : 2 a 4 credits par execution, quel que soit NB_PAR_EXECUTION.
Si /recherche renvoie deja les dirigeants, l'etape 2 est sautee.

SORTIE
------
data/restaurants_metz.csv  : le fichier de prospection (append)
data/state.json            : page courante, pour ne pas repasser sur les memes
data/debug_pappers.json    : (si DEBUG) echantillons bruts des reponses API
"""

import csv
import json
import os
import sys
from datetime import date
from pathlib import Path

import requests

# ===========================================================================
# CONFIGURATION
# ===========================================================================

# Laisser le placeholder : la cle vient du secret GitHub PAPPERS_API_KEY
CLE_API = "COLLE_TA_CLE_PAPPERS_ICI"

CODE_NAF = "56.10A"       # 56.10A = restauration traditionnelle
CODE_POSTAL = "57000"     # Metz
NB_PAR_EXECUTION = 5      # nouveaux prospects par run

EXCLURE_RADIEES = True    # ignore les societes radiees / inactives
DEBUG = True              # ecrit des echantillons bruts (False ensuite)

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

appels = 0
debug = {}


def get(endpoint, **params):
    """Appel API avec compteur et messages d'erreur lisibles."""
    global appels
    r = requests.get(f"{API}/{endpoint}", headers=HEADERS,
                     params=params, timeout=30)

    if r.status_code == 401:
        sys.exit("Cle API refusee (401) : verifie qu'elle est active sur Pappers.")
    if r.status_code == 404:
        print(f"[info] endpoint {endpoint} introuvable (404), on continue sans.")
        return None
    if r.status_code == 429:
        sys.exit("Quota depasse (429) : plus de credits ou trop de requetes.")
    if r.status_code >= 400:
        sys.exit(f"Erreur {r.status_code} sur {endpoint} : {r.text[:300]}")

    appels += 1
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
# Extraction
# ---------------------------------------------------------------------------

def nom_personne(d):
    return (
        d.get("nom_complet")
        or " ".join(filter(None, [d.get("prenom"), d.get("nom")])).strip()
        or d.get("denomination", "")
    )


def formater_dirigeants(liste):
    out = []
    for d in liste or []:
        if not isinstance(d, dict):
            continue
        entree = " | ".join(x for x in [nom_personne(d), d.get("qualite", "")] if x)
        if entree and entree not in out:
            out.append(entree)
    return " ; ".join(out)


def dirigeants_de(item):
    """Cherche les dirigeants sous les differents noms de champ possibles."""
    for cle in ("representants", "dirigeants", "representants_legaux"):
        if item.get(cle):
            return formater_dirigeants(item[cle])
    return ""


def est_active(item):
    statut = (item.get("statut_consolide") or "").lower()
    rcs = (item.get("statut_rcs") or "").lower()
    if "inactif" in statut or "radi" in rcs:
        return False
    if item.get("entreprise_cessee"):
        return False
    return True


def extraire(item):
    siege = item.get("siege") or {}
    return {
        "date_ajout": date.today().isoformat(),
        "siren": item.get("siren", ""),
        "siret_siege": siege.get("siret", "") or item.get("siret", ""),
        "nom": (
            item.get("nom_entreprise")
            or item.get("denomination")
            or item.get("nom_complet", "")
        ),
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
# Etape 2 : dirigeants en UN seul appel
# ---------------------------------------------------------------------------

def table_dirigeants(sirens):
    """Retourne {siren: 'Nom | Qualite ; ...'} ; dict vide si la route echoue."""
    if not sirens:
        return {}

    res = get(
        "recherche-dirigeants",
        code_naf=CODE_NAF,
        code_postal=CODE_POSTAL,
        par_page=100,
        page=1,
    )
    if not res:
        return {}

    if DEBUG and res.get("resultats"):
        debug["recherche_dirigeants"] = res["resultats"][0]

    par_siren = {}
    for d in res.get("resultats", []):
        siren = d.get("siren") or (d.get("entreprise") or {}).get("siren")
        if not siren or siren not in sirens:
            continue
        par_siren.setdefault(siren, []).append(d)

    return {s: formater_dirigeants(v) for s, v in par_siren.items()}


# ---------------------------------------------------------------------------

def main():
    DATA.mkdir(exist_ok=True)
    state = charger_state()
    vus = sirens_deja_vus()
    page = state.get("page", 1)

    retenues = []
    tours = 0

    # --- Etape 1 : /recherche -----------------------------------------------
    while len(retenues) < NB_PAR_EXECUTION and tours < 10:
        tours += 1
        res = get(
            "recherche",
            code_naf=CODE_NAF,
            code_postal=CODE_POSTAL,
            entreprise_cessee="false",
            precision="standard",
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
                print(f"[skip] {item.get('nom_entreprise', siren)} : radiee/inactive")
                vus.add(siren)
                continue
            retenues.append(extraire(item))
            vus.add(siren)

    if not retenues:
        print("Aucune nouvelle entreprise trouvee - rien n'est ecrit.")
        print(f"Appels API : {appels}")
        return

    # --- Etape 2 : dirigeants manquants -------------------------------------
    manquants = {l["siren"] for l in retenues if not l["dirigeants"]}
    if manquants:
        print(f"[info] dirigeants absents pour {len(manquants)} societe(s), "
              f"appel unique a /recherche-dirigeants...")
        table = table_dirigeants(manquants)
        for ligne in retenues:
            if not ligne["dirigeants"]:
                ligne["dirigeants"] = table.get(ligne["siren"], "")

    # --- Ecriture -----------------------------------------------------------
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

    # --- Rapport ------------------------------------------------------------
    print(f"\n{len(retenues)} entreprise(s) ajoutee(s) dans {CSV_PATH} :\n")
    for e in retenues:
        print(f"  {e['nom']} ({e['siren']})")
        print(f"    {e['adresse']}, {e['code_postal']} {e['ville']}")
        print(f"    dirigeants : {e['dirigeants'] or 'non renseignes'}")
    sans = sum(1 for e in retenues if not e["dirigeants"])
    if sans:
        print(f"\n[warn] {sans} ligne(s) sans dirigeant.")
    print(f"\nAppels API (= credits environ) : {appels}")
    print(f"Prochaine execution : page {page}")


if __name__ == "__main__":
    main()
