#!/usr/bin/env python3
"""
Collecte des restaurants de Metz via l'API Pappers v2.

Optimisation credits : on tire le maximum de /recherche (1 credit par appel)
et on n'appelle /entreprise (1 credit par societe) que si les dirigeants
sont absents du resultat de recherche.

Mettre ENRICHIR = False pour interdire totalement les appels /entreprise
(colonne dirigeants potentiellement vide, mais cout garanti = 1 credit/run).

Le premier run ecrit data/debug_recherche.json : le JSON brut du premier
resultat, pour verifier quels champs /recherche renvoie reellement.
"""

import csv
import json
import os
import sys
import time
from datetime import date
from pathlib import Path

import requests

# ---------------------------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------------------------

CLE_API = "cde2d8aff2614968747c3e5e858aa0bd077085f1bbc265a5"

CODE_NAF = "56.10A"      # restauration traditionnelle
CODE_POSTAL = "57000"    # Metz
NB_PAR_EXECUTION = 5

ENRICHIR = True          # False = jamais d'appel /entreprise
DEBUG = True             # False une fois les champs verifies

# ---------------------------------------------------------------------------

API = "https://api.pappers.fr/v2"
KEY = os.environ.get("PAPPERS_API_KEY") or CLE_API

if not KEY or KEY == "COLLE_TA_CLE_PAPPERS_ICI":
    sys.exit(
        "Cle API manquante : definis PAPPERS_API_KEY dans l'environnement "
        "ou renseigne CLE_API en haut du fichier."
    )

HEADERS = {"api-key": KEY}
DATA = Path("data")
CSV_PATH = DATA / "restaurants_metz.csv"
STATE_PATH = DATA / "state.json"
DEBUG_PATH = DATA / "debug_recherche.json"

CHAMPS = [
    "date_ajout", "siren", "siret_siege", "nom", "forme_juridique",
    "code_naf", "libelle_naf", "date_creation", "adresse", "code_postal",
    "ville", "effectif", "chiffre_affaires", "dirigeants", "source",
]

credits_estimes = 0


def get(endpoint, **params):
    global credits_estimes
    r = requests.get(f"{API}/{endpoint}", headers=HEADERS, params=params, timeout=30)
    if r.status_code == 401:
        sys.exit("Cle API refusee (401) - verifie qu'elle est active sur Pappers.")
    if r.status_code == 429:
        sys.exit("Quota depasse (429) - plus de credits ou trop de requetes.")
    r.raise_for_status()
    credits_estimes += 1
    return r.json()


def charger_state():
    if STATE_PATH.exists():
        return json.loads(STATE_PATH.read_text())
    return {"page": 1}


def sirens_deja_vus():
    if not CSV_PATH.exists():
        return set()
    with CSV_PATH.open(encoding="utf-8") as f:
        return {ligne["siren"] for ligne in csv.DictReader(f)}


def formater_dirigeants(source):
    """Cherche les dirigeants sous les differents noms possibles."""
    liste = (
        source.get("representants")
        or source.get("dirigeants")
        or []
    )
    out = []
    for r in liste:
        if not isinstance(r, dict):
            continue
        nom = r.get("nom_complet") or " ".join(
            filter(None, [r.get("prenom"), r.get("nom")])
        )
        qualite = r.get("qualite", "")
        naissance = r.get("date_de_naissance_formate", "")
        entree = " | ".join(filter(None, [nom.strip(), qualite, naissance]))
        if entree:
            out.append(entree)
    return " ; ".join(out)


def extraire(source, siren, source_label):
    siege = source.get("siege") or {}
    finances = (source.get("finances") or [{}])[0]
    return {
        "date_ajout": date.today().isoformat(),
        "siren": siren,
        "siret_siege": siege.get("siret", "") or source.get("siret", ""),
        "nom": (
            source.get("nom_entreprise")
            or source.get("denomination")
            or source.get("nom_complet", "")
        ),
        "forme_juridique": source.get("forme_juridique", ""),
        "code_naf": source.get("code_naf", "") or siege.get("code_naf", ""),
        "libelle_naf": source.get("libelle_code_naf", ""),
        "date_creation": source.get("date_creation_formate", ""),
        "adresse": siege.get("adresse_ligne_1", ""),
        "code_postal": siege.get("code_postal", ""),
        "ville": siege.get("ville", ""),
        "effectif": source.get("effectif", ""),
        "chiffre_affaires": finances.get("chiffre_affaires", ""),
        "dirigeants": formater_dirigeants(source),
        "source": source_label,
    }


def main():
    DATA.mkdir(exist_ok=True)
    state = charger_state()
    vus = sirens_deja_vus()
    page = state.get("page", 1)

    nouvelles = []
    tentatives = 0
    debug_ecrit = False

    while len(nouvelles) < NB_PAR_EXECUTION and tentatives < 10:
        tentatives += 1
        res = get(
            "recherche",
            code_naf=CODE_NAF,
            code_postal=CODE_POSTAL,
            entreprise_cessee="false",
            par_page=10,
            page=page,
        )
        resultats = res.get("resultats", [])
        if not resultats:
            print(f"Plus de resultats a la page {page}, arret.")
            break

        if DEBUG and not debug_ecrit:
            DEBUG_PATH.write_text(
                json.dumps(resultats[0], indent=2, ensure_ascii=False),
                encoding="utf-8",
            )
            debug_ecrit = True
            print(f"[debug] premier resultat brut ecrit dans {DEBUG_PATH}")

        for item in resultats:
            if len(nouvelles) >= NB_PAR_EXECUTION:
                break
            siren = item.get("siren")
            if not siren or siren in vus:
                continue

            ligne = extraire(item, siren, "recherche")

            # Appel detail UNIQUEMENT si la recherche n'a pas donne les dirigeants
            if ENRICHIR and not ligne["dirigeants"]:
                detail = get("entreprise", siren=siren)
                ligne = extraire(detail, siren, "entreprise")
                time.sleep(0.3)

            nouvelles.append(ligne)
            vus.add(siren)

        page += 1

    if not nouvelles:
        print("Aucune nouvelle entreprise trouvee - rien n'est ecrit.")
        print(f"Credits consommes (estimation) : {credits_estimes}")
        return

    nouveau_fichier = not CSV_PATH.exists()
    with CSV_PATH.open("a", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=CHAMPS)
        if nouveau_fichier:
            w.writeheader()
        w.writerows(nouvelles)

    state["page"] = page
    STATE_PATH.write_text(json.dumps(state, indent=2))

    print(f"{len(nouvelles)} entreprises ajoutees a {CSV_PATH} :")
    for e in nouvelles:
        dirigeants = e["dirigeants"] or "dirigeants non renseignes"
        print(f"  - {e['nom']} ({e['siren']}) [{e['source']}] - {dirigeants}")
    print(f"Credits consommes (estimation) : {credits_estimes}")


if __name__ == "__main__":
    main()
