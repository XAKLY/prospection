#!/usr/bin/env python3
"""
Collecte 5 restaurants de Metz par exécution via l'API Pappers v2.
Écrit/complète data/restaurants_metz.csv et mémorise la page dans data/state.json
pour ne pas retomber sur les mêmes entreprises chaque jour.

Usage : PAPPERS_API_KEY=xxx python collecte_pappers.py
"""

import csv
import json
import os
import sys
import time
from datetime import date
from pathlib import Path

import requests

API = "https://api.pappers.fr/v2"
KEY = os.environ.get("PAPPERS_API_KEY")
if not KEY:
    sys.exit("PAPPERS_API_KEY manquante dans l'environnement")

HEADERS = {"api-key": KEY}
DATA = Path("data")
CSV_PATH = DATA / "restaurants_metz.csv"
STATE_PATH = DATA / "state.json"

CHAMPS = [
    "date_ajout", "siren", "siret_siege", "nom", "forme_juridique",
    "code_naf", "libelle_naf", "date_creation", "adresse", "code_postal",
    "ville", "effectif", "chiffre_affaires", "dirigeants",
]


def get(endpoint, **params):
    r = requests.get(f"{API}/{endpoint}", headers=HEADERS, params=params, timeout=30)
    r.raise_for_status()
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


def formater_dirigeants(entreprise):
    out = []
    for r in entreprise.get("representants", []):
        nom = r.get("nom_complet") or " ".join(
            filter(None, [r.get("prenom"), r.get("nom")])
        )
        qualite = r.get("qualite", "")
        naissance = r.get("date_de_naissance_formate", "")
        out.append(" | ".join(filter(None, [nom, qualite, naissance])))
    return " ; ".join(out)


def main():
    DATA.mkdir(exist_ok=True)
    state = charger_state()
    vus = sirens_deja_vus()
    page = state.get("page", 1)

    nouvelles = []
    tentatives = 0

    # On avance de page en page jusqu'à obtenir 5 entreprises inédites
    while len(nouvelles) < 5 and tentatives < 10:
        tentatives += 1
        res = get(
            "recherche",
            code_naf="56.10A",        # restauration traditionnelle
            code_postal="57000",      # Metz
            entreprise_cessee="false",
            par_page=10,
            page=page,
        )
        resultats = res.get("resultats", [])
        if not resultats:
            print(f"Plus de résultats à la page {page}, arrêt.")
            break

        for item in resultats:
            if len(nouvelles) >= 5:
                break
            siren = item.get("siren")
            if not siren or siren in vus:
                continue

            detail = get("entreprise", siren=siren)
            siege = detail.get("siege", {}) or {}
            finances = (detail.get("finances") or [{}])[0]

            nouvelles.append({
                "date_ajout": date.today().isoformat(),
                "siren": siren,
                "siret_siege": siege.get("siret", ""),
                "nom": detail.get("nom_entreprise") or detail.get("denomination", ""),
                "forme_juridique": detail.get("forme_juridique", ""),
                "code_naf": detail.get("code_naf", ""),
                "libelle_naf": detail.get("libelle_code_naf", ""),
                "date_creation": detail.get("date_creation_formate", ""),
                "adresse": siege.get("adresse_ligne_1", ""),
                "code_postal": siege.get("code_postal", ""),
                "ville": siege.get("ville", ""),
                "effectif": detail.get("effectif", ""),
                "chiffre_affaires": finances.get("chiffre_affaires", ""),
                "dirigeants": formater_dirigeants(detail),
            })
            vus.add(siren)
            time.sleep(0.3)   # on reste poli avec l'API

        page += 1

    if not nouvelles:
        print("Aucune nouvelle entreprise trouvée — rien n'est écrit.")
        return

    nouveau_fichier = not CSV_PATH.exists()
    with CSV_PATH.open("a", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=CHAMPS)
        if nouveau_fichier:
            w.writeheader()
        w.writerows(nouvelles)

    state["page"] = page
    STATE_PATH.write_text(json.dumps(state, indent=2))

    print(f"{len(nouvelles)} entreprises ajoutées à {CSV_PATH} :")
    for e in nouvelles:
        print(f"  - {e['nom']} ({e['siren']}) — {e['dirigeants'] or 'dirigeants non renseignés'}")


if __name__ == "__main__":
    main()
