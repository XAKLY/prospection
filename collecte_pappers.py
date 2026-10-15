#!/usr/bin/env python3
"""
Collecte de prospects (restaurants Metz) via l'API Pappers v2.

COMPORTEMENT
------------
- Ajoute NB_PAR_EXECUTION NOUVEAUX restaurants a chaque execution.
- Deduplication par SIREN sur data/restaurants_metz.csv.
- Recupere pour chaque societe :
    * le DIRIGEANT PRINCIPAL : nom, prenom, poste (qualite), adresse
    * la liste complete des dirigeants
    * e-mail, telephone, site web (si Pappers les fournit)
    * annee d'ouverture, effectif, chiffre d'affaires + annee, resultat
  -> colonnes pensees pour l'import dans Odoo CRM (lead + contact).
- Genere data/restaurants_metz.xlsx : lignes du run en VERT.

COUT
----
  /recherche  = 2 credits par appel (20 resultats)
  /entreprise = 1 credit par societe (dirigeants, finances, contacts)
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

CLE_API = "COLLE_TA_CLE_PAPPERS_ICI"   # ou secret GitHub PAPPERS_API_KEY

CODE_NAF = "56.10A"
CODE_POSTAL = "57000"
NB_PAR_EXECUTION = 5

EXCLURE_RADIEES = True
SIEGE_DANS_LA_ZONE = True
ENRICHIR_ENTREPRISE = True
DEBUG = False                # True : ecrit debug_pappers.json (a utiliser si
                             # une colonne reste vide, pour voir les vrais champs)

COULEUR_NOUVEAU = "C6EFCE"

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
    "date_ajout", "siren", "siret_siege", "nom", "raison_sociale",
    "forme_juridique", "code_naf", "libelle_naf", "date_creation",
    "annee_ouverture", "adresse", "code_postal", "ville", "effectif",
    "chiffre_affaires", "annee_ca", "resultat",
    "email", "telephone", "site_web", "statut",
    "dirigeant_nom", "dirigeant_prenom", "dirigeant_poste",
    "dirigeant_adresse", "dirigeant_code_postal", "dirigeant_ville",
    "dirigeants",
]

cout = 0
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
    if not CSV_PATH.exists():
        return []
    try:
        with CSV_PATH.open(encoding="utf-8", newline="") as f:
            return [dict(l) for l in csv.DictReader(f)]
    except Exception as e:
        print(f"[warn] lecture du CSV impossible ({e}), pas de deduplication.")
        return []


# ---------------------------------------------------------------------------
# Dirigeants
# ---------------------------------------------------------------------------

INDICES_PERSONNE = {
    "nom", "prenom", "nom_complet", "qualite", "denomination",
    "prenom_usuel", "nom_usage", "fonction",
}

# Ordre de preference pour choisir le "dirigeant principal"
PRIORITE_POSTES = [
    "gerant", "gérant", "president", "président", "directeur general",
    "directeur général", "dirigeant", "associe unique", "associé unique",
    "pouvoir d'engager",
]


def ressemble_a_une_personne(d):
    return isinstance(d, dict) and bool(INDICES_PERSONNE & set(d.keys()))


def nom_personne(d):
    for cle in ("nom_complet", "denomination", "nom_usage"):
        if d.get(cle):
            return str(d[cle]).strip()
    morceaux = [d.get("prenom") or d.get("prenom_usuel"), d.get("nom")]
    return " ".join(str(m) for m in morceaux if m).strip()


def poste_de(d):
    return str(d.get("qualite") or d.get("fonction") or "").strip()


def liste_representants(item):
    """Retourne la liste brute des dirigeants/representants d'un item."""
    for cle, valeur in item.items():
        if not isinstance(valeur, list) or not valeur:
            continue
        if "representant" not in cle and "dirigeant" not in cle:
            continue
        pers = [d for d in valeur if ressemble_a_une_personne(d)]
        if pers:
            return pers
    return []


def formater(liste):
    out = []
    for d in liste:
        nom = nom_personne(d)
        if not nom:
            continue
        entree = " | ".join(x for x in [nom, poste_de(d)] if x)
        if entree not in out:
            out.append(entree)
    return " ; ".join(out)


def score_dirigeant(d):
    """Plus petit = meilleur. Prefere une personne physique avec un poste cle."""
    morale = 1 if d.get("personne_morale") else 0
    poste = poste_de(d).lower()
    rang = len(PRIORITE_POSTES)
    for i, mot in enumerate(PRIORITE_POSTES):
        if mot in poste:
            rang = i
            break
    return (morale, rang)


def dirigeant_principal(item):
    reps = liste_representants(item)
    vide = {k: "" for k in (
        "dirigeant_nom", "dirigeant_prenom", "dirigeant_poste",
        "dirigeant_adresse", "dirigeant_code_postal", "dirigeant_ville")}
    if not reps:
        return vide
    d = sorted(reps, key=score_dirigeant)[0]
    nom = d.get("nom") or ""
    prenom = d.get("prenom") or d.get("prenom_usuel") or ""
    if not nom and not prenom:
        # personne morale ou champ nom_complet seul
        nom = nom_personne(d)
    return {
        "dirigeant_nom": str(nom).strip(),
        "dirigeant_prenom": str(prenom).strip(),
        "dirigeant_poste": poste_de(d),
        "dirigeant_adresse": d.get("adresse_ligne_1") or "",
        "dirigeant_code_postal": d.get("code_postal") or "",
        "dirigeant_ville": d.get("ville") or "",
    }


# ---------------------------------------------------------------------------
# Finances / contacts
# ---------------------------------------------------------------------------

def derniere_finance(item):
    """Dernier exercice publie (chiffre d'affaires, resultat, annee)."""
    fin = item.get("finances")
    if not isinstance(fin, list) or not fin:
        return {}
    fin = [f for f in fin if isinstance(f, dict)]
    if not fin:
        return {}
    avec_ca = [f for f in fin if f.get("chiffre_affaires") is not None]
    base = avec_ca or fin
    return max(base, key=lambda f: f.get("annee") or 0)


def premier(item, siege, *cles):
    """Premiere valeur non vide parmi plusieurs cles, au niveau societe puis siege."""
    for src in (item, siege):
        for c in cles:
            v = src.get(c)
            if v:
                return str(v).strip()
    return ""


def annee_ouverture(item):
    d = item.get("date_creation") or item.get("date_creation_formate") or ""
    d = str(d)
    # formats : 2024-03-01 ou 01/03/2024
    if len(d) >= 4 and d[:4].isdigit():
        return d[:4]
    if len(d) >= 4 and d[-4:].isdigit():
        return d[-4:]
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
    fin = derniere_finance(item)
    nom = (item.get("nom_entreprise")
           or item.get("denomination")
           or item.get("nom_complet", ""))
    ligne = {
        "date_ajout": date.today().isoformat(),
        "siren": item.get("siren", ""),
        "siret_siege": siege.get("siret", "") or item.get("siret", ""),
        "nom": nom,
        "raison_sociale": item.get("denomination") or nom,
        "forme_juridique": item.get("forme_juridique", ""),
        "code_naf": item.get("code_naf", "") or siege.get("code_naf", ""),
        "libelle_naf": item.get("libelle_code_naf", ""),
        "date_creation": (item.get("date_creation_formate", "")
                          or item.get("date_creation", "")),
        "annee_ouverture": annee_ouverture(item),
        "adresse": (siege.get("adresse_ligne_1", "")
                    or item.get("adresse_ligne_1", "")),
        "code_postal": siege.get("code_postal", "") or item.get("code_postal", ""),
        "ville": siege.get("ville", "") or item.get("ville", ""),
        "effectif": item.get("effectif", "") or item.get("tranche_effectif", ""),
        "chiffre_affaires": fin.get("chiffre_affaires", "") if fin else "",
        "annee_ca": fin.get("annee", "") if fin else "",
        "resultat": fin.get("resultat", "") if fin else "",
        "email": premier(item, siege, "email", "adresse_email"),
        "telephone": premier(item, siege, "telephone", "numero_telephone"),
        "site_web": premier(item, siege, "site_internet", "site_web"),
        "statut": item.get("statut_consolide", "") or item.get("statut_rcs", ""),
        "dirigeants": formater(liste_representants(item)),
    }
    ligne.update(dirigeant_principal(item))
    return ligne


# ---------------------------------------------------------------------------
# Export Excel
# ---------------------------------------------------------------------------

def ecrire_xlsx(lignes, sirens_nouveaux):
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

    for nom_col in ("siren", "siret_siege", "telephone"):
        idx = CHAMPS.index(nom_col) + 1
        for row in ws.iter_rows(min_row=2, min_col=idx, max_col=idx):
            row[0].number_format = "@"
            row[0].value = str(row[0].value or "")

    largeurs = {
        "date_ajout": 12, "siren": 12, "siret_siege": 16, "nom": 30,
        "raison_sociale": 30, "forme_juridique": 30, "code_naf": 9,
        "libelle_naf": 24, "date_creation": 13, "annee_ouverture": 10,
        "adresse": 34, "code_postal": 11, "ville": 18, "effectif": 22,
        "chiffre_affaires": 14, "annee_ca": 8, "resultat": 12,
        "email": 30, "telephone": 16, "site_web": 26, "statut": 9,
        "dirigeant_nom": 18, "dirigeant_prenom": 16, "dirigeant_poste": 40,
        "dirigeant_adresse": 30, "dirigeant_code_postal": 11,
        "dirigeant_ville": 16, "dirigeants": 50,
    }
    for i, nom_col in enumerate(CHAMPS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = largeurs.get(nom_col, 15)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(XLSX_PATH)


def ecrire_csv(lignes):
    """Reecrit tout le CSV (migre automatiquement l'ancien format de colonnes)."""
    with CSV_PATH.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=CHAMPS, restval="",
                           extrasaction="ignore")
        w.writeheader()
        w.writerows(lignes)


# ---------------------------------------------------------------------------

def main():
    DATA.mkdir(exist_ok=True)
    state = charger_state()
    existantes = lire_csv()
    vus = {l["siren"] for l in existantes if l.get("siren")}
    page = state.get("page", 1)

    retenues = []
    tours = 0
    page_epuisee = True

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

            # Detail : 1 credit -> dirigeants, finances, contacts
            if ENRICHIR_ENTREPRISE:
                detail = get("entreprise", 1, siren=siren)
                if detail:
                    if DEBUG and "entreprise" not in debug:
                        debug["entreprise"] = detail
                    # le detail prime pour les champs vides
                    for cle, valeur in extraire(detail).items():
                        if cle == "date_ajout":
                            continue
                        if valeur not in ("", None) and ligne.get(cle) in ("", None):
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

    toutes = existantes + retenues
    ecrire_csv(toutes)
    ecrire_xlsx(toutes, {e["siren"] for e in retenues})

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
        print(f"    dirigeant : {e['dirigeant_prenom']} {e['dirigeant_nom']} "
              f"- {e['dirigeant_poste'] or 'poste non renseigne'}")
        print(f"    contact   : {e['email'] or '-'} / {e['telephone'] or '-'}")
        ca = e["chiffre_affaires"]
        print(f"    CA        : {ca if ca != '' else '-'} ({e['annee_ca'] or '-'})")
    sans = sum(1 for e in retenues if not e["dirigeant_nom"])
    if sans:
        print(f"\n[warn] {sans} ligne(s) sans dirigeant.")
    print(f"\nFichiers : {CSV_PATH} et {XLSX_PATH} (nouvelles lignes en vert)")
    print(f"Cout estime : {cout} credits")
    print(f"Prochaine execution : page {page}")


if __name__ == "__main__":
    main()
