# Devis Axione au format du devis Circet : articles et prix des deux BPU Moselle Fibre RHD57
# (Maintenance et VDR Optique), à l'attention de Nicolas FRUTOS.
#
#   python3 devis/axione.py <BPU Maintenance.xlsx> <BPU VDR Optique.xlsx> <Exemple_devis_LOSANGE.xlsx> <sortie.xlsx>
import re, sys
import openpyxl
from modele_circet import COEF, libelle, lignes, prix, propre, remplir

SRC_MAINT, SRC_OPT, MODELE, OUT = sys.argv[1:5]

ORANGE = "FF6600"  # bandeau « lot » du BPU Maintenance

# Unités du BPU raccourcies comme dans le modèle (Forf., ML, U…) pour tenir dans la colonne « U ».
UNITES = {"u": "U", "U": "U", "ml": "ML", "Forfait": "Forf.", "Forfait par opération": "Forf.",
          "Forfait par intervention": "Forf.", "Forfait par heure": "Heure", "Forfait global Mensuel": "Mois",
          "Par commande": "Cde", "Par mètre (hauteur nacelle)": "m (haut.)", "u (à la fibre)": "Fibre",
          "coef. appliqué au forfait": COEF}

# ---------- Lecture des BPU ----------
articles = []      # dict par prestation proposée dans le devis
majorations = []   # coefficients MC-07 / MC-08

ws = openpyxl.load_workbook(SRC_MAINT).worksheets[0]
lot = famille = ""
for r in range(2, ws.max_row + 1):
    a, code = ws.cell(r, 1).value, ws.cell(r, 2).value
    if a and not code:
        if ws.cell(r, 1).fill.fgColor.rgb == "FF" + ORANGE:
            lot = propre(a)
        else:
            famille = propre(a)
        continue
    if not code:
        continue
    d = dict(bpu="Maintenance", lot=lot, famille=famille, code=propre(code),
             desig=propre(ws.cell(r, 3).value), descr=" / ".join(lignes(ws.cell(r, 7).value)),
             unite=propre(ws.cell(r, 5).value), pbpu=prix(ws.cell(r, 6).value),
             comm=propre(ws.cell(r, 4).value), src=f"Maintenance BPU ST · BPU ligne {r}")
    if d["unite"].startswith("coef"):
        majorations.append(d)
        continue
    articles.append(d)
    if d["code"] == "RA-ON" and "0,93€/ml" in str(ws.cell(r, 7).value).replace(" ", ""):
        # Le descriptif de RA-ON prévoit 0,93 €/ml au-delà de 300 ml : on en fait une ligne.
        articles.append(dict(d, desig=d["desig"] + " – ml au-delà de 300 ml", unite="ml",
                             pbpu=0.93, descr="Mètre linéaire au-delà des 300 ml compris dans le forfait",
                             comm="D'après le descriptif de RA-ON (« jusqu'à 300ml, 0,93€/ml au-delà »)"))

ws = openpyxl.load_workbook(SRC_OPT).worksheets[0]
lot = famille = commun = ""
for r in range(4, ws.max_row + 1):
    a, b = ws.cell(r, 1).value, ws.cell(r, 2).value
    if isinstance(a, str) and " - " in a:
        if re.match(r"^\d+ - ", a.strip()):
            lot, famille = propre(a), ""
        else:
            famille = propre(a)
        commun = ""
        continue
    if a is None:
        if b:
            commun = " ".join(lignes(b))  # descriptif commun à la famille
        continue
    code = str(a).strip()
    l = lignes(b)
    desig, reste = l[0].rstrip(" :"), l[1:]
    if reste and reste[0].startswith("Prix pour"):
        desig, reste = f"{desig} – {reste[0]}", reste[1:]
    unite = propre(ws.cell(r, 3).value) or "U"
    if unite == "fofait":  # coquille du BPU (4.6.1, 4.6.2)
        unite = "Forfait"
    if propre(ws.cell(r, 4).value) == "à la fibre":  # mesures et raccordements comptés à la fibre
        unite = f"{unite} (à la fibre)"
    articles.append(dict(bpu="VDR Optique", lot=lot, famille=famille or lot, code=code, desig=desig,
                         descr=" ".join(reste) or commun, unite=unite, pbpu=prix(ws.cell(r, 5).value),
                         comm=propre(ws.cell(r, 4).value),
                         src=f"BPU ST VDR Optique · BPU MISSION 3 Optique ligne {r}"))

assert [m["code"] for m in majorations] == ["MC-07", "MC-08"]
MAJ_LIB = {"MC-07": "MC-07 · Majoration par technicien supplémentaire – sur la ligne au-dessus",
           "MC-08": "MC-08 · Majoration intervention en HNO – sur la ligne au-dessus"}
catalogue = []
for d in articles + majorations:
    maj = d in majorations
    catalogue.append(dict(categorie=f"{d['bpu']} – {d['famille']}",
                          article=MAJ_LIB[d["code"]] if maj else libelle(d["code"], d["desig"]),
                          unite=UNITES[d["unite"]], prix=d["pbpu"], unite_bpu=d["unite"],
                          descriptif=" / ".join(x for x in [d["desig"] if maj else "", d["descr"]] if x),
                          source=d["src"]))

remplir(MODELE, OUT, onglet="BPU RHD57", catalogue=catalogue,
        titre_bpu="Bordereau de Prix Unitaires — Moselle Fibre RHD57 — Axione (Maintenance et VDR Optique)",
        client="AXIONE", sous_titre="Moselle Fibre – RHD57",
        entete={"B8": "Bordereau de prix unitaire Moselle Fibre – RHD57",
                "B11": "Département : 57",
                "B12": "Objet : Intervention suite au ticket ",
                "B13": "Devis n° : DEVAX57-",
                "B14": "Affaire suivie par : C LECOMTE",
                "B15": "A l'attention de : Nicolas FRUTOS",
                "B17": "Réf. ticket Axione : ",
                "B19": "PRESTATIONS SELON BPU MOSELLE FIBRE – RHD57 (Maintenance et VDR Optique)",
                "H41": "Pour Axione"},
        aide=[
            ("section", "Renseigner l'en-tête"),
            ("texte", "Dans l'onglet \"Devis\", complète les cases du haut : Département, Objet, Devis n°, Affaire "
                      "suivie par, A l'attention de (Nicolas FRUTOS chez Axione), puis la référence du ticket dans la "
                      "bande bleue."),
            ("section", "Ajouter les lignes de prestations"),
            ("texte", "Clique sur une cellule \"Article\" : un menu déroulant propose tous les articles du BPU Moselle "
                      "Fibre RHD57, sous la forme « code · désignation » (Maintenance puis VDR Optique). Choisis "
                      "l'article : l'unité et le P.U.H.T se remplissent tout seuls. Indique la quantité, le Montant HT "
                      "se calcule."),
            ("texte", "Astuce : tape le début du code (ex. MC-T1 ou 4.3.2) pour retrouver l'article plus vite."),
            ("section", "Articles « Sur devis »"),
            ("texte", "Pour les articles sans prix dans le BPU (RA-RO, MC-T6, PU-01 à PU-06) ou hors BPU, le P.U.H.T "
                      "affiche « Sur devis » sur fond jaune : tape ton prix directement dans la case."),
            ("section", "Majorations MC-07 et MC-08"),
            ("texte", "Choisis la ligne de majoration juste sous l'intervention concernée, avec la même quantité. Son "
                      "P.U.H.T vaut le P.U.H.T de la ligne au-dessus × (coefficient − 1) : par exemple MC-08 (HNO, "
                      "coef. 1,4078) ajoute 40,78 % au forfait. Les coefficients sont dans l'onglet \"BPU RHD57\".")])
print(f"{len(articles)} articles ({sum(d['bpu'] == 'Maintenance' for d in articles)} maintenance, "
      f"{sum(d['bpu'] == 'VDR Optique' for d in articles)} optique) + {len(majorations)} majorations -> {OUT}")
