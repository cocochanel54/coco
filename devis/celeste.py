# Devis Celeste au format du devis Circet : articles et prix de l'étude BPU fibre optique Celeste
# (onglet ETUDE_PRIX_CELESTE, prix de la colonne v3 du 15/05/2025).
#
#   python3 devis/celeste.py <BPU_FO_CELESTE.xlsx> <Exemple_devis_LOSANGE.xlsx> <sortie.xlsx>
#
# Le BPU est fourni en .xls : le convertir d'abord en .xlsx (Excel « Enregistrer sous », ou
# soffice --headless --convert-to xlsx), le script lit les valeurs calculées des formules.
import re, sys
import openpyxl
from modele_circet import COEF, libelle, lignes, prix, propre, remplir

SRC, MODELE, OUT = sys.argv[1:4]

COL_PRIX, DATE = 9, "15/05/2025"  # colonne I : prix v3
VERSION = f"v3 du {DATE}"
UNITES = {"ml": "ML", "m/L": "ML", "U": "U", "un": "U", "": ""}
# En-têtes de groupe dont les articles ne portent que la capacité (« 6FO », « 288 FO »…).
GROUPES = {"Pose en chambre d'une boite de protection d'épissure y compris":
               "Pose BPE en chambre, préparation du câble entrant",
           "Préparation et épissurage d'un câble optique dans une BPE existante":
               "Épissurage d'un câble dans une BPE existante",
           "Pose en chambre d'une boite de protection d'épissure avec préparation d'un câble laissé en passage":
               "Pose BPE en chambre, câble laissé en passage (lovage)"}

ws = openpyxl.load_workbook(SRC, data_only=True).worksheets[0]
catalogue, section, groupe, unite_groupe, majoration = [], "", "", "", None
for r in range(6, ws.max_row + 1):
    a, b, c, d, e, j = (ws.cell(r, k).value for k in (1, 2, 3, 4, 5, 10))
    code, px, l = propre(b), ws.cell(r, COL_PRIX).value, lignes(d)
    if not code and px is None:
        m = re.search(r"major\w* à (\d+)\s*%", str(d))
        if m:  # « * Tarif majoré à 30% la nuit »
            majoration = (int(m.group(1)), r)
        elif l and l[0].startswith("(") and catalogue:  # précision de l'article au-dessus
            catalogue[-1]["descriptif"] = (catalogue[-1]["descriptif"] + " " + " ".join(l)).strip()
        elif a and l:
            groupe = l[0]
        elif propre(a) or propre(c) or l:
            section, groupe = " – ".join(x for x in [propre(a), propre(c), " ".join(l)] if x), ""
        continue
    if "Plus d'actualité" in str(j):  # sous-tubage : plus de prix
        continue
    if not isinstance(j, str) or code == "ASTREINTE-REGION":  # note d'astreinte recopiée de SAV-F2
        j = None
    desig, reste = (l[0], l[1:]) if l else (propre(c) or code.split("-")[-1].capitalize(), [])
    if re.fullmatch(r"\d+\s*FO", desig):
        court = next((v for k, v in GROUPES.items() if groupe.startswith(k)), groupe)
        desig, reste = f"{court} – {desig.replace(' ', '')}", [groupe] + reste
    unite_bpu = propre(e) or (unite_groupe if re.search(r"\d+FO$", desig) else "")
    unite = "Forf." if code == "FOR-IMM" else UNITES[unite_bpu]  # FOR-IMM : forfait noté « m/L » dans le BPU
    unite_groupe = unite_bpu
    catalogue.append(dict(categorie=section, article=libelle(code, desig[:1].upper() + desig[1:]), unite=unite,
                          prix=prix(px) if px is not None else "Sur devis", unite_bpu=unite_bpu,
                          descriptif=" ".join(reste + ([f"Note BPU : {' '.join(lignes(j))}"] if j else []))
                                         .replace("", "–"),
                          source=f"BPU FO CELESTE Circet · ligne {r} · prix {VERSION}"))

if majoration:
    pct, r = majoration
    catalogue.append(dict(categorie="Majoration", article=f"Majoration de nuit +{pct} % – sur la ligne au-dessus",
                          unite=COEF, prix=1 + pct / 100, unite_bpu="", source=f"BPU FO CELESTE Circet · ligne {r}",
                          descriptif=f"« Tarif majoré à {pct}% la nuit » : P.U.H.T de la ligne au-dessus × {pct} %."))

remplir(MODELE, OUT, onglet="BPU CELESTE", catalogue=catalogue,
        titre_bpu=f"Bordereau de Prix Unitaires — Travaux Clients fibre optique — Celeste ({VERSION})",
        client="CELESTE", sous_titre="BPU Travaux Clients fibre optique",
        entete={"B8": "Bordereau de prix unitaire Travaux Clients – Celeste",
                "B11": "Département : ",
                "B12": "Objet : ",
                "B13": "Devis n° : DEVCEL-",
                "B14": "Affaire suivie par : C LECOMTE",
                "B15": "A l'attention de : ",
                "B17": "Réf. commande Celeste : ",
                "B19": f"PRESTATIONS SELON BPU FIBRE OPTIQUE CELESTE (version du {DATE})",
                "H41": "Pour Celeste"},
        aide=[
            ("section", "Renseigner l'en-tête"),
            ("texte", "Dans l'onglet \"Devis\", complète les cases du haut : Département, Objet, Devis n°, Affaire "
                      "suivie par, A l'attention de (ton contact chez Celeste), puis la référence de la commande dans "
                      "la bande bleue."),
            ("section", "Ajouter les lignes de prestations"),
            ("texte", "Clique sur une cellule \"Article\" : un menu déroulant propose tous les articles du BPU fibre "
                      f"optique Celeste ({VERSION}), sous la forme « code · désignation ». Choisis l'article : "
                      "l'unité et le P.U.H.T se remplissent tout seuls. Indique la quantité, le Montant HT se calcule."),
            ("texte", "Astuce : tape le début du code (ex. R-BPE ou SAV-F1) pour retrouver l'article plus vite."),
            ("section", "Articles « Sur devis »"),
            ("texte", "Pour les articles sans prix dans le BPU (FOR_ECH_MIG, MAT-FOURN-PLU-VALU, GC-RFO-FORAGE, "
                      "E-DEPL) ou hors BPU, le P.U.H.T affiche « Sur devis » sur fond jaune : tape ton prix "
                      "directement dans la case."),
            ("section", "Majoration de nuit"),
            ("texte", "Le BPU prévoit un tarif majoré de 30 % la nuit : ajoute la ligne « Majoration de nuit +30 % » "
                      "juste sous la prestation concernée, avec la même quantité. Son P.U.H.T vaut 30 % du P.U.H.T de "
                      "la ligne au-dessus. Pour le SAV, prends directement les articles HNO (SAV-F0-HNO, SAV-F1-HNO, "
                      "SAV-F2-HNO).")])
print(f"{len(catalogue)} articles dont {sum(d['prix'] == 'Sur devis' for d in catalogue)} sur devis -> {OUT}")
