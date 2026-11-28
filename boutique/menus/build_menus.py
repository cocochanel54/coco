#!/usr/bin/env python3
"""Génère le planificateur « Menus de la semaine & liste de courses »
(version vierge à vendre + version remplie pour les photos).

Usage : python3 build_menus.py
"""
import sys
import datetime as dt
from pathlib import Path

from openpyxl.formatting.rule import CellIsRule, FormulaRule

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from commun import *  # noqa: E402,F403

OUT = Path(__file__).parent
NAME = "Menus de la semaine & liste de courses"

RAYONS = ["Fruits et légumes", "Boucherie et poisson", "Crèmerie", "Boulangerie", "Épicerie salée",
          "Épicerie sucrée", "Surgelés", "Boissons", "Hygiène et maison"]
UNITES = ["g", "ml", "pièce", "boîte", "paquet", "tranche", "gousse", "bouteille", "pot", "sachet"]

# (ingrédient, rayon, unité, achat régulier dans la version exemple)
INGREDIENTS = [
    ("Pommes de terre", "Fruits et légumes", "g", None), ("Oignon", "Fruits et légumes", "pièce", None),
    ("Ail", "Fruits et légumes", "gousse", None), ("Carotte", "Fruits et légumes", "pièce", None),
    ("Courgette", "Fruits et légumes", "pièce", None), ("Aubergine", "Fruits et légumes", "pièce", None),
    ("Poivron", "Fruits et légumes", "pièce", None), ("Tomate", "Fruits et légumes", "pièce", None),
    ("Poireau", "Fruits et légumes", "pièce", None), ("Brocoli", "Fruits et légumes", "pièce", None),
    ("Salade verte", "Fruits et légumes", "pièce", None), ("Citron", "Fruits et légumes", "pièce", None),
    ("Bananes", "Fruits et légumes", "pièce", 6), ("Pommes", "Fruits et légumes", "pièce", 6),
    ("Bœuf haché", "Boucherie et poisson", "g", None), ("Poulet entier", "Boucherie et poisson", "pièce", None),
    ("Blanc de poulet", "Boucherie et poisson", "g", None), ("Lardons", "Boucherie et poisson", "g", None),
    ("Jambon", "Boucherie et poisson", "tranche", None), ("Pavé de saumon", "Boucherie et poisson", "pièce", None),
    ("Œufs", "Crèmerie", "pièce", None), ("Lait", "Crèmerie", "ml", None),
    ("Crème fraîche", "Crèmerie", "ml", None), ("Beurre", "Crèmerie", "g", None),
    ("Gruyère râpé", "Crèmerie", "g", None), ("Parmesan", "Crèmerie", "g", None),
    ("Fromage en tranches", "Crèmerie", "tranche", None), ("Yaourts", "Crèmerie", "pièce", 8),
    ("Pâte brisée", "Crèmerie", "pièce", None),
    ("Baguette", "Boulangerie", "pièce", 4), ("Pain de mie", "Boulangerie", "paquet", None),
    ("Pâtes", "Épicerie salée", "g", None), ("Riz", "Épicerie salée", "g", None),
    ("Semoule", "Épicerie salée", "g", None), ("Feuilles de lasagne", "Épicerie salée", "g", None),
    ("Coulis de tomate", "Épicerie salée", "g", None), ("Tomates concassées", "Épicerie salée", "boîte", None),
    ("Haricots rouges", "Épicerie salée", "boîte", None), ("Maïs", "Épicerie salée", "boîte", None),
    ("Thon", "Épicerie salée", "boîte", None), ("Lait de coco", "Épicerie salée", "boîte", None),
    ("Bouillon cube", "Épicerie salée", "pièce", None), ("Huile d'olive", "Épicerie salée", "bouteille", None),
    ("Farine", "Épicerie sucrée", "g", None), ("Sucre", "Épicerie sucrée", "g", None),
    ("Café", "Épicerie sucrée", "paquet", 1), ("Confiture", "Épicerie sucrée", "pot", None),
    ("Légumes surgelés", "Surgelés", "sachet", None),
    ("Eau (pack)", "Boissons", "pièce", 1), ("Jus d'orange", "Boissons", "bouteille", None),
    ("Papier toilette", "Hygiène et maison", "paquet", None),
    ("Liquide vaisselle", "Hygiène et maison", "bouteille", None),
    ("Lessive", "Hygiène et maison", "bouteille", None),
]

# Recettes pour 4 personnes : (nom, [(ingrédient, quantité), ...])
RECETTES = [
    ("Pâtes bolognaise", [("Pâtes", 400), ("Bœuf haché", 400), ("Coulis de tomate", 500), ("Oignon", 1),
                          ("Carotte", 1), ("Parmesan", 50)]),
    ("Poulet rôti et pommes de terre", [("Poulet entier", 1), ("Pommes de terre", 1000), ("Ail", 3)]),
    ("Gratin dauphinois", [("Pommes de terre", 1200), ("Crème fraîche", 400), ("Lait", 500), ("Ail", 1),
                           ("Gruyère râpé", 100)]),
    ("Quiche lorraine", [("Pâte brisée", 1), ("Lardons", 200), ("Œufs", 3), ("Crème fraîche", 200),
                         ("Lait", 200), ("Gruyère râpé", 80), ("Salade verte", 1)]),
    ("Chili con carne", [("Bœuf haché", 500), ("Haricots rouges", 1), ("Tomates concassées", 1), ("Oignon", 1),
                         ("Poivron", 1), ("Riz", 300)]),
    ("Curry de poulet et riz", [("Blanc de poulet", 600), ("Lait de coco", 1), ("Oignon", 1), ("Courgette", 1),
                                ("Riz", 300)]),
    ("Soupe de légumes", [("Carotte", 4), ("Pommes de terre", 400), ("Poireau", 2), ("Oignon", 1),
                          ("Bouillon cube", 2)]),
    ("Omelette et salade", [("Œufs", 8), ("Gruyère râpé", 50), ("Salade verte", 1), ("Tomate", 2)]),
    ("Saumon, riz et brocoli", [("Pavé de saumon", 4), ("Riz", 300), ("Brocoli", 1), ("Citron", 1)]),
    ("Lasagnes", [("Feuilles de lasagne", 250), ("Bœuf haché", 500), ("Coulis de tomate", 500), ("Lait", 500),
                  ("Beurre", 50), ("Farine", 50), ("Gruyère râpé", 100), ("Oignon", 1)]),
    ("Hachis parmentier", [("Pommes de terre", 1000), ("Bœuf haché", 500), ("Lait", 200), ("Beurre", 50),
                           ("Oignon", 1), ("Gruyère râpé", 80)]),
    ("Croque-monsieur et salade", [("Pain de mie", 1), ("Jambon", 4), ("Fromage en tranches", 8), ("Beurre", 30),
                                   ("Salade verte", 1)]),
    ("Ratatouille et semoule", [("Courgette", 2), ("Aubergine", 1), ("Poivron", 2), ("Tomate", 4), ("Oignon", 1),
                                ("Semoule", 300)]),
    ("Pâtes carbonara", [("Pâtes", 400), ("Lardons", 200), ("Œufs", 4), ("Parmesan", 80)]),
    ("Crêpes", [("Farine", 250), ("Œufs", 4), ("Lait", 500), ("Beurre", 50), ("Sucre", 50)]),
    ("Salade composée", [("Salade verte", 1), ("Tomate", 3), ("Maïs", 1), ("Thon", 2), ("Œufs", 4)]),
]

SEMAINE_EXEMPLE = [
    ("Salade composée", "Soupe de légumes", "Faire la soupe en double"),
    ("Pâtes bolognaise", "Omelette et salade", None),
    ("Croque-monsieur et salade", "Curry de poulet et riz", None),
    ("Restes", "Quiche lorraine", None),
    ("Saumon, riz et brocoli", "Pâtes carbonara", None),
    ("Restaurant", "Lasagnes", "Anniversaire de Léo"),
    ("Poulet rôti et pommes de terre", "Crêpes", "Crêpes : préparer la pâte le matin"),
]
JOURS = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]

# Emplacements
I0, I1 = 6, 205         # Ingrédients : 200 lignes
R0, R1 = 6, 805         # Recettes : 800 lignes (une par ingrédient)
S0, S1 = 8, 14          # Semaine : lundi → dimanche
L0, L1 = 5, 154         # Liste de courses : 150 lignes
PLANNING = f"Semaine!$C${S0}:$D${S1}"
ING = lambda col: f"Ingrédients!${col}${I0}:${col}${I1}"  # noqa: E731
REC = lambda col: f"Recettes!${col}${R0}:${col}${R1}"     # noqa: E731
FACTOR = "Listes!$B$2"


def build(example: bool):
    wb = new_workbook(NAME)
    guide = wb.active
    guide.title = "Mode d'emploi"
    sem = wb.create_sheet("Semaine")
    lst = wb.create_sheet("Liste de courses")
    rec = wb.create_sheet("Recettes")
    ing = wb.create_sheet("Ingrédients")
    lis = wb.create_sheet("Listes")

    build_listes(lis)
    build_ingredients(ing, example)
    build_recettes(rec)
    build_semaine(sem, example)
    build_liste(lst, example)
    build_guide(guide, NAME, "Choisissez vos repas de la semaine : la liste de courses se fait toute seule, "
                             "rangée par rayon.", GUIDE_STEPS, GUIDE_TIPS)
    lis.sheet_state = "hidden"
    wb.active = 0
    return wb


def build_listes(ws):
    setup(ws, C_MUTED)
    put(ws, "A1", "Recettes (sans doublon)", f=font(bold=True))
    for k in range(1, 201):
        ws[f"A{k + 1}"] = f'=IFERROR(INDEX({REC("A")},MATCH({k},{REC("H")},0)),"")'
    put(ws, "B1", "Coefficient personnes", f=font(bold=True))
    ws["B2"] = ("=IF(OR(N(Semaine!$B$4)=0,N(Recettes!$B$3)=0),1,Semaine!$B$4/Recettes!$B$3)")
    put(ws, "D1", "Rayons", f=font(bold=True))
    for i, r in enumerate(RAYONS):
        ws[f"D{i + 2}"] = r
    put(ws, "E1", "Unités", f=font(bold=True))
    for i, u in enumerate(UNITES):
        ws[f"E{i + 2}"] = u


def build_ingredients(ws, example):
    setup(ws, C_MUTED, {"A": 26, "B": 22, "C": 11, "D": 16, "E": 16, "F": 10, "G": 10})
    title(ws, "A1", "Mes ingrédients", "Le catalogue de vos ingrédients, chacun avec son rayon. "
                                       "Ajoutez les nouveaux à la fin de la liste, sans laisser de ligne vide.", "A2")
    header_row(ws, 5, 1, ["Ingrédient", "Rayon", "Unité", "Achat régulier (chaque semaine)",
                          "À acheter cette semaine"], height=32)
    rayon_n = len(RAYONS)
    for i, r in enumerate(range(I0, I1 + 1)):
        name, rayon, unit, regular = INGREDIENTS[i] if i < len(INGREDIENTS) else (None,) * 4
        put(ws, f"A{r}", name, bg=C_INPUT, border=BOX)
        put(ws, f"B{r}", rayon, bg=C_INPUT, border=BOX)
        put(ws, f"C{r}", unit, bg=C_INPUT, border=BOX, align=CENTER)
        put(ws, f"D{r}", regular if example else None, bg=C_INPUT, border=BOX, align=CENTER)
        # G (caché) : besoin brut = recettes de la semaine + achat régulier
        ws[f"G{r}"] = f'=IF(A{r}="",0,SUMIF({REC("B")},A{r},{REC("G")})+N(D{r}))'
        put(ws, f"E{r}", f'=IF(OR(A{r}="",G{r}=0),"",IF(OR(C{r}="g",C{r}="ml"),ROUND(G{r},0),'
                         f'ROUNDUP(ROUND(G{r},3),0)))',
            bg=C_CALC, border=BOX, align=CENTER, f=font(bold=True, color=C_GREEN))
        # F (caché) : clé de tri = rayon puis ligne
        ws[f"F{r}"] = (f'=IF(N(E{r})>0,IFERROR(MATCH(B{r},Listes!$D$2:$D${rayon_n + 1},0),99)*1000+ROW(),"")')
    ws.column_dimensions["F"].hidden = True
    ws.column_dimensions["G"].hidden = True
    list_validation(ws, f"B{I0}:B{I1}", f"Listes!$D$2:$D${rayon_n + 1}",
                    "Choisissez un rayon dans la liste.")
    list_validation(ws, f"C{I0}:C{I1}", f"Listes!$E$2:$E${len(UNITES) + 1}",
                    "Unité inhabituelle : vérifiez qu'elle est bien écrite.", strict=False)
    ws.freeze_panes = f"A{I0}"
    fit_width(ws)


def build_recettes(ws):
    setup(ws, C_CORAL, {"A": 32, "B": 26, "C": 11, "D": 11})
    title(ws, "A1", "Mes recettes", "Une ligne par ingrédient, avec le nom de la recette sur chaque ligne. "
                                    "Les quantités sont dans l'unité de l'onglet Ingrédients.", "A2")
    put(ws, "A3", "Quantités prévues pour (personnes)", f=font(bold=True), align=RIGHT)
    put(ws, "B3", 4, bg=C_INPUT, border=BOX, align=CENTER, f=font(11, True))
    header_row(ws, 5, 1, ["Recette", "Ingrédient", "Quantité", "Unité"])

    rows = [(name, ing, q) for name, items in RECETTES for ing, q in items]
    ws["H5"] = 0
    for i, r in enumerate(range(R0, R1 + 1)):
        name, ingr, qty = rows[i] if i < len(rows) else (None, None, None)
        put(ws, f"A{r}", name, bg=C_INPUT, border=BOX)
        put(ws, f"B{r}", ingr, bg=C_INPUT, border=BOX)
        put(ws, f"C{r}", qty, bg=C_INPUT, border=BOX, align=CENTER)
        put(ws, f"D{r}", f'=IF(B{r}="","",IFERROR(INDEX({ING("C")},MATCH(B{r},{ING("A")},0)),"?"))',
            bg=C_CALC, border=BOX, align=CENTER, f=font(9, color=C_MUTED))
        # Colonnes cachées : E = nombre de fois au menu, G = quantité à acheter, H = n° de recette unique
        ws[f"E{r}"] = f'=IF(A{r}="",0,COUNTIF({PLANNING},A{r}))'
        ws[f"G{r}"] = f'=IF(ISNUMBER(C{r}),C{r}*E{r}*{FACTOR},0)'
        ws[f"H{r}"] = f'=IF(A{r}="","",IF(COUNTIF($A${R0}:A{r},A{r})=1,MAX($H$5:H{r - 1})+1,""))'
    for col in "EFGH":
        ws.column_dimensions[col].hidden = True

    new_recipe = f'AND($A{R0}<>"",$A{R0}<>$A{R0 - 1})'
    top = Border(top=Side(style="medium", color=C_CORAL))
    ws.conditional_formatting.add(f"B{R0}:D{R1}", FormulaRule(formula=[new_recipe], border=top))
    ws.conditional_formatting.add(f"A{R0}:A{R1}", FormulaRule(formula=[new_recipe], border=top,
                                                              font=Font(name=FONT, bold=True, color=C_DARK)))
    ws.conditional_formatting.add(f"A{R0}:A{R1}", FormulaRule(formula=[f'AND($A{R0}<>"",$A{R0}=$A{R0 - 1})'],
                                                              font=Font(name=FONT, color="B8BEC6")))
    list_validation(ws, f"B{R0}:B{R1}", ING("A"),
                    "Choisissez un ingrédient de la liste. Pour un nouvel ingrédient, "
                    "ajoutez-le d'abord dans l'onglet Ingrédients.")
    ws.freeze_panes = f"A{R0}"
    fit_width(ws)


def build_semaine(ws, example):
    setup(ws, C_GREEN, {"A": 14, "B": 12, "C": 32, "D": 32, "E": 30})
    ws.merge_cells("A1:E1")
    put(ws, "A1", "Menus de la semaine", f=font(18, True, C_DARK))
    ws.row_dimensions[1].height = 30
    put(ws, "A3", "Semaine du", f=font(bold=True), align=RIGHT)
    put(ws, "B3", dt.date(2027, 1, 4) if example else None, bg=C_INPUT, border=BOX, fmt=DATE, align=CENTER,
        f=font(bold=True))
    put(ws, "A4", "Personnes", f=font(bold=True), align=RIGHT)
    put(ws, "B4", 4, bg=C_INPUT, border=BOX, align=CENTER, f=font(bold=True))

    tile(ws, "D3", "D4", "REPAS PRÉVUS", f'=COUNTA(C{S0}:D{S1})&" / 14"', C_GREEN, C_GREEN_L, fmt="General")
    tile(ws, "E3", "E4", "ARTICLES À ACHETER", f"=COUNT({ING('F')})", C_CORAL, C_CORAL_L, fmt="0")
    ws.row_dimensions[3].height = 20
    ws.row_dimensions[4].height = 32

    header_row(ws, S0 - 1, 1, ["Jour", "Date", "Midi", "Soir", "Notes"], height=24)
    for i, r in enumerate(range(S0, S1 + 1)):
        put(ws, f"A{r}", JOURS[i], f=font(11, True, C_WHITE), bg=C_DARK if i < 5 else C_GREEN, border=BOX,
            align=LEFT)
        put(ws, f"B{r}", f'=IF($B$3="","",$B$3-WEEKDAY($B$3,2)+{i + 1})', bg=C_CALC, border=BOX,
            fmt=DATE_SHORT, align=CENTER, f=font(color=C_MUTED))
        midi, soir, note = SEMAINE_EXEMPLE[i] if example else (None, None, None)
        put(ws, f"C{r}", midi, bg=C_INPUT, border=BOX, align=WRAP, f=font(11))
        put(ws, f"D{r}", soir, bg=C_INPUT, border=BOX, align=WRAP, f=font(11))
        put(ws, f"E{r}", note, bg=C_INPUT, border=BOX, align=WRAP, f=font(9, italic=True, color=C_MUTED))
        ws.row_dimensions[r].height = 36
    list_validation(ws, f"C{S0}:D{S1}", "Listes!$A$2:$A$201",
                    "Ce plat n'est pas dans vos recettes : il n'ajoutera rien à la liste de courses.",
                    strict=False)
    put(ws, f"A{S1 + 2}", "Choisissez les plats dans la liste, ou écrivez « Restes », « Restaurant »… "
                          "La liste de courses se met à jour toute seule.", f=font(9, italic=True, color=C_MUTED))
    fit_width(ws, landscape=True)


def build_liste(ws, example):
    setup(ws, C_GREEN, {"A": 20, "B": 28, "C": 11, "D": 10, "E": 6, "F": 3, "G": 28, "H": 6})
    ws.merge_cells("A1:H1")
    put(ws, "A1", '="Liste de courses"&IF(Semaine!$B$3="","",'
                  '" – semaine du "&TEXT(DAY(Semaine!$B$3-WEEKDAY(Semaine!$B$3,2)+1),"00")&"/"&'
                  'TEXT(MONTH(Semaine!$B$3-WEEKDAY(Semaine!$B$3,2)+1),"00"))',
        f=font(18, True, C_DARK))
    ws.row_dimensions[1].height = 30
    put(ws, "A2", f"Se remplit toute seule à partir des menus. Cochez {CHECK} ce qui est dans le panier.",
        f=font(9, italic=True, color=C_MUTED))
    header_row(ws, L0 - 1, 1, ["Rayon", "Article", "Quantité", "Unité", CHECK])
    header_row(ws, L0 - 1, 7, ["Autres courses", CHECK], color=C_CORAL)

    for k, r in enumerate(range(L0, L1 + 1), start=1):
        # Colonnes cachées : L = clé de tri, M = ligne dans Ingrédients, K = rayon
        ws[f"L{r}"] = f'=IFERROR(SMALL({ING("F")},{k}),"")'
        ws[f"M{r}"] = f'=IF(L{r}="","",MOD(L{r},1000)-{I0 - 1})'
        ws[f"K{r}"] = (f'=IF(M{r}="","",IF(INDEX({ING("B")},M{r})="","Autres",INDEX({ING("B")},M{r})))')
        put(ws, f"A{r}", f'=IF(K{r}=K{r - 1},"",K{r})', f=font(9, True, C_DARK), border=BOX)
        put(ws, f"B{r}", f'=IF(M{r}="","",INDEX({ING("A")},M{r}))', bg=C_CALC, border=BOX)
        put(ws, f"C{r}", f'=IF(M{r}="","",INDEX({ING("E")},M{r}))', bg=C_CALC, border=BOX, align=CENTER,
            f=font(bold=True))
        put(ws, f"D{r}", f'=IF(M{r}="","",INDEX({ING("C")},M{r}))', bg=C_CALC, border=BOX, align=CENTER,
            f=font(9, color=C_MUTED))
        put(ws, f"E{r}", None, bg=C_INPUT, border=BOX, align=CENTER, f=font(12, True, C_GREEN))
    for r in range(L0, L0 + 40):
        put(ws, f"G{r}", None, bg=C_INPUT, border=BOX)
        put(ws, f"H{r}", None, bg=C_INPUT, border=BOX, align=CENTER, f=font(12, True, C_GREEN))
    for col in "KLM":
        ws.column_dimensions[col].hidden = True

    check_column(ws, f"E{L0}:E{L1}", f"B{L0}:D{L1}", L0)
    check_column(ws, f"H{L0}:H{L0 + 39}", f"G{L0}:G{L0 + 39}", L0)
    ws.conditional_formatting.add(f"A{L0}:E{L1}", FormulaRule(formula=[f'$A{L0}<>""'],
                                                              border=Border(top=Side(style="medium", color=C_DARK))))
    if example:
        for r in (L0, L0 + 1, L0 + 2):
            ws[f"E{r}"] = CHECK
        ws[f"G{L0}"] = "Piles"
        ws[f"G{L0 + 1}"] = "Shampoing"
        ws[f"G{L0 + 2}"] = "Fleurs pour samedi"
    unlock(ws, f"E{L0}:E{L1}", f"G{L0}:H{L0 + 39}")
    ws.protection.sheet = True
    ws.freeze_panes = f"A{L0}"
    fit_width(ws)


GUIDE_STEPS = [
    ("Semaine", "Indiquez la date et le nombre de personnes, puis choisissez les plats du midi et du soir "
                "dans la liste. Vous pouvez aussi écrire « Restes » ou « Restaurant »."),
    ("Liste de courses", f"Elle se remplit toute seule, rangée par rayon, avec les quantités adaptées au nombre "
                         f"de personnes. Cochez {CHECK} au fur et à mesure dans le magasin."),
    ("Recettes", "16 recettes de tous les jours sont déjà prêtes. Ajoutez les vôtres : une ligne par "
                 "ingrédient, avec le nom de la recette sur chaque ligne."),
    ("Ingrédients", "Le catalogue des ingrédients avec leur rayon. Notez dans « Achat régulier » ce que vous "
                    "achetez chaque semaine (pain, yaourts…) : ça s'ajoute tout seul à la liste."),
]
GUIDE_TIPS = [
    f"Nouvelle semaine : changez la date, choisissez les nouveaux plats et effacez les {CHECK} de la liste de courses.",
    "Un nouvel ingrédient s'ajoute à la fin de l'onglet Ingrédients ; il sera rangé tout seul dans son rayon "
    "sur la liste de courses.",
    "Les quantités en grammes et en millilitres sont arrondies, les autres (pièces, boîtes…) sont arrondies "
    "au-dessus.",
    "Les plats écrits à la main (« Restes », « Restaurant »…) n'ajoutent rien à la liste de courses.",
    "L'onglet Liste de courses est protégé pour éviter d'effacer un calcul. Pour le modifier : "
    "menu Révision > Ôter la protection de la feuille (pas de mot de passe).",
    "Fonctionne avec Excel, LibreOffice et Google Sheets. Ouvrez la liste sur votre téléphone dans le magasin !",
]


if __name__ == "__main__":
    for example, fname in [(False, "Menus_et_courses.xlsx"), (True, "Menus_et_courses_exemple.xlsx")]:
        wb = build(example)
        wb.save(OUT / fname)
        print("écrit :", OUT / fname)
