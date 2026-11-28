#!/usr/bin/env python3
"""Génère le planificateur de budget : une version vierge (à vendre)
et une version remplie d'exemples (pour les photos de l'annonce).

Usage : python3 build_budget.py [année]      (par défaut 2027)
"""
import random
import sys
import datetime as dt
from pathlib import Path

from openpyxl.formatting.rule import CellIsRule, DataBarRule
from openpyxl.chart import BarChart, LineChart, Reference

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from commun import *  # noqa: E402,F403  (couleurs, formats, put, tile, build_guide…)

YEAR = int(sys.argv[1]) if len(sys.argv) > 1 else 2027
OUT = Path(__file__).parent

TOP = Border(top=Side(style="medium", color=C_DARK))

# Catégories proposées, avec le budget mensuel utilisé dans la version exemple
REVENUS = [("Salaire", 2150), ("Prime / 13e mois", 0), ("Aides (CAF…)", 110),
           ("Revenus complémentaires", 0), ("Remboursements", 0), ("Autres revenus", 0)]
DEPENSES = [("Loyer / crédit immobilier", 720), ("Charges", 40), ("Électricité / gaz", 85),
            ("Eau", 30), ("Internet / téléphone", 45), ("Assurances", 65),
            ("Courses alimentaires", 380), ("Restaurants / livraisons", 80),
            ("Carburant / transports", 140), ("Entretien voiture", 30),
            ("Santé / pharmacie", 25), ("Abonnements", 30), ("Sport / loisirs", 60),
            ("Vêtements", 50), ("Beauté / coiffeur", 25), ("Maison / déco", 30),
            ("Enfants", 0), ("Animaux", 0), ("Cadeaux", 30), ("Vacances", 80),
            ("Impôts", 0), ("Crédit conso", 0), ("Épargne", 200), ("Divers", 40)]

# Emplacements dans l'onglet Paramètres
P_REV = list(range(8, 18))     # 10 lignes de revenus
P_DEP = list(range(21, 51))    # 30 lignes de dépenses
REV_RANGE = "'Paramètres'!$A$8:$A$17"
DEP_RANGE = "'Paramètres'!$A$21:$A$50"
YEAR_REF = "'Paramètres'!$B$3"
SOLDE_REF = "'Paramètres'!$B$4"

# Onglet Transactions
T_FIRST, T_LAST = 5, 2004


def trng(col):
    return f"Transactions!${col}${T_FIRST}:${col}${T_LAST}"


def sumifs_month(cat_cell, month_expr):
    return (f'SUMIFS({trng("D")},{trng("C")},{cat_cell},'
            f'{trng("A")},">="&DATE({YEAR_REF},{month_expr},1),'
            f'{trng("A")},"<"&DATE({YEAR_REF},{month_expr}+1,1))')


# ---------- Données d'exemple ----------
def example_transactions(year):
    rnd = random.Random(42)
    rows = []

    def add(m, d, desc, cat, amount):
        rows.append((dt.date(year, m, d), desc, cat, round(amount, 2)))

    for m in range(1, 13):
        add(m, 28, "Salaire", "Salaire", 2150)
        add(m, 5, "Aide au logement", "Aides (CAF…)", 110)
        add(m, 3, "Loyer", "Loyer / crédit immobilier", 720)
        add(m, 3, "Charges de l'immeuble", "Charges", 40)
        add(m, 10, "Électricité", "Électricité / gaz", 85)
        add(m, 12, "Eau", "Eau", 29.5)
        add(m, 8, "Box internet + forfait mobile", "Internet / téléphone", 44.99)
        add(m, 6, "Assurance habitation + auto", "Assurances", 64.5)
        add(m, 15, "Streaming vidéo", "Abonnements", 13.49)
        add(m, 15, "Musique en ligne", "Abonnements", 10.99)
        add(m, 2, "Salle de sport", "Sport / loisirs", 29.9)
        add(m, 27, "Virement vers le livret", "Épargne", 200)
        for _ in range(5):
            add(m, rnd.randint(1, 28), rnd.choice(["Supermarché", "Marché", "Primeur", "Supermarché"]),
                "Courses alimentaires", rnd.uniform(38, 105))
        for _ in range(rnd.randint(1, 3)):
            add(m, rnd.randint(1, 28), rnd.choice(["Restaurant", "Livraison repas", "Café entre amis"]),
                "Restaurants / livraisons", rnd.uniform(12, 48))
        for _ in range(3):
            add(m, rnd.randint(1, 28), "Plein d'essence", "Carburant / transports", rnd.uniform(42, 62))
        if rnd.random() < 0.7:
            add(m, rnd.randint(1, 28), "Pharmacie", "Santé / pharmacie", rnd.uniform(8, 32))
            add(m, rnd.randint(1, 28), "Remboursement Sécu", "Remboursements", rnd.uniform(15, 25))
        if rnd.random() < 0.6:
            add(m, rnd.randint(1, 28), rnd.choice(["Cinéma", "Sortie", "Concert"]),
                "Sport / loisirs", rnd.uniform(12, 40))
        if rnd.random() < 0.5:
            add(m, rnd.randint(1, 28), "Vêtements", "Vêtements", rnd.uniform(25, 90))
        if m % 2 == 0:
            add(m, rnd.randint(1, 28), "Coiffeur", "Beauté / coiffeur", 28)
        if rnd.random() < 0.4:
            add(m, rnd.randint(1, 28), "Petit équipement maison", "Maison / déco", rnd.uniform(15, 60))
        for _ in range(rnd.randint(1, 2)):
            add(m, rnd.randint(1, 28), "Dépense imprévue", "Divers", rnd.uniform(8, 35))
        if rnd.random() < 0.3:
            add(m, rnd.randint(1, 28), "Cadeau d'anniversaire", "Cadeaux", rnd.uniform(20, 45))
    add(4, 18, "Révision voiture", "Entretien voiture", 185)
    add(10, 9, "Pneus", "Entretien voiture", 140)
    add(7, 20, "Acompte location vacances", "Vacances", 250)
    add(8, 4, "Vacances", "Vacances", 640)
    add(12, 15, "Prime de fin d'année", "Prime / 13e mois", 1200)
    add(12, 18, "Cadeaux de Noël", "Cadeaux", 260)
    add(6, 14, "Vente d'objets d'occasion", "Revenus complémentaires", 85)
    rows.sort(key=lambda r: r[0])
    return rows


# ---------- Construction du classeur ----------
def build(example: bool) -> Workbook:
    wb = new_workbook(f"Planificateur de budget {YEAR}")

    guide = wb.active
    guide.title = "Mode d'emploi"
    par = wb.create_sheet("Paramètres")
    tra = wb.create_sheet("Transactions")
    moi = wb.create_sheet("Mois")
    ann = wb.create_sheet("Année")
    lis = wb.create_sheet("Listes")

    for ws, color in [(guide, C_DARK), (par, C_MUTED), (tra, C_CORAL), (moi, C_GREEN),
                      (ann, C_DARK), (lis, C_MUTED)]:
        ws.sheet_view.showGridLines = False
        ws.sheet_properties.tabColor = color

    build_listes(lis)
    build_parametres(par, example)
    build_transactions(tra, example)
    build_mois(moi)
    build_annee(ann)
    build_guide(guide, f"Planificateur de budget {YEAR}", GUIDE_SUBTITLE, GUIDE_STEPS, GUIDE_TIPS)
    for ws in (par, tra):
        fit_width(ws)
    lis.sheet_state = "hidden"
    wb.active = 0
    return wb


def build_listes(ws):
    put(ws, "A1", "Catégories (calcul)", f=font(bold=True))
    for i in range(1, 41):
        r = i + 1
        ws[f"A{r}"] = (f'=IF({i}<=$C$2,INDEX({REV_RANGE},{i}),'
                       f'IF({i}-$C$2<=$C$3,INDEX({DEP_RANGE},{i}-$C$2),""))')
    put(ws, "C1", "Nb revenus / dépenses", f=font(bold=True))
    ws["C2"] = f"=COUNTA({REV_RANGE})"
    ws["C3"] = f"=COUNTA({DEP_RANGE})"
    put(ws, "D1", "Mois", f=font(bold=True))
    for i, m in enumerate(MONTHS):
        ws[f"D{i + 2}"] = m
    put(ws, "E1", "Mois choisi", f=font(bold=True))
    ws["E2"] = "=IFERROR(MATCH(Mois!$B$3,$D$2:$D$13,0),1)"
    put(ws, "F1", "Mois avec des saisies", f=font(bold=True))
    ws["F2"] = "=MAX(1,SUMPRODUCT(--((Année!$B$36:$M$36+Année!$B$69:$M$69)>0)))"


def build_parametres(ws, example):
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 4
    ws.column_dimensions["D"].width = 30
    ws.column_dimensions["E"].width = 16

    put(ws, "A1", "Paramètres", f=font(18, True, C_DARK))
    ws.row_dimensions[1].height = 30

    put(ws, "A3", "Année du budget", f=font(bold=True))
    put(ws, "B3", YEAR, bg=C_INPUT, border=BOX, align=CENTER, f=font(11, True))
    put(ws, "D3", "Changez l'année : tout le classeur suit.", f=font(9, italic=True, color=C_MUTED))
    put(ws, "A4", "Solde de départ sur le compte", f=font(bold=True))
    put(ws, "B4", 1250 if example else 0, bg=C_INPUT, border=BOX, fmt=EUR, align=RIGHT)
    put(ws, "D4", "L'argent sur votre compte au 1er janvier.", f=font(9, italic=True, color=C_MUTED))

    def section(row, title, color, light, items, rows):
        ws.merge_cells(f"A{row - 2}:B{row - 2}")
        put(ws, f"A{row - 2}", title, f=font(12, True, color))
        put(ws, f"A{row - 1}", "Catégorie", f=font(bold=True, color=C_WHITE), bg=color, border=BOX)
        put(ws, f"B{row - 1}", "Budget prévu par mois", f=font(bold=True, color=C_WHITE),
            bg=color, border=BOX, align=CENTER)
        for i, r in enumerate(rows):
            name, budget = items[i] if i < len(items) else (None, None)
            put(ws, f"A{r}", name, bg=C_INPUT, border=BOX)
            put(ws, f"B{r}", budget if (example and budget) else None,
                bg=C_INPUT, border=BOX, fmt=EUR, align=RIGHT)

    section(P_REV[0], "Mes revenus", C_GREEN, C_GREEN_L, REVENUS, P_REV)
    section(P_DEP[0], "Mes dépenses", C_CORAL, C_CORAL_L, DEPENSES, P_DEP)

    # Résumé du budget prévu
    ws.merge_cells("D7:E7")
    put(ws, "D7", "Résumé de mon budget prévu", f=font(bold=True, color=C_WHITE), bg=C_DARK, align=CENTER)
    lines = [("Revenus prévus / mois", f"=SUM(B{P_REV[0]}:B{P_REV[-1]})", C_GREEN),
             ("Dépenses prévues / mois", f"=SUM(B{P_DEP[0]}:B{P_DEP[-1]})", C_CORAL),
             ("Reste prévu / mois", "=E8-E9", C_DARK)]
    for i, (label, formula, color) in enumerate(lines):
        r = 8 + i
        put(ws, f"D{r}", label, bg=C_CALC, border=BOX)
        put(ws, f"E{r}", formula, bg=C_CALC, border=BOX, fmt=EUR, align=RIGHT, f=font(bold=True, color=color))
    ws.conditional_formatting.add("E10", CellIsRule(operator="lessThan", formula=["0"],
                                                    font=Font(name=FONT, bold=True, color=C_CORAL)))

    notes = ["Astuces :",
             "• Renommez, ajoutez ou effacez des catégories librement.",
             "• Pas de ligne vide au milieu d'une liste.",
             "• Chaque catégorie doit avoir un nom unique.",
             "• Laissez le budget vide si vous ne voulez pas en fixer."]
    for i, n in enumerate(notes):
        put(ws, f"D{12 + i}", n, f=font(9, i == 0, C_MUTED, italic=i > 0))
    ws.freeze_panes = "A6"


def build_transactions(ws, example):
    widths = {"A": 13, "B": 40, "C": 28, "D": 14, "E": 13}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    put(ws, "A1", "Transactions", f=font(18, True, C_DARK))
    ws.row_dimensions[1].height = 30
    put(ws, "A2", "Une ligne par dépense ou revenu. Le montant est toujours positif : "
                  "c'est la catégorie qui indique s'il s'agit d'un revenu ou d'une dépense.",
        f=font(9, italic=True, color=C_MUTED))

    headers = ["Date", "Description", "Catégorie", "Montant", "Type"]
    for i, h in enumerate(headers):
        col = "ABCDE"[i]
        put(ws, f"{col}4", h, f=font(bold=True, color=C_WHITE), bg=C_DARK, border=BOX,
            align=CENTER if i != 1 else LEFT)
    ws.row_dimensions[4].height = 22

    if example:
        data = example_transactions(YEAR)
    else:
        data = [(dt.date(YEAR, 1, 3), "Exemple : courses de la semaine (ligne à effacer)",
                 "Courses alimentaires", 85.40)]

    for r in range(T_FIRST, T_LAST + 1):
        row = data[r - T_FIRST] if r - T_FIRST < len(data) else (None, None, None, None)
        put(ws, f"A{r}", row[0], bg=C_INPUT, border=BOX, fmt=DATE, align=CENTER)
        put(ws, f"B{r}", row[1], bg=C_INPUT, border=BOX)
        put(ws, f"C{r}", row[2], bg=C_INPUT, border=BOX)
        put(ws, f"D{r}", row[3], bg=C_INPUT, border=BOX, fmt=EUR, align=RIGHT)
        put(ws, f"E{r}",
            f'=IF(C{r}="","",IF(COUNTIF({REV_RANGE},C{r})>0,"Revenu",'
            f'IF(COUNTIF({DEP_RANGE},C{r})>0,"Dépense","Catégorie inconnue")))',
            bg=C_CALC, border=BOX, align=CENTER, f=font(9, color=C_MUTED))

    rng = f"E{T_FIRST}:E{T_LAST}"
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Revenu"'],
                                                  font=Font(name=FONT, size=9, bold=True, color=C_GREEN)))
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Dépense"'],
                                                  font=Font(name=FONT, size=9, bold=True, color=C_CORAL)))
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Catégorie inconnue"'],
                                                  font=Font(name=FONT, size=9, bold=True, color="B91C1C")))

    dv_cat = DataValidation(type="list", formula1="Listes!$A$2:$A$41", allow_blank=True,
                            showErrorMessage=True, errorTitle="Catégorie",
                            error="Choisissez une catégorie dans la liste. "
                                  "Vous pouvez en ajouter dans l'onglet Paramètres.")
    dv_cat.add(f"C{T_FIRST}:C{T_LAST}")
    dv_amt = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0",
                            allow_blank=True, showErrorMessage=True, errorTitle="Montant",
                            error="Indiquez un montant positif (sans le signe moins).")
    dv_amt.add(f"D{T_FIRST}:D{T_LAST}")
    dv_date = DataValidation(type="date", operator="greaterThan", formula1="1",
                             allow_blank=True, showErrorMessage=True, errorTitle="Date",
                             error="Indiquez une date, par exemple 03/01/" + str(YEAR) + ".")
    dv_date.add(f"A{T_FIRST}:A{T_LAST}")
    for dv in (dv_cat, dv_amt, dv_date):
        ws.add_data_validation(dv)

    ws.auto_filter.ref = f"A4:E{T_LAST}"
    ws.freeze_panes = "A5"


def build_mois(ws):
    for col, w in {"A": 30, "B": 15, "C": 15, "D": 18, "E": 13}.items():
        ws.column_dimensions[col].width = w
    ws.merge_cells("A1:E1")
    put(ws, "A1", f'="Mon budget de "&LOWER(B3)&" "&{YEAR_REF}', f=font(18, True, C_DARK))
    ws.row_dimensions[1].height = 30

    put(ws, "A3", "Choisissez le mois  ▸", f=font(11, True), align=RIGHT)
    ws.merge_cells("B3:C3")
    put(ws, "B3", "Janvier", f=font(12, True, C_DARK), bg=C_INPUT, align=CENTER, border=BOX)
    ws["C3"].border = BOX
    ws["B3"].protection = Protection(locked=False)
    dv = DataValidation(type="list", formula1="Listes!$D$2:$D$13", allow_blank=False,
                        showErrorMessage=True, error="Choisissez un mois dans la liste.")
    dv.add("B3")
    ws.add_data_validation(dv)
    ws.row_dimensions[3].height = 24

    M = "Listes!$E$2"
    r0, r1 = 9, 18     # revenus
    d0, d1 = 22, 51    # dépenses
    tot_r, tot_d = r1 + 1, d1 + 1

    tile(ws, "A5", "A6", "REVENUS DU MOIS", f"=C{tot_r}", C_GREEN, C_GREEN_L)
    tile(ws, "B5", "B6", "DÉPENSES DU MOIS", f"=C{tot_d}", C_CORAL, C_CORAL_L, merge=("B5:C5", "B6:C6"))
    tile(ws, "D5", "D6", "IL ME RESTE", f"=C{tot_r}-C{tot_d}", C_DARK, C_LIGHT, merge=("D5:E5", "D6:E6"))
    ws.row_dimensions[5].height = 18
    ws.row_dimensions[6].height = 34

    def header(row, cols, color):
        for i, h in enumerate(cols):
            col = "ABCDE"[i]
            put(ws, f"{col}{row}", h, f=font(bold=True, color=C_WHITE), bg=color, border=BOX,
                align=LEFT if i == 0 else CENTER)
        ws.row_dimensions[row].height = 20

    header(r0 - 1, ["REVENUS", "Prévu", "Réel", "Écart"], C_GREEN)
    for i, r in enumerate(range(r0, r1 + 1)):
        p = P_REV[i]
        put(ws, f"A{r}", f"=IF('Paramètres'!A{p}=\"\",\"\",'Paramètres'!A{p})", bg=C_CALC, border=BOX)
        put(ws, f"B{r}", f"=IF(A{r}=\"\",\"\",'Paramètres'!B{p})", bg=C_CALC, border=BOX, fmt=EUR)
        put(ws, f"C{r}", f'=IF(A{r}="","",{sumifs_month(f"A{r}", M)})', bg=C_CALC, border=BOX,
            fmt=EUR, f=font(bold=True))
        put(ws, f"D{r}", f'=IF(A{r}="","",C{r}-B{r})', bg=C_CALC, border=BOX, fmt=EUR)
    put(ws, f"A{tot_r}", "Total revenus", f=font(bold=True), bg=C_GREEN_L, border=BOX)
    for col in "BC":
        put(ws, f"{col}{tot_r}", f"=SUM({col}{r0}:{col}{r1})", f=font(bold=True), bg=C_GREEN_L,
            border=BOX, fmt=EUR)
    put(ws, f"D{tot_r}", f"=C{tot_r}-B{tot_r}", f=font(bold=True), bg=C_GREEN_L, border=BOX, fmt=EUR)

    header(d0 - 1, ["DÉPENSES", "Prévu", "Réel", "Reste à dépenser", "% utilisé"], C_CORAL)
    for i, r in enumerate(range(d0, d1 + 1)):
        p = P_DEP[i]
        put(ws, f"A{r}", f"=IF('Paramètres'!A{p}=\"\",\"\",'Paramètres'!A{p})", bg=C_CALC, border=BOX)
        put(ws, f"B{r}", f"=IF(A{r}=\"\",\"\",'Paramètres'!B{p})", bg=C_CALC, border=BOX, fmt=EUR)
        put(ws, f"C{r}", f'=IF(A{r}="","",{sumifs_month(f"A{r}", M)})', bg=C_CALC, border=BOX,
            fmt=EUR, f=font(bold=True))
        put(ws, f"D{r}", f'=IF(A{r}="","",B{r}-C{r})', bg=C_CALC, border=BOX, fmt=EUR)
        put(ws, f"E{r}", f'=IF(A{r}="","",IF(B{r}=0,"",C{r}/B{r}))', bg=C_CALC, border=BOX,
            fmt=PCT, align=CENTER)
    put(ws, f"A{tot_d}", "Total dépenses", f=font(bold=True), bg=C_CORAL_L, border=BOX)
    for col in "BC":
        put(ws, f"{col}{tot_d}", f"=SUM({col}{d0}:{col}{d1})", f=font(bold=True), bg=C_CORAL_L,
            border=BOX, fmt=EUR)
    put(ws, f"D{tot_d}", f"=B{tot_d}-C{tot_d}", f=font(bold=True), bg=C_CORAL_L, border=BOX, fmt=EUR)
    put(ws, f"E{tot_d}", f'=IF(B{tot_d}=0,"",C{tot_d}/B{tot_d})', f=font(bold=True), bg=C_CORAL_L,
        border=BOX, fmt=PCT, align=CENTER)

    red = Font(name=FONT, bold=True, color=C_CORAL)
    ws.conditional_formatting.add(f"D{d0}:D{tot_d}", CellIsRule(operator="lessThan", formula=["0"], font=red))
    ws.conditional_formatting.add(f"E{d0}:E{tot_d}", CellIsRule(operator="greaterThan", formula=["1"], font=red))
    ws.conditional_formatting.add(f"E{d0}:E{d1}", DataBarRule(start_type="num", start_value=0, end_type="num",
                                                             end_value=1, color="8FB8C2"))
    ws.conditional_formatting.add("D6", CellIsRule(operator="lessThan", formula=["0"],
                                                   font=Font(name=FONT, size=16, bold=True, color=C_CORAL)))
    ws.freeze_panes = "A8"
    ws.protection.sheet = True
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def build_annee(ws):
    ws.column_dimensions["A"].width = 28
    for i in range(12):
        ws.column_dimensions["BCDEFGHIJKLM"[i]].width = 11
    ws.column_dimensions["N"].width = 13
    ws.column_dimensions["O"].width = 13

    put(ws, "A1", f'="Bilan de l\'année "&{YEAR_REF}', f=font(18, True, C_DARK))
    ws.row_dimensions[1].height = 30

    HDR, R0, R1, RT = 24, 26, 35, 36
    D0, D1, DT = 39, 68, 69
    REST, SOLDE = 71, 72

    tile(ws, "B3", "B4", "REVENUS DE L'ANNÉE", f"=N{RT}", C_GREEN, C_GREEN_L, merge=("B3:D3", "B4:D4"))
    tile(ws, "E3", "E4", "DÉPENSES DE L'ANNÉE", f"=N{DT}", C_CORAL, C_CORAL_L, merge=("E3:G3", "E4:G4"))
    tile(ws, "H3", "H4", "RESTE SUR L'ANNÉE", f"=N{REST}", C_DARK, C_LIGHT, merge=("H3:J3", "H4:J4"))
    tile(ws, "K3", "K4", "SOLDE EN FIN D'ANNÉE", f"=M{SOLDE}", C_DARK, C_LIGHT, merge=("K3:M3", "K4:M4"))
    ws.row_dimensions[3].height = 18
    ws.row_dimensions[4].height = 34

    # En-têtes du tableau
    put(ws, f"A{HDR - 1}", "Détail mois par mois", f=font(12, True, C_DARK))
    cols = ["Catégorie"] + MONTHS_SHORT + ["Total", "Moyenne / mois"]
    for i, h in enumerate(cols):
        col = ws.cell(row=HDR, column=i + 1)
        put(ws, col.coordinate, h, f=font(9 if i == 14 else 10, True, C_WHITE), bg=C_DARK, border=BOX,
            align=LEFT if i == 0 else CENTER)
    ws.row_dimensions[HDR].height = 22

    def section_label(row, text, color, light):
        for c in range(1, 16):
            ws.cell(row=row, column=c).fill = fill(light)
        put(ws, f"A{row}", text, f=font(10, True, color), bg=light)

    def cat_rows(rows, prows):
        for i, r in enumerate(rows):
            p = prows[i]
            put(ws, f"A{r}", f"=IF('Paramètres'!A{p}=\"\",\"\",'Paramètres'!A{p})", border=BOX)
            for m in range(1, 13):
                c = ws.cell(row=r, column=m + 1)
                put(ws, c.coordinate, f'=IF($A{r}="","",{sumifs_month(f"$A{r}", m)})', border=BOX, fmt=EUR0)
            put(ws, f"N{r}", f'=IF($A{r}="","",SUM(B{r}:M{r}))', border=BOX, fmt=EUR0, f=font(bold=True))
            put(ws, f"O{r}", f'=IF($A{r}="","",N{r}/Listes!$F$2)', border=BOX, fmt=EUR0, f=font(color=C_MUTED))

    def total_row(row, first, last, label, color, light):
        put(ws, f"A{row}", label, f=font(bold=True, color=color), bg=light, border=BOX)
        for c in range(2, 15):
            cell = ws.cell(row=row, column=c)
            L = cell.column_letter
            put(ws, cell.coordinate, f"=SUM({L}{first}:{L}{last})", f=font(bold=True, color=color),
                bg=light, border=BOX, fmt=EUR0)
        put(ws, f"O{row}", f"=N{row}/Listes!$F$2", f=font(bold=True, color=color), bg=light, border=BOX, fmt=EUR0)

    section_label(R0 - 1, "REVENUS", C_GREEN, C_GREEN_L)
    cat_rows(range(R0, R1 + 1), P_REV)
    total_row(RT, R0, R1, "Total revenus", C_GREEN, C_GREEN_L)
    section_label(D0 - 1, "DÉPENSES", C_CORAL, C_CORAL_L)
    cat_rows(range(D0, D1 + 1), P_DEP)
    total_row(DT, D0, D1, "Total dépenses", C_CORAL, C_CORAL_L)

    put(ws, f"A{REST}", "Reste (revenus – dépenses)", f=font(bold=True, color=C_WHITE), bg=C_DARK, border=BOX)
    put(ws, f"A{SOLDE}", "Solde du compte en fin de mois", f=font(bold=True, color=C_DARK), bg=C_LIGHT, border=BOX)
    for m in range(1, 13):
        L = "BCDEFGHIJKLM"[m - 1]
        put(ws, f"{L}{REST}", f"={L}{RT}-{L}{DT}", f=font(bold=True, color=C_WHITE), bg=C_DARK, border=BOX, fmt=EUR0)
        prev = SOLDE_REF if m == 1 else f"{'BCDEFGHIJKLM'[m - 2]}{SOLDE}"
        put(ws, f"{L}{SOLDE}", f"={prev}+{L}{REST}", f=font(bold=True, color=C_DARK), bg=C_LIGHT,
            border=BOX, fmt=EUR0)
    put(ws, f"N{REST}", f"=SUM(B{REST}:M{REST})", f=font(bold=True, color=C_WHITE), bg=C_DARK, border=BOX, fmt=EUR0)
    put(ws, f"O{REST}", f"=N{REST}/Listes!$F$2", f=font(bold=True, color=C_WHITE), bg=C_DARK, border=BOX, fmt=EUR0)
    put(ws, f"N{SOLDE}", f"=M{SOLDE}", f=font(bold=True, color=C_DARK), bg=C_LIGHT, border=BOX, fmt=EUR0)
    put(ws, f"O{SOLDE}", None, bg=C_LIGHT, border=BOX)
    ws.conditional_formatting.add(f"B{REST}:O{REST}", CellIsRule(operator="lessThan", formula=["0"],
                                                                fill=fill(C_CORAL)))
    ws.conditional_formatting.add(f"B{SOLDE}:N{SOLDE}", CellIsRule(operator="lessThan", formula=["0"],
                                                                  font=Font(name=FONT, bold=True, color=C_CORAL)))
    put(ws, f"A{SOLDE + 2}", "La moyenne est calculée sur les mois où vous avez déjà saisi quelque chose.",
        f=font(9, italic=True, color=C_MUTED))

    # Graphique 1 : revenus et dépenses par mois
    ch = BarChart()
    ch.type = "col"
    ch.grouping = "clustered"
    ch.title = "Revenus et dépenses par mois"
    ch.add_data(Reference(ws, min_col=1, max_col=13, min_row=RT, max_row=RT), from_rows=True, titles_from_data=True)
    ch.add_data(Reference(ws, min_col=1, max_col=13, min_row=DT, max_row=DT), from_rows=True, titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, max_col=13, min_row=HDR, max_row=HDR))
    for s, color in zip(ch.series, [C_GREEN, C_CORAL]):
        s.graphicalProperties.solidFill = color
        s.graphicalProperties.line.solidFill = color
    ch.gapWidth = 60
    ch.y_axis.number_format = '#,##0" €"'
    ch.legend.position = "b"
    style_chart(ch)
    ch.width, ch.height = 16, 7.6
    ws.add_chart(ch, "A6")

    # Graphique 2 : solde du compte
    lc = LineChart()
    lc.title = "Évolution du solde du compte"
    lc.add_data(Reference(ws, min_col=1, max_col=13, min_row=SOLDE, max_row=SOLDE), from_rows=True,
                titles_from_data=True)
    lc.set_categories(Reference(ws, min_col=2, max_col=13, min_row=HDR, max_row=HDR))
    s = lc.series[0]
    s.graphicalProperties.line.solidFill = C_DARK
    s.graphicalProperties.line.width = 32000
    s.smooth = False
    lc.y_axis.number_format = '#,##0" €"'
    style_chart(lc)
    lc.legend = None
    lc.width, lc.height = 16, 7.6
    ws.add_chart(lc, "H6")

    ws.freeze_panes = "B1"
    ws.protection.sheet = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


GUIDE_SUBTITLE = "Suivez vos revenus et vos dépenses mois par mois, en quelques minutes par semaine."
GUIDE_STEPS = [
    ("Paramètres", "Indiquez l'année et l'argent sur votre compte au 1er janvier. Adaptez la liste "
                   "des catégories et notez le budget que vous prévoyez chaque mois pour chacune."),
    ("Transactions", "Notez chaque dépense et chaque revenu : la date, une description, la catégorie "
                     "(à choisir dans la liste) et le montant. Le montant est toujours positif."),
    ("Mois", "Choisissez un mois en haut de la page pour comparer ce que vous aviez prévu et ce que "
             "vous avez vraiment dépensé, catégorie par catégorie."),
    ("Année", "Retrouvez le bilan de l'année, les graphiques et l'évolution du solde de votre compte."),
]
GUIDE_TIPS = [
    "Si vous renommez une catégorie, corrigez aussi les anciennes transactions qui portent l'ancien nom.",
    "Une transaction marquée « Catégorie inconnue » n'est comptée nulle part : choisissez une catégorie de la liste.",
    "Les onglets Mois et Année sont protégés pour éviter d'effacer un calcul par erreur. "
    "Pour les modifier : menu Révision > Ôter la protection de la feuille (pas de mot de passe).",
    "Fonctionne avec Excel, LibreOffice et Google Sheets (Fichier > Importer).",
    "Pour l'année suivante : gardez une copie du fichier, videz l'onglet Transactions et changez l'année dans Paramètres.",
]


if __name__ == "__main__":
    for example, name in [(False, f"Budget_{YEAR}.xlsx"), (True, f"Budget_{YEAR}_exemple.xlsx")]:
        wb = build(example)
        path = OUT / name
        wb.save(path)
        print("écrit :", path)
