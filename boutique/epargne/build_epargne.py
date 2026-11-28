#!/usr/bin/env python3
"""Génère « Objectifs d'épargne & défi 52 semaines »
(version vierge à vendre + version remplie pour les photos).

Usage : python3 build_epargne.py
"""
import sys
import datetime as dt
from pathlib import Path

from openpyxl.formatting.rule import CellIsRule, DataBarRule, FormulaRule

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from commun import *  # noqa: E402,F403

OUT = Path(__file__).parent
NAME = "Objectifs d'épargne & défi 52 semaines"
YEAR = 2027

G0, G1 = 9, 20       # Objectifs : 12 lignes
V0, V1 = 6, 1005     # Versements : 1000 lignes
W0, W1 = 13, 38      # Défi : 26 lignes × 2 blocs

OBJECTIFS_EXEMPLE = [("Fonds d'urgence", 3000, dt.date(2027, 12, 31)),
                     ("Vacances d'été", 1200, dt.date(2027, 6, 30)),
                     ("Nouveau téléphone", 600, dt.date(2027, 4, 30)),
                     ("Cadeaux de Noël", 400, dt.date(2027, 11, 30)),
                     ("Nouvelle voiture", 8000, dt.date(2028, 12, 31))]
MENSUEL_EXEMPLE = {"Fonds d'urgence": 150, "Vacances d'été": 200, "Nouveau téléphone": 150,
                   "Cadeaux de Noël": 30, "Nouvelle voiture": 100}


def versements_exemple():
    rows = []
    for m in range(1, 5):
        for goal, amount in MENSUEL_EXEMPLE.items():
            rows.append((dt.date(YEAR, m, 5), goal, amount, "Virement automatique"))
    rows.append((dt.date(YEAR, 2, 17), "Fonds d'urgence", -80, "Réparation de la machine à laver"))
    rows.append((dt.date(YEAR, 3, 22), "Vacances d'été", 120, "Vente d'un vieux vélo"))
    rows.sort(key=lambda r: r[0])
    return rows


def build(example: bool):
    wb = new_workbook(NAME)
    guide = wb.active
    guide.title = "Mode d'emploi"
    obj = wb.create_sheet("Objectifs")
    ver = wb.create_sheet("Versements")
    defi = wb.create_sheet("Défi 52 semaines")
    build_objectifs(obj, example)
    build_versements(ver, example)
    build_defi(defi, example)
    build_guide(guide, NAME, "Fixez vos objectifs, suivez chaque euro mis de côté et relevez le défi "
                             "des 52 semaines.", GUIDE_STEPS, GUIDE_TIPS)
    wb.active = 0
    return wb


def build_objectifs(ws, example):
    setup(ws, C_GREEN, {"A": 28, "B": 14, "C": 13, "D": 14, "E": 13, "F": 14, "G": 16, "H": 15})
    title(ws, "A1", "Mes objectifs d'épargne", "Notez vos projets : le classeur calcule combien mettre de "
                                                "côté chaque mois pour y arriver à temps.", "A2")
    tot_d, tot_b = f"SUM(D{G0}:D{G1})", f"SUM(B{G0}:B{G1})"
    tile(ws, "A4", "A5", "TOTAL ÉPARGNÉ", f"={tot_d}", C_GREEN, C_GREEN_L)
    tile(ws, "B4", "B5", "OBJECTIF TOTAL", f"={tot_b}", C_DARK, C_LIGHT, merge=("B4:C4", "B5:C5"))
    tile(ws, "D4", "D5", "PROGRESSION", f"=IF({tot_b}=0,0,MIN(1,{tot_d}/{tot_b}))", C_DARK, C_LIGHT,
         merge=("D4:E4", "D5:E5"), fmt=PCT)
    tile(ws, "F4", "F5", "À METTRE DE CÔTÉ PAR MOIS", f"=SUM(G{G0}:G{G1})", C_CORAL, C_CORAL_L,
         merge=("F4:H4", "F5:H5"))
    ws.row_dimensions[4].height = 18
    ws.row_dimensions[5].height = 34

    header_row(ws, G0 - 1, 1, ["Objectif", "Montant visé", "Date visée", "Déjà épargné", "Reste",
                               "Progression", "À mettre de côté / mois", "État"], height=32)
    for i, r in enumerate(range(G0, G1 + 1)):
        if example:
            name, target, date = OBJECTIFS_EXEMPLE[i] if i < len(OBJECTIFS_EXEMPLE) else (None,) * 3
        else:
            name, target, date = (("Exemple : vacances d'été", 1200, dt.date(YEAR, 6, 30)) if i == 0
                                  else (None,) * 3)
        put(ws, f"A{r}", name, bg=C_INPUT, border=BOX, f=font(bold=True))
        put(ws, f"B{r}", target, bg=C_INPUT, border=BOX, fmt=EUR0)
        put(ws, f"C{r}", date, bg=C_INPUT, border=BOX, fmt=DATE, align=CENTER)
        put(ws, f"D{r}", f'=IF(A{r}="","",SUMIF(Versements!$B${V0}:$B${V1},A{r},Versements!$C${V0}:$C${V1}))',
            bg=C_CALC, border=BOX, fmt=EUR0, f=font(bold=True, color=C_GREEN))
        put(ws, f"E{r}", f'=IF(A{r}="","",MAX(0,N(B{r})-D{r}))', bg=C_CALC, border=BOX, fmt=EUR0)
        put(ws, f"F{r}", f'=IF(OR(A{r}="",N(B{r})=0),"",MAX(0,MIN(1,D{r}/B{r})))', bg=C_CALC, border=BOX,
            fmt=PCT, align=CENTER, f=font(bold=True))
        months_left = f"MAX(1,(YEAR(C{r})-YEAR(TODAY()))*12+MONTH(C{r})-MONTH(TODAY()))"
        put(ws, f"G{r}", f'=IF(A{r}="","",IF(OR(N(E{r})=0,C{r}=""),"",E{r}/{months_left}))',
            bg=C_CALC, border=BOX, fmt=EUR0)
        put(ws, f"H{r}", f'=IF(A{r}="","",IF(AND(N(B{r})>0,D{r}>=B{r}),"Atteint !",'
                         f'IF(AND(C{r}<>"",C{r}<TODAY()),"Date dépassée","En cours")))',
            bg=C_CALC, border=BOX, align=CENTER, f=font(9, color=C_MUTED))
        ws.row_dimensions[r].height = 22
    ws.conditional_formatting.add(f"F{G0}:F{G1}", DataBarRule(start_type="num", start_value=0, end_type="num",
                                                             end_value=1, color="7CC4A8"))
    ws.conditional_formatting.add(f"H{G0}:H{G1}", CellIsRule(operator="equal", formula=['"Atteint !"'],
                                                             fill=fill(C_GREEN),
                                                             font=Font(name=FONT, bold=True, color=C_WHITE)))
    ws.conditional_formatting.add(f"H{G0}:H{G1}", CellIsRule(operator="equal", formula=['"Date dépassée"'],
                                                             font=Font(name=FONT, bold=True, color=C_CORAL)))
    put(ws, f"A{G1 + 2}", "« À mettre de côté / mois » = ce qu'il reste à épargner, divisé par le nombre de mois "
                          "jusqu'à la date visée.", f=font(9, italic=True, color=C_MUTED))
    fit_width(ws, landscape=True)


def build_versements(ws, example):
    setup(ws, C_DARK, {"A": 13, "B": 28, "C": 14, "D": 40})
    title(ws, "A1", "Mes versements", "Notez chaque somme mise de côté. Pour un retrait, mettez un montant "
                                      "négatif (par exemple -50).", "A2")
    header_row(ws, V0 - 1, 1, ["Date", "Objectif", "Montant", "Note"])
    data = versements_exemple() if example else [(dt.date(YEAR, 1, 5), "Exemple : vacances d'été", 100,
                                                  "Ligne d'exemple, à effacer")]
    for i, r in enumerate(range(V0, V1 + 1)):
        d, goal, amount, note = data[i] if i < len(data) else (None,) * 4
        put(ws, f"A{r}", d, bg=C_INPUT, border=BOX, fmt=DATE, align=CENTER)
        put(ws, f"B{r}", goal, bg=C_INPUT, border=BOX)
        put(ws, f"C{r}", amount, bg=C_INPUT, border=BOX, fmt=EUR)
        put(ws, f"D{r}", note, bg=C_INPUT, border=BOX, f=font(9, italic=True, color=C_MUTED))
    list_validation(ws, f"B{V0}:B{V1}", f"Objectifs!$A${G0}:$A${G1}",
                    "Choisissez un objectif de la liste (à créer dans l'onglet Objectifs).")
    ws.conditional_formatting.add(f"C{V0}:C{V1}", CellIsRule(operator="lessThan", formula=["0"],
                                                             font=Font(name=FONT, bold=True, color=C_CORAL)))
    ws.auto_filter.ref = f"A{V0 - 1}:D{V1}"
    ws.freeze_panes = f"A{V0}"
    fit_width(ws)


def build_defi(ws, example):
    setup(ws, C_CORAL, {"A": 8, "B": 12, "C": 13, "D": 6, "E": 4, "F": 8, "G": 12, "H": 13, "I": 6})
    ws.merge_cells("A1:I1")
    put(ws, "A1", "Défi 52 semaines", f=font(18, True, C_DARK))
    ws.row_dimensions[1].height = 30
    ws.merge_cells("A2:I2")
    put(ws, "A2", f"Chaque semaine, mettez de côté un peu plus que la semaine d'avant. Cochez {CHECK} quand "
                  "c'est fait.", f=font(9, italic=True, color=C_MUTED))

    done = (f'SUMPRODUCT((D{W0}:D{W1}<>"")*C{W0}:C{W1})+SUMPRODUCT((I{W0}:I{W1}<>"")*H{W0}:H{W1})')
    weeks = f"(COUNTA(D{W0}:D{W1})+COUNTA(I{W0}:I{W1}))"
    tile(ws, "A4", "A5", "DÉJÀ ÉPARGNÉ", f"={done}", C_GREEN, C_GREEN_L, merge=("A4:D4", "A5:D5"))
    tile(ws, "F4", "F5", "OBJECTIF DU DÉFI", f"=SUM(C{W0}:C{W1})+SUM(H{W0}:H{W1})", C_CORAL, C_CORAL_L,
         merge=("F4:I4", "F5:I5"))
    ws.row_dimensions[4].height = 18
    ws.row_dimensions[5].height = 34
    ws.merge_cells("A6:I6")
    put(ws, "A6", f'={weeks}&" semaines réussies sur 52"', f=font(11, True, C_DARK), align=CENTER)
    ws.row_dimensions[6].height = 22

    settings = [(8, "Date de départ", dt.date(YEAR, 1, 4), DATE, "Un lundi, par exemple."),
                (9, "Mise de la semaine 1", 1, EUR, "1 € → 1 378 € en un an ; 2 € → 2 756 €."),
                (10, "Sens du défi", "Croissant", None,
                 "Décroissant : on commence par la plus grosse semaine.")]
    for row, label, value, fmt, hint in settings:
        ws.merge_cells(f"A{row}:B{row}")
        put(ws, f"A{row}", label, f=font(bold=True), align=RIGHT)
        ws.merge_cells(f"C{row}:D{row}")
        put(ws, f"C{row}", value, bg=C_INPUT, border=BOX, fmt=fmt, align=CENTER, f=font(bold=True))
        ws[f"D{row}"].border = BOX
        ws.merge_cells(f"F{row}:I{row}")
        put(ws, f"F{row}", hint, f=font(8, italic=True, color=C_MUTED))
    list_validation(ws, "C10", '"Croissant,Décroissant"', "Choisissez Croissant ou Décroissant.")

    header_row(ws, W0 - 1, 1, ["Sem.", "Date", "Montant", CHECK], color=C_DARK)
    header_row(ws, W0 - 1, 6, ["Sem.", "Date", "Montant", CHECK], color=C_DARK)
    for i, r in enumerate(range(W0, W1 + 1)):
        for col_sem, col_date, col_amt, col_chk, n in (("A", "B", "C", "D", i + 1), ("F", "G", "H", "I", i + 27)):
            put(ws, f"{col_sem}{r}", n, bg=C_CALC, border=BOX, align=CENTER, f=font(bold=True, color=C_DARK))
            put(ws, f"{col_date}{r}", f'=IF($C$8="","",$C$8+7*({col_sem}{r}-1))', bg=C_CALC, border=BOX,
                fmt=DATE_SHORT, align=CENTER, f=font(9, color=C_MUTED))
            put(ws, f"{col_amt}{r}", f'=$C$9*IF($C$10="Décroissant",53-{col_sem}{r},{col_sem}{r})',
                bg=C_CALC, border=BOX, fmt=EUR, align=RIGHT)
            put(ws, f"{col_chk}{r}", CHECK if (example and n <= 14) else None, bg=C_INPUT, border=BOX,
                align=CENTER, f=font(12, True, C_GREEN))
    for chk, rng in (("D", f"A{W0}:C{W1}"), ("I", f"F{W0}:H{W1}")):
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'${chk}{W0}<>""'], fill=fill(C_GREEN_L),
                                                       font=Font(name=FONT, bold=True, color=C_GREEN)))
        dv = DataValidation(type="list", formula1=f'"{CHECK}"', allow_blank=True, showErrorMessage=False)
        dv.add(f"{chk}{W0}:{chk}{W1}")
        ws.add_data_validation(dv)
    fit_width(ws)


GUIDE_STEPS = [
    ("Objectifs", "Notez vos projets (vacances, voiture, fonds d'urgence…), le montant visé et la date. "
                  "Le classeur calcule combien mettre de côté chaque mois pour y arriver."),
    ("Versements", "À chaque fois que vous mettez de l'argent de côté, notez la date, l'objectif et le montant. "
                   "La progression se met à jour toute seule."),
    ("Défi 52 semaines", f"Choisissez la date de départ et la mise de la première semaine, puis cochez {CHECK} "
                         "chaque semaine réussie. Avec 1 € au départ, vous aurez 1 378 € au bout d'un an."),
]
GUIDE_TIPS = [
    "Pour un retrait (une dépense imprévue payée avec l'épargne), notez un montant négatif dans Versements.",
    "Si vous renommez un objectif, corrigez aussi les anciens versements qui portent l'ancien nom.",
    "Le défi 52 semaines se suit à part. S'il sert un de vos objectifs, notez aussi chaque semaine réussie "
    "dans l'onglet Versements.",
    "« Date dépassée » s'affiche quand la date visée est passée sans que l'objectif soit atteint : "
    "il suffit de choisir une nouvelle date.",
    "Effacez les lignes d'exemple avant de commencer.",
    "Fonctionne avec Excel, LibreOffice et Google Sheets (Fichier > Importer).",
]


if __name__ == "__main__":
    for example, fname in [(False, "Epargne_defi_52_semaines.xlsx"), (True, "Epargne_defi_52_semaines_exemple.xlsx")]:
        wb = build(example)
        wb.save(OUT / fname)
        print("écrit :", OUT / fname)
