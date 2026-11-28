#!/usr/bin/env python3
"""Génère « Comptes à deux » : dépenses communes d'un couple ou d'une colocation,
et qui doit combien à qui (version vierge à vendre + version remplie pour les photos).

Usage : python3 build_comptes.py
"""
import random
import sys
import datetime as dt
from pathlib import Path

from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import DataBarRule

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from commun import *  # noqa: E402,F403

OUT = Path(__file__).parent
NAME = "Comptes à deux"
YEAR = 2027

CATEGORIES = ["Courses", "Loyer", "Factures", "Restaurants", "Sorties", "Vacances", "Maison", "Transports",
              "Animaux", "Enfants", "Cadeaux", "Santé", "Abonnements", "Divers"]
C0, C1 = 16, 35        # Réglages : 20 catégories
D0, D1 = 5, 1004       # Dépenses : 1000 lignes
M0 = 11                # Bilan : première ligne des mois
P1, P2 = "Réglages!$B$3", "Réglages!$B$4"
MODES = "Moitié-moitié,Selon les revenus,Pourcentage libre"
REMB = "Remboursement"


def dep(col):
    return f"Dépenses!${col}${D0}:${col}${D1}"


def depenses_exemple():
    rnd = random.Random(7)
    a, b = "Camille", "Sam"
    rows = []
    for m in range(1, 4):
        rows.append((dt.date(YEAR, m, 2), "Loyer", "Loyer", 950, a, None))
        rows.append((dt.date(YEAR, m, 8), "Électricité", "Factures", 78, b, None))
        rows.append((dt.date(YEAR, m, 10), "Box internet", "Factures", 34.99, b, None))
        rows.append((dt.date(YEAR, m, 15), "Streaming", "Abonnements", 13.49, a, None))
        for _ in range(4):
            rows.append((dt.date(YEAR, m, rnd.randint(1, 28)), "Courses de la semaine", "Courses",
                         round(rnd.uniform(55, 140), 2), rnd.choice([a, b]), None))
        rows.append((dt.date(YEAR, m, rnd.randint(1, 28)), "Restaurant", "Restaurants",
                     round(rnd.uniform(35, 70), 2), rnd.choice([a, b]), None))
        rows.append((dt.date(YEAR, m, rnd.randint(1, 28)), "Cinéma", "Sorties", 23, rnd.choice([a, b]), None))
    rows.append((dt.date(YEAR, 1, 20), "Cadeau pour la mère de Sam", "Cadeaux", 60, a, 0))
    rows.append((dt.date(YEAR, 2, 14), "Week-end à la mer", "Vacances", 420, b, None))
    rows.append((dt.date(YEAR, 2, 28), "Virement à Camille", REMB, 200, b, None))
    rows.append((dt.date(YEAR, 3, 12), "Aspirateur", "Maison", 149, a, None))
    rows.sort(key=lambda r: r[0])
    return rows


def build(example: bool):
    wb = new_workbook(NAME)
    guide = wb.active
    guide.title = "Mode d'emploi"
    reg = wb.create_sheet("Réglages")
    dps = wb.create_sheet("Dépenses")
    bil = wb.create_sheet("Bilan")
    lis = wb.create_sheet("Listes")
    build_listes(lis)
    build_reglages(reg, example)
    build_depenses(dps, example)
    build_bilan(bil)
    build_guide(guide, NAME, "Pour un couple ou une colocation : notez les dépenses communes, "
                             "le classeur dit qui doit combien à qui.", GUIDE_STEPS, GUIDE_TIPS)
    lis.sheet_state = "hidden"
    wb.active = 0
    return wb


def build_listes(ws):
    setup(ws, C_MUTED)
    put(ws, "A1", "Catégories", f=font(bold=True))
    put(ws, "B1", "Nombre", f=font(bold=True))
    ws["B2"] = f"=COUNTA(Réglages!$A${C0}:$A${C1})"
    for k in range(1, 22):
        ws[f"A{k + 1}"] = (f'=IF({k}<=$B$2,INDEX(Réglages!$A${C0}:$A${C1},{k}),'
                           f'IF({k}=$B$2+1,"{REMB}",""))')
    put(ws, "C1", "Prénoms", f=font(bold=True))
    ws["C2"] = f"={P1}"
    ws["C3"] = f"={P2}"
    put(ws, "E1", "Solde de la personne 1", f=font(bold=True))
    ws["E2"] = f"=SUMIF({dep('E')},{P1},{dep('D')})-SUM({dep('G')})"


def build_reglages(ws, example):
    setup(ws, C_MUTED, {"A": 38, "B": 18, "C": 3, "D": 52})
    title(ws, "A1", "Réglages")
    hint = dict(f=font(9, italic=True, color=C_MUTED))
    put(ws, "A3", "Prénom de la personne 1", f=font(bold=True))
    put(ws, "B3", "Camille" if example else "Personne 1", bg=C_INPUT, border=BOX, align=CENTER, f=font(bold=True))
    put(ws, "A4", "Prénom de la personne 2", f=font(bold=True))
    put(ws, "B4", "Sam" if example else "Personne 2", bg=C_INPUT, border=BOX, align=CENTER, f=font(bold=True))

    put(ws, "A6", "Comment partager les dépenses ?", f=font(bold=True))
    put(ws, "B6", "Selon les revenus" if example else "Moitié-moitié", bg=C_INPUT, border=BOX, align=CENTER)
    list_validation(ws, "B6", f'"{MODES}"', "Choisissez une option dans la liste.")
    put(ws, "D6", "Moitié-moitié, selon les revenus de chacun, ou un pourcentage libre.", **hint)
    put(ws, "A7", '="Revenus mensuels de "&B3')
    put(ws, "B7", 2400 if example else None, bg=C_INPUT, border=BOX, fmt=EUR0)
    put(ws, "D7", "Seulement pour « Selon les revenus ».", **hint)
    put(ws, "A8", '="Revenus mensuels de "&B4')
    put(ws, "B8", 1800 if example else None, bg=C_INPUT, border=BOX, fmt=EUR0)
    put(ws, "A9", '="Part de "&B3&" (pourcentage libre)"')
    put(ws, "B9", None, bg=C_INPUT, border=BOX, fmt=PCT, align=CENTER)
    put(ws, "D9", "Seulement pour « Pourcentage libre » (par exemple 60 %).", **hint)

    put(ws, "A11", '="Part de "&B3', f=font(bold=True, color=C_DARK), bg=C_LIGHT, border=BOX)
    put(ws, "B11", '=IF(B6="Selon les revenus",IF(N(B7)+N(B8)=0,0.5,N(B7)/(N(B7)+N(B8))),'
                   'IF(B6="Pourcentage libre",N(B9),0.5))',
        f=font(11, True, C_DARK), bg=C_LIGHT, border=BOX, fmt=PCT, align=CENTER)
    put(ws, "A12", '="Part de "&B4', f=font(bold=True, color=C_DARK), bg=C_LIGHT, border=BOX)
    put(ws, "B12", "=1-B11", f=font(11, True, C_DARK), bg=C_LIGHT, border=BOX, fmt=PCT, align=CENTER)
    put(ws, "D11", "Le partage utilisé pour chaque dépense, sauf si vous en indiquez un autre sur la ligne.", **hint)

    put(ws, f"A{C0 - 2}", "Catégories de dépenses", f=font(12, True, C_DARK))
    header_row(ws, C0 - 1, 1, ["Catégorie"])
    for i, r in enumerate(range(C0, C1 + 1)):
        put(ws, f"A{r}", CATEGORIES[i] if i < len(CATEGORIES) else None, bg=C_INPUT, border=BOX)
    put(ws, f"D{C0}", f"La catégorie « {REMB} » existe toujours : utilisez-la quand l'un", **hint)
    put(ws, f"D{C0 + 1}", "rembourse l'autre. Pas de ligne vide au milieu de la liste.", **hint)
    fit_width(ws)


def build_depenses(ws, example):
    setup(ws, C_CORAL, {"A": 12, "B": 32, "C": 18, "D": 12, "E": 14, "F": 15, "G": 13, "H": 13})
    title(ws, "A1", "Dépenses communes", f"Une ligne par dépense commune. Quand l'un rembourse l'autre, "
                                         f"choisissez la catégorie « {REMB} ».", "A2")
    header_row(ws, D0 - 1, 1, ["Date", "Description", "Catégorie", "Montant", "Payé par",
                               f'="Part de "&{P1}&" si différente"', f'="Pour "&{P1}', f'="Pour "&{P2}'],
               height=32)
    data = depenses_exemple() if example else [(dt.date(YEAR, 1, 3), "Exemple : courses (ligne à effacer)",
                                                "Courses", 85.40, "Personne 1", None)]
    for i, r in enumerate(range(D0, D1 + 1)):
        d, desc, cat, amount, payer, share = data[i] if i < len(data) else (None,) * 6
        put(ws, f"A{r}", d, bg=C_INPUT, border=BOX, fmt=DATE, align=CENTER)
        put(ws, f"B{r}", desc, bg=C_INPUT, border=BOX)
        put(ws, f"C{r}", cat, bg=C_INPUT, border=BOX)
        put(ws, f"D{r}", amount, bg=C_INPUT, border=BOX, fmt=EUR)
        put(ws, f"E{r}", payer, bg=C_INPUT, border=BOX, align=CENTER)
        put(ws, f"F{r}", share, bg=C_INPUT, border=BOX, fmt=PCT, align=CENTER)
        put(ws, f"G{r}", f'=IF(OR(N(D{r})=0,E{r}=""),"",D{r}*IF(C{r}="{REMB}",IF(E{r}={P1},0,1),'
                         f'IF(F{r}<>"",F{r},Réglages!$B$11)))', bg=C_CALC, border=BOX, fmt=EUR)
        put(ws, f"H{r}", f'=IF(G{r}="","",D{r}-G{r})', bg=C_CALC, border=BOX, fmt=EUR)
    list_validation(ws, f"C{D0}:C{D1}", "Listes!$A$2:$A$22", "Choisissez une catégorie dans la liste.")
    list_validation(ws, f"E{D0}:E{D1}", "Listes!$C$2:$C$3", "Choisissez qui a payé dans la liste.")
    dv = DataValidation(type="decimal", operator="between", formula1="0", formula2="1", allow_blank=True,
                        showErrorMessage=True, error="Indiquez un pourcentage entre 0 % et 100 %.")
    dv.add(f"F{D0}:F{D1}")
    ws.add_data_validation(dv)
    ws.auto_filter.ref = f"A{D0 - 1}:H{D1}"
    ws.freeze_panes = f"A{D0}"
    fit_width(ws, landscape=True)


def build_bilan(ws):
    setup(ws, C_GREEN, {c: 15 for c in "ABCDEF"} | {"G": 3, "H": 20, "I": 13, "J": 10})
    title(ws, "A1", "Qui doit combien ?")
    ws.merge_cells("A3:F4")
    bal = "Listes!$E$2"
    put(ws, "A3", f'=IF(ABS({bal})<0.005,"Vous êtes quittes !",IF({bal}>0,'
                  f'{P2}&" doit "&FIXED({bal},2)&" € à "&{P1},{P1}&" doit "&FIXED(-{bal},2)&" € à "&{P2}))',
        f=font(16, True, C_WHITE), bg=C_DARK, align=CENTER)
    ws.row_dimensions[3].height = 24
    ws.row_dimensions[4].height = 24

    not_remb = f'{dep("C")},"<>{REMB}"'
    tile(ws, "A6", "A7", "DÉPENSES COMMUNES", f"=SUMIF({dep('C')},\"<>{REMB}\",{dep('D')})", C_DARK, C_LIGHT,
         merge=("A6:B6", "A7:B7"))
    tile(ws, "C6", "C7", f'="PAYÉ PAR "&UPPER({P1})', f"=SUMIFS({dep('D')},{dep('E')},{P1},{not_remb})",
         C_GREEN, C_GREEN_L, merge=("C6:D6", "C7:D7"))
    tile(ws, "E6", "E7", f'="PAYÉ PAR "&UPPER({P2})', f"=SUMIFS({dep('D')},{dep('E')},{P2},{not_remb})",
         C_CORAL, C_CORAL_L, merge=("E6:F6", "E7:F7"))
    ws.row_dimensions[6].height = 18
    ws.row_dimensions[7].height = 34
    put(ws, "A8", "Depuis le début, sans compter les remboursements.", f=font(8, italic=True, color=C_MUTED))

    put(ws, "A9", "Année", f=font(bold=True), align=RIGHT)
    put(ws, "B9", YEAR, bg=C_INPUT, border=BOX, align=CENTER, f=font(bold=True))
    unlock(ws, "B9")
    header_row(ws, M0 - 1, 1, ["Mois", "Dépenses communes", f'="Payé par "&{P1}', f'="Payé par "&{P2}',
                               f'="Part de "&{P1}', f'="Part de "&{P2}'], height=32)
    for m in range(1, 13):
        r = M0 + m - 1
        when = f'{dep("A")},">="&DATE($B$9,{m},1),{dep("A")},"<"&DATE($B$9,{m}+1,1),{not_remb}'
        put(ws, f"A{r}", MONTHS[m - 1], border=BOX, f=font(bold=True))
        put(ws, f"B{r}", f"=SUMIFS({dep('D')},{when})", bg=C_CALC, border=BOX, fmt=EUR0)
        put(ws, f"C{r}", f"=SUMIFS({dep('D')},{dep('E')},{P1},{when})", bg=C_CALC, border=BOX, fmt=EUR0)
        put(ws, f"D{r}", f"=SUMIFS({dep('D')},{dep('E')},{P2},{when})", bg=C_CALC, border=BOX, fmt=EUR0)
        put(ws, f"E{r}", f"=SUMIFS({dep('G')},{when})", bg=C_CALC, border=BOX, fmt=EUR0)
        put(ws, f"F{r}", f"=SUMIFS({dep('H')},{when})", bg=C_CALC, border=BOX, fmt=EUR0)
    rt = M0 + 12
    put(ws, f"A{rt}", "Total", f=font(bold=True, color=C_DARK), bg=C_LIGHT, border=BOX)
    for col in "BCDEF":
        put(ws, f"{col}{rt}", f"=SUM({col}{M0}:{col}{rt - 1})", f=font(bold=True, color=C_DARK), bg=C_LIGHT,
            border=BOX, fmt=EUR0)

    header_row(ws, M0 - 1, 8, ["Par catégorie (année choisie)", "Total", "Part"], height=32)
    tot = f"SUM($I${M0}:$I${M0 + 19})"
    for i in range(20):
        r = M0 + i
        put(ws, f"H{r}", f'=IF(Réglages!A{C0 + i}="","",Réglages!A{C0 + i})', border=BOX)
        put(ws, f"I{r}", f'=IF(H{r}="","",SUMIFS({dep("D")},{dep("C")},H{r},{dep("A")},">="&DATE($B$9,1,1),'
                         f'{dep("A")},"<"&DATE($B$9+1,1,1)))', bg=C_CALC, border=BOX, fmt=EUR0)
        put(ws, f"J{r}", f'=IF(OR(H{r}="",N(I{r})=0),"",I{r}/{tot})', bg=C_CALC, border=BOX, fmt=PCT,
            align=CENTER)
    ws.conditional_formatting.add(f"J{M0}:J{M0 + 19}", DataBarRule(start_type="num", start_value=0,
                                                                  end_type="num", end_value=1, color="8FB8C2"))

    ch = BarChart()
    ch.type = "col"
    ch.title = "Qui a payé quoi, mois par mois"
    ch.add_data(Reference(ws, min_col=3, max_col=4, min_row=M0 - 1, max_row=M0 + 11), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=M0, max_row=M0 + 11))
    for s, color in zip(ch.series, [C_GREEN, C_CORAL]):
        s.graphicalProperties.solidFill = color
        s.graphicalProperties.line.solidFill = color
    ch.gapWidth = 60
    ch.y_axis.number_format = '#,##0" €"'
    ch.legend.position = "b"
    style_chart(ch)
    ch.width, ch.height = 17, 7.5
    ws.add_chart(ch, f"A{rt + 2}")
    ws.protection.sheet = True
    fit_width(ws, landscape=True)
    ws.page_setup.fitToHeight = 1


GUIDE_STEPS = [
    ("Réglages", "Écrivez vos deux prénoms et choisissez comment partager : moitié-moitié, selon les revenus "
                 "de chacun, ou un pourcentage libre. Adaptez la liste des catégories si besoin."),
    ("Dépenses", "Notez chaque dépense commune : la date, une description, la catégorie, le montant et qui a "
                 "payé. Le classeur calcule la part de chacun."),
    ("Bilan", "Voyez tout de suite qui doit combien à qui, le détail mois par mois et les dépenses par "
              "catégorie."),
]
GUIDE_TIPS = [
    f"Pour équilibrer les comptes, celui qui doit fait un virement à l'autre : notez-le dans Dépenses avec la "
    f"catégorie « {REMB} » et « Payé par » = celui qui rembourse.",
    "Une dépense à partager autrement (un cadeau pour la famille de l'un, par exemple) : remplissez « Part de… "
    "si différente ». 0 % = l'autre paie tout, 100 % = c'est pour vous seul.",
    "Si vous changez un prénom dans Réglages, corrigez aussi la colonne « Payé par » des anciennes dépenses.",
    "L'onglet Bilan est protégé pour éviter d'effacer un calcul (sauf la case Année). Pour le modifier : "
    "menu Révision > Ôter la protection de la feuille (pas de mot de passe).",
    "Effacez la ligne d'exemple avant de commencer.",
    "Fonctionne avec Excel, LibreOffice et Google Sheets : partagez le fichier pour que chacun note ses dépenses.",
]


if __name__ == "__main__":
    for example, fname in [(False, "Comptes_a_deux.xlsx"), (True, "Comptes_a_deux_exemple.xlsx")]:
        wb = build(example)
        wb.save(OUT / fname)
        print("écrit :", OUT / fname)
