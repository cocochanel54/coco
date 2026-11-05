# Construit le fichier Excel d'attachement à remplir chaque mois, un par sous-traitant,
# dans le style du suivi maintenance (titre rouge, colonnes du suivi, NB Ticket en vert, case jaune, TOTAL HT).
#
#   python3 build_modele.py                      -> Attachement_sous-traitant.xlsx (modèle, nom à remplir)
#   python3 build_modele.py CET "LOR" "HK"       -> un fichier par sous-traitant : Attachement CET.xlsx, ...
#
# Chaque mois : coller l'export des interventions du sous-traitant dans l'onglet « Export » (en A1, avec
# la ligne des titres). L'onglet « Attachement » se remplit et calcule tout seul, avec un forfait d'équipes
# par projet (le projet vient de la colonne Agence, voir PROJETS) :
#   nombre d'équipes du projet = tickets au forfait du projet ÷ jours ouvrés du mois ÷ 4,
#   montant forfait du projet = équipes × 12 000 € ;
#   le forfait couvre les tickets classiques et les travaux de 8h à 18h (du lundi au vendredi hors jours fériés) ;
#   en dehors, le ticket est sur devis et peut entraîner une plus-value, saisie à la main.
import sys, datetime as dt, pathlib
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter as L

ICI = pathlib.Path(__file__).parent
NB_TICKETS = 300
EUR = '#,##0.00\\ "€"'

# Projets : code, libellé, base mensuelle par équipe, tickets demandés par jour et par équipe, agences de l'export
PROJETS = [
    ('P2200154', 'CIRCET NANCY RES', 12000, 4, ['CIRCET NANCY (RES)']),
    ('P2100474', 'CIRCET 54 RES', 12000, 4, ['CIRCET 54 (RES)']),
    ('P1800267', 'CIRCET GLAUCOS RES + CIRCET RESOPTIC RES', 12000, 4, ['CIRCET GLAUCOS (RES)', 'CIRCET RESOPTIC (RES)']),
]
LIGNES_AGENCES = 16          # place dans le tableau Agences -> projet
INCONNUE, SANS_AGENCE = 'Agence inconnue', 'Sans agence'

FT = 'Calibri'
fin = Side(style='thin', color='A6A6A6')
fort = Side(style='medium', color='000000')
BORD = Border(top=fin, bottom=fin, left=fin, right=fin)
fond = lambda c: PatternFill('solid', fgColor=c)
GRIS, FONCE, VERT, JAUNE, SAISIE = 'BFBFBF', '767171', 'C6E0B4', 'FFFF00', 'FFF2CC'
centre = Alignment(horizontal='center', vertical='center', wrap_text=True)
f9, f9b = Font(name=FT, size=9), Font(name=FT, size=9, bold=True)
note = Font(name=FT, size=9, italic=True, color='595959')

# Colonnes de la liste des tickets. Celles avec un titre sont lues dans l'export (par leur titre).
COLS = [('A', 'N°', 20), ('B', 'Description', 42), ('C', 'EQUIPE', 16), ('D', 'Agence', 18),
        ('F', 'Client', 16), ('G', 'SITE', 18), ('H', 'Date de début', 15), ('I', 'Clôturée', 15), ('J', 'Type inter.', 20),
        ('K', 'Commentaire facturation', 22), ('L', 'Materiel utilisé', 16), ('M', 'CERFA/ASSURANCE', 16),
        ('R', 'Commentaire', 24)]
C_AGENCE, C_PROJET, C_DEB = 'D', 'E', 'H'
C_F, C_HF, C_MOT, C_PV, C_AIDE = 'N', 'O', 'P', 'Q', 'S'
DERNIERE_COL = 18            # R


def paques(y):
    a, b, c = y % 19, y // 100, y % 100
    d, e = b // 4, b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i, k = c // 4, c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    mois = (h + l - 7 * m + 114) // 31
    jour = (h + l - 7 * m + 114) % 31 + 1
    return dt.date(y, mois, jour)


def feries(annees):
    out = []
    for y in annees:
        p = paques(y)
        out += [(dt.date(y, 1, 1), "Jour de l'an"), (p + dt.timedelta(1), 'Lundi de Pâques'), (dt.date(y, 5, 1), 'Fête du Travail'),
                (dt.date(y, 5, 8), 'Victoire 1945'), (p + dt.timedelta(39), 'Ascension'), (p + dt.timedelta(50), 'Lundi de Pentecôte'),
                (dt.date(y, 7, 14), 'Fête nationale'), (dt.date(y, 8, 15), 'Assomption'), (dt.date(y, 11, 1), 'Toussaint'),
                (dt.date(y, 11, 11), 'Armistice 1918'), (dt.date(y, 12, 25), 'Noël')]
    return sorted(out)


def cle(x):
    """Clé de comparaison d'une agence : sans parenthèses, espaces ni tirets, en majuscules (« CIRCET 54 (RES) » = « CIRCET 54 RES »)."""
    return f'UPPER(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE({x},"(",""),")","")," ",""),"-",""))'


def nom_global(wb, nom, ref):
    wb.defined_names[nom] = DefinedName(nom, attr_text=ref)


def cellule(ws, ref, valeur=None, font=None, fill=None, fmt=None, align=None, border=BORD):
    c = ws[ref]
    if valeur is not None:
        c.value = valeur
    if font:
        c.font = font
    if fill:
        c.fill = fill
    if fmt:
        c.number_format = fmt
    if align:
        c.alignment = align
    if border:
        c.border = border
    return c


def construire(sous_traitant, chemin):
    n = len(PROJETS)
    R_ENT = 4                                   # en-tête du récapitulatif par projet
    R_P1 = R_ENT + 1                            # premier projet
    R_INC = R_P1 + n                            # tickets non classés
    R_TOT = R_INC + 1
    R_NOTE = R_TOT + 1
    R_T1 = R_NOTE + 2                           # en-têtes de la liste (2 lignes)
    P0 = R_T1 + 2                               # premier ticket
    P1 = P0 + NB_TICKETS - 1                    # dernier ticket
    rng = lambda col: f'${col}${P0}:${col}${P1}'

    wb = openpyxl.Workbook()
    at = wb.active
    at.title = 'Attachement'
    ex = wb.create_sheet('Export')
    rg = wb.create_sheet('Réglages')
    at.sheet_properties.tabColor = '70AD47'
    ex.sheet_properties.tabColor = '5B9BD5'

    # ================================================================== Réglages
    for col, w in zip('ABCD', (44, 40, 24, 30)):
        rg.column_dimensions[col].width = w
    cellule(rg, 'A1', "Réglages de l'attachement", Font(name=FT, size=14, bold=True), border=None)
    cellule(rg, 'A2', "Chaque mois : 1) efface l'onglet Export (Ctrl+A puis Suppr) ; 2) colle en A1 de l'onglet Export l'export des "
            "interventions du sous-traitant, avec sa ligne de titres ; 3) l'onglet Attachement se remplit et calcule tout seul, "
            "projet par projet ; 4) pour un ticket hors forfait sur devis, saisis la plus-value colonne Q de l'Attachement ; "
            "5) si une agence apparaît « Agence inconnue », ajoute-la dans le tableau « Agences de l'export » ci-dessous ; "
            "6) enregistre sous « Attachement <sous-traitant> <mois>.xlsx » et envoie-le.",
            Font(name=FT, size=10), align=Alignment(wrap_text=True, vertical='top'), border=None)
    rg.merge_cells('A2:D4')
    for r in (2, 3, 4):
        rg.row_dimensions[r].height = 20

    params = [
        (6, 'ST', 'Sous-traitant', sous_traitant, None, "Nom affiché dans le titre de l'attachement."),
        (7, 'MOIS', "Mois de l'attachement",
         f'=IF(COUNT(Attachement!{rng(C_AIDE)})=0,"",DATE(YEAR(MEDIAN(Attachement!{rng(C_AIDE)})),MONTH(MEDIAN(Attachement!{rng(C_AIDE)})),1))',
         'mmmm yyyy', "Lu dans l'export. Tu peux taper une date, ex. 01/09/2026."),
        (8, 'JOURS', 'Jours ouvrés du mois', '=IF(MOIS="",0,NETWORKDAYS(MOIS,EOMONTH(MOIS,0),FERIES))', '0',
         'Lundi au vendredi hors jours fériés. Tu peux taper un autre nombre.'),
        (9, 'HDEB', 'Début des heures ouvrables', dt.time(8, 0), 'hh:mm', ''),
        (10, 'HFIN', 'Fin des heures ouvrables', dt.time(18, 0), 'hh:mm', '18:00 pile reste au forfait.'),
        (11, 'WE', 'Week-ends et jours fériés hors forfait', 'OUI', None, 'OUI ou NON.'),
        (12, 'LIGNE_TITRES', "Ligne des titres dans l'onglet Export", 1, '0', "Si l'export a des lignes au-dessus de ses titres."),
        (13, 'MOIS_TXT', 'Mois en lettres (calculé)',
         '=IF(MOIS="","",CHOOSE(MONTH(MOIS),"JANVIER","FÉVRIER","MARS","AVRIL","MAI","JUIN","JUILLET","AOÛT","SEPTEMBRE",'
         '"OCTOBRE","NOVEMBRE","DÉCEMBRE")&" "&YEAR(MOIS))', None, ''),
    ]
    for r, nom, lib, val, fmt, rem in params:
        cellule(rg, f'A{r}', lib, Font(name=FT, size=10, bold=True), fond(GRIS))
        cellule(rg, f'B{r}', val, Font(name=FT, size=11, bold=True), fond(SAISIE if nom != 'MOIS_TXT' else 'FFFFFF'), fmt,
                Alignment(horizontal='right'))
        cellule(rg, f'C{r}', rem, note, border=None)
        rg.merge_cells(f'C{r}:D{r}')
        nom_global(wb, nom, f"'Réglages'!$B${r}")
    dv = DataValidation(type='list', formula1='"OUI,NON"', allow_blank=False)
    rg.add_data_validation(dv)
    dv.add('B11')

    # projets
    r = 15
    cellule(rg, f'A{r}', "Projets : un forfait d'équipes par projet", Font(name=FT, size=12, bold=True), border=None)
    r += 1
    for col, t in zip('ABCD', ['Code projet', 'Agences', 'Base mensuelle par équipe', 'Tickets demandés par jour et par équipe']):
        cellule(rg, f'{col}{r}', t, Font(name=FT, size=10, bold=True), fond(VERT), align=centre)
    rg.row_dimensions[r].height = 30
    lignes_projet = []
    for code, lib, base, tpj, _ in PROJETS:
        r += 1
        lignes_projet.append(r)
        cellule(rg, f'A{r}', code, Font(name=FT, size=11, bold=True), fond(SAISIE))
        cellule(rg, f'B{r}', lib, Font(name=FT, size=10), fond(SAISIE))
        cellule(rg, f'C{r}', base, Font(name=FT, size=10), fond(SAISIE), EUR)
        cellule(rg, f'D{r}', tpj, Font(name=FT, size=10), fond(SAISIE), '0.##')
    dv_proj = DataValidation(type='list', formula1=f'$A${lignes_projet[0]}:$A${lignes_projet[-1]}', allow_blank=True)
    rg.add_data_validation(dv_proj)

    # agences -> projet
    r += 2
    cellule(rg, f'A{r}', "Agences de l'export : à quel projet elles appartiennent", Font(name=FT, size=12, bold=True), border=None)
    r += 1
    for col, t in zip('ABC', ["Agence (telle qu'écrite dans l'export)", 'Projet', 'Clé de comparaison (calculée)']):
        cellule(rg, f'{col}{r}', t, Font(name=FT, size=10, bold=True), fond(VERT), align=centre)
    cellule(rg, f'D{r}', 'Parenthèses, espaces, tirets et majuscules ne comptent pas : « CIRCET 54 RES » = « CIRCET 54 (RES) ».',
            note, align=Alignment(wrap_text=True, vertical='center'), border=None)
    rg.row_dimensions[r].height = 30
    agences = [(a, code) for code, _, _, _, ags in PROJETS for a in ags]
    a0 = r + 1
    for i in range(LIGNES_AGENCES):
        r += 1
        a, code = agences[i] if i < len(agences) else (None, None)
        cellule(rg, f'A{r}', a, Font(name=FT, size=10), fond(SAISIE))
        cellule(rg, f'B{r}', code, Font(name=FT, size=10, bold=True), fond(SAISIE))
        cellule(rg, f'C{r}', f'=IF(OR(A{r}="",B{r}=""),"#",{cle(f"A{r}")})', Font(name=FT, size=9, color='808080'))
        dv_proj.add(f'B{r}')
    nom_global(wb, 'AG_PROJ', f"'Réglages'!$B${a0}:$B${r}")
    nom_global(wb, 'AG_CLE', f"'Réglages'!$C${a0}:$C${r}")

    # colonnes lues dans l'export
    r += 2
    cellule(rg, f'A{r}', "Colonnes lues dans l'export", Font(name=FT, size=12, bold=True), border=None)
    r += 1
    for col, t in zip('ABCD', ["Colonne de l'attachement", "Titre cherché dans l'export", 'Trouvée en colonne n°', '']):
        cellule(rg, f'{col}{r}', t, Font(name=FT, size=10, bold=True), fond(VERT))
    rang_col = {}
    for lettre, titre, _ in COLS:
        r += 1
        rang_col[lettre] = r
        cellule(rg, f'A{r}', f'Colonne {lettre}', Font(name=FT, size=10))
        cellule(rg, f'B{r}', f'=Attachement!{lettre}{R_T1}', Font(name=FT, size=10))
        cellule(rg, f'C{r}', f'=IFERROR(MATCH(B{r},INDEX(Export!$1:$30,LIGNE_TITRES,0),0),0)', Font(name=FT, size=10))
        cellule(rg, f'D{r}', f'=IF(C{r}>0,"OK","Absente de l\'export")', Font(name=FT, size=10), border=None)
    r += 1
    cellule(rg, f'A{r}', f"Si une colonne est absente, écris dans l'onglet Attachement (ligne {R_T1}) le titre exact de la colonne de l'export.",
            note, border=None)

    # jours fériés
    r += 2
    cellule(rg, f'A{r}', 'Jours fériés', Font(name=FT, size=12, bold=True), border=None)
    cellule(rg, f'B{r}', '(ajoute des dates à la suite si besoin)', note, border=None)
    liste = feries(range(2026, 2033))
    f0 = r + 1
    for d, nm in liste:
        r += 1
        rg.cell(r, 1, d).number_format = 'dd/mm/yyyy'
        rg.cell(r, 2, nm)
    nom_global(wb, 'FERIES', f"'Réglages'!$A${f0}:$A${r + 20}")

    # ================================================================== Attachement
    larg = {l: w for l, _, w in COLS}
    larg.update({C_PROJET: 14, C_F: 13, C_HF: 13, C_MOT: 24, C_PV: 15, C_AIDE: 14})
    for l, w in larg.items():
        at.column_dimensions[l].width = w
    at.column_dimensions[C_AIDE].hidden = True
    der = L(DERNIERE_COL)

    cellule(at, 'A1', 'SUIVI MAINTENANCE RESEAU 54\nLOSANGE', Font(name=FT, size=12, bold=True, color='FF0000'), align=centre, border=None)
    at.merge_cells(f'A1:{der}1')
    at.row_dimensions[1].height = 36
    cellule(at, 'A2', '="Prestation "&ST&" pour le Contrat maintenance "&MOIS_TXT', Font(name=FT, size=11, bold=True), align=centre, border=None)
    at.merge_cells(f'A2:{der}2')
    at.row_dimensions[2].height = 30
    for c in range(1, DERNIERE_COL + 1):
        at.cell(1, c).border = Border(bottom=fort)
        at.cell(2, c).border = Border(bottom=fort)
    at.row_dimensions[3].height = 10

    # ---- récapitulatif par projet (colonnes A à L)
    titres = ['Projet', 'Agences', 'Tickets au forfait', 'Jours ouvrés', 'Tickets demandés / jour / équipe', 'Tickets par jour',
              "NOMBRE D'ÉQUIPES", 'Base mensuelle / équipe', 'Montant forfait', 'Tickets hors forfait (sur devis)',
              'Plus-values sur devis', 'TOTAL HT']
    for i, t in enumerate(titres):
        cellule(at, f'{L(i + 1)}{R_ENT}', t, Font(name=FT, size=9, bold=True), fond(JAUNE if i == 6 else VERT), align=centre)
    cellule(at, f'M{R_ENT}', '="Équipes = tickets au forfait ÷ jours ÷ tickets demandés"', note,
            align=Alignment(wrap_text=True, vertical='center'), border=None)
    at.merge_cells(f'M{R_ENT}:O{R_ENT}')
    at.row_dimensions[R_ENT].height = 36

    proj, forf, hf, mot, pv = rng(C_PROJET), rng(C_F), rng(C_HF), rng(C_MOT), rng(C_PV)
    v10 = Font(name=FT, size=10)
    for i, lp in enumerate(lignes_projet):
        r = R_P1 + i
        cellule(at, f'A{r}', f"='Réglages'!A{lp}", Font(name=FT, size=11, bold=True), align=Alignment(vertical='center'))
        cellule(at, f'B{r}', f"='Réglages'!B{lp}", v10, align=Alignment(vertical='center', wrap_text=True))
        cellule(at, f'C{r}', f'=COUNTIFS({proj},$A{r},{forf},1)', v10, fmt='0')
        cellule(at, f'D{r}', '=JOURS', v10, fmt='0')
        cellule(at, f'E{r}', f"='Réglages'!D{lp}", v10, fmt='0.##')
        cellule(at, f'F{r}', f'=IF(D{r}>0,C{r}/D{r},0)', v10, fmt='0.00')
        cellule(at, f'G{r}', f'=ROUND(IF(AND(D{r}>0,E{r}>0),C{r}/D{r}/E{r},0),2)', Font(name=FT, size=13, bold=True), fond(JAUNE), '0.00',
                Alignment(horizontal='center', vertical='center'))
        cellule(at, f'H{r}', f"='Réglages'!C{lp}", v10, fmt=EUR)
        cellule(at, f'I{r}', f'=G{r}*H{r}', Font(name=FT, size=10, bold=True), fmt=EUR)
        cellule(at, f'J{r}', f'=COUNTIFS({proj},$A{r},{hf},1)', v10, fmt='0')
        cellule(at, f'K{r}', f'=SUMIFS({pv},{proj},$A{r},{hf},1)', v10, fmt=EUR)
        cellule(at, f'L{r}', f'=I{r}+K{r}', Font(name=FT, size=10, bold=True), fmt=EUR)
        at.row_dimensions[r].height = 22
    r = R_INC
    cellule(at, f'A{r}', 'Non classés', Font(name=FT, size=10, bold=True), align=Alignment(vertical='center'))
    cellule(at, f'B{r}', 'Agence absente du tableau des Réglages : à classer', Font(name=FT, size=9, italic=True),
            align=Alignment(vertical='center', wrap_text=True))
    cellule(at, f'C{r}', f'=COUNTIFS({proj},"{INCONNUE}",{forf},1)+COUNTIFS({proj},"{SANS_AGENCE}",{forf},1)', v10, fmt='0')
    cellule(at, f'J{r}', f'=COUNTIFS({proj},"{INCONNUE}",{hf},1)+COUNTIFS({proj},"{SANS_AGENCE}",{hf},1)', v10, fmt='0')
    for col in 'DEFGHIKL':
        cellule(at, f'{col}{r}', None, v10)
    at.conditional_formatting.add(f'A{r}:L{r}', FormulaRule(formula=[f'$C${r}+$J${r}>0'], fill=fond('FFC7CE'), font=Font(color='9C0006', bold=True)))
    at.conditional_formatting.add(f'A{r}:L{r}', FormulaRule(formula=[f'$C${r}+$J${r}=0'], font=Font(color='A6A6A6')))
    r = R_TOT
    cellule(at, f'A{r}', 'TOTAL', Font(name=FT, size=12, bold=True, color='FFFFFF'), fond(FONCE), align=Alignment(horizontal='right', vertical='center'))
    cellule(at, f'B{r}', None, fill=fond(FONCE))
    at.merge_cells(f'A{r}:B{r}')
    som = lambda col, a=R_P1, b=R_INC: f'=SUM({col}{a}:{col}{b})'
    cellule(at, f'C{r}', som('C'), Font(name=FT, size=10, bold=True), fmt='0')
    for col in 'DEFH':
        cellule(at, f'{col}{r}', None)
    cellule(at, f'G{r}', som('G', b=R_INC - 1), Font(name=FT, size=13, bold=True), fond(JAUNE), '0.00', Alignment(horizontal='center', vertical='center'))
    cellule(at, f'I{r}', som('I', b=R_INC - 1), Font(name=FT, size=10, bold=True), fmt=EUR)
    cellule(at, f'J{r}', som('J'), Font(name=FT, size=10, bold=True), fmt='0')
    cellule(at, f'K{r}', som('K', b=R_INC - 1), Font(name=FT, size=10, bold=True), fmt=EUR)
    cellule(at, f'L{r}', som('L', b=R_INC - 1), Font(name=FT, size=13, bold=True), fond(JAUNE), EUR, Alignment(horizontal='center', vertical='center'))
    at.row_dimensions[r].height = 28

    doublons = f'COUNTIF({mot},"Doublon")'
    autres = f'COUNTIF({mot},"Autre mois")'
    cellule(at, f'A{R_NOTE}',
            '="Forfait par projet : tickets classiques et travaux commencés entre "&TEXT(HDEB,"hh:mm")&" et "&TEXT(HFIN,"hh:mm")'
            '&IF(WE="OUI",", du lundi au vendredi hors jours fériés","")&". Nombre d\'équipes = tickets au forfait du projet ÷ jours ouvrés ÷ '
            'tickets demandés par jour et par équipe (arrondi au centième) ; montant forfait = nombre d\'équipes × base mensuelle. '
            'En dehors de ces heures, le ticket est sur devis et peut entraîner une plus-value (colonne Q). Un même N° de ticket n\'est compté qu\'une fois."'
            f'&IF({doublons}+{autres}=0,""," Non comptés ce mois-ci : "&{doublons}&" doublon(s), "&{autres}&" ticket(s) d\'un autre mois.")',
            note, align=Alignment(wrap_text=True, vertical='center'), border=None)
    at.merge_cells(f'A{R_NOTE}:{der}{R_NOTE}')
    at.row_dimensions[R_NOTE].height = 30

    # ---- en-têtes de la liste des tickets
    ent = Font(name=FT, size=9, bold=True)
    r2 = R_T1 + 1
    for lettre, titre, _ in COLS + [(C_PROJET, 'Projet', 0)]:
        cellule(at, f'{lettre}{R_T1}', titre, ent, fond(GRIS) if lettre == C_PROJET else None, align=centre)
        cellule(at, f'{lettre}{r2}', None, align=centre)
        at.merge_cells(f'{lettre}{R_T1}:{lettre}{r2}')
    at[f'A{R_T1}'].comment = Comment("Les titres de cette ligne sont ceux cherchés dans l'export. "
                                     "Si une colonne de l'export porte un autre nom, écris-le ici.", 'Attachement')
    cellule(at, f'{C_F}{R_T1}', 'NB Ticket', ent, fond(VERT), align=centre)
    cellule(at, f'{C_HF}{R_T1}', None, fill=fond(VERT))
    cellule(at, f'{C_MOT}{R_T1}', None, fill=fond(VERT))
    at.merge_cells(f'{C_F}{R_T1}:{C_MOT}{R_T1}')
    cellule(at, f'{C_F}{r2}', '="FORFAIT "&TEXT(HDEB,"hh:mm")&"-"&TEXT(HFIN,"hh:mm")', ent, fond(VERT), align=centre)
    cellule(at, f'{C_HF}{r2}', 'HORS FORFAIT', ent, fond(VERT), align=centre)
    cellule(at, f'{C_MOT}{r2}', 'Motif hors forfait', ent, fond(VERT), align=centre)
    cellule(at, f'{C_PV}{R_T1}', 'Plus-value sur devis (€)', ent, fond(VERT), align=centre)
    cellule(at, f'{C_PV}{r2}', None, fill=fond(VERT))
    at.merge_cells(f'{C_PV}{R_T1}:{C_PV}{r2}')
    at.row_dimensions[R_T1].height = 24
    at.row_dimensions[r2].height = 30

    # ---- lignes de tickets
    def lecture(lettre, r):
        rc = f"'Réglages'!$C${rang_col[lettre]}"
        src = f'INDEX(Export!$A:$DZ,LIGNE_TITRES+ROW()-{P0 - 1},{rc})'
        if lettre == 'A':
            return f'=IF({rc}=0,"",IF({src}="","",{src}))'
        return f'=IF(OR($A{r}="",{rc}=0),"",IF({src}="","",{src}))'

    for r in range(P0, P1 + 1):
        for lettre, _, _ in COLS:
            date = lettre in (C_DEB, 'I')
            cellule(at, f'{lettre}{r}', lecture(lettre, r), f9, fmt='dd/mm/yyyy hh:mm' if date else None,
                    align=Alignment(vertical='center', horizontal='center' if date else None))
        ag = f'${C_AGENCE}{r}'
        cellule(at, f'{C_PROJET}{r}',
                f'=IF($A{r}="","",IF({ag}="","{SANS_AGENCE}",IFERROR(INDEX(AG_PROJ,MATCH({cle(ag)},AG_CLE,0)),"{INCONNUE}")))',
                f9b, align=Alignment(horizontal='center', vertical='center'))
        at[f'{C_AIDE}{r}'] = f'=IF($A{r}="","",IF(ISNUMBER(${C_DEB}{r}),${C_DEB}{r},IFERROR(--${C_DEB}{r},"")))'
        at[f'{C_AIDE}{r}'].number_format = 'dd/mm/yyyy hh:mm'
        d = f'${C_AIDE}{r}'
        hm = f'ROUND(MOD({d},1)*1440,0)'
        cellule(at, f'{C_MOT}{r}',
                f'=IF($A{r}="","",IF(COUNTIF($A${P0}:$A{r},$A{r})>1,"Doublon",IF({d}="","",'
                f'IF(AND(MOIS<>"",OR({d}<MOIS,{d}>=EOMONTH(MOIS,0)+1)),"Autre mois",'
                f'IF(AND(WE="OUI",WEEKDAY({d},2)=6),"Samedi",IF(AND(WE="OUI",WEEKDAY({d},2)=7),"Dimanche",'
                f'IF(AND(WE="OUI",COUNTIF(FERIES,INT({d}))>0),"Jour férié",'
                f'IF({hm}<ROUND(HDEB*1440,0),"Avant "&TEXT(HDEB,"hh:mm")&" (début "&TEXT({hm}/1440,"hh:mm")&")",'
                f'IF({hm}>ROUND(HFIN*1440,0),"Après "&TEXT(HFIN,"hh:mm")&" (début "&TEXT({hm}/1440,"hh:mm")&")","")))))))))',
                f9, align=Alignment(vertical='center'))
        m = f'${C_MOT}{r}'
        cellule(at, f'{C_F}{r}', f'=IF($A{r}="","",IF({m}="",1,""))', f9, align=Alignment(horizontal='center', vertical='center'))
        cellule(at, f'{C_HF}{r}', f'=IF(OR($A{r}="",{m}="",{m}="Doublon",{m}="Autre mois"),"",1)', f9,
                align=Alignment(horizontal='center', vertical='center'))
        cellule(at, f'{C_PV}{r}', None, f9b, fmt=EUR)

    plage = f'A{P0}:{der}{P1}'
    at.conditional_formatting.add(f'{C_HF}{P0}:{C_MOT}{P1}', FormulaRule(
        formula=[f'${C_HF}{P0}=1'], fill=fond('FFC7CE'), font=Font(color='9C0006', bold=True)))
    at.conditional_formatting.add(f'{C_PV}{P0}:{C_PV}{P1}', FormulaRule(formula=[f'${C_HF}{P0}=1'], fill=fond(SAISIE)))
    # montant saisi sur un ticket au forfait : barré, il n'est pas compté
    at.conditional_formatting.add(f'{C_PV}{P0}:{C_PV}{P1}', FormulaRule(
        formula=[f'AND(${C_PV}{P0}<>"",${C_HF}{P0}<>1)'], font=Font(color='A6A6A6', strike=True)))
    at.conditional_formatting.add(f'{C_PROJET}{P0}:{C_PROJET}{P1}', FormulaRule(
        formula=[f'OR(${C_PROJET}{P0}="{INCONNUE}",${C_PROJET}{P0}="{SANS_AGENCE}")'], fill=fond('FFC7CE'), font=Font(color='9C0006')))
    at.conditional_formatting.add(plage, FormulaRule(
        formula=[f'OR(${C_MOT}{P0}="Doublon",${C_MOT}{P0}="Autre mois")'], font=Font(color='A6A6A6', italic=True)))
    at.auto_filter.ref = f'A{r2}:{der}{P1}'

    # impression : paysage, une page de large, titres répétés, zone limitée aux tickets présents
    at.page_setup.orientation = 'landscape'
    at.page_setup.paperSize = at.PAPERSIZE_A4
    at.page_setup.fitToWidth = 1
    at.page_setup.fitToHeight = 0
    at.sheet_properties.pageSetUpPr.fitToPage = True
    at.print_title_rows = f'{R_T1}:{r2}'
    at.page_margins.left = at.page_margins.right = 0.3
    at.page_margins.top = at.page_margins.bottom = 0.5
    at.defined_names['_xlnm.Print_Area'] = DefinedName(
        '_xlnm.Print_Area', localSheetId=0,
        attr_text=f'OFFSET(Attachement!$A$1,0,0,{P0 - 1}+SUMPRODUCT(--(Attachement!$A${P0}:$A${P1}<>"")),{DERNIERE_COL})')
    at.sheet_view.zoomScale = 90

    wb.active = 0
    wb.calculation.fullCalcOnLoad = True
    wb.save(chemin)
    return P0


if __name__ == '__main__':
    noms = sys.argv[1:]
    if not noms:
        construire('NOM DU SOUS-TRAITANT', ICI / 'Attachement_sous-traitant.xlsx')
        print('Attachement_sous-traitant.xlsx')
    for nm in noms:
        f = ICI / f'Attachement {nm}.xlsx'
        construire(nm, f)
        print(f.name)
