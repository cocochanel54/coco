# Construit le fichier Excel d'attachement à remplir chaque mois, un par sous-traitant,
# dans le style du suivi maintenance (titre rouge, colonnes du suivi, NB Ticket en vert, case jaune, TOTAL HT).
#
#   python3 build_modele.py                      -> Attachement_sous-traitant.xlsx (modèle, nom à remplir)
#   python3 build_modele.py CET "LOR" "HK"       -> un fichier par sous-traitant : Attachement CET.xlsx, ...
#
# Chaque mois : coller l'export des interventions du sous-traitant dans l'onglet « Export » (en A1, avec
# la ligne des titres). L'onglet « Attachement » se remplit et calcule tout seul, avec un forfait d'équipes
# par projet (le projet vient de la colonne Agence, voir PROJETS) :
#   nombre d'équipes du projet = (nombre de tickets + bonif tickets) du projet ÷ jours ouvrés du mois ÷ 4,
#   montant forfait du projet = équipes × 12 000 € ;
#   le forfait couvre les tickets classiques et les travaux de 8h à 18h (du lundi au vendredi hors jours fériés) ;
#   en dehors, le ticket est sur devis et peut entraîner une plus-value, saisie à la main ;
#   un ticket de 2h30 ou plus (dans les heures ouvrables) (Soldée - Début) est « À vérifier » : après vérification, le nombre
#   de tickets bonifiés se saisit à la main et s'ajoute au nombre de tickets du projet ;
#   une malfaçon (N° qui commence par MAL, ou type qui contient MALFA) n'est pas un ticket : elle compte en IRRIS, 50 € l'unité.
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

# Projets : code, libellé, base mensuelle par équipe, tickets demandés par jour et par équipe, prix par prod (rempli = projet
# payé à l'unité au lieu du forfait d'équipes), mots qui reconnaissent
# l'agence de l'export (l'agence doit contenir le mot ; espaces, parenthèses, tirets et majuscules ne comptent pas)
PROJETS = [
    ('P2200154', 'CIRCET NANCY RES', 12000, 4, None, ['NANCY']),
    ('P2100474', 'CIRCET 54 RES', 12000, 4, None, ['CIRCET 54']),
    ('P1800267', 'CIRCET GLAUCOS RES + CIRCET RESOPTIC RES', 12000, 4, None, ['GLAUCOS', 'RESOPTIC (RES)']),
    ('P2200155', 'CIRCET RESOPTIC PROD-RACCO (à la prod)', None, None, 500, ['RESOPTIC (PROD-RACCO)']),
]
LIGNES_AGENCES = 16          # place dans le tableau des mots qui reconnaissent les agences
INCONNUE, SANS_AGENCE = 'Agence inconnue', 'Sans agence'
# autres titres acceptés dans l'export si le titre exact n'y est pas (* = n'importe quels caractères)
AUTRES_TITRES = {'A': ('N°*', 'Num*'), 'B': ('*descri*', ''), 'C': ('*techn*', '*quipe*'), 'D': ('*agence*', ''),
                 'F': ('*client*', ''), 'G': ('*site*', ''), 'H': ('*d*but*', ''), 'I': ('*sold*', '*cl*tur*'),
                 'J': ('*type*', ''), 'K': ('*factur*', ''), 'L': ('*mat*riel*', ''), 'M': ('*cerfa*', '*constat*')}

FT = 'Calibri'
fin = Side(style='thin', color='A6A6A6')
fort = Side(style='medium', color='000000')
BORD = Border(top=fin, bottom=fin, left=fin, right=fin)
fond = lambda c: PatternFill('solid', fgColor=c)
GRIS, FONCE, VERT, JAUNE, SAISIE, IRIS_FOND = 'BFBFBF', '767171', 'C6E0B4', 'FFFF00', 'FFF2CC', 'DDEBF7'
centre = Alignment(horizontal='center', vertical='center', wrap_text=True)
f9, f9b = Font(name=FT, size=9), Font(name=FT, size=9, bold=True)
note = Font(name=FT, size=9, italic=True, color='595959')

# Colonnes de la liste des tickets. Celles avec un titre sont lues dans l'export (par leur titre).
COLS = [('A', 'N°', 20), ('B', 'Description', 42), ('C', 'Nom technicien', 18), ('D', 'Agence', 18),
        ('F', 'Client', 16), ('G', 'SITE', 18), ('H', 'Début', 15), ('I', 'Soldée', 15), ('J', 'Type inter.', 20),
        ('K', 'Commentaire facturation', 22), ('L', "Du matériel à t'il été utlisé?", 18), ('M', 'Cerfa / Plainte/ Constat', 18),
        ('W', 'Commentaire', 24)]
C_AGENCE, C_PROJET, C_DEB, C_CLO, C_TYPE = 'D', 'E', 'H', 'I', 'J'
C_SANS_PROD = ('L', 'M')     # matériel et CERFA : jamais sur un projet payé à la prod
# NB Ticket : forfait, bonif ticket (saisie après vérification), durée, bonif proposée ; puis hors forfait sur devis
# NOMBRE DE TICKETS ; bonif : demandée par le sous-traitant, durée, accordée par le donneur d'ordre, commentaire ;
# IRRIS ; hors forfait sur devis
C_F, C_NBB, C_DUR, C_ACC, C_VCOM = 'N', 'O', 'P', 'Q', 'R'
C_IRIS, C_HF, C_MOT, C_PV = 'S', 'T', 'U', 'V'
C_AIDE, C_AIDE2, C_RET, C_STAT, C_PROD = 'X', 'Y', 'Z', 'AA', 'AB'   # cachées : début et fin en nombres, bonif retenue,
                                                                  # statut du ticket, ticket d'un projet payé à la prod
DERNIERE_COL = 23            # W
A_VERIFIER, VERIFIE = 'À vérifier', 'Vérifié'


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
    cellule(rg, 'A2', "Chaque mois : 1) efface l'onglet Export (Ctrl+A puis Suppr) ; 2) colle en A1 de l'onglet Export (jamais dans "
            "l'onglet Attachement) l'export des "
            "interventions du sous-traitant, avec sa ligne de titres ; 3) l'onglet Attachement se remplit et calcule tout seul, "
            "projet par projet ; 4) envoie le fichier au sous-traitant : il indique la bonif demandée colonne O pour ses tickets de 2h30 "
            "ou plus ; 5) à son retour, indique colonne Q le nombre de tickets que tu lui accordes et, si besoin, un commentaire "
            "colonne R ; seule la bonif accordée compte ; pour un ticket hors forfait sur devis, la plus-value se saisit colonne V ; "
            "6) le projet est trouvé tout seul à partir de l'agence ; si une agence apparaît « Agence inconnue », ajoute un mot "
            "qui la reconnaît dans le tableau « Agences de l'export » ci-dessous ; "
            "7) enregistre sous « Attachement <sous-traitant> <mois>.xlsx » et envoie-le. Pars chaque mois du fichier vide : "
            "les bonifs et plus-values saisies restent sinon sur les lignes.",
            Font(name=FT, size=10), align=Alignment(wrap_text=True, vertical='top'), border=None)
    rg.merge_cells('A2:D4')
    for r in (2, 3, 4):
        rg.row_dimensions[r].height = 24

    params = [
        (6, 'ST', 'Sous-traitant', sous_traitant, None, "Nom affiché dans le titre de l'attachement."),
        (7, 'MOIS', "Mois de l'attachement",
         f'=IF(COUNT(Attachement!{rng(C_AIDE2)})>0,DATE(YEAR(MEDIAN(Attachement!{rng(C_AIDE2)})),MONTH(MEDIAN(Attachement!{rng(C_AIDE2)})),1),'
         f'IF(COUNT(Attachement!{rng(C_AIDE)})>0,DATE(YEAR(MEDIAN(Attachement!{rng(C_AIDE)})),MONTH(MEDIAN(Attachement!{rng(C_AIDE)})),1),""))',
         'mmmm yyyy', "Lu dans l'export, d'après la date « Soldée » (un ticket compte dans le mois où il est soldé). "
         "Tu peux taper une date, ex. 01/10/2026."),
        (8, 'JOURS', 'Jours ouvrés du mois', '=IF(MOIS="",0,NETWORKDAYS(MOIS,EOMONTH(MOIS,0),FERIES))', '0',
         'Lundi au vendredi hors jours fériés. Tu peux taper un autre nombre.'),
        (9, 'HDEB', 'Début des heures ouvrables', dt.time(8, 0), 'hh:mm', ''),
        (10, 'HFIN', 'Fin des heures ouvrables', dt.time(18, 0), 'hh:mm', '18:00 pile reste au forfait.'),
        (11, 'WE', 'Week-ends et jours fériés hors forfait', 'OUI', None, 'OUI ou NON.'),
        (12, 'BONIF', 'Bonif de ticket à partir de', dt.time(2, 30), '[h]:mm',
         'Repère : à partir de cette durée (Soldée - Début), le sous-traitant peut demander une bonif. Tu décides du nombre accordé.'),
        (13, 'PRIX_IRIS', "Prix d'une malfaçon (IRRIS)", 50, EUR, "Les malfaçons ne comptent pas comme tickets : elles sont payées à l'unité."),
        (14, 'MALF_NUM', 'Malfaçon si le N° commence par', 'MAL', None, 'Ex. MAL-20260722T0465#1.'),
        (15, 'MALF_TYPE', 'ou si le Type inter. contient', 'MALFA', None, 'Ex. MAINTENANCE - REPRISE MALFACON IRRIS.'),
        (16, 'LIGNE_TITRES', "Ligne des titres dans l'onglet Export", 1, '0', "Si l'export a des lignes au-dessus de ses titres."),
        (17, 'MOIS_TXT', 'Mois en lettres (calculé)',
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
    r = 19
    cellule(rg, f'A{r}', "Projets : un forfait d'équipes par projet", Font(name=FT, size=12, bold=True), border=None)
    r += 1
    for col, t in zip('ABCDE', ['Code projet', 'Agences', 'Base mensuelle par équipe', 'Tickets demandés par jour et par équipe',
                                'Prix par prod (rempli = payé à l\'unité)']):
        cellule(rg, f'{col}{r}', t, Font(name=FT, size=10, bold=True), fond(VERT), align=centre)
    rg.row_dimensions[r].height = 30
    lignes_projet = []
    for code, lib, base, tpj, prix, _ in PROJETS:
        r += 1
        lignes_projet.append(r)
        cellule(rg, f'A{r}', code, Font(name=FT, size=11, bold=True), fond(SAISIE))
        cellule(rg, f'B{r}', lib, Font(name=FT, size=10), fond(SAISIE))
        cellule(rg, f'C{r}', base, Font(name=FT, size=10), fond(SAISIE), EUR)
        cellule(rg, f'D{r}', tpj, Font(name=FT, size=10), fond(SAISIE), 'General')
        cellule(rg, f'E{r}', prix, Font(name=FT, size=10, bold=True), fond(SAISIE), EUR)
    nom_global(wb, 'PROJ_CODE', f"'Réglages'!$A${lignes_projet[0]}:$A${lignes_projet[-1]}")
    nom_global(wb, 'PROJ_PRIX', f"'Réglages'!$E${lignes_projet[0]}:$E${lignes_projet[-1]}")
    dv_proj = DataValidation(type='list', formula1=f'$A${lignes_projet[0]}:$A${lignes_projet[-1]}', allow_blank=True)
    rg.add_data_validation(dv_proj)

    # agences -> projet
    r += 2
    cellule(rg, f'A{r}', "Agences de l'export : le projet est trouvé automatiquement", Font(name=FT, size=12, bold=True), border=None)
    r += 1
    for col, t in zip('ABC', ["Mot contenu dans l'agence", 'Projet', 'Mot comparé (calculé)']):
        cellule(rg, f'{col}{r}', t, Font(name=FT, size=10, bold=True), fond(VERT), align=centre)
    cellule(rg, f'D{r}', "Une agence qui contient le mot va dans le projet : « CIRCET NANCY (RES) » contient NANCY → P2200154. "
            "Espaces, parenthèses, tirets et majuscules ne comptent pas. Si deux mots conviennent, la ligne la plus basse l'emporte.",
            note, align=Alignment(wrap_text=True, vertical='center'), border=None)
    rg.row_dimensions[r].height = 80
    agences = [(a, code) for code, _, _, _, _, ags in PROJETS for a in ags]
    a0 = r + 1
    for i in range(LIGNES_AGENCES):
        r += 1
        a, code = agences[i] if i < len(agences) else (None, None)
        cellule(rg, f'A{r}', a, Font(name=FT, size=10), fond(SAISIE))
        cellule(rg, f'B{r}', code, Font(name=FT, size=10, bold=True), fond(SAISIE))
        cellule(rg, f'C{r}', f'=IF(OR(A{r}="",B{r}=""),"§§§",{cle(f"A{r}")})', Font(name=FT, size=9, color='808080'))
        dv_proj.add(f'B{r}')
    nom_global(wb, 'AG_PROJ', f"'Réglages'!$B${a0}:$B${r}")
    nom_global(wb, 'AG_CLE', f"'Réglages'!$C${a0}:$C${r}")

    # colonnes lues dans l'export : titre exact, sinon deux autres titres acceptés (* = n'importe quels caractères)
    r += 2
    cellule(rg, f'A{r}', "Colonnes lues dans l'export", Font(name=FT, size=12, bold=True), border=None)
    r += 1
    for col, t in zip('ABCDEF', ["Colonne de l'attachement", "Titre cherché dans l'export", 'Autre titre accepté',
                                 'Autre titre accepté', 'Trouvée en colonne n°', "Colonne de l'export utilisée"]):
        cellule(rg, f'{col}{r}', t, Font(name=FT, size=10, bold=True), fond(VERT), align=centre)
    rg.row_dimensions[r].height = 30
    rang_col = {}
    entetes = 'INDEX(Export!$1:$30,LIGNE_TITRES,0)'
    for lettre, titre, _ in COLS:
        r += 1
        rang_col[lettre] = r
        a1, a2 = AUTRES_TITRES.get(lettre, ('', ''))
        cellule(rg, f'A{r}', f'Colonne {lettre}', Font(name=FT, size=10))
        cellule(rg, f'B{r}', f'=Attachement!{lettre}{R_T1}', Font(name=FT, size=10))
        cellule(rg, f'C{r}', a1 or None, Font(name=FT, size=10), fond(SAISIE))
        cellule(rg, f'D{r}', a2 or None, Font(name=FT, size=10), fond(SAISIE))
        essai = lambda x: f'MATCH(IF({x}="","§§§",{x}),{entetes},0)'
        cellule(rg, f'E{r}', f'=IFERROR(MATCH(B{r},{entetes},0),IFERROR({essai(f"C{r}")},IFERROR({essai(f"D{r}")},0)))',
                Font(name=FT, size=10), align=Alignment(horizontal='center'))
        cellule(rg, f'F{r}', f'=IF(E{r}>0,"OK : « "&INDEX(Export!$1:$30,LIGNE_TITRES,E{r})&" »","Absente de l\'export")',
                Font(name=FT, size=10))
    at_ = rang_col
    rg.conditional_formatting.add(f'F{rang_col["A"]}:F{r}', FormulaRule(formula=[f'$E{rang_col["A"]}=0'], font=Font(color='C00000', bold=True)))
    r += 1
    cellule(rg, f'A{r}', f"Si une colonne est absente, écris dans l'onglet Attachement (ligne {R_T1}) le titre exact de la colonne de l'export, "
            "ou un autre titre accepté ici (* remplace n'importe quels caractères : *sold* trouve « Date soldée »).",
            note, border=None)
    rg.column_dimensions['E'].width = 14
    rg.column_dimensions['F'].width = 40

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
    larg.update({C_PROJET: 14, C_F: 13, C_NBB: 14, C_DUR: 16, C_ACC: 14, C_VCOM: 28, C_IRIS: 11, C_HF: 13, C_MOT: 24, C_PV: 15,
                 C_AIDE: 14, C_AIDE2: 14, C_RET: 8, C_STAT: 14, C_PROD: 6})
    for l, w in larg.items():
        at.column_dimensions[l].width = w
    at.column_dimensions[C_AIDE].hidden = True
    at.column_dimensions[C_AIDE2].hidden = True
    at.column_dimensions[C_RET].hidden = True
    at.column_dimensions[C_STAT].hidden = True
    at.column_dimensions[C_PROD].hidden = True
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
    at.row_dimensions[3].height = 18
    # alerte si l'export a été collé dans cet onglet (les formules des tickets ont été effacées)
    # sinon, alerte si les dates manquent (pas de mois, pas de jours ouvrés, pas de hors forfait ni de bonif)
    msg_deb = ("ATTENTION : la colonne « Début » n'est pas trouvée dans l'export : mois, jours ouvrés, hors forfait et bonif "
               f"ne peuvent pas être calculés. Écris en {C_DEB}{R_T1} le titre exact de cette colonne dans l'export (voir l'onglet Réglages).")
    msg_fin = ("ATTENTION : la colonne « Soldée » n'est pas trouvée dans l'export : la durée et la bonif ne peuvent pas être calculées. "
               f"Écris en {C_CLO}{R_T1} le titre exact de cette colonne dans l'export (voir l'onglet Réglages).")
    msg_jours = "ATTENTION : jours ouvrés = 0. Tape le mois dans l'onglet Réglages (Mois de l'attachement, ex. 01/10/2026)."
    q = lambda t: '"' + t.replace('"', '""') + '"'
    verif_dates = (f"IF('Réglages'!$E${rang_col[C_DEB]}=0,{q(msg_deb)},IF('Réglages'!$E${rang_col[C_CLO]}=0,{q(msg_fin)},"
                   f'IF(AND(JOURS=0,SUM({C_F}{P0}:{C_F}{P1})>0),{q(msg_jours)},"")))')
    cellule(at, 'A3', f'=IF(AND(_xlfn.ISFORMULA(A{P0}),_xlfn.ISFORMULA({C_PROJET}{P0}),_xlfn.ISFORMULA({C_F}{P0})),{verif_dates},"ATTENTION : l\'export a été collé dans l\'onglet Attachement et a effacé ses formules. Reprends le fichier vide et colle l\'export dans l\'onglet « Export », cellule A1.")',
            Font(name=FT, size=11, bold=True, color='C00000'), align=Alignment(horizontal='center', vertical='center'), border=None)
    at.merge_cells(f'A3:{der}3')

    # ---- récapitulatif par projet (colonnes A à P)
    titres = ['Projet', 'Agences', 'Nombre de tickets', 'Bonif accordées', 'Tickets retenus', 'Jours ouvrés',
              'Tickets demandés / jour / équipe', 'Tickets par jour', "NOMBRE D'ÉQUIPES", 'Base mensuelle / équipe (ou prix par prod)',
              'Montant forfait (ou prods)',
              'IRRIS (malfaçons)', 'Montant IRRIS', 'Tickets hors forfait (sur devis)', 'Plus-values sur devis', 'TOTAL HT']
    for i, t in enumerate(titres):
        coul = JAUNE if t == "NOMBRE D'ÉQUIPES" else IRIS_FOND if t.startswith(('IRRIS', 'Montant IRRIS')) else VERT
        cellule(at, f'{L(i + 1)}{R_ENT}', t, Font(name=FT, size=9, bold=True), fond(coul), align=centre)
    cellule(at, f'Q{R_ENT}', 'Tickets retenus = nombre de tickets + bonif accordées. Équipes = tickets retenus ÷ jours ÷ tickets demandés. '
            'TOTAL HT = montant forfait + montant IRIS + plus-values.', note, align=Alignment(wrap_text=True, vertical='center'), border=None)
    at.merge_cells(f'Q{R_ENT}:{der}{R_ENT}')
    at.row_dimensions[R_ENT].height = 36

    proj, forf, hf, mot, pv = rng(C_PROJET), rng(C_F), rng(C_HF), rng(C_MOT), rng(C_PV)
    nbb, ret, iris = rng(C_NBB), rng(C_RET), rng(C_IRIS)
    v10, v10b = Font(name=FT, size=10), Font(name=FT, size=10, bold=True)
    gros = Font(name=FT, size=13, bold=True)
    mil = Alignment(horizontal='center', vertical='center')
    for i, lp in enumerate(lignes_projet):
        r = R_P1 + i
        cellule(at, f'A{r}', f"='Réglages'!A{lp}", Font(name=FT, size=11, bold=True), align=Alignment(vertical='center'))
        cellule(at, f'B{r}', f"='Réglages'!B{lp}", v10, align=Alignment(vertical='center', wrap_text=True))
        cellule(at, f'C{r}', f'=COUNTIFS({proj},$A{r},{forf},1)', v10, fmt='0')
        cellule(at, f'D{r}', f'=SUMIFS({ret},{proj},$A{r})', v10, fmt='0')
        cellule(at, f'E{r}', f'=C{r}+D{r}', v10b, fmt='0')
        cellule(at, f'F{r}', '=JOURS', v10, fmt='0')
        unite = f"N('Réglages'!E{lp})>0"          # projet payé à la prod
        cellule(at, f'G{r}', f"=IF({unite},\"\",'Réglages'!D{lp})", v10, fmt='General')
        cellule(at, f'H{r}', f'=IF({unite},"",IF(F{r}>0,E{r}/F{r},0))', v10, fmt='0.00')
        cellule(at, f'I{r}', f'=IF({unite},"à la prod",ROUND(IF(AND(F{r}>0,G{r}>0),E{r}/F{r}/G{r},0),2))', gros, fond(JAUNE), '0.00', mil)
        cellule(at, f'J{r}', f"=IF({unite},'Réglages'!E{lp},'Réglages'!C{lp})", v10, fmt=EUR)
        cellule(at, f'K{r}', f'=IF({unite},E{r}*J{r},I{r}*J{r})', v10b, fmt=EUR)
        cellule(at, f'L{r}', f'=COUNTIFS({proj},$A{r},{iris},1)', v10, fmt='0')
        cellule(at, f'M{r}', f'=L{r}*PRIX_IRIS', v10b, fmt=EUR)
        cellule(at, f'N{r}', f'=COUNTIFS({proj},$A{r},{hf},1)', v10, fmt='0')
        cellule(at, f'O{r}', f'=SUMIFS({pv},{proj},$A{r},{hf},1)', v10, fmt=EUR)
        cellule(at, f'P{r}', f'=K{r}+M{r}+O{r}', v10b, fmt=EUR)
        at.row_dimensions[r].height = 22
    r = R_INC
    inc = lambda crit: f'COUNTIFS({proj},"{INCONNUE}",{crit},1)+COUNTIFS({proj},"{SANS_AGENCE}",{crit},1)'
    cellule(at, f'A{r}', 'Non classés', v10b, align=Alignment(vertical='center'))
    cellule(at, f'B{r}', 'Agence absente du tableau des Réglages : à classer', Font(name=FT, size=9, italic=True),
            align=Alignment(vertical='center', wrap_text=True))
    cellule(at, f'C{r}', '=' + inc(forf), v10, fmt='0')
    cellule(at, f'D{r}', f'=SUMIFS({ret},{proj},"{INCONNUE}")+SUMIFS({ret},{proj},"{SANS_AGENCE}")', v10, fmt='0')
    cellule(at, f'E{r}', f'=C{r}+D{r}', v10, fmt='0')
    cellule(at, f'L{r}', '=' + inc(iris), v10, fmt='0')
    cellule(at, f'N{r}', '=' + inc(hf), v10, fmt='0')
    for col in 'FGHIJKMOP':
        cellule(at, f'{col}{r}', None, v10)
    at.conditional_formatting.add(f'A{r}:P{r}', FormulaRule(formula=[f'$E${r}+$L${r}+$N${r}>0'], fill=fond('FFC7CE'), font=Font(color='9C0006', bold=True)))
    at.conditional_formatting.add(f'A{r}:P{r}', FormulaRule(formula=[f'$E${r}+$L${r}+$N${r}=0'], font=Font(color='A6A6A6')))
    r = R_TOT
    cellule(at, f'A{r}', 'TOTAL', Font(name=FT, size=12, bold=True, color='FFFFFF'), fond(FONCE), align=Alignment(horizontal='right', vertical='center'))
    cellule(at, f'B{r}', None, fill=fond(FONCE))
    at.merge_cells(f'A{r}:B{r}')
    som = lambda col, a=R_P1, b=R_INC: f'=SUM({col}{a}:{col}{b})'
    for col in 'CDELN':
        cellule(at, f'{col}{r}', som(col), v10b, fmt='0')
    for col in 'FGHJ':
        cellule(at, f'{col}{r}', None)
    cellule(at, f'I{r}', som('I', b=R_INC - 1), gros, fond(JAUNE), '0.00', mil)
    for col in 'KMO':
        cellule(at, f'{col}{r}', som(col, b=R_INC - 1), v10b, fmt=EUR)
    cellule(at, f'P{r}', som('P', b=R_INC - 1), gros, fond(JAUNE), EUR, mil)
    at.row_dimensions[r].height = 28
    a_traiter = f'COUNTIFS({nbb},">0",{rng(C_ACC)},"")'
    cellule(at, f'Q{r}', f'=IF({a_traiter}=0,"",{a_traiter}&" bonif(s) demandée(s) à traiter : nombre accordé colonne {C_ACC}")',
            Font(name=FT, size=10, bold=True, color='9C5700'), align=Alignment(wrap_text=True, vertical='center'), border=None)
    at.merge_cells(f'Q{r}:{der}{r}')

    doublons = f'COUNTIF({mot},"Doublon")'
    autres = f'COUNTIF({mot},"Autre mois")'
    cellule(at, f'A{R_NOTE}',
            '="Forfait par projet : tickets classiques et travaux commencés entre "&TEXT(HDEB,"hh:mm")&" et "&TEXT(HFIN,"hh:mm")'
            '&IF(WE="OUI",", du lundi au vendredi hors jours fériés","")&". Un ticket de "&TEXT(BONIF,"hh:mm")&" ou plus (Soldée - Début) '
            'compte un ticket de plus par tranche de "&TEXT(BONIF,"hh:mm")&" : le sous-traitant demande sa bonif (colonne O), seule la bonif accordée (colonne Q) compte. Nombre d\'équipes = tickets retenus du projet ÷ jours ouvrés ÷ '
            'tickets demandés par jour et par équipe (arrondi au centième) ; montant forfait = nombre d\'équipes × base mensuelle ; '
            'projet payé à la prod (prix par prod dans Réglages) : montant = nombre de prods × prix par prod, quelle que soit l\'heure, sans bonif, sans matériel ni CERFA. '
            'En dehors des heures, le ticket est sur devis et peut entraîner une plus-value (colonne V). Les malfaçons ne comptent pas comme tickets : '
            'elles sont payées en IRRIS, "&TEXT(PRIX_IRIS,"0")&" € l\'unité. Un ticket compte dans le mois où il est soldé. '
            'Un même N° de ticket n\'est compté qu\'une fois."'
            f'&IF({doublons}+{autres}=0,""," Non comptés ce mois-ci : "&{doublons}&" doublon(s), "&{autres}&" ticket(s) d\'un autre mois.")',
            note, align=Alignment(wrap_text=True, vertical='center'), border=None)
    at.merge_cells(f'A{R_NOTE}:{der}{R_NOTE}')
    at.row_dimensions[R_NOTE].height = 40

    # ---- en-têtes de la liste des tickets
    ent = Font(name=FT, size=9, bold=True)
    r2 = R_T1 + 1
    for lettre, titre, _ in COLS + [(C_PROJET, 'Projet', 0)]:
        cellule(at, f'{lettre}{R_T1}', titre, ent, fond(GRIS) if lettre == C_PROJET else None, align=centre)
        cellule(at, f'{lettre}{r2}', None, align=centre)
        at.merge_cells(f'{lettre}{R_T1}:{lettre}{r2}')
    at[f'A{R_T1}'].comment = Comment("Les titres de cette ligne sont ceux cherchés dans l'export. "
                                     "Si une colonne de l'export porte un autre nom, écris-le ici.", 'Attachement')
    cellule(at, f'{C_F}{R_T1}', 'NOMBRE DE TICKETS', ent, fond(VERT), align=centre)
    cellule(at, f'{C_F}{r2}', None, fill=fond(VERT))
    at.merge_cells(f'{C_F}{R_T1}:{C_F}{r2}')
    for debut, fin_, titre, coul in ((C_NBB, C_VCOM, 'Bonif', VERT),
                                     (C_HF, C_PV, 'Hors forfait (sur devis)', 'F8CBAD')):
        cellule(at, f'{debut}{R_T1}', titre, ent, fond(coul), align=centre)
        for k in range(openpyxl.utils.column_index_from_string(debut) + 1, openpyxl.utils.column_index_from_string(fin_) + 1):
            cellule(at, f'{L(k)}{R_T1}', None, fill=fond(coul))
        at.merge_cells(f'{debut}{R_T1}:{fin_}{R_T1}')
    cellule(at, f'{C_IRIS}{R_T1}', 'IRRIS (malfaçon)', ent, fond(IRIS_FOND), align=centre)
    cellule(at, f'{C_IRIS}{r2}', None, fill=fond(IRIS_FOND))
    at.merge_cells(f'{C_IRIS}{R_T1}:{C_IRIS}{r2}')
    sous = [(C_NBB, 'BONIF DEMANDÉE (sous-traitant)', JAUNE),
            (C_DUR, 'Durée', VERT), (C_ACC, 'BONIF ACCORDÉE', JAUNE), (C_VCOM, 'Commentaire bonif', VERT),
            (C_HF, 'HORS FORFAIT', 'F8CBAD'), (C_MOT, 'Motif', 'F8CBAD'), (C_PV, 'Plus-value (€)', 'F8CBAD')]
    for col, t, coul in sous:
        cellule(at, f'{col}{r2}', t, ent, fond(coul), align=centre)
    at[f'{C_NBB}{r2}'].comment = Comment("À remplir par le sous-traitant : nombre de tickets en plus demandés pour ce ticket "
                                         "(tickets de 2h30 ou plus).", 'Attachement')
    at[f'{C_ACC}{r2}'].comment = Comment("À remplir par toi : nombre de tickets en plus que tu accordes (0 si tu refuses). "
                                         "Seul ce nombre compte. Explique dans « Commentaire bonif » si besoin.", 'Attachement')
    at.row_dimensions[R_T1].height = 24
    at.row_dimensions[r2].height = 36

    # ---- lignes de tickets
    def lecture(lettre, r):
        rc = f"'Réglages'!$E${rang_col[lettre]}"
        src = f'INDEX(Export!$A:$DZ,LIGNE_TITRES+ROW()-{P0 - 1},{rc})'
        if lettre == 'A':
            return f'=IF({rc}=0,"",IF({src}="","",{src}))'
        if lettre in C_SANS_PROD:          # jamais de matériel ni de CERFA sur un projet payé à la prod
            return f'=IF($A{r}="","",IF(${C_PROD}{r}=1,"—",IF({rc}=0,"",IF({src}="","",{src}))))'
        return f'=IF(OR($A{r}="",{rc}=0),"",IF({src}="","",{src}))'

    en_nombre = lambda col, r: f'=IF($A{r}="","",IF(ISNUMBER(${col}{r}),${col}{r},IFERROR(--${col}{r},"")))'
    # pas de bordure fixe sur les 300 lignes : elle vient d'une mise en forme conditionnelle (lignes remplies seulement)
    cel = lambda *a, **k: cellule(*a, **{**k, 'border': None})
    for r in range(P0, P1 + 1):
        for lettre, _, _ in COLS:
            date = lettre in (C_DEB, C_CLO)
            cel(at, f'{lettre}{r}', lecture(lettre, r), f9, fmt='dd/mm/yyyy hh:mm' if date else None,
                    align=Alignment(vertical='center', horizontal='center' if date else None))
        ag = f'${C_AGENCE}{r}'
        cel(at, f'{C_PROJET}{r}',
                f'=IF($A{r}="","",IF({ag}="","{SANS_AGENCE}",IFERROR(LOOKUP(2^15,SEARCH(AG_CLE,{cle(ag)}),AG_PROJ),"{INCONNUE}")))',
                f9b, align=mil)
        at[f'{C_AIDE}{r}'] = en_nombre(C_DEB, r)
        at[f'{C_AIDE2}{r}'] = en_nombre(C_CLO, r)
        d = f'${C_AIDE}{r}'
        f_ = f'${C_AIDE2}{r}'
        hm = f'ROUND(MOD({d},1)*1440,0)'
        malf = (f'OR(AND(MALF_NUM<>"",LEFT($A{r},LEN(MALF_NUM))=MALF_NUM),'
                f'AND(MALF_TYPE<>"",ISNUMBER(SEARCH(MALF_TYPE,${C_TYPE}{r}))))')
        at[f'{C_PROD}{r}'] = f'=IF($A{r}="",0,IF(COUNTIFS(PROJ_CODE,${C_PROJET}{r},PROJ_PRIX,">0")>0,1,0))'
        at[f'{C_STAT}{r}'] = (
            f'=IF($A{r}="","",IF(COUNTIF($A${P0}:$A{r},$A{r})>1,"Doublon",'
            f'IF(AND(MOIS<>"",IF({f_}<>"",OR({f_}<MOIS,{f_}>=EOMONTH(MOIS,0)+1),AND({d}<>"",OR({d}<MOIS,{d}>=EOMONTH(MOIS,0)+1)))),"Autre mois",IF({malf},"IRIS",IF(OR({d}="",${C_PROD}{r}=1),"",'
            f'IF(AND(WE="OUI",WEEKDAY({d},2)=6),"Samedi",IF(AND(WE="OUI",WEEKDAY({d},2)=7),"Dimanche",'
            f'IF(AND(WE="OUI",COUNTIF(FERIES,INT({d}))>0),"Jour férié",'
            f'IF({hm}<ROUND(HDEB*1440,0),"Avant "&TEXT(HDEB,"hh:mm")&" (début "&TEXT({hm}/1440,"hh:mm")&")",'
            f'IF({hm}>ROUND(HFIN*1440,0),"Après "&TEXT(HFIN,"hh:mm")&" (début "&TEXT({hm}/1440,"hh:mm")&")",""))))))))))')
        st = f'${C_STAT}{r}'
        cel(at, f'{C_MOT}{r}', f'=IF(OR({st}="",{st}="IRIS"),"",{st})', f9, align=Alignment(vertical='center'))
        cel(at, f'{C_F}{r}', f'=IF($A{r}="","",IF({st}="",1,""))', f9, align=mil)
        cel(at, f'{C_IRIS}{r}', f'=IF({st}="IRIS",1,"")', f9b, align=mil)
        cel(at, f'{C_HF}{r}', f'=IF(OR($A{r}="",{st}="",{st}="Doublon",{st}="Autre mois",{st}="IRIS"),"",1)', f9, align=mil)
        fin_ = f'${C_AIDE2}{r}'
        cel(at, f'{C_DUR}{r}', f'=IF(OR({d}="",{fin_}=""),"",IF({fin_}>{d},{fin_}-{d},""))', f9, fmt='[h]:mm', align=mil)
        cel(at, f'{C_NBB}{r}', None, f9b, align=mil)
        cel(at, f'{C_ACC}{r}', None, f9b, align=mil)
        cel(at, f'{C_VCOM}{r}', None, f9, align=Alignment(vertical='center', wrap_text=True))
        at[f'{C_RET}{r}'] = f'=IF(AND(${C_F}{r}=1,${C_PROD}{r}<>1,ISNUMBER(${C_ACC}{r})),${C_ACC}{r},0)'
        cel(at, f'{C_PV}{r}', None, f9b, fmt=EUR)

    plage = f'A{P0}:{der}{P1}'
    at.conditional_formatting.add(f'{C_HF}{P0}:{C_MOT}{P1}', FormulaRule(
        formula=[f'${C_HF}{P0}=1'], fill=fond('FFC7CE'), font=Font(color='9C0006', bold=True)))
    at.conditional_formatting.add(f'{C_IRIS}{P0}:{C_IRIS}{P1}', FormulaRule(
        formula=[f'${C_IRIS}{P0}=1'], fill=fond(IRIS_FOND), font=Font(color='1F4E79', bold=True)))
    # bonif : demande (O) en jaune clair sur les tickets de 2h30 ou plus ; orange tant que rien n'est accordé ;
    # accordée (Q) verte si tout est accordé, rouge clair si moins que demandé (commentaire en orange tant qu'il est vide)
    dem, acc, com, dur = f'${C_NBB}{P0}', f'${C_ACC}{P0}', f'${C_VCOM}{P0}', f'${C_DUR}{P0}'
    eligible = f'AND(${C_F}{P0}=1,${C_PROD}{P0}<>1,{dur}<>"",IFERROR(ROUND({dur}*1440,0)>=ROUND(BONIF*1440,0),FALSE))'
    zone = lambda c: f'{c}{P0}:{c}{P1}'
    regle = lambda c, f, **k: at.conditional_formatting.add(zone(c), FormulaRule(formula=[f], **k))
    regle(C_NBB, f'AND(ISNUMBER({dem}),{acc}="")', fill=fond('FFEB9C'), font=Font(color='9C5700', bold=True))
    regle(C_NBB, f'AND(ISNUMBER({dem}),ISNUMBER({acc}))', fill=fond('F2F2F2'))
    regle(C_NBB, eligible, fill=fond(SAISIE))
    regle(C_ACC, f'AND(ISNUMBER({dem}),{acc}="")', fill=fond('FFEB9C'))
    regle(C_ACC, f'AND(ISNUMBER({acc}),ISNUMBER({dem}),{acc}<{dem})', fill=fond('FFC7CE'), font=Font(color='9C0006', bold=True))
    regle(C_ACC, f'ISNUMBER({acc})', fill=fond('C6EFCE'), font=Font(color='006100', bold=True))
    regle(C_VCOM, f'AND(ISNUMBER({acc}),ISNUMBER({dem}),{acc}<{dem},{com}="")', fill=fond('FFEB9C'))
    at.conditional_formatting.add(f'{C_PV}{P0}:{C_PV}{P1}', FormulaRule(formula=[f'${C_HF}{P0}=1'], fill=fond(SAISIE)))
    # montant saisi sur un ticket au forfait : barré, il n'est pas compté
    at.conditional_formatting.add(f'{C_PV}{P0}:{C_PV}{P1}', FormulaRule(
        formula=[f'AND(${C_PV}{P0}<>"",${C_HF}{P0}<>1)'], font=Font(color='A6A6A6', strike=True)))
    at.conditional_formatting.add(f'{C_PROJET}{P0}:{C_PROJET}{P1}', FormulaRule(
        formula=[f'OR(${C_PROJET}{P0}="{INCONNUE}",${C_PROJET}{P0}="{SANS_AGENCE}")'], fill=fond('FFC7CE'), font=Font(color='9C0006')))
    at.conditional_formatting.add(plage, FormulaRule(
        formula=[f'OR(${C_MOT}{P0}="Doublon",${C_MOT}{P0}="Autre mois")'], font=Font(color='A6A6A6', italic=True)))
    # en dernier : les règles de couleur ci-dessus passent avant (LibreOffice n'applique que la première règle vraie)
    at.conditional_formatting.add(plage, FormulaRule(formula=[f'$A{P0}<>""'], border=BORD))
    dv_bonif = DataValidation(type='decimal', operator='greaterThanOrEqual', formula1='0', allow_blank=True,
                              errorTitle='Bonif demandée', error='Saisis un nombre de tickets en plus : 0, 1, 2…')
    at.add_data_validation(dv_bonif)
    dv_bonif.add(f'{C_NBB}{P0}:{C_NBB}{P1}')
    dv_acc = DataValidation(type='decimal', operator='greaterThanOrEqual', formula1='0', allow_blank=True,
                            errorTitle='Bonif accordée', error='Saisis le nombre de tickets accordés : 0, 1, 2…')
    at.add_data_validation(dv_acc)
    dv_acc.add(f'{C_ACC}{P0}:{C_ACC}{P1}')
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
