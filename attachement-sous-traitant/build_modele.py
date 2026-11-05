# Construit le fichier Excel d'attachement à remplir chaque mois, un par sous-traitant,
# dans le style du suivi maintenance (titre rouge, colonnes du suivi, NB Ticket en vert, case jaune, TOTAL HT).
#
#   python3 build_modele.py                      -> Attachement_sous-traitant.xlsx (modèle, nom à remplir)
#   python3 build_modele.py CET "LOR" "HK"       -> un fichier par sous-traitant : Attachement CET.xlsx, ...
#
# Chaque mois : coller l'export des interventions du sous-traitant dans l'onglet « Export » (en A1, avec
# la ligne des titres). L'onglet « Attachement » se remplit et calcule tout seul :
#   nombre d'équipes = tickets au forfait ÷ jours ouvrés du mois ÷ 4, montant = équipes × 12 000 €,
#   le forfait couvre les tickets classiques et les travaux de 8h à 18h (du lundi au vendredi hors jours fériés) ;
#   en dehors, le ticket est sur devis et peut entraîner une plus-value, saisie à la main colonne P.
import sys, datetime as dt, pathlib
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.comments import Comment

ICI = pathlib.Path(__file__).parent
PREMIERE, DERNIERE = 18, 317            # lignes des tickets dans l'onglet Attachement (300 tickets)
EUR = '#,##0.00\\ "€"'

FT = 'Calibri'
fin = Side(style='thin', color='A6A6A6')
fort = Side(style='medium', color='000000')
BORD = Border(top=fin, bottom=fin, left=fin, right=fin)
fond = lambda c: PatternFill('solid', fgColor=c)
GRIS, FONCE, VERT, JAUNE, SAISIE = 'BFBFBF', '767171', 'C6E0B4', 'FFFF00', 'FFF2CC'
centre = Alignment(horizontal='center', vertical='center', wrap_text=True)

# colonnes de l'attachement : (lettre, titre cherché dans l'export, largeur)
COLS = [('A', 'N°', 20), ('B', 'Description', 42), ('C', 'EQUIPE', 16), ('D', 'Agence', 16), ('E', 'Client', 18),
        ('F', 'SITE', 18), ('G', 'Date de début', 15), ('H', 'Clôturée', 15), ('I', 'Type inter.', 20),
        ('J', 'Commentaire facturation', 22), ('K', 'Materiel utilisé', 16), ('L', 'CERFA/ASSURANCE', 16),
        ('Q', 'Commentaire', 24)]


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


def nom_global(wb, nom, ref):
    wb.defined_names[nom] = DefinedName(nom, attr_text=ref)


def construire(sous_traitant, chemin):
    wb = openpyxl.Workbook()
    at = wb.active
    at.title = 'Attachement'
    ex = wb.create_sheet('Export')
    rg = wb.create_sheet('Réglages')
    at.sheet_properties.tabColor = '70AD47'
    ex.sheet_properties.tabColor = '5B9BD5'

    # ------------------------------------------------------------------ Réglages
    rg.column_dimensions['A'].width = 46
    rg.column_dimensions['B'].width = 26
    rg.column_dimensions['C'].width = 22
    rg.column_dimensions['D'].width = 60
    rg['A1'] = "Réglages de l'attachement"
    rg['A1'].font = Font(name=FT, size=14, bold=True)
    rg['A2'] = ("Chaque mois : 1) efface l'onglet Export (Ctrl+A puis Suppr) ; 2) colle en A1 de l'onglet Export l'export des "
                "interventions du sous-traitant, avec sa ligne de titres ; 3) l'onglet Attachement se remplit et calcule tout seul ; "
                "4) pour un ticket hors forfait sur devis, saisis le montant de la plus-value colonne P de l'Attachement ; "
                "5) enregistre sous « Attachement <sous-traitant> <mois>.xlsx » et envoie-le.")
    rg['A2'].alignment = Alignment(wrap_text=True, vertical='top')
    rg['A2'].font = Font(name=FT, size=10)
    rg.merge_cells('A2:D4')
    for r in (2, 3, 4):
        rg.row_dimensions[r].height = 18

    params = [
        # ligne, nom, libellé, valeur, format, remarque
        (6, 'ST', 'Sous-traitant', sous_traitant, None, 'Nom affiché dans le titre de l\'attachement.'),
        (7, 'MOIS', "Mois de l'attachement",
         f'=IF(COUNT(Attachement!$R${PREMIERE}:$R${DERNIERE})=0,"",DATE(YEAR(MEDIAN(Attachement!$R${PREMIERE}:$R${DERNIERE})),'
         f'MONTH(MEDIAN(Attachement!$R${PREMIERE}:$R${DERNIERE})),1))', 'mmmm yyyy',
         "Lu dans l'export (dates de début). Tu peux taper une date, ex. 01/09/2026."),
        (8, 'JOURS', 'Jours ouvrés du mois', '=IF(MOIS="",0,NETWORKDAYS(MOIS,EOMONTH(MOIS,0),FERIES))', '0',
         'Lundi au vendredi hors jours fériés (liste plus bas). Tu peux taper un autre nombre.'),
        (9, 'TPJ', 'Tickets demandés par jour et par équipe', 4, '0.##', ''),
        (10, 'BASE', 'Base mensuelle par équipe', 12000, EUR, ''),
        (11, 'HDEB', 'Début des heures ouvrables', dt.time(8, 0), 'hh:mm', ''),
        (12, 'HFIN', 'Fin des heures ouvrables', dt.time(18, 0), 'hh:mm', 'Un ticket commencé à 18:00 pile reste au forfait.'),
        (13, 'WE', 'Week-ends et jours fériés hors forfait', 'OUI', None, 'OUI ou NON.'),
        (14, 'LIGNE_TITRES', "Ligne des titres dans l'onglet Export", 1, '0', 'Si l\'export a des lignes au-dessus de ses titres.'),
        (15, 'MOIS_TXT', 'Mois en lettres (calculé)',
         '=IF(MOIS="","",CHOOSE(MONTH(MOIS),"JANVIER","FÉVRIER","MARS","AVRIL","MAI","JUIN","JUILLET","AOÛT","SEPTEMBRE",'
         '"OCTOBRE","NOVEMBRE","DÉCEMBRE")&" "&YEAR(MOIS))', None, ''),
    ]
    for r, nom, lib, val, fmt, rem in params:
        rg.cell(r, 1, lib).font = Font(name=FT, size=10, bold=True)
        rg.cell(r, 1).fill = fond(GRIS)
        rg.cell(r, 1).border = BORD
        c = rg.cell(r, 2, val)
        c.font = Font(name=FT, size=11, bold=True)
        c.fill = fond(SAISIE if nom not in ('MOIS_TXT',) else 'FFFFFF')
        c.border = BORD
        c.alignment = Alignment(horizontal='right')
        if fmt:
            c.number_format = fmt
        rg.cell(r, 3, rem).font = Font(name=FT, size=9, italic=True, color='595959')
        rg.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
        nom_global(wb, nom, f"'Réglages'!$B${r}")
    dv = DataValidation(type='list', formula1='"OUI,NON"', allow_blank=False)
    rg.add_data_validation(dv)
    dv.add('B13')

    # colonnes de l'export
    rg['A18'] = "Colonnes lues dans l'export"
    rg['A18'].font = Font(name=FT, size=12, bold=True)
    for c, t in zip('ABCD', ["Colonne de l'attachement", "Titre cherché dans l'export", "Trouvée en colonne n°", '']):
        cell = rg[f'{c}19']
        cell.value = t
        cell.font = Font(name=FT, size=10, bold=True)
        cell.fill = fond(VERT)
        cell.border = BORD
    rang_col = {}
    for i, (lettre, titre, _) in enumerate(COLS):
        r = 20 + i
        rang_col[lettre] = r
        rg.cell(r, 1, f'Colonne {lettre}').border = BORD
        rg.cell(r, 2, f'=Attachement!{lettre}16').border = BORD
        rg.cell(r, 3, f'=IFERROR(MATCH(B{r},INDEX(Export!$1:$30,LIGNE_TITRES,0),0),0)').border = BORD
        rg.cell(r, 4, f'=IF(C{r}>0,"OK","Absente de l\'export : la colonne reste vide (tu peux la remplir à la main)")')
        for k in (1, 2, 3, 4):
            rg.cell(r, k).font = Font(name=FT, size=10)
    fin_cols = 20 + len(COLS)
    rg.cell(fin_cols, 1, "Si une colonne est absente, écris dans l'onglet Attachement (ligne 16) le titre exact de la colonne de l'export.").font = Font(name=FT, size=9, italic=True, color='595959')

    # jours fériés
    r0 = fin_cols + 2
    rg.cell(r0, 1, 'Jours fériés').font = Font(name=FT, size=12, bold=True)
    liste = feries(range(2026, 2033))
    for i, (d, n) in enumerate(liste):
        rg.cell(r0 + 1 + i, 1, d).number_format = 'dd/mm/yyyy'
        rg.cell(r0 + 1 + i, 2, n)
    derniere_ferie = r0 + 1 + len(liste) + 20          # place pour en ajouter
    nom_global(wb, 'FERIES', f"'Réglages'!$A${r0 + 1}:$A${derniere_ferie}")
    rg.cell(r0, 2, '(ajoute des dates à la suite si besoin)').font = Font(name=FT, size=9, italic=True, color='595959')

    # ------------------------------------------------------------------ Attachement
    larg = {l: w for l, _, w in COLS}
    larg.update({'M': 13, 'N': 13, 'O': 24, 'P': 15, 'R': 14})
    for l, w in larg.items():
        at.column_dimensions[l].width = w
    at.column_dimensions['R'].hidden = True

    at['A1'] = 'SUIVI MAINTENANCE RESEAU 54\nLOSANGE'
    at['A1'].font = Font(name=FT, size=12, bold=True, color='FF0000')
    at['A1'].alignment = centre
    at.merge_cells('A1:Q1')
    at.row_dimensions[1].height = 36
    at['A2'] = '="Prestation "&ST&" pour le Contrat maintenance "&MOIS_TXT'
    at['A2'].font = Font(name=FT, size=11, bold=True)
    at['A2'].alignment = centre
    at.merge_cells('A2:Q2')
    at.row_dimensions[2].height = 30
    for c in range(1, 18):
        at.cell(1, c).border = Border(bottom=fort)
        at.cell(2, c).border = Border(bottom=fort)
    at.row_dimensions[3].height = 8

    # bloc de calcul (lignes 4 à 14) : libellés gris à droite, valeurs en M:N
    P_, D_ = PREMIERE, DERNIERE
    bloc = [
        (4, '="Tickets au forfait ("&TEXT(HDEB,"hh:mm")&" - "&TEXT(HFIN,"hh:mm")&IF(WE="OUI",", du lundi au vendredi hors jours fériés","")&")"',
         f'=SUM(M{P_}:M{D_})', '0', 'n'),
        (5, 'Jours ouvrés du mois', '=JOURS', '0', 'n'),
        (6, 'Tickets demandés par jour et par équipe', '=TPJ', '0.##', 'n'),
        (7, 'Tickets par jour (tickets au forfait ÷ jours)', '=IF(M5>0,M4/M5,0)', '0.00', 'n'),
        (8, "NOMBRE D'ÉQUIPES À PAYER  =  tickets au forfait ÷ jours ÷ tickets demandés par jour", '=ROUND(IF(AND(M5>0,M6>0),M4/M5/M6,0),2)', '0.00', 'jaune'),
        (9, 'Base mensuelle par équipe', '=BASE', EUR, 'n'),
        (10, "Montant forfait (nombre d'équipes × base mensuelle)", '=M8*M9', EUR, 'gras'),
        (11, '="Tickets hors forfait, sur devis (avant "&TEXT(HDEB,"hh:mm")&", après "&TEXT(HFIN,"hh:mm")&IF(WE="OUI",", week-ends et jours fériés","")&")"',
         f'=SUM(N{P_}:N{D_})', '0', 'n'),
        (12, 'Plus-values sur devis des tickets hors forfait (colonne P)', f'=SUMIF(N{P_}:N{D_},1,P{P_}:P{D_})', EUR, 'gras'),
        (13, 'TOTAL HT (forfait + plus-values)', '=M10+M12', EUR, 'total'),
    ]
    for r, lib, f, fmt, sty in bloc:
        lab = at.cell(r, 1, lib)
        at.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12)
        val = at.cell(r, 13, f)
        at.merge_cells(start_row=r, start_column=13, end_row=r, end_column=14)
        val.number_format = fmt
        lab_fill, lab_font = fond(GRIS), Font(name=FT, size=10, bold=True)
        val_fill, val_font, h = None, Font(name=FT, size=10), 16
        if sty == 'gras':
            val_font = Font(name=FT, size=10, bold=True)
        if sty == 'jaune':
            lab_fill, lab_font = fond(JAUNE), Font(name=FT, size=11, bold=True)
            val_fill, val_font, h = fond(JAUNE), Font(name=FT, size=16, bold=True), 30
        if sty == 'total':
            lab_fill, lab_font = fond(FONCE), Font(name=FT, size=12, bold=True, color='FFFFFF')
            val_fill, val_font, h = fond(JAUNE), Font(name=FT, size=13, bold=True), 26
        for c in range(1, 13):
            at.cell(r, c).fill = lab_fill
            at.cell(r, c).border = BORD
        lab.font = lab_font
        lab.alignment = Alignment(horizontal='right', vertical='center')
        for c in (13, 14):
            at.cell(r, c).border = BORD
            if val_fill:
                at.cell(r, c).fill = val_fill
        val.font = val_font
        val.alignment = Alignment(horizontal='center' if sty in ('jaune', 'total') else 'right', vertical='center')
        at.row_dimensions[r].height = h
    at['O8'] = '=IF(COUNTIF(O18:O317,"Doublon")+COUNTIF(O18:O317,"Autre mois")=0,"",COUNTIF(O18:O317,"Doublon")&" doublon(s) et "&COUNTIF(O18:O317,"Autre mois")&" ticket(s) d\'un autre mois non comptés")'
    at['O8'].font = Font(name=FT, size=9, italic=True, color='9C0006')
    at['O8'].alignment = Alignment(wrap_text=True, vertical='center')
    at.merge_cells('O8:Q8')

    at['A15'] = ('="Le forfait couvre les tickets classiques et les travaux commencés entre "&TEXT(HDEB,"hh:mm")&" et "&TEXT(HFIN,"hh:mm")'
                 '&IF(WE="OUI",", du lundi au vendredi hors jours fériés","")&". Nombre d\'équipes = tickets au forfait ÷ jours ouvrés ÷ "&TPJ'
                 '&" tickets demandés par jour et par équipe (arrondi au centième) ; montant forfait = nombre d\'équipes × base mensuelle. '
                 'En dehors de ces heures, le ticket est sur devis et peut entraîner une plus-value (colonne P). Un même N° de ticket n\'est compté qu\'une fois."')
    at['A15'].font = Font(name=FT, size=9, italic=True, color='595959')
    at['A15'].alignment = Alignment(wrap_text=True, vertical='center')
    at.merge_cells('A15:Q15')
    at.row_dimensions[15].height = 30

    # en-têtes (lignes 16-17)
    ent = Font(name=FT, size=9, bold=True)
    for lettre, titre, _ in COLS:
        c = at[f'{lettre}16']
        c.value = titre
        at.merge_cells(f'{lettre}16:{lettre}17')
        for r in (16, 17):
            at[f'{lettre}{r}'].border = BORD
            at[f'{lettre}{r}'].alignment = centre
        c.font = ent
    at['A16'].comment = Comment("Les titres de cette ligne sont ceux cherchés dans l'export. "
                                "Si une colonne de l'export porte un autre nom, écris-le ici.", 'Attachement')
    at['M16'] = 'NB Ticket'
    at.merge_cells('M16:O16')
    at['M17'] = '="FORFAIT "&TEXT(HDEB,"hh:mm")&"-"&TEXT(HFIN,"hh:mm")'
    at['N17'] = 'HORS FORFAIT'
    at['O17'] = 'Motif hors forfait'
    at['P16'] = 'Plus-value sur devis (€)'
    at.merge_cells('P16:P17')
    for a in ('M16', 'N16', 'O16', 'M17', 'N17', 'O17', 'P16', 'P17'):
        at[a].fill = fond(VERT)
        at[a].border = BORD
        at[a].alignment = centre
        at[a].font = ent
    at.row_dimensions[16].height = 24
    at.row_dimensions[17].height = 30

    # lignes de tickets
    def lecture(lettre, r):
        rc = f"'Réglages'!$C${rang_col[lettre]}"
        src = f'INDEX(Export!$A:$DZ,LIGNE_TITRES+ROW()-{PREMIERE - 1},{rc})'
        if lettre == 'A':
            return f'=IF({rc}=0,"",IF({src}="","",{src}))'
        return f'=IF(OR($A{r}="",{rc}=0),"",IF({src}="","",{src}))'

    d9 = Font(name=FT, size=9)
    for r in range(PREMIERE, DERNIERE + 1):
        for lettre, _, _ in COLS:
            c = at[f'{lettre}{r}']
            c.value = lecture(lettre, r)
            c.font = d9
            c.alignment = Alignment(vertical='center', horizontal='center' if lettre in 'GH' else None)
            if lettre in 'GH':
                c.number_format = 'dd/mm/yyyy hh:mm'
        at[f'R{r}'] = f'=IF($A{r}="","",IF(ISNUMBER($G{r}),$G{r},IFERROR(--$G{r},"")))'
        at[f'R{r}'].number_format = 'dd/mm/yyyy hh:mm'
        hm = f'ROUND(MOD($R{r},1)*1440,0)'
        at[f'O{r}'] = (
            f'=IF($A{r}="","",IF(COUNTIF($A${PREMIERE}:$A{r},$A{r})>1,"Doublon",IF($R{r}="","",'
            f'IF(AND(MOIS<>"",OR($R{r}<MOIS,$R{r}>=EOMONTH(MOIS,0)+1)),"Autre mois",'
            f'IF(AND(WE="OUI",WEEKDAY($R{r},2)=6),"Samedi",IF(AND(WE="OUI",WEEKDAY($R{r},2)=7),"Dimanche",'
            f'IF(AND(WE="OUI",COUNTIF(FERIES,INT($R{r}))>0),"Jour férié",'
            f'IF({hm}<ROUND(HDEB*1440,0),"Avant "&TEXT(HDEB,"hh:mm")&" (début "&TEXT({hm}/1440,"hh:mm")&")",'
            f'IF({hm}>ROUND(HFIN*1440,0),"Après "&TEXT(HFIN,"hh:mm")&" (début "&TEXT({hm}/1440,"hh:mm")&")","")))))))))')
        at[f'M{r}'] = f'=IF($A{r}="","",IF($O{r}="",1,""))'
        at[f'N{r}'] = f'=IF(OR($A{r}="",$O{r}="",$O{r}="Doublon",$O{r}="Autre mois"),"",1)'
        for a in ('M', 'N'):
            at[f'{a}{r}'].font = d9
            at[f'{a}{r}'].alignment = Alignment(horizontal='center', vertical='center')
        at[f'O{r}'].font = d9
        at[f'P{r}'].font = Font(name=FT, size=9, bold=True)
        at[f'P{r}'].number_format = EUR
    plage = f'A{PREMIERE}:Q{DERNIERE}'
    for ligne in at.iter_rows(min_row=PREMIERE, max_row=DERNIERE, min_col=1, max_col=17):
        for c in ligne:
            c.border = BORD
    at.conditional_formatting.add(f'N{PREMIERE}:O{DERNIERE}', FormulaRule(
        formula=[f'$N{PREMIERE}=1'], fill=fond('FFC7CE'), font=Font(color='9C0006', bold=True)))
    at.conditional_formatting.add(f'P{PREMIERE}:P{DERNIERE}', FormulaRule(formula=[f'$N{PREMIERE}=1'], fill=fond(SAISIE)))
    # montant saisi sur un ticket au forfait : barré, il n'est pas compté
    at.conditional_formatting.add(f'P{PREMIERE}:P{DERNIERE}', FormulaRule(
        formula=[f'AND($P{PREMIERE}<>"",$N{PREMIERE}<>1)'], font=Font(color='A6A6A6', strike=True)))
    at.conditional_formatting.add(plage, FormulaRule(
        formula=[f'OR($O{PREMIERE}="Doublon",$O{PREMIERE}="Autre mois")'], font=Font(color='A6A6A6', italic=True)))
    at.auto_filter.ref = f'A17:Q{DERNIERE}'

    # impression : paysage, une page de large, titres répétés, zone limitée aux tickets présents
    at.page_setup.orientation = 'landscape'
    at.page_setup.paperSize = at.PAPERSIZE_A4
    at.page_setup.fitToWidth = 1
    at.page_setup.fitToHeight = 0
    at.sheet_properties.pageSetUpPr.fitToPage = True
    at.print_title_rows = '16:17'
    at.page_margins.left = at.page_margins.right = 0.3
    at.page_margins.top = at.page_margins.bottom = 0.5
    at.defined_names['_xlnm.Print_Area'] = DefinedName(
        '_xlnm.Print_Area', localSheetId=0,
        attr_text=f'OFFSET(Attachement!$A$1,0,0,{PREMIERE - 1}+SUMPRODUCT(--(Attachement!$A${PREMIERE}:$A${DERNIERE}<>"")),17)')
    at.sheet_view.zoomScale = 90

    wb.active = 0
    wb.calculation.fullCalcOnLoad = True
    wb.save(chemin)


if __name__ == '__main__':
    noms = sys.argv[1:]
    if not noms:
        construire('NOM DU SOUS-TRAITANT', ICI / 'Attachement_sous-traitant.xlsx')
        print('Attachement_sous-traitant.xlsx')
    for n in noms:
        f = ICI / f'Attachement {n}.xlsx'
        construire(n, f)
        print(f.name)
