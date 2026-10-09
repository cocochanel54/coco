# Construit Bareme_points.xlsx : barème de points (temps standards) par famille de tâches,
# à partir de l'onglet BPU de Journee_agent.xlsm. Aucun prix n'est repris.
import re, sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment
from openpyxl.workbook.defined_name import DefinedName

SRC, OUT = sys.argv[1], sys.argv[2]

# Familles : (nom, ce qui est compté, temps estimé en minutes par unité, remarque)
FAMILLES = [
    ("Ticket intervention",               "par ticket",        120, "Ticket de maintenance : déplacement local, diagnostic, réparation simple."),
    ("Malfaçon aérien",                    "par ticket",        150, "Reprise en hauteur (échelle ou nacelle)."),
    ("Malfaçon souterrain",                "par ticket",         90, "Reprise en chambre."),
    ("Heure de technicien",                "par heure",          60, "Prestations facturées à l'heure : 1 h = 1 h de travail. En binôme, chaque technicien saisit sa ligne."),
    ("City Fast N1 · site (PM, DME)",      "par intervention",  180, "Délai (urgent, standard) et horaire ne changent pas le travail : seule l'heure de nuit est majorée (voir Réglages)."),
    ("City Fast N2 · BPE en chambre",      "par intervention",  240, ""),
    ("City Fast N3 · câble en conduite",   "par intervention",  360, "Tirage de câble : souvent une grosse demi-journée."),
    ("SAV niveau 0",                       "par intervention",   90, ""),
    ("SAV niveau 1",                       "par intervention",  180, "Inclut le niveau 0."),
    ("SAV niveau 2",                       "par intervention",  300, "Inclut le niveau 1."),
    ("Raccordement entreprise en immeuble","par raccordement",  360, "46.x, 47.x, 59.1 FTTE, super forfait FTTE et lien supplémentaire."),
    ("Raccordement entreprise avec extension (Z50 à Z200)", "par raccordement", 600, "48.x à 53.x et super forfaits Z100 / Z200. Le câble posé en plus se compte au mètre."),
    ("Raccordement FTTH Pro",              "par raccordement",  180, "60.1 à 60.5, 62.9 (création ou déplacement de PTO)."),
    ("Petite intervention équipement",     "par intervention",   60, "CPE, ONT, tiroir, résiliation, upgrade simple."),
    ("Mise en service / upgrade",          "par intervention",  120, "56.1 upgrade, 60.8 mise en service FTTE."),
    ("Visite technique",                   "par visite",         90, "61.1 à 61.3."),
    ("Câble posé ou tiré",                 "par mètre",           1, "Conduite, aérien, façade, immeuble, extension. 1 min/m ≈ 400 m par jour."),
    ("Aiguillage / dépose / portage de câble", "par mètre",     0.5, ""),
    ("Chemin de câble, goulotte, gaine",   "par mètre",           5, "Fixation au mur ou au plafond."),
    ("Percement simple",                   "par percement",      30, "Cloison, parpaing."),
    ("Carottage / percement spécifique",   "par percement",      90, "Béton armé, traversée d'étage."),
    ("Soudure à la fibre",                 "par fibre",           3, "Soudures supplémentaires."),
    ("Boîte d'épissure (joint) jusqu'à 48 FO",   "par boîte",   120, "Joints droits ou blancs (lovage, tubes)."),
    ("Boîte d'épissure (joint) 72 à 144 FO",     "par boîte",   240, ""),
    ("Boîte d'épissure (joint) 192 à 288 FO",    "par boîte",   360, ""),
    ("Boîte d'épissure (joint) 432 FO et plus",  "par boîte",   600, "Souvent plus d'une journée."),
    ("Rentrée / piquage dans une BPE",     "par intervention",  120, "17.1, 17.3, 17.5, 58.x."),
    ("Fermeture / fixation de BPE",        "par boîte",          30, ""),
    ("Audit BPE jusqu'à 288 FO",           "par boîte",          90, ""),
    ("Audit BPE plus de 288 FO",           "par boîte",         150, ""),
    ("Récolement lors de modification",    "par boîte",          45, ""),
    ("Tête de câble jusqu'à 24 FO",        "par tête",          120, ""),
    ("Tête de câble 36 à 72 FO",           "par tête",          240, ""),
    ("Tête de câble 96 à 144 FO",          "par tête",          360, ""),
    ("Mesure de liaison client",           "par mesure",         90, "20.1."),
    ("Mesure de liaison BPE à BPE",        "par mesure",        360, "20.2 à 20.4."),
    ("Mesure réflectométrique",            "par fibre",          30, ""),
    ("Jarretière, coupleur, breakout",     "par unité",          15, ""),
    ("Tiroir PM / baie / départ électrique","par unité",        180, "20.8, 20.9, 20.19, 22.5."),
    ("Pompage de chambre",                 "par chambre",        45, ""),
    ("Pompage de chambre plafonnée",       "par chambre",        90, ""),
    ("Poteau / travaux en hauteur intérieur","par unité",        90, "62.4 élévateur, 62.6 armement de poteau."),
    ("Gros travaux / génie civil",         "par unité",         480, "Hydrocurage, réparation de fourreaux, chambre, armoire, NRO, shelter, élagage. Souvent sous-traité : à ajuster."),
    ("Étude simple / DOE",                 "par étude",         120, "Travail de bureau ou de terrain léger."),
    ("Étude complexe",                     "par étude",         300, "2.3, 3.2, 4.2, 8.3."),
    ("Étude grande (jusqu'à 500 m)",       "par étude",         600, "2.4."),
    ("Étude au mètre (au-delà de 500 m)",  "par mètre",         0.5, "2.5."),
    ("Survey / réunion / audit (½ journée)","par unité",        210, "5.1."),
    ("Test d'étanchéité de boîte",         "par boîte",          20, "6.3."),
    ("Sans temps (matériel, plus-value)",  "—",                   0, "Lignes de prix seules : location de nacelle, plus-value de câble, fourniture, rendez-vous, majoration hotline. Ne comptent pas de travail."),
    ("Sur devis : temps à saisir",         "—",                   0, "Désamiantage, rocade datacenter : saisir un temps spécifique dans « Prestations »."),
    ("Absence (non comptée)",              "par journée",         0, "Congés, absences : la journée est retirée des jours travaillés, ce n'est pas un manque de production."),
]

def num(code):
    m = re.search(r'(\d+)\.(\d+)$', code)
    return (int(m.group(1)), int(m.group(2))) if m else (None, None)

def fo(lib):
    m = re.search(r'(\d+)\s*FO', lib)
    return int(m.group(1)) if m else None

def famille(projet, code, lib):
    c = code.upper()
    if projet in ('CONGÉS', 'ABSENCE'): return "Absence (non comptée)"
    if c.endswith('-TICKET'): return "Ticket intervention"
    if 'MALF-AERIEN' in c: return "Malfaçon aérien"
    if 'MALF-SOUT' in c: return "Malfaçon souterrain"
    if c in ('MC-T1', 'MC-T2', 'MC-T3', 'MC-T4', 'MC-T5'): return "Heure de technicien"
    if c == 'MC-T6': return "Sans temps (matériel, plus-value)"
    if c.startswith('CF-N1'): return "City Fast N1 · site (PM, DME)"
    if c.startswith('CF-N2'): return "City Fast N2 · BPE en chambre"
    if c.startswith('CF-N3'): return "City Fast N3 · câble en conduite"
    if c.startswith('SAV-F0'): return "SAV niveau 0"
    if c.startswith('SAV-F1'): return "SAV niveau 1"
    if c.startswith('SAV-F2'): return "SAV niveau 2"
    if c in ('COV-SF-FTTE', 'COV-SF-LIEN'): return "Raccordement entreprise en immeuble"
    if c in ('COV-SF-Z100', 'COV-SF-Z200'): return "Raccordement entreprise avec extension (Z50 à Z200)"
    if c.startswith('COV-ETU'):
        n = c.replace('COV-ETU-', '')
        if n == '2.4': return "Étude grande (jusqu'à 500 m)"
        if n == '2.5': return "Étude au mètre (au-delà de 500 m)"
        if n in ('2.3', '3.2', '4.2', '8.3'): return "Étude complexe"
        if n == '5.1': return "Survey / réunion / audit (½ journée)"
        if n == '6.3': return "Test d'étanchéité de boîte"
        return "Étude simple / DOE"
    a, b = num(code)
    if a is None: return None
    if a in (46, 47) or (a, b) == (59, 1): return "Raccordement entreprise en immeuble"
    if 48 <= a <= 53: return "Raccordement entreprise avec extension (Z50 à Z200)"
    if a == 55: return "Câble posé ou tiré" if b <= 4 else "Sans temps (matériel, plus-value)"
    if (a, b) in ((56, 2), (56, 3), (57, 1), (60, 6), (60, 7), (20, 6), (20, 7)): return "Petite intervention équipement"
    if (a, b) in ((56, 1), (60, 8)): return "Mise en service / upgrade"
    if (a, b) in ((56, 4), (63, 1), (63, 2)): return "Heure de technicien"
    if a == 58: return "Rentrée / piquage dans une BPE"
    if (a, b) in ((59, 2), (59, 3), (60, 3), (61, 4), (12, 11)): return "Sans temps (matériel, plus-value)" if (a, b) != (59, 3) else "Jarretière, coupleur, breakout"
    if a == 60 or (a, b) == (62, 9): return "Raccordement FTTH Pro"
    if a == 61: return "Visite technique"
    if (a, b) == (62, 1): return "Câble posé ou tiré"
    if (a, b) == (62, 2): return "Chemin de câble, goulotte, gaine"
    if (a, b) == (62, 3): return "Carottage / percement spécifique"
    if (a, b) in ((62, 4), (62, 6)): return "Poteau / travaux en hauteur intérieur"
    if (a, b) in ((62, 5), (62, 7), (62, 8), (11, 11), (11, 14)): return "Gros travaux / génie civil"
    if (a, b) in ((62, 10), (11, 12)): return "Pompage de chambre"
    if (a, b) == (11, 13): return "Pompage de chambre plafonnée"
    if a == 11 and b <= 8: return "Câble posé ou tiré" if b not in (6,) else "Aiguillage / dépose / portage de câble"
    if a == 11: return "Aiguillage / dépose / portage de câble"
    if a == 12:
        if b in (1, 2, 3, 4, 10): return "Chemin de câble, goulotte, gaine"
        if b == 5: return "Percement simple"
        if b == 6: return "Carottage / percement spécifique"
        if b in (7, 8, 12, 13, 14): return "Câble posé ou tiré"
        if b in (9, 15): return "Sur devis : temps à saisir"
    if a in (14, 15, 16):
        if 'soudure' in lib.lower(): return "Soudure à la fibre"
        n = fo(lib) or 6
        if n <= 48: return "Boîte d'épissure (joint) jusqu'à 48 FO"
        if n <= 144: return "Boîte d'épissure (joint) 72 à 144 FO"
        if n <= 288: return "Boîte d'épissure (joint) 192 à 288 FO"
        return "Boîte d'épissure (joint) 432 FO et plus"
    if a == 17:
        if b in (2, 4, 6): return "Soudure à la fibre"
        if b == 7: return "Fermeture / fixation de BPE"
        return "Rentrée / piquage dans une BPE"
    if a == 18:
        if b == 7: return "Récolement lors de modification"
        return "Audit BPE jusqu'à 288 FO" if b <= 3 else "Audit BPE plus de 288 FO"
    if a == 19:
        n = fo(lib)
        if n <= 24: return "Tête de câble jusqu'à 24 FO"
        if n <= 72: return "Tête de câble 36 à 72 FO"
        return "Tête de câble 96 à 144 FO"
    if a == 20:
        if b == 1: return "Mesure de liaison client"
        if b in (2, 3, 4): return "Mesure de liaison BPE à BPE"
        if b == 5: return "Mesure réflectométrique"
        if b in (8, 9, 19): return "Tiroir PM / baie / départ électrique"
        return "Jarretière, coupleur, breakout"
    if a in (21, 22, 23):
        if (a, b) in ((21, 4), (22, 2), (23, 2)): return "Câble posé ou tiré"
        if (a, b) == (22, 5): return "Tiroir PM / baie / départ électrique"
        return "Gros travaux / génie civil"
    return None

# ---------- lecture du BPU (sans les prix) ----------
bpu = openpyxl.load_workbook(SRC)['BPU']
lignes = []
for r in range(6, bpu.max_row + 1):
    p, contrat, code, des, unite, lib = (bpu.cell(r, c).value for c in (1, 2, 3, 4, 5, 7))
    if not p or not code: continue
    lib = lib or des
    f = famille(p, code, lib)
    if f is None: raise SystemExit(f'Prestation non classée : {code} {lib}')
    hno = 'oui' if re.search(r'\bHNO\b', lib) else ''
    lignes.append((p, contrat if contrat and contrat != '—' else '', code, lib, unite or '', f, hno))
noms = [f[0] for f in FAMILLES]
assert all(l[5] in noms for l in lignes)

# ---------- styles ----------
ARIAL = 'Arial'
F = lambda **k: Font(name=ARIAL, size=k.pop('size', 10), **k)
TITRE = F(size=14, bold=True, color='1F3864')
ENTETE = F(bold=True, color='FFFFFF'); FOND_ENTETE = PatternFill('solid', fgColor='1F3864')
SAISIE = PatternFill('solid', fgColor='FFF2CC'); BLEU = F(color='0000FF')
GRIS = PatternFill('solid', fgColor='F2F2F2')
fin = Side(style='thin', color='BFBFBF'); BORD = Border(left=fin, right=fin, top=fin, bottom=fin)
WRAP = Alignment(wrap_text=True, vertical='top')

wb = openpyxl.Workbook()

# ---------- Mode d'emploi + réglages ----------
ws = wb.active; ws.title = "Mode d'emploi"
ws['A1'] = 'BARÈME DE POINTS — MESURER LE TRAVAIL DES TECHNICIENS, PAS LE PRIX'; ws['A1'].font = TITRE
texte = [
    "Chaque type de tâche vaut des points selon le temps et la difficulté réels, pas selon ce qu'elle rapporte.",
    "La facturation ne change pas : les euros restent dans ton Excel pour les clients. Les points servent seulement à comparer les techniciens entre eux, de façon juste.",
    "",
    "CE QUE TU AS À FAIRE",
    "1. Onglet « Familles » : corrige les temps de la colonne D (cases jaunes). Ce sont des estimations de départ, à valider avec les techniciens.",
    "2. Onglet « Prestations » : si une prestation ne correspond pas à sa famille, change sa famille (liste en colonne F) ou mets un temps à part en colonne H (cases jaunes).",
    "3. Les réglages ci-dessous (cases jaunes) : valeur d'un point, durée de la journée, majoration de nuit.",
    "4. Onglet « Exemple » : une journée type pour voir le calcul. Tu peux y changer les codes et les quantités.",
    "",
    "Cases jaunes, chiffres en bleu = à saisir ou à corriger. Chiffres en noir = calculés, ne pas modifier.",
]
for i, t in enumerate(texte, start=3):
    ws.cell(i, 1, t).font = F(bold=t.isupper() and t != '')
ws['A14'] = 'RÉGLAGES'; ws['A14'].font = F(bold=True, size=11, color='1F3864')
reglages = [
    ('Minutes de travail pour 1 point', 15, 'MinParPoint', '0', "1 point = 15 minutes de travail normal. Proposition à valider."),
    ('Heures de travail dans une journée', 7, 'HeuresJour', '0.0', "Journée de référence pour calculer le taux. Proposition à valider."),
    ('Points attendus pour une journée', '=HeuresJour*60/MinParPoint', 'PointsJour', '0.0', "Calculé : une journée normale vaut ce nombre de points (100 %)."),
    ('Majoration pour une intervention de nuit (HNO)', 0.25, 'MajHNO', '0%', "Proposition : +25 % sur les prestations de nuit, pour la pénibilité. Mets 0 % pour ne rien majorer."),
    ('Bonus « intervention difficile »', 0.5, 'BonusDifficile', '0%', "Pour plus tard : bonus que le technicien demande et que tu valides. Pas encore utilisé dans le calcul."),
]
for i, (lab, val, nom, fmt, com) in enumerate(reglages, start=15):
    ws.cell(i, 1, lab).font = F()
    c = ws.cell(i, 2, val); c.number_format = fmt; c.border = BORD
    if isinstance(val, str): c.font = F()
    else: c.font = BLEU; c.fill = SAISIE
    ws.cell(i, 3, com).font = F(italic=True, color='595959')
    wb.defined_names[nom] = DefinedName(nom, attr_text=f"'Mode d''emploi'!$B${i}")
ws.column_dimensions['A'].width = 48; ws.column_dimensions['B'].width = 12; ws.column_dimensions['C'].width = 95
ws['A21'] = "POURQUOI C'EST PLUS JUSTE"; ws['A21'].font = F(bold=True, size=11, color='1F3864')
pourquoi = [
    "• Un technicien est comparé à la journée de référence, pas aux tarifs du client : 28 points sur une journée de 7 h = 100 %, quel que soit le projet.",
    "• Les tâches longues ou difficiles rapportent autant de points que le temps qu'elles demandent, même si elles sont mal payées.",
    "• Les congés et absences sont retirés des jours travaillés : ils ne font pas baisser le taux.",
    "• Au bout d'un mois, on pourra comparer ces temps avec la réalité et ajuster le barème.",
]
for i, t in enumerate(pourquoi, start=22): ws.cell(i, 1, t).font = F()

# ---------- Familles ----------
wf = wb.create_sheet('Familles')
wf['A1'] = 'FAMILLES DE TÂCHES ET TEMPS STANDARDS'; wf['A1'].font = TITRE
wf['A2'] = "Corrige seulement la colonne D (minutes par unité). Les points et le nombre de prestations se calculent seuls."; wf['A2'].font = F(italic=True, color='595959')
ent = ['Famille', 'Ce qui est compté', 'Exemples de prestations', 'Temps estimé (min / unité)', 'Points / unité', 'Nb de prestations', 'Remarque']
for j, h in enumerate(ent, 1):
    c = wf.cell(4, j, h); c.font = ENTETE; c.fill = FOND_ENTETE; c.alignment = Alignment(wrap_text=True, vertical='center'); c.border = BORD
for i, (nom, compte, minutes, rem) in enumerate(FAMILLES, start=5):
    ex = [l for l in lignes if l[5] == nom]
    exemples = ' · '.join(dict.fromkeys(l[3] for l in ex[:3]))
    vals = [nom, compte, exemples, minutes, f'=D{i}/MinParPoint', f'=COUNTIF(Prestations!$F$5:$F${4+len(lignes)},A{i})', rem]
    for j, v in enumerate(vals, 1):
        c = wf.cell(i, j, v); c.font = F(); c.border = BORD; c.alignment = WRAP
    wf.cell(i, 1).font = F(bold=True)
    d = wf.cell(i, 4); d.font = BLEU; d.fill = SAISIE; d.number_format = '0.0'
    wf.cell(i, 5).number_format = '0.00'
nfam_last = 4 + len(FAMILLES)
wf.cell(5, 4).comment = Comment("Estimation de départ proposée par Claude, à corriger avec l'expérience des techniciens.", 'Claude')
for col, w in zip('ABCDEFG', (40, 16, 60, 13, 10, 11, 60)): wf.column_dimensions[col].width = w
wf.freeze_panes = 'B5'

# ---------- Prestations ----------
wp = wb.create_sheet('Prestations')
wp['A1'] = 'PRESTATIONS DU BPU ET LEURS POINTS'; wp['A1'].font = TITRE
wp['A2'] = "Change la famille (colonne F) ou mets un temps à part (colonne H) si besoin. Colonne I : « oui » = intervention de nuit, majorée."; wp['A2'].font = F(italic=True, color='595959')
ent = ['Projet', 'Contrat / réseau', 'Code BPU', 'Prestation', 'Unité', 'Famille', 'Temps famille (min)', 'Temps à part (min)', 'Nuit (HNO)', 'Temps retenu (min)', 'Points']
for j, h in enumerate(ent, 1):
    c = wp.cell(4, j, h); c.font = ENTETE; c.fill = FOND_ENTETE; c.alignment = Alignment(wrap_text=True, vertical='center'); c.border = BORD
for i, (p, contrat, code, lib, unite, fam, hno) in enumerate(lignes, start=5):
    vals = [p, contrat, code, lib, unite, fam,
            f'=IFERROR(INDEX(Familles!$D$5:$D${nfam_last},MATCH(F{i},Familles!$A$5:$A${nfam_last},0)),"")',
            None, hno,
            f'=IF(AND(H{i}="",G{i}=""),"",IF(H{i}<>"",H{i},G{i})*IF(I{i}="oui",1+MajHNO,1))',
            f'=IF(J{i}="","",J{i}/MinParPoint)']
    for j, v in enumerate(vals, 1):
        c = wp.cell(i, j, v); c.font = F(); c.border = BORD
    for j in (6, 8, 9):
        wp.cell(i, j).fill = SAISIE; wp.cell(i, j).font = BLEU
    for j in (7, 10): wp.cell(i, j).number_format = '0.0'
    wp.cell(i, 11).number_format = '0.00'
last = 4 + len(lignes)
dv = DataValidation(type='list', formula1=f'=Familles!$A$5:$A${nfam_last}', allow_blank=False)
dv.error = 'Choisis une famille de la liste (onglet Familles).'; wp.add_data_validation(dv); dv.add(f'F5:F{last}')
dv2 = DataValidation(type='list', formula1='"oui"', allow_blank=True); wp.add_data_validation(dv2); dv2.add(f'I5:I{last}')
for col, w in zip('ABCDEFGHIJK', (11, 16, 15, 58, 22, 40, 11, 11, 9, 11, 9)): wp.column_dimensions[col].width = w
wp.freeze_panes = 'E5'; wp.auto_filter.ref = f'A4:K{last}'

# ---------- Exemple ----------
we = wb.create_sheet('Exemple')
we['A1'] = "EXEMPLE : DEUX JOURNÉES COMPARÉES AVEC LES POINTS"; we['A1'].font = TITRE
we['A2'] = "Change les codes (colonne B) et les quantités (colonne D) pour tester. Les points se calculent avec le barème."; we['A2'].font = F(italic=True, color='595959')
ent = ['Technicien', 'Code BPU', 'Prestation', 'Quantité', 'Points / unité', 'Points']
exemple = [('Technicien A', 'COVAGE-TICKET', 3), ('Technicien B', 'ALTITUDE-MALF-AERIEN', 2), ('Technicien B', 'LOSANGE-MALF-SOUT', 1),
           ('', '', None), ('', '', None), ('', '', None)]
for j, h in enumerate(ent, 1):
    c = we.cell(4, j, h); c.font = ENTETE; c.fill = FOND_ENTETE; c.border = BORD
for i, (t, code, q) in enumerate(exemple, start=5):
    vals = [t or None, code or None, f'=IF(B{i}="","",IFERROR(INDEX(Prestations!$D$5:$D${last},MATCH(B{i},Prestations!$C$5:$C${last},0)),"Code inconnu"))', q,
            f'=IF(B{i}="","",IFERROR(INDEX(Prestations!$K$5:$K${last},MATCH(B{i},Prestations!$C$5:$C${last},0)),0))', f'=IF(B{i}="","",N(D{i})*E{i})']
    for j, v in enumerate(vals, 1):
        c = we.cell(i, j, v); c.font = F(); c.border = BORD
    for j in (1, 2, 4): we.cell(i, j).font = BLEU; we.cell(i, j).fill = SAISIE
    we.cell(i, 5).number_format = '0.00'; we.cell(i, 6).number_format = '0.0'
r0 = 5 + len(exemple) + 1
we.cell(r0, 1, 'Résultat de la journée').font = F(bold=True, size=11, color='1F3864')
for j, h in enumerate(['Technicien', 'Points de la journée', 'Points attendus', 'Taux'], 1):
    c = we.cell(r0 + 1, j, h); c.font = ENTETE; c.fill = FOND_ENTETE; c.border = BORD
for k, t in enumerate(['Technicien A', 'Technicien B']):
    r = r0 + 2 + k
    vals = [t, f'=SUMIFS($F$5:$F${4+len(exemple)},$A$5:$A${4+len(exemple)},A{r})', '=PointsJour', f'=IF(C{r}>0,B{r}/C{r},"")']
    for j, v in enumerate(vals, 1):
        c = we.cell(r, j, v); c.font = F(); c.border = BORD
    we.cell(r, 2).number_format = '0.0'; we.cell(r, 3).number_format = '0.0'; we.cell(r, 4).number_format = '0%'
we.cell(r0 + 5, 1, "A a fait 3 tickets COVAGE, bien payés. B a fait 3 malfaçons, mal payées mais longues (dont 2 en hauteur). En euros, A rapporte plus de deux fois plus que B et passerait pour le meilleur.").font = F(italic=True, color='595959')
we.cell(r0 + 6, 1, "Avec les points, B a en réalité travaillé plus longtemps que A : chacun est jugé sur son temps de travail, pas sur le tarif du client.").font = F(italic=True, color='595959')
we.cell(r0 + 7, 1, "Lignes 8 à 10 libres : ajoute un technicien, un code BPU et une quantité pour tester d'autres journées.").font = F(italic=True, color='595959')
for col, w in zip('ABCDEF', (16, 16, 58, 10, 13, 10)): we.column_dimensions[col].width = w

for s in wb.worksheets: s.sheet_view.showGridLines = False
wb.save(OUT)
print(len(lignes), 'prestations,', len(FAMILLES), 'familles')
