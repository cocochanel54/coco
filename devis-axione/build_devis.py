# Construit Devis_BPU_Axione_RHD57.xlsx : un devis dont les prix viennent des deux BPU
# Moselle Fibre RHD57 (Maintenance et VDR Optique). On choisit la prestation dans une liste,
# on saisit la quantité, le reste se calcule.
#
#   python3 build_devis.py <BPU Maintenance.xlsx> <BPU VDR Optique.xlsx> <sortie.xlsx>
import re, sys
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.comments import Comment
from openpyxl.workbook.defined_name import DefinedName

SRC_MAINT, SRC_OPT, OUT = sys.argv[1], sys.argv[2], sys.argv[3]

EURO = '_-* #,##0.00\\ "€"_-;\\-* #,##0.00\\ "€"_-;_-* "-"??\\ "€"_-;_-@_-'  # format des BPU d'origine
COEF = '0.0000'
NB_LIGNES = 30

ORANGE, BLEU = "FF6600", "0093BB"  # couleurs des bandeaux des BPU d'origine
F = lambda **k: Font(name="Arial", size=k.pop("size", 10), **k)
fill = lambda c: PatternFill("solid", fgColor=c)
SAISIE = fill("FFF2CC")
fin = Side(style="thin", color="A6A6A6")
CADRE = Border(left=fin, right=fin, top=fin, bottom=fin)
GAUCHE = Alignment(horizontal="left", vertical="center", wrap_text=True)
CENTRE = Alignment(horizontal="center", vertical="center", wrap_text=True)
DROITE = Alignment(horizontal="right", vertical="center")


def propre(t):
    return re.sub(r"[ \t]+", " ", str(t or "")).strip()


def lignes(t):
    return [propre(l) for l in str(t or "").split("\n") if propre(l)]


def prix(v):
    # Les BPU écrivent « sur devis », « Devis », « devis BPU »… : une seule écriture dans le devis.
    if v is None or isinstance(v, (int, float)):
        return v
    return "Sur devis"


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


def libelle(d, n=95):
    t = d["desig"]
    if len(t) > n:
        t = t[:n].rsplit(" ", 1)[0].rstrip(" ,;-–") + "…"
    return f"{d['code']} · {t}"


for d in articles:
    d["lib"] = libelle(d)
    assert d["pbpu"] is not None, d["code"]
libs = [d["lib"] for d in articles]
assert len(libs) == len(set(libs)), "libellés en double"
N = len(articles) + 1  # dernière ligne du tableau BPU

MAJ_LIB = {"MC-07": "MC-07 · Technicien supplémentaire", "MC-08": "MC-08 · Intervention en HNO"}
assert [m["code"] for m in majorations] == ["MC-07", "MC-08"]

wb = openpyxl.Workbook()
dv_ws = wb.active
dv_ws.title = "Devis"
bpu = wb.create_sheet("BPU")
maj = wb.create_sheet("Majorations")

# ---------- Onglet BPU ----------
ent = ["BPU", "Lot", "Famille", "Code", "Désignation", "Descriptif", "Unité",
       "Prix BPU\n€ HT / u", "Commentaire", "Libellé dans le devis", "Source"]
larg = [12, 30, 30, 9, 60, 50, 16, 14, 40, 70, 42]
for c, (t, w) in enumerate(zip(ent, larg), 1):
    x = bpu.cell(1, c, t)
    x.font, x.fill, x.alignment, x.border = F(bold=True, color="FFFFFF"), fill(BLEU), CENTRE, CADRE
    bpu.column_dimensions[x.column_letter].width = w
bpu.row_dimensions[1].height = 30
bpu["H1"].comment = Comment("Maintenance : colonne « Prix unitaire Euro(s) / u ».\n"
                            "VDR Optique : colonne « PU ».", "Devis")
for i, d in enumerate(articles, 2):
    vals = [d["bpu"], d["lot"], d["famille"], d["code"], d["desig"], d["descr"], d["unite"],
            d["pbpu"], d["comm"], d["lib"], d["src"]]
    for c, v in enumerate(vals, 1):
        x = bpu.cell(i, c, v)
        x.font = F(color="0000FF") if c == 8 else F()
        x.border = CADRE
        x.alignment = Alignment(vertical="center", wrap_text=c == 5)
        if c == 8:
            x.number_format = EURO
            x.alignment = Alignment(horizontal="right", vertical="center")
    if d["bpu"] == "VDR Optique":
        for c in range(1, 5):
            bpu.cell(i, c).fill = fill("EAF6FA")
bpu.freeze_panes = "E2"
bpu.auto_filter.ref = f"A1:K{N}"

# ---------- Onglet Majorations ----------
for c, (t, w) in enumerate(zip(["Code", "Libellé dans le devis", "Coefficient", "Unité", "Commentaire", "Source"],
                               [9, 36, 13, 26, 60, 42]), 1):
    x = maj.cell(1, c, t)
    x.font, x.fill, x.alignment, x.border = F(bold=True, color="FFFFFF"), fill(BLEU), CENTRE, CADRE
    maj.column_dimensions[x.column_letter].width = w
for i, m in enumerate(majorations, 2):
    vals = [m["code"], MAJ_LIB[m["code"]], m["pbpu"], m["unite"],
            m["desig"] + (f" ({m['comm']})" if m["comm"] else ""), m["src"]]
    for c, v in enumerate(vals, 1):
        x = maj.cell(i, c, v)
        x.font, x.border = (F(color="0000FF") if c == 3 else F()), CADRE
        x.alignment = Alignment(vertical="center", wrap_text=c == 5)
    maj.cell(i, 3).number_format = COEF
x = maj.cell(5, 1, "Le montant de la ligne du devis est multiplié par le coefficient "
               "(ex. 1,4078 = +40,78 %).")
x.font, x.alignment = F(italic=True, color="595959"), GAUCHE
maj.merge_cells("A5:F5")
maj.row_dimensions[5].height = 30

wb.defined_names["ListePrestations"] = DefinedName("ListePrestations", attr_text=f"BPU!$J$2:$J${N}")
wb.defined_names["ListeMajorations"] = DefinedName("ListeMajorations", attr_text="Majorations!$B$2:$B$3")
wb.defined_names["TauxTVA"] = DefinedName("TauxTVA", attr_text="Devis!$N$2")
wb.defined_names["Autoliq"] = DefinedName("Autoliq", attr_text="Devis!$N$3")

# ---------- Onglet Devis ----------
ws = dv_ws
for col, w in zip("ABCDEFGHIJKLMN", [5, 72, 22, 9, 13, 13, 29, 15, 6, 6, 3, 34, 4, 26]):
    ws.column_dimensions[col].width = w
ws.column_dimensions["J"].hidden = ws.column_dimensions["K"].hidden = True  # calculs intermédiaires
ws.sheet_view.showGridLines = False


def saisie(ref, v=None, **font):
    x = ws[ref]
    x.value = v
    x.fill, x.font, x.alignment = SAISIE, F(color="0000FF", **font), GAUCHE
    return x


def bandeau(rng, t, couleur=BLEU):
    ws.merge_cells(rng)
    x = ws[rng.split(":")[0]]
    x.value, x.font, x.fill, x.alignment = t, F(bold=True, color="FFFFFF"), fill(couleur), GAUCHE


# En-tête : émetteur à gauche, références du devis à droite
emetteur = ["Raison sociale", "Adresse", "Code postal – Ville", "SIRET : ", "N° TVA intracommunautaire : ",
            "Tél. :                     E-mail : "]
for i, t in enumerate(emetteur, 1):
    ws.merge_cells(f"A{i}:B{i}")
    saisie(f"A{i}", t, bold=i == 1, size=13 if i == 1 else 10)
ws.merge_cells("F1:H1")
ws["F1"] = "DEVIS"
ws["F1"].font, ws["F1"].alignment = F(bold=True, size=22, color=BLEU), DROITE
ws.row_dimensions[1].height = 30
refs = [("N° de devis", "DEV-2026-001"), ("Date", datetime(2026, 10, 21)), ("Validité (jours)", 30),
        ("Valable jusqu'au", "=IF(OR(H3=\"\",H4=\"\"),\"\",H3+H4)"), ("Réf. client / commande", None)]
for i, (lab, v) in enumerate(refs, 2):
    ws.merge_cells(f"F{i}:G{i}")
    ws[f"F{i}"] = lab
    ws[f"F{i}"].font, ws[f"F{i}"].alignment = F(bold=True), DROITE
    if lab == "Valable jusqu'au":
        ws[f"H{i}"] = v
        ws[f"H{i}"].font, ws[f"H{i}"].alignment = F(), Alignment(horizontal="center")
    else:
        saisie(f"H{i}", v).alignment = Alignment(horizontal="center", vertical="center")
    ws[f"H{i}"].border = CADRE
    if lab in ("Date", "Valable jusqu'au"):
        ws[f"H{i}"].number_format = "dd/mm/yyyy"

# Client et chantier
bandeau("A8:B8", "CLIENT")
bandeau("C8:H8", "CHANTIER")
for i, t in enumerate(["Société", "Contact", "Adresse", "Code postal – Ville"], 9):
    ws.merge_cells(f"A{i}:B{i}")
    saisie(f"A{i}", t)
chantier = [("Projet / réseau", "Moselle Fibre – RHD57"),
            ("Objet", "Exemple : raccordements, SAV et pose de câble (à remplacer)"),
            ("Site / commune", None),
            ("Base de prix", "BPU Moselle Fibre – RHD57 (Maintenance et VDR Optique)")]
for i, (lab, v) in enumerate(chantier, 9):
    ws[f"C{i}"] = lab
    ws[f"C{i}"].font = F(bold=True)
    ws.merge_cells(f"D{i}:H{i}")
    if lab == "Base de prix":
        ws[f"D{i}"] = v
        ws[f"D{i}"].font = F()
    else:
        saisie(f"D{i}", v)
for r in range(9, 13):
    for c in "ABCDEFGH":
        ws[f"{c}{r}"].border = CADRE

# Tableau des prestations
H0, L1 = 14, 15
LN = L1 + NB_LIGNES - 1
for c, t in zip("ABCDEFGH", ["N°", "Prestation (code · désignation BPU)", "Unité", "Qté",
                             "PU HT BPU", "PU HT\nsur devis", "Majoration", "Montant HT"]):
    x = ws[f"{c}{H0}"]
    x.value, x.font, x.fill, x.alignment, x.border = t, F(bold=True, color="FFFFFF"), fill(BLEU), CENTRE, CADRE
ws.row_dimensions[H0].height = 30
ws["J14"], ws["K14"] = "lig", "prix"

rng = lambda col: f"BPU!${col}$2:${col}${N}"
for r in range(L1, LN + 1):
    ws[f"J{r}"] = f'=IF($B{r}="","",IFERROR(MATCH($B{r},ListePrestations,0),0))'
    ws[f"K{r}"] = f'=IF(OR($J{r}="",$J{r}=0),"",INDEX({rng("H")},$J{r}))'
    ws[f"A{r}"] = f'=IF($B{r}="","",COUNTA($B${L1}:$B{r}))'
    ws[f"C{r}"] = f'=IF(OR($J{r}="",$J{r}=0),"",INDEX({rng("G")},$J{r})&"")'
    ws[f"E{r}"] = f'=IF($B{r}="","",IF($J{r}=0,"Hors BPU",IF(ISNUMBER($K{r}),ROUND($K{r},2),$K{r})))'
    coef = (f'IF($G{r}="",1,IFERROR(INDEX(Majorations!$C$2:$C$3,'
            f'MATCH($G{r},ListeMajorations,0)),1))')
    ws[f"H{r}"] = (f'=IF($B{r}="","",IF(AND($F{r}="",NOT(ISNUMBER($E{r}))),"PU à saisir",'
                   f'ROUND(N($D{r})*IF($F{r}<>"",$F{r},$E{r})*{coef},2)))')
    for c in "ABCDEFGH":
        x = ws[f"{c}{r}"]
        x.border = CADRE
        x.font = F()
        x.alignment = Alignment(horizontal="center", vertical="center") if c in "ACDG" else Alignment(vertical="center")
    for c in "BDFG":
        ws[f"{c}{r}"].fill = SAISIE
        ws[f"{c}{r}"].font = F(color="0000FF")
    ws[f"B{r}"].alignment = Alignment(vertical="center", wrap_text=True)
    ws[f"G{r}"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    for c in "EFH":
        ws[f"{c}{r}"].number_format = EURO
        ws[f"{c}{r}"].alignment = Alignment(horizontal="right", vertical="center")
    ws[f"H{r}"].font = F(bold=True)

v = DataValidation(type="list", formula1="=ListePrestations", allow_blank=True, errorStyle="warning",
                   errorTitle="Prestation hors BPU",
                   error="Cette prestation n'est pas dans le BPU. Continuer pour la saisir quand même "
                         "(indiquer alors un PU HT sur devis).",
                   promptTitle="Prestation", prompt="Choisir dans la liste (ou taper le code pour la retrouver).")
v.add(f"B{L1}:B{LN}")
m = DataValidation(type="list", formula1="=ListeMajorations", allow_blank=True,
                   promptTitle="Majoration", prompt="Facultatif : MC-07 ou MC-08 multiplie le montant de la ligne.")
m.add(f"G{L1}:G{LN}")
q = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True,
                   error="Quantité positive attendue.")
q.add(f"D{L1}:D{LN}")
for x in (v, m, q):
    ws.add_data_validation(x)
ws.conditional_formatting.add(
    f"F{L1}:F{LN}",
    FormulaRule(formula=[f'AND($B{L1}<>"",NOT(ISNUMBER($E{L1})),$F{L1}="")'], fill=fill("F8CBAD")))
ws.conditional_formatting.add(
    f"H{L1}:H{LN}", FormulaRule(formula=[f'$H{L1}="PU à saisir"'], font=Font(name="Arial", color="C00000", bold=True)))

# Lignes d'exemple
pick = lambda code: next(d["lib"] for d in articles if d["code"] == code)
exemples = [(pick("RA-F1"), 2, None), (pick("MC-01"), 1, MAJ_LIB["MC-08"]), (pick("MC-T1"), 3, None),
            (pick("4.1.1"), 250, None), (pick("4.3.2"), 1, None), (pick("4.8.1"), 24, None)]
for r, (l, qte, mj) in enumerate(exemples, L1):
    ws[f"B{r}"], ws[f"D{r}"], ws[f"G{r}"] = l, qte, mj

# Totaux
T = LN + 2
for i, (lab, f) in enumerate([
        ("Total HT", f"=SUM(H{L1}:H{LN})"),
        ('=IF(Autoliq="Oui","TVA (autoliquidation)","TVA "&TauxTVA*100&" %")',
         f'=IF(Autoliq="Oui",0,ROUND(H{T}*TauxTVA,2))'),
        ("Total TTC", f"=H{T}+H{T + 1}")]):
    r = T + i
    ws[f"G{r}"], ws[f"H{r}"] = lab, f
    ws[f"G{r}"].font, ws[f"G{r}"].alignment = F(bold=True), DROITE
    ws[f"H{r}"].font, ws[f"H{r}"].number_format, ws[f"H{r}"].border = F(bold=True), EURO, CADRE
for c in "GH":
    ws[f"{c}{T + 2}"].fill = fill(BLEU)
    ws[f"{c}{T + 2}"].font = F(bold=True, color="FFFFFF", size=11)

notes = [
    (f'=IF(COUNTIF(H{L1}:H{LN},"PU à saisir")>0,"Attention : "&COUNTIF(H{L1}:H{LN},"PU à saisir")'
     f'&" ligne(s) sur devis sans prix, à compléter dans la colonne « PU HT sur devis ».","")',
     F(bold=True, color="C00000")),
    (f'=IF(COUNTA(G{L1}:G{LN})=0,"","Majorations BPU appliquées au montant de la ligne : '
     f'MC-07 × "&FIXED(Majorations!$C$2,4)&" ; MC-08 × "&FIXED(Majorations!$C$3,4)&".")', F(italic=True)),
    ('=IF(Autoliq="Oui","Autoliquidation : TVA due par le preneur (article 283-2 nonies du CGI).","")',
     F(italic=True))]
for i, (f, fo) in enumerate(notes):
    r = T + i
    ws.merge_cells(f"A{r}:F{r}")
    ws[f"A{r}"], ws[f"A{r}"].font, ws[f"A{r}"].alignment = f, fo, GAUCHE

# Conditions
C0 = T + 4
bandeau(f"A{C0}:H{C0}", "CONDITIONS")
conditions = [
    "Prix unitaires HT issus du BPU Moselle Fibre – RHD57 (Maintenance et VDR Optique).",
    "Hors fourniture, sauf mention contraire du BPU.",
    "Conditions de paiement : à préciser.",
    '="Devis valable jusqu\'au "&IF(H5="","…",DAY(H5)&"/"&RIGHT("0"&MONTH(H5),2)&"/"&YEAR(H5))&"."']
for i, t in enumerate(conditions, C0 + 1):
    ws.merge_cells(f"A{i}:H{i}")
    if t.startswith("="):
        ws[f"A{i}"], ws[f"A{i}"].font, ws[f"A{i}"].alignment = t, F(), GAUCHE
    else:
        saisie(f"A{i}", t)

# Signatures
S0 = C0 + len(conditions) + 2
for rng_, t in [(f"A{S0}:B{S0}", "Pour l'entreprise – date, nom et signature"),
                (f"E{S0}:H{S0}", "Bon pour accord du client – date, nom, signature et cachet")]:
    ws.merge_cells(rng_)
    x = ws[rng_.split(":")[0]]
    x.value, x.font, x.alignment = t, F(bold=True), GAUCHE
for rng_ in [f"A{S0 + 1}:B{S0 + 5}", f"E{S0 + 1}:H{S0 + 5}"]:
    ws.merge_cells(rng_)
    for row in ws[rng_]:
        for x in row:
            x.border = CADRE
FIN = S0 + 5

# Paramètres et mode d'emploi (hors zone d'impression)
ws.merge_cells("L1:N1")
ws["L1"], ws["L1"].font, ws["L1"].fill = "PARAMÈTRES", F(bold=True, color="FFFFFF"), fill(ORANGE)
params = [("Taux de TVA", 0.2), ("Autoliquidation TVA", "Non")]
for i, (lab, val) in enumerate(params, 2):
    ws[f"L{i}"] = lab
    ws[f"L{i}"].font = F(bold=True)
    saisie(f"N{i}", val).border = CADRE
ws["N2"].number_format = "0%"
o = DataValidation(type="list", formula1='"Non,Oui"', allow_blank=False)
o.add("N3")
ws.add_data_validation(o)
ws["N3"].comment = Comment("Oui : TVA à 0 et mention d'autoliquidation (sous-traitance de travaux). "
                           "À valider avec votre comptable.", "Devis")

ws.merge_cells("L6:N6")
ws["L6"], ws["L6"].font, ws["L6"].fill = "COMMENT REMPLIR", F(bold=True, color="FFFFFF"), fill(ORANGE)
aide = [
    "Cases jaunes = à remplir. Tout le reste se calcule.",
    "Prestation : choisir dans la liste (code · désignation). Unité et prix viennent de l'onglet BPU.",
    "Qté : en ml, unités, heures, fibres… selon l'unité affichée.",
    "PU HT sur devis : seulement pour les lignes « Sur devis » (case orange) ou pour forcer un prix.",
    "Majoration : MC-07 (technicien supplémentaire) ou MC-08 (HNO) multiplie le montant de la ligne "
    "par le coefficient du BPU (onglet Majorations).",
    "Les 6 lignes remplies sont un exemple : effacer les colonnes Prestation, Qté et Majoration.",
    f"{NB_LIGNES} lignes prévues. Pour en ajouter : insérer des lignes au milieu du tableau, "
    "puis recopier une ligne remplie vers le bas.",
    "Seules les colonnes A à H s'impriment (paysage, une page de large)."]
for i, t in enumerate(aide, 7):
    # Texte sur une ligne qui déborde à droite : ne change pas la hauteur des lignes du devis.
    ws[f"L{i}"], ws[f"L{i}"].font = "• " + t, F()

# Mise en page
ws.print_area = f"A1:H{FIN}"
ws.print_title_rows = f"{H0}:{H0}"
ws.page_setup.orientation = "landscape"
ws.page_setup.paperSize = ws.PAPERSIZE_A4
ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.page_margins.left = ws.page_margins.right = 0.4
ws.page_margins.top = ws.page_margins.bottom = 0.5
ws.oddFooter.center.text = "Page &P / &N"
ws.oddFooter.center.font = "Arial"
bpu.print_area = f"A1:I{N}"
bpu.print_title_rows = "1:1"
for s in (bpu, maj):
    s.page_setup.orientation = "landscape"
    s.page_setup.fitToWidth, s.page_setup.fitToHeight = 1, 0
    s.sheet_properties.pageSetUpPr.fitToPage = True

wb.save(OUT)
print(f"{len(articles)} prestations ({sum(d['bpu'] == 'Maintenance' for d in articles)} maintenance, "
      f"{sum(d['bpu'] == 'VDR Optique' for d in articles)} optique), {len(majorations)} majorations -> {OUT}")
