# Remplit le modèle de devis Circet (Exemple_devis_LOSANGE.xlsx) pour un client : on garde l'en-tête,
# le tableau, le logo Circet et la signature du modèle, et on remplace le BPU, l'en-tête et le mode
# d'emploi. Utilisé par axione.py et celeste.py.
import math, re
from copy import copy
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule

L1, LN = 21, 32  # lignes d'articles du devis modèle
COEF = "Coef."   # unité des lignes de majoration


def propre(t):
    return re.sub(r"[ \t]+", " ", str(t or "")).strip()


def lignes(t):
    return [propre(l) for l in str(t or "").split("\n") if propre(l)]


def prix(v):
    # Les BPU écrivent « sur devis », « Devis », « devis BPU »… : une seule écriture dans le devis.
    if v is None or isinstance(v, (int, float)):
        return v
    return "Sur devis"


def libelle(code, desig, n=95):
    if len(desig) > n:
        desig = desig[:n].rsplit(" ", 1)[0].rstrip(" ,;-–") + "…"
    return f"{code} · {desig}" if code else desig


def remplir(modele, sortie, *, onglet, titre_bpu, catalogue, entete, client, sous_titre, aide):
    """catalogue : dicts categorie, article, unite, prix, unite_bpu, descriptif, source (unite == COEF pour
    une majoration, prix = coefficient). entete : {cellule de l'onglet Devis: texte}. aide : (« section »
    ou « texte », texte) propres au client, placés avant les rubriques communes du mode d'emploi."""
    arts = [d["article"] for d in catalogue]
    assert len(arts) == len(set(arts)), "articles en double"
    assert all(d["prix"] is not None for d in catalogue)

    wb = openpyxl.load_workbook(modele)
    dv, bpu, mode = wb["Devis"], wb["BPU 2026"], wb["Mode d'emploi"]

    # ---------- Onglet BPU : même présentation que « BPU 2026 » ----------
    bpu.title = onglet
    fin = 3 + len(catalogue)
    bpu["A1"] = titre_bpu
    bpu["A2"] = ("Catalogue de référence — sert de source aux menus déroulants de l'onglet \"Devis\". "
                 "Ne pas modifier l'ordre des colonnes B, C et D.")
    ent = {c: copy(bpu[f"{c}3"]._style) for c in "ABCD"}
    cel = {c: copy(bpu[f"{c}4"]._style) for c in "ABCD"}
    bpu.delete_rows(4, bpu.max_row)
    for c, t, w in [("A", "Catégorie", 48), ("B", "Article", 95), ("C", "Unité", 10), ("D", "PU HT", 14),
                    ("E", "Unité BPU", 26), ("F", "Descriptif", 70), ("G", "Source", 46)]:
        bpu[f"{c}3"] = t
        bpu[f"{c}3"]._style = copy(ent.get(c, ent["B"]))
        bpu.column_dimensions[c].width = w
    for i, d in enumerate(catalogue, 4):
        vals = {"A": d["categorie"], "B": d["article"], "C": d["unite"], "D": d["prix"],
                "E": d["unite_bpu"], "F": d["descriptif"], "G": d["source"]}
        for c, v in vals.items():
            bpu[f"{c}{i}"] = v
            bpu[f"{c}{i}"]._style = copy(cel.get(c, cel["B"]))
        if d["unite"] == COEF:
            bpu[f"D{i}"].number_format = "0.0000"
    bpu.auto_filter.ref = f"A3:G{fin}"
    bpu.freeze_panes = "A4"

    # ---------- Onglet Devis : en-tête du client ----------
    for ref, t in entete.items():
        dv[ref] = t
    # Le logo Losange laisse la place au nom du client (le logo Circet et la signature restent).
    dv._images = [im for im in dv._images if im.anchor._from.col != 7]
    dv.merge_cells("H2:K4")
    dv["H2"] = client
    dv["H2"].font = Font(name="Aptos Narrow", size=22, bold=True, color="1F497D")
    dv["H2"].alignment = Alignment(horizontal="right", vertical="center")
    dv.merge_cells("H5:K5")
    dv["H5"] = sous_titre
    dv["H5"].font = Font(name="Aptos Narrow", size=10, italic=True, color="808080")
    dv["H5"].alignment = Alignment(horizontal="right", vertical="center")

    # Lignes d'articles : unité et prix depuis le BPU. Un article « Sur devis » (ou hors BPU) laisse
    # le P.U.H.T à taper. Une majoration vaut P.U.H.T de la ligne au-dessus × (coef − 1).
    plage = lambda c: f"'{onglet}'!${c}$4:${c}${fin}"
    for r in range(L1, LN + 1):
        m = f"MATCH($B{r},{plage('B')},0)"
        dv[f"H{r}"] = f'=IF($B{r}="","",IFERROR(INDEX({plage("C")},{m})&"",""))'
        dv[f"I{r}"] = (f'=IF($B{r}="","",IFERROR(IF(INDEX({plage("C")},{m})="{COEF}",'
                       f'ROUND(I{r - 1}*(INDEX({plage("D")},{m})-1),2),ROUND(INDEX({plage("D")},{m}),2)),"Sur devis"))')
        dv[f"K{r}"] = f'=IF($B{r}="","",IF(ISNUMBER(I{r}),ROUND(I{r}*N(J{r}),2),"PU à saisir"))'
        dv[f"H{r}"]._style = copy(dv[f"H{L1 + 1}"]._style)
        dv[f"H{r}"].alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
    v = DataValidation(type="list", formula1=plage("B"), allow_blank=True, errorStyle="warning",
                       errorTitle="Article hors BPU",
                       error="Cet article n'est pas dans le BPU. Continuer pour le saisir quand même, "
                             "puis taper le P.U.H.T.")
    v.add(f"B{L1}:B{LN}")
    dv.add_data_validation(v)
    dv.conditional_formatting.add(f"I{L1}:I{LN}", FormulaRule(formula=[f'I{L1}="Sur devis"'],
                                                             fill=PatternFill("solid", fgColor="FFFF00")))
    dv.conditional_formatting.add(f"K{L1}:K{LN}", FormulaRule(formula=[f'K{L1}="PU à saisir"'],
                                                             font=Font(color="C00000", bold=True)))

    # ---------- Mode d'emploi ----------
    st = {k: copy(mode[ref]._style) for k, ref in [("titre", "A1"), ("section", "A3"), ("texte", "A4")]}
    mode.delete_rows(1, mode.max_row)
    texte = [("titre", "Comment utiliser ce classeur")] + aide + [
        ("section", "Le total"),
        ("texte", "\"Sous Total HT\" additionne automatiquement toutes les lignes remplies."),
        ("section", "Pour un nouveau devis"),
        ("texte", "Clic droit sur l'onglet \"Devis\" → Déplacer ou copier → cocher \"Créer une copie\" → renomme "
                  "l'onglet avec le n° de ticket ou de commande. Ça garde une trace de chaque devis dans le même "
                  f"classeur, avec l'onglet \"{onglet}\" partagé par tous."),
        ("section", "Mettre à jour les prix"),
        ("texte", f"Si le BPU change, modifie uniquement l'onglet \"{onglet}\" (colonnes C et D) — tous les devis "
                  "qui utilisent un article existant se mettent à jour tout seuls.")]
    r, n = 1, 0
    for k, t in texte:
        if k == "section":
            n += 1
            t = f"{n}. {t}"
            if n > 1:
                mode.row_dimensions[r].height = 9.75
                r += 1
        mode.cell(r, 1, t)._style = copy(st[k])
        mode.row_dimensions[r].height = 24 if k != "texte" else 15 * math.ceil(len(t) / 105) + 6
        r += 1
        if k == "titre":
            mode.row_dimensions[r].height = 9.75
            r += 1

    wb.save(sortie)
