"""Style et briques communes à tous les modèles de la boutique.

Tous les produits partagent les mêmes couleurs, la même police et la même
page « Mode d'emploi » : la boutique a ainsi un style reconnaissable.
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule

FONT = "Arial"
C_DARK = "1F4E5F"
C_LIGHT = "E3EFF1"
C_GREEN = "2E8B6E"
C_GREEN_L = "E4F3EE"
C_CORAL = "D9644A"
C_CORAL_L = "FBEAE5"
C_INPUT = "FFF6E0"
C_CALC = "F3F4F6"
C_TEXT = "1F2937"
C_MUTED = "6B7280"
C_BORDER = "D1D5DB"
C_WHITE = "FFFFFF"

EUR = '#,##0.00 "€";-#,##0.00 "€";"-"'
EUR0 = '#,##0 "€";-#,##0 "€";"-"'
EUR_TILE = '#,##0 "€";-#,##0 "€";0 "€"'
PCT = "0%"
DATE = "dd/mm/yyyy"
DATE_SHORT = "dd/mm"

thin = Side(style="thin", color=C_BORDER)
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)

LEFT = Alignment(horizontal="left", vertical="center")
CENTER = Alignment(horizontal="center", vertical="center")
RIGHT = Alignment(horizontal="right", vertical="center")
WRAP = Alignment(horizontal="left", vertical="center", wrap_text=True)

MONTHS = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet",
          "Août", "Septembre", "Octobre", "Novembre", "Décembre"]
MONTHS_SHORT = ["Janv.", "Févr.", "Mars", "Avr.", "Mai", "Juin", "Juil.",
                "Août", "Sept.", "Oct.", "Nov.", "Déc."]
CHECK = "✓"


def font(size=10, bold=False, color=C_TEXT, italic=False, strike=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic, strike=strike)


def fill(color):
    return PatternFill("solid", start_color=color, end_color=color)


def put(ws, ref, value=None, *, f=None, bg=None, fmt=None, align=None, border=None):
    c = ws[ref]
    if value is not None:
        c.value = value
    c.font = f or font()
    if bg:
        c.fill = fill(bg)
    if fmt:
        c.number_format = fmt
    if align:
        c.alignment = align
    if border:
        c.border = border
    return c


def setup(ws, tab_color, widths=None):
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = tab_color
    for col, w in (widths or {}).items():
        ws.column_dimensions[col].width = w


def title(ws, ref, text, subtitle=None, sub_ref=None):
    put(ws, ref, text, f=font(18, True, C_DARK))
    ws.row_dimensions[ws[ref].row].height = 30
    if subtitle:
        put(ws, sub_ref, subtitle, f=font(9, italic=True, color=C_MUTED))


def header_row(ws, row, first_col, labels, color=C_DARK, height=22):
    for i, label in enumerate(labels):
        c = ws.cell(row=row, column=first_col + i)
        put(ws, c.coordinate, label, f=font(bold=True, color=C_WHITE), bg=color, border=BOX,
            align=Alignment(horizontal="left" if i == 0 else "center", vertical="center", wrap_text=True))
    ws.row_dimensions[row].height = height


def tile(ws, label_ref, value_ref, label, formula, color, light, merge=None, fmt=EUR_TILE):
    """Une « tuile » de chiffre clé : petit titre coloré + grand chiffre."""
    if merge:
        ws.merge_cells(merge[0])
        ws.merge_cells(merge[1])
    put(ws, label_ref, label, f=font(9, True, C_WHITE), bg=color, align=CENTER)
    put(ws, value_ref, formula, f=font(16, True, color), bg=light, align=CENTER, fmt=fmt)
    if merge:
        for rng in merge:
            for row in ws[rng]:
                for c in row:
                    c.fill = fill(color if c.row == ws[label_ref].row else light)


def unlock(ws, *refs):
    for ref in refs:
        cells = ws[ref]
        if not isinstance(cells, tuple):
            cells = ((cells,),)
        elif cells and not isinstance(cells[0], tuple):
            cells = (cells,)
        for row in cells:
            for c in row:
                c.protection = Protection(locked=False)


def list_validation(ws, rng, formula, error=None, strict=True):
    dv = DataValidation(type="list", formula1=formula, allow_blank=True,
                        showErrorMessage=bool(error), error=error,
                        errorStyle="stop" if strict else "information")
    dv.add(rng)
    ws.add_data_validation(dv)
    return dv


def check_column(ws, rng, row_range, first_row):
    """Colonne « fait » : liste déroulante ✓, et la ligne se barre quand c'est coché."""
    dv = DataValidation(type="list", formula1=f'"{CHECK}"', allow_blank=True, showErrorMessage=False)
    dv.add(rng)
    ws.add_data_validation(dv)
    col = rng.split(":")[0].rstrip("0123456789")
    ws.conditional_formatting.add(row_range, FormulaRule(formula=[f'${col}{first_row}<>""'],
                                                         font=Font(name=FONT, strike=True, color="9CA3AF")))


def fit_width(ws, landscape=False):
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def style_chart(ch):
    ch.roundedCorners = False
    if ch.legend is not None:
        ch.legend.overlay = False
    if ch.title is not None:
        ch.title.overlay = False
    ch.x_axis.delete = False
    ch.y_axis.delete = False


def build_guide(ws, name, subtitle, steps, tips):
    """Onglet « Mode d'emploi », identique d'un produit à l'autre."""
    setup(ws, C_DARK, {"A": 3, "B": 6, "C": 95})
    put(ws, "B2", name, f=font(22, True, C_DARK))
    ws.row_dimensions[2].height = 36
    put(ws, "B3", subtitle, f=font(11, color=C_MUTED))

    row = 5
    put(ws, f"B{row}", "Comment l'utiliser", f=font(13, True, C_DARK))
    for i, (tab, text) in enumerate(steps):
        r = row + 1 + i
        put(ws, f"B{r}", i + 1, f=font(14, True, C_WHITE), bg=C_DARK, align=CENTER)
        c = put(ws, f"C{r}", None, align=WRAP)
        c.value = CellRichText(TextBlock(InlineFont(rFont=FONT, b=True, sz=11, color=C_DARK), f"Onglet {tab} : "),
                               TextBlock(InlineFont(rFont=FONT, sz=10, color=C_TEXT), text))
        ws.row_dimensions[r].height = 34

    row += len(steps) + 2
    put(ws, f"B{row}", "Légende", f=font(13, True, C_DARK))
    put(ws, f"B{row + 1}", None, bg=C_INPUT, border=BOX)
    put(ws, f"C{row + 1}", "  Case à remplir", f=font(10))
    put(ws, f"B{row + 2}", None, bg=C_CALC, border=BOX)
    put(ws, f"C{row + 2}", "  Calculé automatiquement : ne pas modifier", f=font(10))
    ws.row_dimensions[row + 1].height = 20
    ws.row_dimensions[row + 2].height = 20

    row += 4
    put(ws, f"B{row}", "Bon à savoir", f=font(13, True, C_DARK))
    for i, t in enumerate(tips):
        r = row + 1 + i
        put(ws, f"B{r}", "•", f=font(12, True, C_CORAL), align=CENTER)
        put(ws, f"C{r}", t, f=font(10), align=WRAP)
        ws.row_dimensions[r].height = 32 if len(t) > 100 else 18
    fit_width(ws)


def new_workbook(product_title):
    wb = Workbook()
    wb.properties.creator = "Modèles de la boutique"
    wb.properties.title = product_title
    wb.calculation.fullCalcOnLoad = True
    return wb
