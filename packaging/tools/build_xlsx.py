"""Construit packaging_engine_mms_v2.xlsx : classeur d'origine + feuilles Packaging Calculator et Mixed Pack Helper."""
import sys, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

src, dst = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(src)
bold = Font(name="Arial", bold=True)
normal = Font(name="Arial")
hdr_fill = PatternFill("solid", fgColor="DDEBF7")
inp_fill = PatternFill("solid", fgColor="FFFFCC")
thin = Side(style="thin", color="999999")
box = Border(left=thin, right=thin, top=thin, bottom=thin)

def style_hdr(ws, cells):
    for c in cells:
        ws[c].font = bold; ws[c].fill = hdr_fill; ws[c].border = box
        ws[c].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

# ---------- Packaging Calculator ----------
if "Packaging Calculator" in wb.sheetnames: del wb["Packaging Calculator"]
ws = wb.create_sheet("Packaging Calculator")
ws["A1"] = "Marge vide All Seasons"; ws["B1"] = 0.05
ws["A2"] = "Marge vide Summer";      ws["B2"] = 0.20
for c in ("A1", "A2"): ws[c].font = bold
for c in ("B1", "B2"):
    ws[c].number_format = "0%"; ws[c].fill = inp_fill; ws[c].border = box; ws[c].font = Font(name="Arial", color="0000FF")
ws["C1"] = "← % du volume de la caisse laissé vide (marge en haut). Modifiable."
ws["C2"] = "← plus large en été (gel / conservation). Modifiable."
ws["C1"].font = Font(name="Arial", italic=True, color="666666"); ws["C2"].font = Font(name="Arial", italic=True, color="666666")
ws["A4"] = "Item"; ws["B4"] = "Quantity"; ws["C4"] = "Suggested Package All Seasons"; ws["D4"] = "Suggested Package Summer"; ws["E4"] = "Info"
style_hdr(ws, ["A4", "B4", "C4", "D4", "E4"])
# lignes d'exemple (SKU de la feuille Produits)
examples = [("9004", 1), ("9004", 10), ("9124", 3), ("9130", 2), ("9081", 2), ("9150", 5), ("9491", 1)]
for i, (sku, q) in enumerate(examples, start=5):
    ws.cell(i, 1, sku).font = normal; ws.cell(i, 2, q).font = normal
    for c in range(1, 6): ws.cell(i, c).border = box
ws["A3"] = "Mode d'emploi : SKU (ou nom exact du produit) en colonne A, quantité en B, puis bouton « Remplir ». C, D, E sont remplies par la macro (format qté:caisse, séparateur « _ »)."
ws["A3"].font = Font(name="Arial", italic=True, color="666666")
ws.column_dimensions["A"].width = 24; ws.column_dimensions["B"].width = 11
ws.column_dimensions["C"].width = 36; ws.column_dimensions["D"].width = 36; ws.column_dimensions["E"].width = 60
ws.freeze_panes = "A5"

# ---------- Mixed Pack Helper ----------
if "Mixed Pack Helper" in wb.sheetnames: del wb["Mixed Pack Helper"]
ws = wb.create_sheet("Mixed Pack Helper")
for col, h in zip("ABC", ["Order ID", "Item", "Quantity"]): ws[f"{col}1"] = h
for col, h in zip("EFGH", ["Order ID", "Suggested Package All Seasons", "Suggested Package Summer", "Info"]): ws[f"{col}1"] = h
style_hdr(ws, ["A1", "B1", "C1", "E1", "F1", "G1", "H1"])
rows = [(1001, "9004", 10), (1001, "9130", 2), (1001, "9124", 3), (1002, "9491", 5), (1002, "9004", 1), (1003, "9124", 5), (1003, "9130", 1)]
for i, (o, s, q) in enumerate(rows, start=2):
    ws.cell(i, 1, o); ws.cell(i, 2, s); ws.cell(i, 3, q)
    for c in range(1, 4): ws.cell(i, c).border = box; ws.cell(i, c).font = normal
ws["J4"] = "Mode d'emploi : une ligne par article de la commande (Order ID, SKU, quantité)."
ws["J5"] = "Bouton « Remplir » : un résultat par commande en E:H (marges lues dans Packaging Calculator)."
ws["J4"].font = Font(name="Arial", italic=True, color="666666"); ws["J5"].font = Font(name="Arial", italic=True, color="666666")
for col, w in zip("ABCDEFGH", [12, 14, 11, 3, 12, 36, 36, 60]): ws.column_dimensions[col].width = w
ws.freeze_panes = "A2"

# police Arial sur les nouvelles feuilles (cellules non stylées)
for name in ("Packaging Calculator", "Mixed Pack Helper"):
    for row in wb[name].iter_rows():
        for c in row:
            if c.value is not None and c.font.name != "Arial":
                c.font = Font(name="Arial", bold=c.font.bold, italic=c.font.italic, color=c.font.color)
wb.save(dst)
print("saved", dst)
