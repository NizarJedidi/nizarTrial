"""Compare le référentiel entrepôt (max d'articles par caisse, X = caisse interdite) à l'algorithme.

Usage : python3 compare_entrepot.py <classeur.xlsx|xlsm> [marge_all] [marge_ete] [jeu_cm]
Le référentiel est lu dans la feuille 'Référentiel Entrepôt' du classeur (sinon tests/referentiel_entrepot.csv).
Pour chaque catégorie, le 1er SKU candidat est utilisé ; les colonnes 'Couchable?' / 'Caisses interdites' s'appliquent.
"""
import csv, sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import packer_prototype as pk
import openpyxl

wb = sys.argv[1]
mA = float(sys.argv[2]) if len(sys.argv) > 2 else 0.05
mS = float(sys.argv[3]) if len(sys.argv) > 3 else 0.20
cl = float(sys.argv[4]) if len(sys.argv) > 4 else 0.5
prods, caisses = pk.load(wb)

def read_ref(wb):
    book = openpyxl.load_workbook(wb, data_only=True)
    if "Référentiel Entrepôt" in book.sheetnames:
        ws = book["Référentiel Entrepôt"]
        hdr = [c.value for c in ws[1]]
        rows = []
        for r in ws.iter_rows(min_row=2, values_only=True):
            if not r[0] or r[0].startswith("Lecture") or r[0].startswith("Cellules"): continue
            d = {"categorie": str(r[0]), "sku_candidats": str(r[1] or "")}
            for i, h in enumerate(hdr):
                if h and h.startswith(("Caisse", "ETE")):
                    v = r[i]
                    if v is None: d[h] = ""
                    elif str(v).strip().upper() == "X": d[h] = "X"
                    else:
                        import re as _re; nums = _re.findall(r"\d+", str(v)); d[h] = nums[-1] if nums else ""
            rows.append(d)
        return rows
    ref = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "tests", "referentiel_entrepot.csv")
    return list(csv.DictReader(open(ref, encoding="utf-8"), delimiter=";"))
by_name = {c.name: c for c in caisses}

def max_in_box(sku, c, margin):
    items_all, _, err = pk.build_items(prods, [(sku, 1)])
    if err: return err
    ul, uw, uh = c.l - cl, c.w - cl, c.h * (1 - margin) - cl
    if min(ul, uw, uh) <= 0: return 0
    unit_vol = sum(it.vol for it in items_all)
    hi = int(ul * uw * uh // unit_vol) if unit_vol > 0 else 0
    best = 0
    for q in range(1, hi + 1):
        items, _, _ = pk.build_items(prods, [(sku, q)])
        for it in items: it.done = False
        allp, _ = pk.pack_best(items, ul, uw, uh)
        if allp: best = q
        else: break
    return best

rows = read_ref(wb)
boxes = [k for k in rows[0].keys() if k not in ("categorie", "sku_candidats", "couchable")]
print(f"Réglages : marge All Seasons {mA:.0%}, marge Summer {mS:.0%}, jeu {cl} cm. Format : entrepôt / algorithme (X = interdite).")
print()
print("| Catégorie | SKU testé | " + " | ".join(b.replace("Caisse ", "") for b in boxes) + " |")
print("|---|---|" + "---|" * len(boxes))
n_ok = n_tot = 0
for r in rows:
    skus = [s for s in r["sku_candidats"].split("|") if s and s != "?"]
    cells = []
    for b in boxes:
        want = r[b].strip()
        if not skus:
            cells.append(f"{want} / ?"); continue
        c = by_name[b]
        got = max_in_box(skus[0], c, mS if pk.season_kind(c.season) == "summer" else mA)
        if isinstance(got, str): got = "err"
        if want == "" : cells.append(f"- / {got}"); continue
        mark = "" if (want == "X" or str(got) == want) else (" ⚠" if want != "X" and got < int(want) else " ✦")
        if want != "X":
            n_tot += 1; n_ok += (mark == "")
        cells.append(f"{want} / {got}{mark}")
    name = prods[skus[0]]["name"] if skus else "(à identifier)"
    print(f"| {r['categorie']} | {skus[0] + ' ' + name if skus else '?'} | " + " | ".join(cells) + " |")
print()
print(f"Correspondances exactes (hors X et inconnus) : {n_ok} / {n_tot}")
print("⚠ = l'algorithme en met moins que l'entrepôt ; ✦ = il en met plus. Les dimensions des sous-articles viennent de la feuille Produits.")
