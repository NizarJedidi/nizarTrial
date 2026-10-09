"""Recherche des réglages qui reproduisent le plus de cases du Référentiel Entrepôt.

Usage : python3 search_config.py <classeur.xlsm> [sortie.md]
Grille : marge All Seasons, marge Summer, jeu de sécurité ; 'Couchable?' choisi librement par catégorie
(c'est un réglage par produit). Score = nombre de cases (hors X) où le max calculé = max entrepôt.
"""
import sys, os, re, itertools, functools
from multiprocessing import Pool
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import packer_prototype as pk
import openpyxl

WB = sys.argv[1]
OUT = sys.argv[2] if len(sys.argv) > 2 else None
def _grid(env, default):
    v = os.environ.get(env)
    return [float(x) for x in v.split(",")] if v else default
GRID_A = _grid("SC_GRID_A", [0.0, 0.05, 0.10, 0.15])   # ex. SC_GRID_A=0.05 pour figer la marge hiver
GRID_S = _grid("SC_GRID_S", [0.10, 0.20, 0.30, 0.40])
GRID_J = _grid("SC_GRID_J", [0.0, 0.5, 1.0])

prods, caisses = pk.load(WB)
by_name = {c.name: c for c in caisses}

def read_ref():
    ws = openpyxl.load_workbook(WB, data_only=True)["Référentiel Entrepôt"]
    hdr = [c.value for c in ws[1]]
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r[0] or str(r[0]).startswith(("Lecture", "Cellules")): continue
        d = {"cat": str(r[0]), "skus": [s.strip() for s in str(r[1] or "").split("|") if s.strip()], "cells": {}}
        for i, h in enumerate(hdr):
            if h and h.startswith(("Caisse", "ETE")) and r[i] is not None:
                v = str(r[i]).strip()
                if v.upper() == "X": d["cells"][h] = ("X", None, None)
                else:
                    nums = [int(x) for x in re.findall(r"\d+", v)]
                    if len(nums) >= 2: d["cells"][h] = ("range", nums[0] + 1, nums[1])   # ">5 & <12" -> min 6, max 12
                    elif nums: d["cells"][h] = ("max", None, nums[0])
        rows.append(d)
    return rows

def max_in_box(sku, tip, ul, uw, uh):
    """max q tel que q articles tiennent (recherche binaire, borne volume)."""
    items, _, err = pk.build_items(prods, [(sku, 1)])
    if err or min(ul, uw, uh) <= 0: return 0
    for it in items: it.tip = tip
    unit = sum(it.vol for it in items)
    hi = int(ul * uw * uh // unit)
    def fits(q):
        its, _, _ = pk.build_items(prods, [(sku, q)])
        for it in its: it.tip = tip; it.done = False
        return pk.pack_best(its, ul, uw, uh)[0]
    if hi == 0 or not fits(1): return 0
    lo = 1
    while lo < hi:
        mid = (lo + hi + 1) // 2
        if fits(mid): lo = mid
        else: hi = mid - 1
    return lo

def job(args):
    sku, tip, box, margin, jeu = args
    c = by_name[box]
    ul, uw, uh = c.l - jeu, c.w - jeu, c.h * (1 - margin) - jeu
    return (args, max_in_box(sku, tip, ul, uw, uh))

if __name__ == "__main__":
    rows = read_ref()
    jobs = set()
    for r in rows:
        if not r["skus"]: continue
        sku = r["skus"][0]
        for box in r["cells"]:
            summer = pk.season_kind(by_name[box].season) == "summer"
            for tip in (False, True):
                for m in (GRID_S if summer else GRID_A):
                    for j in GRID_J:
                        jobs.add((sku, tip, box, m, j))
    jobs = sorted(jobs)
    print(f"{len(jobs)} calculs de max par caisse...", file=sys.stderr)
    with Pool() as pool:
        res = dict(pool.imap_unordered(job, jobs, chunksize=4))

    def cell_ok(spec, got, box, qmaxref=None):
        kind, mn, mx = spec
        if kind == "X": return None
        return got == mx

    # score de chaque config globale, Couchable choisi par ligne
    results = []
    for mA, mS, j in itertools.product(GRID_A, GRID_S, GRID_J):
        total = 0; ntot = 0; choice = {}; detail = {}
        for r in rows:
            if not r["skus"]: continue
            sku = r["skus"][0]
            best = None
            for tip in (False, True):
                ok = 0; n = 0; cells = {}
                for box, spec in r["cells"].items():
                    summer = pk.season_kind(by_name[box].season) == "summer"
                    got = res[(sku, tip, box, mS if summer else mA, j)]
                    cells[box] = got
                    if spec[0] != "X":
                        n += 1; ok += (got == spec[2])
                if best is None or ok > best[0]: best = (ok, n, tip, cells)
            total += best[0]; ntot += best[1]; choice[r["cat"]] = best[2]; detail[r["cat"]] = best[3]
        results.append((total, ntot, mA, mS, j, choice, detail))
    results.sort(key=lambda x: (-x[0], x[4], x[2], x[3]))

    out = []
    out.append(f"# Recherche de réglages ({len(results)} configurations, Couchable choisi par catégorie)\n")
    out.append("| Score | Marge All | Marge Été | Jeu (cm) | Couchable = Oui pour |")
    out.append("|---|---|---|---|---|")
    for total, ntot, mA, mS, j, choice, _ in results[:15]:
        out.append(f"| {total} / {ntot} | {mA:.0%} | {mS:.0%} | {j} | {', '.join(k for k, v in choice.items() if v) or '-'} |")
    best = results[0]
    total, ntot, mA, mS, j, choice, detail = best
    out.append(f"\n## Meilleure configuration : marge All {mA:.0%}, marge Été {mS:.0%}, jeu {j} cm — {total} / {ntot}\n")
    boxes = [c.name for c in caisses]
    out.append("| Catégorie | SKU | Couchable | " + " | ".join(b.replace("Caisse ", "") for b in boxes) + " | score |")
    out.append("|---|---|---|" + "---|" * (len(boxes) + 1))
    for r in rows:
        if not r["skus"]: continue
        cells = []; ok = 0; n = 0
        for b in boxes:
            spec = r["cells"].get(b)
            got = detail[r["cat"]].get(b)
            if spec is None: cells.append("-"); continue
            if spec[0] == "X": cells.append(f"X / {got}"); continue
            want = spec[2] if spec[0] == "max" else f"{spec[1]}..{spec[2]}"
            n += 1; good = (got == spec[2]); ok += good
            cells.append(f"{want} / {got}" + ("" if good else (" ⚠" if got < spec[2] else " ✦")))
        out.append(f"| {r['cat']} | {r['skus'][0]} | {'Oui' if choice[r['cat']] else 'Non'} | " + " | ".join(cells) + f" | {ok}/{n} |")
    out.append("\nFormat : entrepôt / algorithme. ⚠ = l'algorithme en met moins, ✦ = il en met plus. X = caisse interdite (gérée par 'Caisses interdites').")
    # marges par saison independantes : meilleur score par (mA, jeu) et par (mS, jeu)
    out.append("\n## Scores par paramètre (les marges hiver et été sont indépendantes)\n")
    out.append("| Jeu | " + " | ".join(f"All {m:.0%}" for m in GRID_A) + " | " + " | ".join(f"Été {m:.0%}" for m in GRID_S) + " |")
    out.append("|---|" + "---|" * (len(GRID_A) + len(GRID_S)))
    for jj in GRID_J:
        rowA = []; rowS = []
        for m in GRID_A:
            rowA.append(max(t for t, _, a, s, j2, _, _ in results if a == m and j2 == jj and s == GRID_S[0]) if True else 0)
        for m in GRID_S:
            rowS.append(max(t for t, _, a, s, j2, _, _ in results if s == m and j2 == jj and a == GRID_A[0]))
        out.append(f"| {jj} | " + " | ".join(str(x) for x in rowA) + " | " + " | ".join(str(x) for x in rowS) + " |")
    text = "\n".join(out)
    print(text)
    if OUT: open(OUT, "w", encoding="utf-8").write(text + "\n")
