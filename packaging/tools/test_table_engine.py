"""Tests du moteur 'Referentiel Entrepot' (prototype Python, meme logique que le VBA).
Usage : python3 test_table_engine.py <classeur.xlsm>"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import packer_prototype as pk
wb=sys.argv[1]
prods, caisses = pk.load(wb); ref = pk.load_ref(wb, caisses)
print(f"{len(ref)} categories chargees :", [r["cat"] for r in ref])
for row in ref:
    sku=sorted(row["skus"])[0]
    if sku not in prods: print(f"\n== {row['cat']} : SKU {sku} absent de Produits"); continue
    mx=max(v for v in row["max"].values())
    qs=sorted({1,2,3,4,5,6,7,8,10,12,13,15,20,21,25,30,50,51,60,70,71,80,81,100,120,121,150,151,160,200,300} & set(range(1,2*mx+2)))
    print(f"\n== {row['cat']} (SKU {sku}) ; plages hiver: " + ", ".join(f"{b.replace('Caisse ','')} {row['min'][b]}-{row['max'][b]}" for b in row['max'] if row['max'][b]>=1 and not b.startswith('ETE')) +
          " ; ete: " + ", ".join(f"{b.replace('ETE Caisse ','')} {row['min'][b]}-{row['max'][b]}" for b in row['max'] if row['max'][b]>=1 and b.startswith('ETE')))
    for q in qs:
        a,s=pk.run_lines_ref(prods,caisses,ref,[(sku,q)],0.05,0.27)
        print(f"  x{q:<4} hiver: {a:34s} ete: {s}")
print("\n== Commandes mixtes")
for lines in [[("9004",3),("9122",2)],[("243",75),("9122",6)],[("243",100),("9122",6)],[("9004",1),("9124",1)],[("9124",2),("9130",1)],[("9290",20),("9004",2)],[("9203",2),("9130",1),("9004",1)]]:
    a,s=pk.run_lines_ref(prods,caisses,ref,lines,0.05,0.27)
    print(f"  {lines} -> hiver: {a} | ete: {s}")
