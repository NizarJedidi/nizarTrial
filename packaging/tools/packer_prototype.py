"""Prototype Python du moteur de packaging (meme logique que vba/modPackaging.bas).
Usage : python3 packer_prototype.py <classeur.xlsx>"""
import itertools, time, sys
import openpyxl

DEFAULT_TIP_OVER = False   # valeur par défaut de 'Couchable?' (6 orientations si Oui, sinon 2 : rotation à plat)
EPS = 1e-9

class Item:
    __slots__=("l","w","h","vol","ref","done","placed","tip")
    def __init__(s,l,w,h,ref,tip=DEFAULT_TIP_OVER):
        s.l,s.w,s.h,s.ref,s.tip=l,w,h,ref,tip; s.vol=l*w*h; s.done=False; s.placed=False

def orientations(it):
    if it.tip:
        return [(it.l,it.w,it.h),(it.w,it.l,it.h),(it.l,it.h,it.w),(it.h,it.l,it.w),(it.w,it.h,it.l),(it.h,it.w,it.l)]
    return [(it.l,it.w,it.h),(it.w,it.l,it.h)]


STRATEGIES=[(0,0,0),(0,1,0),(1,0,0),(1,1,0),(0,0,1),(0,1,1)]
MAXSP=[0]
BIG_N=150  # (sortMode, orientPref, spaceKeyMode)

def sorted_items(items, sortMode):
    if sortMode==0: return sorted(items,key=lambda it:-it.vol)
    return sorted(items,key=lambda it:(-max(it.l,it.w,it.h),-it.vol))

def pack(items, BL, BW, BH, strat=(0,0,0)):
    """Maximal-spaces heuristic. Marque items[i].placed. Retourne (all_placed, placed_vol)."""
    sortMode,orientPref,keyMode=strat
    spaces=[(0.0,0.0,0.0,BL,BW,BH)]
    allp=True; pvol=0.0
    for it in items: it.placed=False
    for it in sorted_items(items,sortMode):
        if it.done: continue
        best=None
        for (sx,sy,sz,sl,sw,sh) in spaces:
            if best is not None and keyMode==0: break   # espaces triés (z,y,x): le premier qui convient est le meilleur
            for (il,iw,ih) in orientations(it):
                if il<=sl+EPS and iw<=sw+EPS and ih<=sh+EPS:
                    op = ih if orientPref==0 else (il*iw)
                    if keyMode==0: key=(sz,sy,sx,op)
                    else: key=(sl*sw*sh - il*iw*ih, sz, sy, sx, op)   # best fit (moins de volume restant dans l'espace)
                    if best is None or key<best[0]:
                        best=(key,(sx,sy,sz,il,iw,ih))
        if best is None:
            allp=False; continue
        px,py,pz,pl,pw,ph=best[1]
        it.placed=True; pvol+=it.vol
        new=[]
        for (sx,sy,sz,sl,sw,sh) in spaces:
            if px>=sx+sl-EPS or px+pl<=sx+EPS or py>=sy+sw-EPS or py+pw<=sy+EPS or pz>=sz+sh-EPS or pz+ph<=sz+EPS:
                new.append((sx,sy,sz,sl,sw,sh)); continue
            if px-sx>EPS: new.append((sx,sy,sz,px-sx,sw,sh))
            if sx+sl-(px+pl)>EPS: new.append((px+pl,sy,sz,sx+sl-(px+pl),sw,sh))
            if py-sy>EPS: new.append((sx,sy,sz,sl,py-sy,sh))
            if sy+sw-(py+pw)>EPS: new.append((sx,py+pw,sz,sl,sy+sw-(py+pw),sh))
            if pz-sz>EPS: new.append((sx,sy,sz,sl,sw,pz-sz))
            if sz+sh-(pz+ph)>EPS: new.append((sx,sy,pz+ph,sl,sw,sz+sh-(pz+ph)))
        pruned=[]
        for i,a in enumerate(new):
            contained=False
            for j,b in enumerate(new):
                if i==j: continue
                if (a[0]>=b[0]-EPS and a[1]>=b[1]-EPS and a[2]>=b[2]-EPS and
                    a[0]+a[3]<=b[0]+b[3]+EPS and a[1]+a[4]<=b[1]+b[4]+EPS and a[2]+a[5]<=b[2]+b[5]+EPS):
                    if a==b and j<i: contained=True; break
                    if a!=b: contained=True; break
            if not contained: pruned.append(a)
        pruned.sort(key=lambda s:(s[2],s[1],s[0]))
        spaces=pruned
        MAXSP[0]=max(MAXSP[0],len(spaces))
    return allp, pvol

def pack_best(items, BL, BW, BH):
    """Essaie toutes les stratégies ; retourne (all, vol) de la meilleure, et laisse items.placed sur la meilleure."""
    bestv=-1; bestflags=None; bestall=False
    n=sum(1 for it in items if not it.done)
    for st in STRATEGIES:
        if n>BIG_N and st[2]==1: continue
        allp,pv=pack(items,BL,BW,BH,st)
        if pv>bestv+EPS:
            bestv=pv; bestall=allp; bestflags=[it.placed for it in items]
        if allp: break
    for it,f in zip(items,bestflags): it.placed=f
    return bestall,bestv

class Caisse:
    def __init__(s,idx,name,season,l,w,h,poids):
        s.idx,s.name,s.season,s.l,s.w,s.h,s.poids=idx,name,season,l,w,h,poids; s.vol=l*w*h

def season_kind(txt):
    t=(txt or "").strip().lower()
    return "summer" if ("summer" in t or "été" in t or "ete" in t) else "all"

def suggest(items, caisses, margin):
    """items: liste d'Item (déjà triés vol desc). caisses: filtrées saison. margin: fraction."""
    for it in items: it.done=False
    remaining=[it for it in items]
    usable=[(c, c.l, c.w, c.h*(1-margin), c.l*c.w*c.h*(1-margin)) for c in caisses]
    usable.sort(key=lambda u:u[4])
    counts={}
    guard=0
    while any(not it.done for it in items):
        guard+=1
        if guard>200: return None,"Trop de caisses (>200)"
        rem_vol=sum(it.vol for it in items if not it.done)
        # 1) une seule caisse suffit ? -> la plus petite
        chosen=None
        for c,ul,uw,uh,uv in usable:
            if rem_vol>uv+EPS: continue
            allp,_=pack_best(items,ul,uw,uh)
            if allp: chosen=c; break
        if chosen:
            counts[chosen.idx]=counts.get(chosen.idx,0)+1
            for it in items:
                if not it.done and it.placed: it.done=True
            break
        # 2) sinon: caisse qui place le plus de volume (égalité -> plus petite caisse)
        bestc=None; bestv=-1; bestplaced=None
        for c,ul,uw,uh,uv in usable:
            _,pv=pack_best(items,ul,uw,uh)
            if pv>bestv+EPS:
                bestv=pv; bestc=c; bestplaced=[it.placed for it in items]
        if bestv<=EPS:
            big=[it.ref for it in items if not it.done]
            return None,"Article trop grand pour toutes les caisses: "+big[0]
        counts[bestc.idx]=counts.get(bestc.idx,0)+1
        for it,p in zip(items,bestplaced):
            if not it.done and p: it.done=True
    return counts,None

def fmt(counts, caisses_by_idx):
    return "_".join(f"{counts[i]}:{caisses_by_idx[i].name}" for i in sorted(counts))

# ---------- chargement ----------
def load(path):
    wb=openpyxl.load_workbook(path,data_only=True)
    prods={}
    ws=wb["Produits"]
    hdr=[str(c.value or "") for c in ws[1]]
    ctip=next((i for i,h in enumerate(hdr) if h.lower().startswith("couchable")),None)
    for r in ws.iter_rows(min_row=2,values_only=True):
        if r[0] is None: continue
        sku=str(r[0]).strip()
        tip=DEFAULT_TIP_OVER
        if ctip is not None and r[ctip] is not None:
            tip=str(r[ctip]).strip().lower() in ("oui","yes","o","y","1","true","vrai")
        subs=[]
        for k in range(3):
            base=4+5*k
            name,q,l,w,h=r[base:base+5]
            if name is None and q is None: continue
            subs.append((name,q,l,w,h))
        prods[sku]=dict(name=r[2],saison=(r[3] or ""),subs=subs,tip=tip)
    caisses=[]
    ws=wb["Caisses"]
    for i,r in enumerate(ws.iter_rows(min_row=2,values_only=True),start=1):
        if r[0] is None: continue
        caisses.append(Caisse(i,r[0],r[1],float(r[3]),float(r[4]),float(r[5]),r[2]))
    return prods,caisses

def build_items(prods, lines):
    """lines: [(sku, qty)] -> items, seasonal(bool), err"""
    items=[]; seasonal=False
    for sku,qty in lines:
        p=prods.get(str(sku).strip())
        if p is None: return None,None,f"SKU inconnu: {sku}"
        if str(p["saison"]).strip().lower()=="oui": seasonal=True
        for (name,q,l,w,h) in p["subs"]:
            if l is None or w is None or h is None or q is None:
                return None,None,f"Dimensions manquantes pour SKU {sku}"
            for _ in range(int(q)*int(qty)):
                items.append(Item(float(l),float(w),float(h),f"{sku}/{name}",p["tip"]))
    items.sort(key=lambda it:-it.vol)
    return items,seasonal,None

def run_lines(prods,caisses,lines,mA,mS):
    items,seasonal,err=build_items(prods,lines)
    if err: return err,err
    byidx={c.idx:c for c in caisses}
    allc=[c for c in caisses if season_kind(c.season)=="all"]
    sumc=[c for c in caisses if season_kind(c.season)=="summer"]
    cA,e=suggest(items,allc,mA); resA=fmt(cA,byidx) if cA else e
    if seasonal:
        cS,e=suggest(items,sumc,mS); resS=fmt(cS,byidx) if cS else e
    else:
        resS=resA
    return resA,resS

if __name__=="__main__":
    prods,caisses=load(sys.argv[1])
    for c in caisses: print(c.idx,c.name,c.season,c.l,c.w,c.h)
    print()
    t=time.time()
    for sku in ["9004","9124","9130","9475","9081","9148","9150","9185","9491","243"]:
        for q in [1,2,5,10]:
            t0=time.time()
            a,s=run_lines(prods,caisses,[(sku,q)],0.05,0.20)
            print(f"{sku:>5} x{q:<3} | {a:45s} | {s:45s} | {time.time()-t0:.2f}s")
    print("\nOrders:")
    for oid,lines in {"1001":[("9290",10),("9130",2),("9124",3)],"1001b":[("9004",10),("9130",2),("9124",3)],"1002":[("9491",5),("9004",1)],"1003":[("9124",5),("9130",1)]}.items():
        a,s=run_lines(prods,caisses,lines,0.05,0.20)
        print(oid,"|",a,"|",s)
    print("total",time.time()-t, "max spaces", MAXSP[0])
