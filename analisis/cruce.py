import json, os, glob, difflib, re
from collections import OrderedDict, defaultdict

def norm(s):
    s=" ".join(str(s or "").strip().upper().split())
    for a,b in zip("ÁÉÍÓÚÜÑ","AEIOUUN"): s=s.replace(a,b)
    s=s.replace("PRODUCCCION","PRODUCCION").replace("BUSINNES","BUSINESS")
    s=re.sub(r'\bCOORDINADOR/?A?\b','COORDINADOR',s)
    s=re.sub(r'\bSUPERVISOR/?A?\b','SUPERVISOR',s)
    s=re.sub(r'\bANALISTA/?A?\b','ANALISTA',s)
    s=s.replace("/O","").replace("/A","")
    return " ".join(s.split())

hc=json.load(open("headcount_adm.json"))
EQ={}
try:
    _e=json.load(open("equivalencias.json"))
    for k,v in _e.get("equivalencias",{}).items(): EQ[k]=v["cargo_hc"]
except FileNotFoundError: pass
mx=json.load(open("matrices.json"))

# cargos evaluados -> lista de (codigo matriz, nombre tal cual)
eval_map=defaultdict(list)
for arch,hojas in mx.items():
    for h in hojas:
        cod=h.get("codigo") or arch
        amat=h.get("meta",{}).get("PUESTO DE TRABAJO A EVALUAR","?")
        if "puestos" not in h: continue
        for pu in h["puestos"]:
            k=norm(pu)
            k=norm(EQ.get(k,k))
            eval_map[k].append((cod,amat,pu))

# pares area-cargo del head count
pares=OrderedDict()
for r in hc:
    k=(r["area"], r["cargo"])
    pares.setdefault(k,{"n":0,"norm":norm(r["cargo"])})
    pares[k]["n"]+=1

print(f"{'ÁREA':<30}{'CARGO':<45}{'N':>3}  ESTADO / MATRIZ")
print("="*130)
ok=falta=0
dudas=[]
for (area,cargo),v in sorted(pares.items()):
    n=v["n"]; cn=v["norm"]
    if cn in eval_map:
        m=", ".join(sorted(set(c for c,_,_ in eval_map[cn])))
        print(f"{area:<30}{cargo:<45}{n:>3}  ✔ {m}")
        ok+=1
    else:
        cand=difflib.get_close_matches(cn, list(eval_map.keys()), n=2, cutoff=0.72)
        if cand:
            c0,amt,orig=eval_map[cand[0]][0]
            print(f"{area:<30}{cargo:<45}{n:>3}  ? posible '{orig}' ({c0})")
            dudas.append((area,cargo,orig,c0))
        else:
            print(f"{area:<30}{cargo:<45}{n:>3}  ✘ SIN MATRIZ")
        falta+=1
print("="*130)
print(f"Pares área-cargo: {len(pares)} | con matriz: {ok} | pendientes: {falta}")

# puestos en matrices que no están en head count
hcn=set(v["norm"] for v in pares.values())
print("\n=== PUESTOS EN MATRICES SIN EQUIVALENTE EXACTO EN HEAD COUNT ===")
for k,vs in sorted(eval_map.items()):
    if k not in hcn:
        cand=difflib.get_close_matches(k, list(hcn), n=1, cutoff=0.72)
        s=f" -> ¿'{cand[0]}'?" if cand else "  (sin candidato)"
        print(f"   {vs[0][2]:<45} [{vs[0][0]}]{s}")
