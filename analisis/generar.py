# -*- coding: utf-8 -*-
import json, re, os
from collections import OrderedDict, defaultdict, Counter
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE=os.path.dirname(os.path.abspath(__file__))
def J(n): return json.load(open(os.path.join(BASE,n),encoding="utf-8"))

def norm(s):
    s=" ".join(str(s or "").strip().upper().split())
    for a,b in zip("ÁÉÍÓÚÜÑ","AEIOUUN"): s=s.replace(a,b)
    s=s.replace("PRODUCCCION","PRODUCCION").replace("BUSINNES","BUSINESS")
    s=re.sub(r'\bCOORDINADOR/?A?\b','COORDINADOR',s)
    s=re.sub(r'\bSUPERVISOR/?A?\b','SUPERVISOR',s)
    s=re.sub(r'\bANALISTA/?A?\b','ANALISTA',s)
    s=s.replace("/O","").replace("/A","")
    return " ".join(s.split())

hc=J("headcount_adm.json"); mx=J("matrices.json"); eqd=J("equivalencias.json")
EQ={k:v["cargo_hc"] for k,v in eqd["equivalencias"].items()}

# ---- puestos evaluados
evald=defaultdict(list)          # norm(cargo_hc) -> [(cod, area_matriz, nombre_original, lineas)]
matrices=[]                      # trazabilidad
for arch,hojas in mx.items():
    for h in hojas:
        if "puestos" not in h: continue
        cod=h.get("codigo") or arch
        amat=h.get("meta",{}).get("PUESTO DE TRABAJO A EVALUAR","(sin dato)")
        matrices.append({"cod":cod,"arch":arch,"area":amat,
                         "puestos":h["puestos"],"lineas":h.get("filas_datos",0)})
        for pu,n in h["puestos"].items():
            k=norm(pu); k=norm(EQ.get(k,k))
            evald[k].append((cod,amat,pu,n))

# ---- universo head count
pares=OrderedDict()
for r in hc:
    k=(r["area"],r["cargo"])
    if k not in pares: pares[k]={"n":0,"norm":norm(r["cargo"]),"jer":set(),"loc":set(),"cc":set()}
    p=pares[k]; p["n"]+=1; p["jer"].add(r["jerarquia"]); p["loc"].add(r["localidad"]); p["cc"].add(r["cc"])

# ---- estilos
AZ="1F4E79"; VD="2E7D32"; RJ="C62828"; NJ="EF6C00"; GR="F2F2F2"
def F(c,b=True,sz=10,col="FFFFFF"): return Font(name="Calibri",bold=b,size=sz,color=col)
fill=lambda c: PatternFill("solid",fgColor=c)
thin=Side(style="thin",color="BFBFBF"); BOR=Border(left=thin,right=thin,top=thin,bottom=thin)
CEN=Alignment(horizontal="center",vertical="center",wrap_text=True)
IZQ=Alignment(horizontal="left",vertical="center",wrap_text=True)

wb=openpyxl.Workbook(); wb.remove(wb.active)

def hoja(nombre, titulo, cols, filas, anchos, congelar="A5"):
    ws=wb.create_sheet(nombre)
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=len(cols))
    c=ws.cell(1,1,titulo); c.font=F(None,True,13); c.fill=fill(AZ); c.alignment=CEN
    ws.row_dimensions[1].height=26
    ws.merge_cells(start_row=2,start_column=1,end_row=2,end_column=len(cols))
    s=ws.cell(2,1,"PLANTACIONES DE BALSA PLANTABAL S.A. / 3A CORE MATERIALS  ·  Tratamiento CAR  ·  Verificación de cobertura IPERC - Personal administrativo")
    s.font=Font(name="Calibri",size=9,italic=True,color="595959"); s.alignment=CEN
    for j,t in enumerate(cols,1):
        c=ws.cell(4,j,t); c.font=F(None,True,10); c.fill=fill(AZ); c.alignment=CEN; c.border=BOR
    ws.row_dimensions[4].height=32
    for i,fl in enumerate(filas,5):
        for j,v in enumerate(fl,1):
            c=ws.cell(i,j,v); c.border=BOR; c.font=Font(name="Calibri",size=10)
            c.alignment=CEN if (isinstance(v,int) or j>len(cols)-0) else IZQ
            if i%2==1: c.fill=fill(GR)
    for j,w in enumerate(anchos,1): ws.column_dimensions[get_column_letter(j)].width=w
    ws.freeze_panes=congelar
    ws.auto_filter.ref=f"A4:{get_column_letter(len(cols))}{4+len(filas)}"
    return ws

# ================= HOJA 1 - ALCANCE
crit=[
 ["1","Fuente del head count","Archivo 'HeadCount - V1.1', Hoja1 (base nominal, 817 registros). La Hoja2 contiene una tabla dinámica cuyos filtros fueron recuperados para replicar el criterio original."],
 ["2","Exclusión por área","Se excluye el área 'Silvicultura - Abastecimiento' (personal de campo). Confirmado por el responsable de la evaluación."],
 ["3","Exclusión por jerarquía","Se excluyen las jerarquías 'BC Workers' y 'BC Leaders' (personal operativo). Se mantienen: BC Adm, Support Staff, Professional, Shift Supervisor, Low / Middle / Top Management."],
 ["4","Alcance geográfico","Se incluyen las localidades Quevedo y Samborondón. Se excluye Juárez (México)."],
 ["5","Exclusión por dependencia","Se excluye al personal con Localidad Principal 'High Point', por corresponder a colaboradores que no permanecen habitualmente en las instalaciones del país."],
 ["6","Cargos de frontera","Se mantienen dentro del alcance cargos como Supervisor de Producción, Supervisor de Mantenimiento, Chofer, Mensajero, Aprendiz, Jefe de Bodega, Técnico de Inventario, Enfermera/o y Médico Ocupacional."],
 ["7","Unidad de análisis","El cruce se realiza por el par ÁREA + CARGO, no por el nombre del cargo de forma aislada, dado que existen cargos presentes en más de un área."],
 ["8","Normalización","Se neutralizan mayúsculas, tildes, dobles espacios, terminaciones de género y errores de tipeo. Las equivalencias no evidentes fueron validadas una a una con el responsable (ver hoja 5)."],
]
hoja("1. Alcance y criterios","ALCANCE Y CRITERIOS DE ANÁLISIS",
     ["N°","Criterio","Descripción"],crit,[6,30,110],"A5")

# ================= HOJA 2 - RESUMEN POR ÁREA
cub=Counter(); tot=Counter(); pt=Counter(); pc=Counter()
for (area,cargo),v in pares.items():
    tot[area]+=1; pt[area]+=v["n"]
    if v["norm"] in evald: cub[area]+=1; pc[area]+=v["n"]
filas=[]
for a in sorted(tot,key=lambda x:(cub[x]-tot[x], x)):
    pend=tot[a]-cub[a]
    filas.append([a,tot[a],cub[a],pend,round(100*cub[a]/tot[a]),pt[a],pc[a],pt[a]-pc[a],
                  "CONFORME" if pend==0 else "CON BRECHAS"])
T,C=sum(tot.values()),sum(cub.values()); PT,PC=sum(pt.values()),sum(pc.values())
filas.append(["TOTAL GENERAL",T,C,T-C,round(100*C/T),PT,PC,PT-PC,""])
ws=hoja("2. Resumen por área","RESUMEN DE COBERTURA POR ÁREA",
  ["Área","Cargos en head count","Cargos analizados","Cargos sin analizar","% cobertura de cargos",
   "N° personas","Personas cubiertas","Personas sin cubrir","Estado"],
  filas,[32,13,13,13,13,11,12,12,16])
for i in range(5,5+len(filas)):
    est=ws.cell(i,9).value
    if est=="CONFORME": ws.cell(i,9).font=F(None,True,10,VD)
    elif est=="CON BRECHAS": ws.cell(i,9).font=F(None,True,10,RJ)
    ws.cell(i,5).number_format='0"%"'
for j in range(1,10):
    c=ws.cell(4+len(filas),j); c.font=F(None,True,10); c.fill=fill("D9E2F3")

# ================= HOJA 3 - MATRIZ CONSOLIDADA
filas=[]
for (area,cargo),v in sorted(pares.items()):
    ev=evald.get(v["norm"],[])
    if ev:
        cods=" / ".join(sorted(set(e[0] for e in ev)))
        noms=" / ".join(sorted(set(e[2] for e in ev)))
        lin=sum(e[3] for e in ev)
        est="ANALIZADO"
        obs=""
        if norm(noms)!=v["norm"] and len(set(norm(e[2]) for e in ev))==1:
            obs="Denominación distinta entre matriz y head count (equivalencia validada)"
        if len(set(e[0] for e in ev))>1:
            obs=(obs+" | " if obs else "")+"Evaluado en más de una matriz"
    else:
        cods=noms=""; lin=0; est="SIN ANALIZAR"; obs="Requiere elaboración de matriz IPERC"
    filas.append([area,cargo,v["n"]," / ".join(sorted(v["jer"]))," / ".join(sorted(v["loc"])),
                  est,cods,noms,lin,obs])
ws=hoja("3. Matriz consolidada","MATRIZ CONSOLIDADA: CARGOS DEL HEAD COUNT vs. CARGOS ANALIZADOS EN LAS MATRICES IPERC",
  ["Área","Cargo según head count","N° personas","Jerarquía","Localidad","Estado",
   "Código de matriz","Denominación en la matriz","Líneas de riesgo","Observación"],
  filas,[30,42,9,18,14,14,16,42,10,48])
for i in range(5,5+len(filas)):
    c=ws.cell(i,6)
    if c.value=="ANALIZADO": c.font=F(None,True,10,VD)
    else: c.font=F(None,True,10,RJ); c.fill=fill("FCE4E4")

# ================= HOJA 4 - BRECHAS
br=[]
for (area,cargo),v in sorted(pares.items(), key=lambda x:(-x[1]["n"],x[0][0],x[0][1])):
    if v["norm"] not in evald:
        br.append([area,cargo,v["n"]," / ".join(sorted(v["jer"]))," / ".join(sorted(v["loc"])),
                   " / ".join(sorted(v["cc"])),"Elaborar matriz IPERC para el puesto","",""])
ws=hoja("4. Brechas","BRECHAS DETECTADAS: CARGOS DEL HEAD COUNT SIN EVALUACIÓN DE RIESGOS",
  ["Área","Cargo sin analizar","N° personas","Jerarquía","Localidad","Centro de costo",
   "Acción requerida","Responsable","Fecha compromiso"],
  br,[28,42,9,18,13,34,38,18,14])
for i in range(5,5+len(br)): ws.cell(i,2).font=F(None,True,10,RJ)

# ================= HOJA 5 - EQUIVALENCIAS Y DEPURACIÓN
eq=[]
for k,v in sorted(eqd["equivalencias"].items()):
    eq.append([v["matriz"],k,v["cargo_hc"],v.get("area_hc",""),"Equivalencia de denominación",
               v.get("nota",""),"Unificar la denominación en la matriz según el head count"])
for k,v in sorted(eqd.get("cargos_nuevos_no_en_headcount",{}).items()):
    eq.append([v["matriz"],k,"(no consta)",v.get("area_hc",""),"Cargo nuevo no registrado",
               v.get("estado",""),v.get("accion","")])
for k,v in sorted(eqd.get("cargos_obsoletos_a_eliminar",{}).items()):
    eq.append([v["matriz"],k,"(no existe)","","Cargo obsoleto",
               v.get("estado",""),v.get("accion","")])
hoja("5. Observaciones","OBSERVACIONES DOCUMENTALES: EQUIVALENCIAS, CARGOS NUEVOS Y CARGOS OBSOLETOS",
  ["Código de matriz","Denominación en la matriz","Cargo en head count","Área","Tipo de hallazgo",
   "Detalle","Acción requerida"],
  eq,[16,40,38,24,26,70,52])

# ================= HOJA 6 - TRAZABILIDAD
tz=[]
for m in sorted(matrices,key=lambda x:x["cod"]):
    tz.append([m["cod"],m["area"],len(m["puestos"]),sum(m["puestos"].values()),
               " · ".join(f"{k} ({v})" for k,v in m["puestos"].items()),
               os.path.basename(m["arch"])])
tz.append(["TOTAL","",sum(t[2] for t in tz),sum(t[3] for t in tz),"",""])
hoja("6. Trazabilidad","TRAZABILIDAD DE LAS MATRICES IPERC ANALIZADAS",
  ["Código","Puesto / proceso evaluado","N° de puestos","Líneas de riesgo",
   "Puestos evaluados (líneas de riesgo por puesto)","Archivo fuente"],
  tz,[16,34,12,13,95,52])

out=os.path.join("/home/user/tratamiento_CAR-s","Cobertura_IPERC_Administrativos_CAR.xlsx")
wb.save(out)
print("Generado:",out)
print(f"Resumen: {C}/{T} cargos ({round(100*C/T)}%) | {PC}/{PT} personas | {len(br)} brechas | {len(eq)} observaciones | {len(matrices)} matrices")
