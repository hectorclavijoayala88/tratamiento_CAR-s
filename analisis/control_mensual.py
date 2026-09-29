# -*- coding: utf-8 -*-
"""
Control mensual de cobertura IPERC - Personal administrativo
Plantabal S.A. / 3A Core Materials

Cruza el head count del mes contra las matrices IPERC y genera el consolidado.
NO requiere que el head count traiga tabla dinamica: localiza automaticamente
la hoja con la base nominal y aplica los criterios definidos.

Uso:
    python3 control_mensual.py --hc HeadCount.xlsx --matrices carpeta_matrices/
    python3 control_mensual.py --hc HeadCount.xlsx --matrices m1.xlsx m2.xlsx
Opcionales:
    --out archivo.xlsx        (por defecto: Cobertura_IPERC_Administrativos_CAR.xlsx)
    --sin-comparar            (no compara contra el mes anterior)
"""
import argparse, glob, json, os, re, sys, datetime, warnings
warnings.filterwarnings("ignore", module="openpyxl")
from collections import OrderedDict, defaultdict, Counter
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE = os.path.dirname(os.path.abspath(__file__))
SNAP = os.path.join(BASE, "snapshots")

# ───────────────────────── criterios del universo ─────────────────────────
AREAS_EXCLUIDAS      = {"SILVICULTURA - ABASTECIMIENTO"}
JERARQUIAS_EXCLUIDAS = {"BC WORKERS", "BC LEADERS"}
LOCALIDADES_INCLUIDAS= {"QUEVEDO", "SAMBORONDON"}
LOCALIDADP_EXCLUIDAS = {"HIGH POINT"}

def norm(s):
    s = " ".join(str(s or "").strip().upper().split())
    for a, b in zip("ÁÉÍÓÚÜÑ", "AEIOUUN"): s = s.replace(a, b)
    s = s.replace("PRODUCCCION", "PRODUCCION").replace("BUSINNES", "BUSINESS")
    s = re.sub(r"\bCOORDINADOR/?A?\b", "COORDINADOR", s)
    s = re.sub(r"\bSUPERVISOR/?A?\b", "SUPERVISOR", s)
    s = re.sub(r"\bANALISTA/?A?\b", "ANALISTA", s)
    s = s.replace("/O", "").replace("/A", "")
    return " ".join(s.split())

# alias tolerantes para los encabezados del head count
ALIAS = {
    "area":       ["AREA", "ÁREA", "DEPARTAMENTO"],
    "cargo":      ["CARGO", "PUESTO", "DENOMINACION DEL CARGO"],
    "jerarquia":  ["JERARQUIA", "JERARQUÍA", "NIVEL", "BANDA"],
    "localidad":  ["LOCALIDAD", "SEDE", "UBICACION"],
    "localidadP": ["LOCALIDADP", "LOCALIDAD P", "LOCALIDAD PRINCIPAL"],
    "cc":         ["CENTRO DE COSTO", "CENTRO DE COSTOS", "CC"],
    "tipo":       ["TIPO DE CARGO", "TIPO CARGO"],
}

def localizar_base(path):
    """Busca en TODAS las hojas la fila de encabezados de la base nominal."""
    wb = openpyxl.load_workbook(path, data_only=True)
    mejor = None
    for ws in wb.worksheets:
        for fila in range(1, min(ws.max_row, 15) + 1):
            vals = {norm(c.value): c.column for c in ws[fila] if c.value not in (None, "")}
            if not vals: continue
            idx, encontrados = {}, 0
            for clave, opciones in ALIAS.items():
                for op in opciones:
                    if op in vals:
                        idx[clave] = vals[op] - 1; encontrados += 1; break
            # la base nominal debe traer al menos area + cargo
            if "area" in idx and "cargo" in idx:
                filas = ws.max_row - fila
                puntaje = (encontrados, filas)
                if mejor is None or puntaje > mejor[0]:
                    mejor = (puntaje, ws, fila, idx)
    if mejor is None:
        sys.exit("ERROR: no se encontro ninguna hoja con columnas 'Area' y 'Cargo'.\n"
                 "       Revisa que el head count incluya la base nominal.")
    (_, ws, fila, idx) = mejor
    faltan = [k for k in ("jerarquia", "localidad") if k not in idx]
    if faltan:
        print(f"  AVISO: no se hallaron las columnas {faltan}. "
              f"Los criterios que dependen de ellas no se aplicaran.")
    return ws, fila, idx

def leer_headcount(path):
    ws, fila, idx = localizar_base(path)
    print(f"  Base nominal: hoja '{ws.title}', encabezados en fila {fila}")
    G = lambda r, k: (str(r[idx[k]]).strip() if k in idx and r[idx[k]] is not None else "")
    total = incl = 0; motivos = Counter(); universo = []
    for r in ws.iter_rows(min_row=fila + 1, values_only=True):
        if not any(v not in (None, "") for v in r): continue
        if not G(r, "cargo"): continue
        total += 1
        if norm(G(r, "area")) in AREAS_EXCLUIDAS:            motivos["Area excluida (Silvicultura)"] += 1; continue
        if norm(G(r, "jerarquia")) in JERARQUIAS_EXCLUIDAS:  motivos["Jerarquia operativa (BC Workers/Leaders)"] += 1; continue
        if "localidad" in idx and norm(G(r, "localidad")) not in LOCALIDADES_INCLUIDAS:
            motivos["Localidad fuera de alcance"] += 1; continue
        if "localidadP" in idx and norm(G(r, "localidadP")) in LOCALIDADP_EXCLUIDAS:
            motivos["Localidad principal High Point"] += 1; continue
        incl += 1
        universo.append({"area": G(r, "area"), "cargo": G(r, "cargo"),
                         "cargo_norm": norm(G(r, "cargo")), "cc": G(r, "cc"),
                         "jerarquia": G(r, "jerarquia"), "tipo": G(r, "tipo"),
                         "localidad": G(r, "localidad"), "localidadP": G(r, "localidadP")})
    print(f"  Registros leidos: {total}  ->  dentro del alcance: {incl}")
    for m, n in motivos.most_common(): print(f"     - excluidos por {m}: {n}")
    return universo

def extraer_matriz(path):
    """Lee una matriz IPERC (GTC 45) y devuelve los puestos evaluados."""
    wb = openpyxl.load_workbook(path, data_only=True)
    salida = []
    for ws in wb.worksheets:
        hrow = hcol = None
        for row in ws.iter_rows(min_row=1, max_row=20, max_col=30):
            for c in row:
                if norm(c.value) == "PUESTO DE TRABAJO": hrow, hcol = c.row, c.column; break
            if hrow: break
        if not hrow: continue                      # hoja de tablas de referencia
        codigo = None
        for row in ws.iter_rows(min_row=1, max_row=3, max_col=30):
            for c in row:
                if isinstance(c.value, str) and re.match(r"^GE-SO-[A-Z]+-D\d", c.value.strip()):
                    codigo = c.value.strip()
        proceso = "(sin dato)"
        for row in ws.iter_rows(min_row=1, max_row=8, max_col=30):
            for c in row:
                if norm(c.value) == "PUESTO DE TRABAJO A EVALUAR":
                    for cc in ws[c.row]:
                        if cc.column > c.column and cc.value not in (None, ""):
                            proceso = str(cc.value).strip(); break
        fill = {}
        for mr in ws.merged_cells.ranges:
            if mr.min_col == hcol:
                v = ws.cell(mr.min_row, mr.min_col).value
                for r in range(mr.min_row, mr.max_row + 1): fill[r] = v
        puestos, actual, nfilas = OrderedDict(), None, 0
        for r in range(hrow + 1, ws.max_row + 1):
            if not any(ws.cell(r, c).value not in (None, "") for c in range(1, min(ws.max_column, 30) + 1)):
                continue
            nfilas += 1
            v = fill.get(r, ws.cell(r, hcol).value)
            if v not in (None, "") and str(v).strip(): actual = str(v).strip()
            if actual: puestos[actual] = puestos.get(actual, 0) + 1
        salida.append({"codigo": codigo or os.path.basename(path), "proceso": proceso,
                       "puestos": puestos, "filas_datos": nfilas,
                       "archivo": os.path.basename(path)})
    return salida

# ───────────────────────── generacion del Excel ─────────────────────────
AZ, VD, RJ, GR = "1F4E79", "2E7D32", "C62828", "F2F2F2"
_thin = Side(style="thin", color="BFBFBF")
BOR = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
CEN = Alignment(horizontal="center", vertical="center", wrap_text=True)
IZQ = Alignment(horizontal="left", vertical="center", wrap_text=True)
FB  = lambda sz=10, col="FFFFFF": Font(name="Calibri", bold=True, size=sz, color=col)
fill= lambda c: PatternFill("solid", fgColor=c)

def hoja(wb, nombre, titulo, cols, filas, anchos):
    ws = wb.create_sheet(nombre)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(cols))
    c = ws.cell(1, 1, titulo); c.font = FB(13); c.fill = fill(AZ); c.alignment = CEN
    ws.row_dimensions[1].height = 26
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(cols))
    s = ws.cell(2, 1, "PLANTACIONES DE BALSA PLANTABAL S.A. / 3A CORE MATERIALS  ·  "
                      "Control de cobertura IPERC - Personal administrativo  ·  "
                      f"Corte {datetime.date.today().strftime('%d/%m/%Y')}")
    s.font = Font(name="Calibri", size=9, italic=True, color="595959"); s.alignment = CEN
    for j, t in enumerate(cols, 1):
        c = ws.cell(4, j, t); c.font = FB(); c.fill = fill(AZ); c.alignment = CEN; c.border = BOR
    ws.row_dimensions[4].height = 32
    for i, fl in enumerate(filas, 5):
        for j, v in enumerate(fl, 1):
            c = ws.cell(i, j, v); c.border = BOR; c.font = Font(name="Calibri", size=10)
            c.alignment = CEN if isinstance(v, int) else IZQ
            if i % 2 == 1: c.fill = fill(GR)
    for j, w in enumerate(anchos, 1): ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "A5"
    if filas: ws.auto_filter.ref = f"A4:{get_column_letter(len(cols))}{4+len(filas)}"
    return ws

def generar(universo, matrices, eqd, salida, cambios=None):
    EQ = {k: v["cargo_hc"] for k, v in eqd.get("equivalencias", {}).items()}
    evald = defaultdict(list)
    for m in matrices:
        for pu, n in m["puestos"].items():
            k = norm(pu); k = norm(EQ.get(k, k))
            evald[k].append((m["codigo"], m["proceso"], pu, n))

    pares = OrderedDict()
    for r in universo:
        k = (r["area"], r["cargo"])
        if k not in pares:
            pares[k] = {"n": 0, "norm": norm(r["cargo"]), "jer": set(), "loc": set(), "cc": set()}
        p = pares[k]; p["n"] += 1
        p["jer"].add(r["jerarquia"]); p["loc"].add(r["localidad"]); p["cc"].add(r["cc"])

    wb = openpyxl.Workbook(); wb.remove(wb.active)

    hoja(wb, "1. Alcance y criterios", "ALCANCE Y CRITERIOS DE ANÁLISIS",
         ["N°", "Criterio", "Descripción"],
         [["1", "Fuente del head count", "Base nominal del head count del mes. La hoja se localiza automáticamente por sus encabezados; no se requiere tabla dinámica."],
          ["2", "Exclusión por área", "Se excluye el área 'Silvicultura - Abastecimiento' (personal de campo)."],
          ["3", "Exclusión por jerarquía", "Se excluyen las jerarquías 'BC Workers' y 'BC Leaders' (personal operativo)."],
          ["4", "Alcance geográfico", "Se incluyen las localidades Quevedo y Samborondón."],
          ["5", "Exclusión por dependencia", "Se excluye al personal con Localidad Principal 'High Point'."],
          ["6", "Cargos de frontera", "Se mantienen dentro del alcance Supervisor de Producción, Supervisor de Mantenimiento, Chofer, Mensajero, Aprendiz, Jefe de Bodega, Técnico de Inventario, Enfermera/o y Médico Ocupacional."],
          ["7", "Unidad de análisis", "El cruce se realiza por el par ÁREA + CARGO, dado que existen cargos presentes en más de un área."],
          ["8", "Normalización", "Se neutralizan mayúsculas, tildes, dobles espacios, terminaciones de género y errores de tipeo. Las equivalencias no evidentes están registradas en la hoja 5."]],
         [6, 30, 110])

    cub = Counter(); tot = Counter(); pt = Counter(); pc = Counter()
    for (area, cargo), v in pares.items():
        tot[area] += 1; pt[area] += v["n"]
        if v["norm"] in evald: cub[area] += 1; pc[area] += v["n"]
    filas = []
    for a in sorted(tot, key=lambda x: (cub[x] - tot[x], x)):
        pend = tot[a] - cub[a]
        filas.append([a, tot[a], cub[a], pend, round(100 * cub[a] / tot[a]),
                      pt[a], pc[a], pt[a] - pc[a], "CONFORME" if pend == 0 else "CON BRECHAS"])
    T, C = sum(tot.values()), sum(cub.values()); PT, PC = sum(pt.values()), sum(pc.values())
    filas.append(["TOTAL GENERAL", T, C, T - C, round(100 * C / T) if T else 0, PT, PC, PT - PC, ""])
    ws = hoja(wb, "2. Resumen por área", "RESUMEN DE COBERTURA POR ÁREA",
              ["Área", "Cargos en head count", "Cargos analizados", "Cargos sin analizar",
               "% cobertura de cargos", "N° personas", "Personas cubiertas", "Personas sin cubrir", "Estado"],
              filas, [32, 13, 13, 13, 13, 11, 12, 12, 16])
    for i in range(5, 5 + len(filas)):
        e = ws.cell(i, 9).value
        if e == "CONFORME":     ws.cell(i, 9).font = FB(10, VD)
        elif e == "CON BRECHAS": ws.cell(i, 9).font = FB(10, RJ)
        ws.cell(i, 5).number_format = '0"%"'
    for j in range(1, 10):
        c = ws.cell(4 + len(filas), j); c.font = FB(); c.fill = fill("D9E2F3")

    filas = []
    for (area, cargo), v in sorted(pares.items()):
        ev = evald.get(v["norm"], [])
        if ev:
            cods = " / ".join(sorted({e[0] for e in ev}))
            noms = " / ".join(sorted({e[2] for e in ev}))
            obs = []
            if len({norm(e[2]) for e in ev}) == 1 and norm(noms) != v["norm"]:
                obs.append("Denominación distinta entre matriz y head count (equivalencia validada)")
            if len({e[0] for e in ev}) > 1: obs.append("Evaluado en más de una matriz")
            filas.append([area, cargo, v["n"], " / ".join(sorted(v["jer"])), " / ".join(sorted(v["loc"])),
                          "ANALIZADO", cods, noms, sum(e[3] for e in ev), " | ".join(obs)])
        else:
            filas.append([area, cargo, v["n"], " / ".join(sorted(v["jer"])), " / ".join(sorted(v["loc"])),
                          "SIN ANALIZAR", "", "", 0, "Requiere elaboración de matriz IPERC"])
    ws = hoja(wb, "3. Matriz consolidada",
              "MATRIZ CONSOLIDADA: CARGOS DEL HEAD COUNT vs. CARGOS ANALIZADOS EN LAS MATRICES IPERC",
              ["Área", "Cargo según head count", "N° personas", "Jerarquía", "Localidad", "Estado",
               "Código de matriz", "Denominación en la matriz", "Líneas de riesgo", "Observación"],
              filas, [30, 42, 9, 18, 14, 14, 16, 42, 10, 48])
    for i in range(5, 5 + len(filas)):
        c = ws.cell(i, 6)
        if c.value == "ANALIZADO": c.font = FB(10, VD)
        else: c.font = FB(10, RJ); c.fill = fill("FCE4E4")

    br = [[a, c, v["n"], " / ".join(sorted(v["jer"])), " / ".join(sorted(v["loc"])),
           " / ".join(sorted(v["cc"])), "Elaborar matriz IPERC para el puesto", "", ""]
          for (a, c), v in sorted(pares.items(), key=lambda x: (-x[1]["n"], x[0][0], x[0][1]))
          if v["norm"] not in evald]
    ws = hoja(wb, "4. Brechas", "BRECHAS DETECTADAS: CARGOS DEL HEAD COUNT SIN EVALUACIÓN DE RIESGOS",
              ["Área", "Cargo sin analizar", "N° personas", "Jerarquía", "Localidad", "Centro de costo",
               "Acción requerida", "Responsable", "Fecha compromiso"],
              br, [28, 42, 9, 18, 13, 34, 38, 18, 14])
    for i in range(5, 5 + len(br)): ws.cell(i, 2).font = FB(10, RJ)

    hcn = {v["norm"] for v in pares.values()}
    eq = []
    for k, v in sorted(eqd.get("equivalencias", {}).items()):
        eq.append([v.get("matriz", ""), k, v["cargo_hc"], v.get("area_hc", ""),
                   "Equivalencia de denominación", v.get("nota", ""),
                   "Unificar la denominación en la matriz según el head count"])
    for k, v in sorted(eqd.get("cargos_nuevos_no_en_headcount", {}).items()):
        eq.append([v.get("matriz", ""), k, "(no consta)", v.get("area_hc", ""),
                   "Cargo nuevo no registrado", v.get("estado", ""), v.get("accion", "")])
    for k, v in sorted(eqd.get("cargos_obsoletos_a_eliminar", {}).items()):
        eq.append([v.get("matriz", ""), k, "(no existe)", "", "Cargo obsoleto",
                   v.get("estado", ""), v.get("accion", "")])
    # puestos evaluados que este mes no encuentran cargo en el head count
    ya = {norm(k) for k in list(eqd.get("equivalencias", {})) +
                          list(eqd.get("cargos_nuevos_no_en_headcount", {})) +
                          list(eqd.get("cargos_obsoletos_a_eliminar", {}))}
    for k, vs in sorted(evald.items()):
        if k not in hcn and norm(vs[0][2]) not in ya:
            eq.append([vs[0][0], vs[0][2], "(no consta)", "", "SIN CLASIFICAR - detectado este mes",
                       "Puesto evaluado en la matriz que no aparece en el head count del mes",
                       "Definir si es cargo nuevo, denominación alterna o cargo suprimido"])
    hoja(wb, "5. Observaciones",
         "OBSERVACIONES DOCUMENTALES: EQUIVALENCIAS, CARGOS NUEVOS Y CARGOS OBSOLETOS",
         ["Código de matriz", "Denominación en la matriz", "Cargo en head count", "Área",
          "Tipo de hallazgo", "Detalle", "Acción requerida"],
         eq, [16, 40, 38, 24, 26, 70, 52])

    tz = [[m["codigo"], m["proceso"], len(m["puestos"]), sum(m["puestos"].values()),
           " · ".join(f"{k} ({v})" for k, v in m["puestos"].items()), m["archivo"]]
          for m in sorted(matrices, key=lambda x: x["codigo"])]
    tz.append(["TOTAL", "", sum(t[2] for t in tz), sum(t[3] for t in tz), "", ""])
    hoja(wb, "6. Trazabilidad", "TRAZABILIDAD DE LAS MATRICES IPERC ANALIZADAS",
         ["Código", "Puesto / proceso evaluado", "N° de puestos", "Líneas de riesgo",
          "Puestos evaluados (líneas de riesgo por puesto)", "Archivo fuente"],
         tz, [16, 34, 12, 13, 95, 52])

    if cambios:
        hoja(wb, "7. Cambios del mes", "CAMBIOS RESPECTO AL CORTE ANTERIOR",
             ["Tipo de cambio", "Área", "Cargo", "N° personas", "Estado de cobertura", "Acción sugerida"],
             cambios, [26, 30, 44, 11, 18, 52])

    wb.save(salida)
    return {"cargos": T, "cubiertos": C, "personas": PT, "personas_cub": PC,
            "brechas": len(br), "observaciones": len(eq), "matrices": len(matrices)}

def comparar(universo):
    """Compara contra el snapshot anterior y devuelve las filas de cambios."""
    os.makedirs(SNAP, exist_ok=True)
    hoy = {f'{r["area"]}||{r["cargo"]}': 0 for r in universo}
    for r in universo: hoy[f'{r["area"]}||{r["cargo"]}'] += 1
    previos = sorted(glob.glob(os.path.join(SNAP, "universo_*.json")))
    cambios = []
    if previos:
        ant = json.load(open(previos[-1], encoding="utf-8"))
        etq = os.path.basename(previos[-1])[9:-5]
        for k, n in sorted(hoy.items()):
            a, c = k.split("||")
            if k not in ant:
                cambios.append(["ALTA - cargo nuevo", a, c, n, "Verificar cobertura",
                                "Confirmar si existe matriz IPERC que cubra el puesto"])
            elif ant[k] != n:
                cambios.append([f"VARIACIÓN de dotación ({ant[k]} → {n})", a, c, n, "",
                                "Sin acción si el cargo ya está cubierto"])
        for k, n in sorted(ant.items()):
            if k not in hoy:
                a, c = k.split("||")
                cambios.append(["BAJA - cargo ya no consta", a, c, n, "",
                                "Evaluar si corresponde retirar el bloque de la matriz IPERC"])
        print(f"  Comparado contra el corte {etq}: {len(cambios)} cambios")
    else:
        print("  No hay corte anterior con el cual comparar (este sera el primero).")
    json.dump(hoy, open(os.path.join(SNAP, f"universo_{datetime.date.today():%Y-%m-%d}.json"), "w"),
              ensure_ascii=False, indent=1)
    return cambios

CATALOGO = os.path.join(BASE, "catalogo_matrices.json")

def cargar_catalogo():
    if not os.path.exists(CATALOGO): return []
    return json.load(open(CATALOGO, encoding="utf-8")).get("matrices", [])

def fecha_catalogo():
    if not os.path.exists(CATALOGO): return "-"
    return json.load(open(CATALOGO, encoding="utf-8")).get("actualizado", "-")

def fusionar_catalogo(nuevas):
    """Incorpora las matrices leidas al catalogo: las de igual codigo se reemplazan."""
    cat = {m["codigo"]: m for m in cargar_catalogo()}
    altas = [m["codigo"] for m in nuevas if m["codigo"] not in cat]
    reemp = [m["codigo"] for m in nuevas if m["codigo"] in cat]
    for m in nuevas: cat[m["codigo"]] = m
    json.dump({"actualizado": datetime.date.today().strftime("%Y-%m-%d"),
               "matrices": list(cat.values())},
              open(CATALOGO, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    if altas: print(f"  Catalogo: {len(altas)} matriz(ces) incorporada(s): {', '.join(altas)}")
    if reemp: print(f"  Catalogo: {len(reemp)} matriz(ces) actualizada(s): {', '.join(reemp)}")
    print(f"  Catalogo guardado con {len(cat)} matrices en total")
    return list(cat.values())

def main():
    ap = argparse.ArgumentParser(description="Control mensual de cobertura IPERC")
    ap.add_argument("--hc", required=True, help="Head count del mes (.xlsx)")
    ap.add_argument("--matrices", nargs="+", default=None,
                    help="Carpeta o archivos de matrices IPERC. Si se omite, se usa el "
                         "catalogo guardado en analisis/catalogo_matrices.json")
    ap.add_argument("--out", default="Cobertura_IPERC_Administrativos_CAR.xlsx")
    ap.add_argument("--sin-comparar", action="store_true")
    a = ap.parse_args()

    print("\n[1/4] Leyendo head count …")
    universo = leer_headcount(a.hc)
    if not universo: sys.exit("ERROR: el universo quedo vacio. Revisa los criterios o el archivo.")

    if a.matrices:
        print("\n[2/4] Leyendo matrices IPERC …")
        rutas = []
        for m in a.matrices:
            rutas += sorted(glob.glob(os.path.join(m, "*.xlsx"))) if os.path.isdir(m) else sorted(glob.glob(m))
        rutas = [r for r in rutas if not os.path.basename(r).startswith("~$")]
        matrices = []
        for r in rutas:
            try:
                ms = extraer_matriz(r)
                if not ms: print(f"  AVISO: {os.path.basename(r)} no parece una matriz IPERC (se omite)")
                for m in ms: print(f"  {m['codigo']:<14} {m['proceso']:<32} {len(m['puestos'])} puestos")
                matrices += ms
            except Exception as e:
                print(f"  ERROR leyendo {os.path.basename(r)}: {e}")
        if not matrices: sys.exit("ERROR: no se pudo leer ninguna matriz.")
        matrices = fusionar_catalogo(matrices)
    else:
        print("\n[2/4] Usando el catalogo de matrices guardado …")
        matrices = cargar_catalogo()
        if not matrices:
            sys.exit("ERROR: no hay catalogo guardado y no se indicaron matrices.\n"
                     "       Ejecuta una vez con --matrices <carpeta> para crearlo.")
        print(f"  Catalogo: {len(matrices)} matrices  (actualizado el {fecha_catalogo()})")
        for m in sorted(matrices, key=lambda x: x["codigo"]):
            print(f"  {m['codigo']:<14} {m['proceso']:<32} {len(m['puestos'])} puestos")

    print("\n[3/4] Comparando con el corte anterior …")
    cambios = [] if a.sin_comparar else comparar(universo)

    print("\n[4/4] Generando el consolidado …")
    eqd = {}
    f = os.path.join(BASE, "equivalencias.json")
    if os.path.exists(f): eqd = json.load(open(f, encoding="utf-8"))
    else: print("  AVISO: no se hallo equivalencias.json; el cruce sera solo por nombre exacto.")
    res = generar(universo, matrices, eqd, a.out, cambios)

    print(f"\n  Archivo generado: {a.out}")
    print(f"  Cobertura: {res['cubiertos']}/{res['cargos']} cargos "
          f"({round(100*res['cubiertos']/res['cargos'])}%)  |  "
          f"{res['personas_cub']}/{res['personas']} personas")
    print(f"  Brechas: {res['brechas']}  |  Observaciones: {res['observaciones']}  |  "
          f"Matrices: {res['matrices']}\n")

if __name__ == "__main__":
    main()
