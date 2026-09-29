# Control de cobertura IPERC — Personal administrativo

Verifica que todos los cargos administrativos del head count cuenten con
identificación de peligros y evaluación de riesgos (matriz IPERC, metodología GTC 45).

**Plantaciones de Balsa Plantabal S.A. / 3A Core Materials**

Nació como respuesta a una CAR y quedó convertido en control mensual.

## Uso mensual

Cada vez que llega el head count nuevo:

```bash
python3 analisis/control_mensual.py \
    --hc "HeadCount_del_mes.xlsx" \
    --matrices carpeta_con_las_matrices/
```

**No requiere que el head count traiga tabla dinámica.** El script localiza por sí
mismo la hoja con la base nominal buscando sus encabezados, sin importar cómo se
llame la hoja ni en qué fila empiecen los datos.

Opciones:

| Opción | Para qué |
|---|---|
| `--matrices carpeta/` | Toma todos los .xlsx de la carpeta |
| `--matrices a.xlsx b.xlsx` | Toma archivos sueltos |
| `--out nombre.xlsx` | Cambia el nombre de salida |
| `--sin-comparar` | Omite la comparación con el corte anterior |

Requiere `openpyxl` (`pip install openpyxl`).

## Qué genera

`Cobertura_IPERC_Administrativos_CAR.xlsx`

| Hoja | Contenido |
|---|---|
| 1. Alcance y criterios | Los 8 criterios que delimitan el universo |
| 2. Resumen por área | Cobertura por área, en cargos y en personas |
| 3. Matriz consolidada | Un renglón por par área + cargo, con la matriz donde fue evaluado |
| 4. Brechas | Cargos sin evaluar, con columnas para el plan de acción |
| 5. Observaciones | Equivalencias, cargos nuevos y cargos obsoletos |
| 6. Trazabilidad | Matrices analizadas, sus puestos y líneas de riesgo |
| 7. Cambios del mes | Altas, bajas y variaciones de dotación frente al corte anterior |

La hoja 7 aparece solo cuando existe un corte anterior con el cual comparar.

## Control de cambios

El script guarda en `analisis/snapshots/` una foto del universo de cada corte y la
compara con la del mes anterior. Detecta en ambos sentidos:

- **Altas** — cargo nuevo en el head count: hay que verificar si alguna matriz lo cubre
- **Bajas** — cargo que desapareció: evaluar si corresponde retirar su bloque de la matriz
- **Variaciones de dotación** — cambió el número de personas en un cargo ya existente
- **Puestos sin clasificar** — un puesto evaluado en una matriz que dejó de constar en
  el head count (aparece en la hoja 5)

## Criterios del universo

Se excluye del head count:

- Área `Silvicultura - Abastecimiento` (personal de campo)
- Jerarquías `BC Workers` y `BC Leaders` (personal operativo)
- Localidades distintas de Quevedo y Samborondón
- Personal con Localidad Principal `High Point`

El cruce se hace por el par **área + cargo**, ya que hay cargos presentes en más de un área.

Para cambiar un criterio, edita las constantes al inicio de `control_mensual.py`
(`AREAS_EXCLUIDAS`, `JERARQUIAS_EXCLUIDAS`, `LOCALIDADES_INCLUIDAS`, `LOCALIDADP_EXCLUIDAS`).

## Equivalencias de denominación

Las matrices no siempre usan la nomenclatura del head count. Las equivalencias
validadas están en `analisis/equivalencias.json` y el script las aplica solo.
Para registrar una nueva, agrégala en la sección `equivalencias`:

```json
"NOMBRE EN LA MATRIZ": {
  "cargo_hc": "CARGO EN EL HEAD COUNT",
  "area_hc": "Área",
  "matriz": "GE-SO-XX-D1",
  "nota": "Motivo de la equivalencia"
}
```

Las secciones `cargos_nuevos_no_en_headcount` y `cargos_obsoletos_a_eliminar`
alimentan la hoja de observaciones.

## Estado del último corte

- **52 de 70 cargos** analizados (74%)
- **76 de 101 personas** cubiertas (75%)
- **15 matrices** procesadas
- Áreas conformes: ADM-L&RH, ESG/Lab, Engicore

## Archivos

| Archivo | Función |
|---|---|
| `analisis/control_mensual.py` | Script único del control mensual |
| `analisis/equivalencias.json` | Equivalencias y hallazgos documentales validados |
| `analisis/snapshots/` | Fotos del universo de cada corte, para comparar |
| `analisis/extraer.py`, `cruce.py`, `generar.py` | Scripts originales del análisis de la CAR |
| `analisis/headcount_adm.json`, `matrices.json` | Datos del corte de la CAR |
