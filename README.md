# Control de cobertura IPERC — Personal administrativo

Verifica que todos los cargos administrativos del head count cuenten con
identificación de peligros y evaluación de riesgos (matriz IPERC, metodología GTC 45).

**Plantaciones de Balsa Plantabal S.A. / 3A Core Materials**

Nació como respuesta a una CAR y quedó convertido en control mensual.

## Uso mensual

Cada vez que llega el head count nuevo, basta con eso:

```bash
python3 analisis/control_mensual.py --hc "HeadCount_del_mes.xlsx"
```

Las matrices no hace falta volver a pasarlas: sus datos quedan guardados en
`analisis/catalogo_matrices.json`. Solo se indican cuando hay una matriz nueva
o cuando alguna se modificó:

```bash
python3 analisis/control_mensual.py \
    --hc "HeadCount_del_mes.xlsx" \
    --matrices matrices_nuevas/
```

En ese caso el catálogo se actualiza solo: las matrices con código nuevo se
incorporan y las de código ya conocido se reemplazan por la versión recién leída.

**No requiere que el head count traiga tabla dinámica.** El script localiza por sí
mismo la hoja con la base nominal buscando sus encabezados, sin importar cómo se
llame la hoja ni en qué fila empiecen los datos.

Opciones:

| Opción | Para qué |
|---|---|
| *(sin `--matrices`)* | Usa el catálogo guardado |
| `--matrices carpeta/` | Lee los .xlsx de la carpeta y actualiza el catálogo |
| `--matrices a.xlsx b.xlsx` | Lee archivos sueltos y actualiza el catálogo |
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
| 8. Cargos por clasificar | Cargos del head count que no figuran en el catálogo |

Las hojas 7 y 8 aparecen solo cuando hay algo que reportar: la 7 cuando existe un
corte anterior con el cual comparar, la 8 cuando algún cargo quedó sin clasificar.

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
- Localidades distintas de Quevedo y Samborondón
- Personal con Localidad Principal `High Point`
- Cargos clasificados como operativos en `analisis/cargos_administrativos.json`

El cruce se hace por el par **área + cargo**, ya que hay cargos presentes en más de un área.

Para cambiar un criterio geográfico o de área, edita las constantes al inicio de
`control_mensual.py` (`AREAS_EXCLUIDAS`, `LOCALIDADES_INCLUIDAS`, `LOCALIDADP_EXCLUIDAS`).

### Por qué no se usa la columna de jerarquía

El universo se delimitaba antes excluyendo las jerarquías `BC Workers` y
`BC Leaders`. **Ese criterio se abandonó** porque el head count no siempre entrega
esa columna de forma confiable. Un archivo recibido en octubre de 2026 la traía
desalineada respecto al cargo: el Gerente General, el Director de Operaciones y
el Director Financiero figuraban como `BC Workers`, los dos únicos
`Top Management` eran un Operador Industrial y un Auxiliar de Producción, y
`TRABAJADOR FORESTAL` se repartía entre ocho jerarquías. Ese archivo se descartó
por completo.

La clasificación se sostiene ahora en el catálogo de cargos, construido a partir
del head count V1.1 y validado cargo por cargo. Da el mismo universo que daba el
filtro por jerarquía sobre ese archivo (101 personas, 70 pares área-cargo), pero
no depende de una columna que puede llegar corrupta.

Un cargo que no figure en el catálogo queda **fuera** del universo y se reporta en
la hoja `8. Cargos por clasificar`, para que la decisión sea explícita y no un
descarte silencioso.

### Control de salud del archivo

Si el head count trae columna de jerarquía, el script la contrasta con el catálogo:
cuando más del 30% de los cargos administrativos aparecen marcados como
`BC Workers` o `BC Leaders`, avisa en pantalla que esa columna llegó inconsistente.
El análisis no se detiene, porque la clasificación no depende de ella, pero el aviso
sirve para reclamar la corrección a quien genera el archivo.

### Clasificar un cargo nuevo

Agrégalo a la lista que corresponda en `analisis/cargos_administrativos.json`:

```json
{
  "administrativos": ["...", "CARGO NUEVO"],
  "operativos": ["...", "OTRO CARGO"]
}
```

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

Head count V1.1:

- **56 de 70 cargos** analizados (80%)
- **81 de 101 personas** cubiertas (80%)
- **17 matrices** procesadas
- Áreas conformes: ADM-L&RH, ESG/Lab, Engicore, QSE & Operational Excellence

## Archivos

| Archivo | Función |
|---|---|
| `analisis/control_mensual.py` | Script único del control mensual |
| `analisis/catalogo_matrices.json` | Puestos extraídos de cada matriz IPERC, para no volver a subirlas |
| `analisis/cargos_administrativos.json` | Clasificación administrativo / operativo de cada cargo |
| `analisis/equivalencias.json` | Equivalencias y hallazgos documentales validados |
| `analisis/snapshots/` | Fotos del universo de cada corte, para comparar |
