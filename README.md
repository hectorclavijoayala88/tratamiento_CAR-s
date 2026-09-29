# Tratamiento CAR — Cobertura IPERC del personal administrativo

Verificación de que todos los cargos administrativos del head count cuentan con
identificación de peligros y evaluación de riesgos (matriz IPERC, metodología GTC 45).

**Plantaciones de Balsa Plantabal S.A. / 3A Core Materials**

## Entregable

`Cobertura_IPERC_Administrativos_CAR.xlsx` — 6 hojas:

| Hoja | Contenido |
|---|---|
| 1. Alcance y criterios | Los 8 criterios que delimitan el universo analizado |
| 2. Resumen por área | Cobertura por área, en cargos y en personas |
| 3. Matriz consolidada | Un renglón por par área + cargo, con la matriz donde fue evaluado |
| 4. Brechas | Cargos sin evaluar, con columnas para el plan de acción |
| 5. Observaciones | Equivalencias de denominación, cargos nuevos y cargos obsoletos |
| 6. Trazabilidad | Las matrices analizadas, sus puestos y líneas de riesgo |

## Estado

- **52 de 70 cargos** analizados (74%)
- **76 de 101 personas** cubiertas (75%)
- **15 matrices** procesadas
- Áreas conformes: ADM-L&RH, ESG/Lab, Engicore

## Criterios del universo

El personal administrativo se define excluyendo del head count:

- Área `Silvicultura - Abastecimiento` (personal de campo)
- Jerarquías `BC Workers` y `BC Leaders` (personal operativo)
- Localidades distintas de Quevedo y Samborondón
- Personal con Localidad Principal `High Point`

El cruce se hace por el par **área + cargo**, ya que hay cargos presentes en más de un área.

## Reproducibilidad

Carpeta `analisis/`:

| Archivo | Función |
|---|---|
| `extraer.py` | Lee las matrices IPERC y extrae los puestos evaluados por bloque |
| `cruce.py` | Cruza el head count contra los puestos evaluados |
| `generar.py` | Genera el archivo Excel consolidado |
| `headcount_adm.json` | Universo administrativo procesado (sin datos personales) |
| `matrices.json` | Puestos extraídos de cada matriz |
| `equivalencias.json` | Decisiones de mapeo validadas una a una |

Para regenerar el consolidado tras incorporar nuevas matrices:

```bash
python3 analisis/extraer.py <matrices.xlsx> > analisis/matrices.json
python3 analisis/generar.py
```

Requiere `openpyxl`.
