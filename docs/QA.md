# QA y control de calidad

## Objetivo

Garantizar que cada cambio del procesador Effi sea reproducible, revisable y seguro antes de llegar a `develop` o `main`.

## Flujo obligatorio

`feature/* -> PR -> develop -> PR -> main`

No se deben hacer pushes directos a `develop` ni `main` cuando los Rulesets estén activos.

## Puertas de calidad

### 1. Tests
Ejecutar:

```bash
pytest -q --cov=src/effi_processor --cov-report=term-missing
```

Cobertura mínima inicial recomendada: 70%. Subir progresivamente a 80%+ cuando el conjunto de pruebas esté maduro.

### 2. Formato y lint

```bash
black --check .
flake8 .
mypy src
```

### 3. Seguridad

```bash
bandit -r src app_effi.py -c pyproject.toml
pip-audit -r requirements.txt
```

### 4. Pre-commit

```bash
pre-commit install
pre-commit run --all-files
```

## Pruebas funcionales prioritarias

1. Match exacto por GTIN.
2. Match por nombre/referencia cuando no existe GTIN.
3. Omisión de coincidencias con confianza insuficiente.
4. Conversión g -> kg.
5. Conversión ml -> L.
6. Conversión de galón y presentaciones equivalentes.
7. Presentaciones tipo `X 12 UNID`.
8. Bonificación `10 + 2`.
9. Descuento individual por línea.
10. IVA incluido.
11. IVA no incluido.
12. Prorrateo exacto de costos.
13. No creación de códigos Effi inexistentes.
14. Hoja principal sin fórmulas.
15. Encabezados Effi exactos.
16. Archivo final <= 5 MB.
17. Omisiones y cálculos documentados en hojas auxiliares.
18. Procesamiento de múltiples facturas conservando origen.

## Casos límite

- cantidad cero;
- cantidades negativas;
- precio cero;
- descuento > 100%;
- GTIN inválido;
- GTIN con ceros iniciales;
- unidades ambiguas;
- producto duplicado en catálogo;
- factura con IVA incluido;
- factura sin IVA;
- OCR incompleto;
- factura escaneada;
- caracteres especiales;
- separadores decimales `,` y `.`;
- archivos vacíos o corruptos.

## Seguridad

No subir al repositorio:
- facturas reales;
- datos personales;
- credenciales;
- tokens;
- archivos `.env`;
- catálogos privados reales;
- salidas de procesamiento con información sensible.

Usar fixtures anonimizados para las pruebas.

## Reglas de salida Effi

La hoja `Plantilla_Importacion` debe contener únicamente valores/texto y respetar exactamente las columnas confirmadas. Los cálculos, conversiones, alertas y omisiones deben quedar fuera de la hoja principal.

El esquema completo de 9 columnas de Effi debe confirmarse contra la plantilla vigente antes de inventar columnas adicionales.
