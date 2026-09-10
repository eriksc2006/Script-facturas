# Effi Processor

Aplicación Python para procesar facturas, documentos y listados de compras y generar una plantilla de importación compatible con Effi, manteniendo una separación estricta entre datos de importación y cálculos/auditoría.

## Descripción

El proyecto está diseñado con enfoque de **Software Quality Assurance (SQA)** y trazabilidad. Sus objetivos principales son:

- Cruzar productos de una factura contra un catálogo maestro Effi.
- Priorizar coincidencias por GTIN y utilizar coincidencia semántica como mecanismo secundario.
- Convertir presentaciones y unidades cuando exista una equivalencia matemática confiable.
- Prorratear costos cuando se detecten unidades bonificadas.
- Mantener descuentos calculados individualmente por artículo.
- Omitir productos que no tengan una coincidencia suficiente en el catálogo.
- Registrar productos omitidos para su creación posterior en Effi.
- Generar una hoja `Plantilla_Importacion` sin fórmulas.
- Generar hojas de auditoría y un archivo de log.
- Validar calidad, formato, pruebas y dependencias mediante GitHub Actions.

> **Principio de seguridad:** ante una coincidencia dudosa, el sistema debe preferir marcar el artículo para revisión antes que asignarle un código Effi incorrecto.

## Requisitos

- Python 3.10 o superior.
- Git.
- Cuenta/repositorio de GitHub para CI/CD.
- Tesseract OCR si se procesan imágenes o PDF escaneados.
- Sistema operativo Windows, macOS o Linux.

## Instalación

```bash
git clone <https://github.com/eriksc2006/Script-Facturas.git>
cd effi-processor

python -m venv .venv
```

### Windows

```powershell
.venv\Scripts\activate
```

### Linux/macOS

```bash
source .venv/bin/activate
```

Instalar dependencias:

```bash
python -m pip install --upgrade pip
pip install -r requirements.txt
```

Copiar la configuración:

```bash
cp .env.example .env
```

En Windows PowerShell:

```powershell
Copy-Item .env.example .env
```

## Uso

La interfaz principal puede ejecutarse con Streamlit:

```bash
streamlit run app_effI.py
```

Si el archivo principal se reorganiza como módulo dentro de `src`, puede ejecutarse según la entrada definida por el proyecto.

### Flujo recomendado

1. Cargar el catálogo maestro Effi.
2. Cargar la factura/listado.
3. Revisar las columnas detectadas.
4. Procesar.
5. Revisar las alertas y coincidencias dudosas.
6. Revisar `Calculos_y_Verificaciones`.
7. Confirmar que los productos omitidos hayan sido creados en Effi antes de reprocesar.
8. Descargar la plantilla final.
9. Conservar el log como evidencia de ejecución.

## Variables de entorno

| Variable | Valor ejemplo | Descripción |
|---|---:|---|
| `EFFI_TAX_RATE` | `0.19` | Tasa de IVA utilizada por defecto. |
| `EFFI_MATCH_THRESHOLD` | `0.82` | Confianza mínima para aceptar una coincidencia automática. |
| `EFFI_LOG_LEVEL` | `INFO` | Nivel de detalle del log. |

**Nunca** suba `.env`, credenciales, tokens o documentos comerciales reales al repositorio.

## Tests

Ejecutar todas las pruebas:

```bash
pytest
```

Con cobertura:

```bash
pytest --cov=src/effi_processor --cov-report=term-missing
```

## Calidad de código

Formatear:

```bash
black src tests
```

Verificar formato:

```bash
black --check src tests
```

Lint:

```bash
flake8 src tests
```

Seguridad:

```bash
bandit -r src -ll
pip-audit -r requirements.txt
```

## GitHub Actions

El workflow `.github/workflows/ci.yml` se ejecuta automáticamente en:

- Pull Requests hacia `main`.
- Pushes hacia `main`.

Comprueba:

1. Black.
2. Flake8.
3. Pytest + cobertura.
4. Bandit.
5. `pip-audit`.
6. Gitleaks para detección de secretos.

Además, las pruebas se ejecutan sobre Python 3.10, 3.11, 3.12 y 3.13.

## Estructura

```text
effi-processor/
├── .github/
│   └── workflows/
│       └── ci.yml
├── src/
│   └── effi_processor/
│       ├── __init__.py
│       └── quality_rules.py
├── tests/
│   └── test_quality_rules.py
├── .env.example
├── .gitignore
├── .pre-commit-config.yaml
├── pyproject.toml
├── requirements.txt
├── requirements-dev.txt
└── README.md
```

## Política de calidad

Una modificación no debe fusionarse a `main` si:

- Rompe las pruebas.
- Introduce errores de linting.
- No cumple el formato Black.
- Introduce vulnerabilidades detectables.
- Expone secretos.
- Modifica accidentalmente la estructura de la plantilla Effi.

## Licencia

Añadir aquí la licencia elegida por el propietario del proyecto.
