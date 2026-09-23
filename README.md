# Effi Processor — Script-Facturas

**Versión:** 0.1.0  
**Propósito:** procesar facturas y listas de compra, relacionarlas con el catálogo maestro de Effi y generar archivos Excel compatibles con la importación de Effi.

## ¿Qué hace?

Effi Processor permite:

- Cargar un catálogo maestro Effi en CSV/XLSX.
- Procesar una o varias facturas o listas de compra.
- Trabajar con PDF, imágenes, Excel, CSV y TXT.
- Extraer información mediante lectura de archivos, OCR (Tesseract) y Vision LLM local (Ollama) para documentos escaneados/imágenes.
- Priorizar coincidencias por GTIN y utilizar coincidencia aproximada cuando no existe GTIN.
- Normalizar presentaciones y unidades.
- Resolver conversiones como g/ml/kg/L, galón y presentaciones equivalentes cuando la información disponible permite hacerlo.
- Prorratear correctamente costos de bonificaciones.
- Aplicar descuentos individuales.
- Manejar IVA configurable.
- Omitir coincidencias de baja confianza para evitar asignaciones incorrectas.
- Generar una plantilla Excel para Effi y archivos de auditoría/verificación.

## Regla de seguridad de datos

El sistema **no debe inventar códigos Effi**. Cuando una coincidencia no alcanza el nivel de confianza establecido, el artículo debe quedar fuera de la plantilla principal y registrarse para revisión.

## Requisitos

- Python >= 3.10
- Tesseract OCR para procesamiento OCR (respaldo)
- Ollama (opcional, recomendado) para Vision LLM local gratuito
- Git
- Windows, macOS o Linux

## Instalación rápida

Crear entorno virtual:

### Windows PowerShell

```powershell
py -3.10 -m venv .venv
.venv\Scripts\Activate.ps1
```

### Windows CMD

```cmd
py -3.10 -m venv .venv
.venv\Scripts\activate.bat
```

### macOS / Linux

```bash
python3 -m venv .venv
source .venv/bin/activate
```

Instalar dependencias:

```bash
python -m pip install --upgrade pip
pip install -r requirements.txt
```

Para desarrollo:

```bash
pip install -r requirements-dev.txt
```

## Tesseract OCR

Debe estar instalado en el sistema y disponer de los idiomas español e inglés.

En Ubuntu/Debian:

```bash
sudo apt update
sudo apt install tesseract-ocr tesseract-ocr-spa tesseract-ocr-eng
```

Comprobar:

```bash
tesseract --version
tesseract --list-langs
```

Debe aparecer, como mínimo, `spa` y `eng`.

> La aplicación acepta `TESSERACT_CMD` / `TESSERACT_PATH` si Tesseract no está en el PATH.

## Vision LLM local (Ollama, gratis)

Para mejorar la lectura de PDF escaneados e imágenes **sin API cloud**, use Ollama en su PC. Las facturas se procesan en `localhost` y no se envían a internet.

1. Instale [Ollama](https://ollama.com).
2. Descargue el modelo por defecto:

```bash
ollama pull moondream
```

3. En la barra lateral de la app elija **Lectura de documentos**:
   - `Auto (Vision si hay Ollama)` — intenta Vision y si falla usa Tesseract
   - `Solo Vision` — exige Ollama
   - `Solo Tesseract` — OCR clásico

Variables opcionales en `.env` (ver `.env.example`):

- `OLLAMA_BASE_URL=http://127.0.0.1:11434`
- `OLLAMA_VISION_MODEL=moondream`

## Ejecutar la aplicación

```bash
streamlit run app_effi.py
```

Streamlit mostrará la dirección local de la aplicación.

## Flujo de trabajo

1. Cargar el catálogo maestro de Effi.
2. Cargar una o varias facturas/documentos.
3. Revisar la vista previa.
4. Ajustar IVA y umbral de coincidencia, si están disponibles.
5. Ejecutar el procesamiento.
6. Revisar productos encontrados, omitidos y verificaciones.
7. Descargar el Excel.
8. Revisar la hoja de auditoría.
9. Crear en Effi los productos omitidos cuando corresponda.
10. Reprocesar si es necesario.
11. Importar el archivo en Effi.

## Excel de salida

La primera hoja debe llamarse exactamente:

`Plantilla_Importacion`

Esta hoja está destinada exclusivamente a la importación y debe contener **valores/texto planos**, sin fórmulas de Excel.

Entre los encabezados confirmados del formato Effi se encuentran:

- `Artículo (ID EFFI | Código de barras GTIN | Serie)`
- `Observación`
- `Cantidad *`
- `Precio ud. *`
- `Valor descuento total. *`
- `Código Effi Impuesto`

> Los 9 encabezados completos deben confirmarse contra la plantilla vigente de importación de Effi antes de publicar una versión definitiva. No se inventan aquí los encabezados restantes.

Los cálculos, conversiones, alertas y omisiones deben quedar fuera de la hoja principal, principalmente en:

`Calculos_y_Verificaciones`

y/o en las hojas de auditoría que realmente estén implementadas.

## Reglas de cálculo

### Bonificaciones

Cuando una compra incluye unidades pagadas y unidades gratuitas, el costo real debe distribuirse sobre el total de unidades recibidas cuando corresponda.

Ejemplo:

- 10 unidades pagadas.
- 2 unidades gratis.
- Total recibido: 12.
- Valor pagado: $100.000.

Costo unitario efectivo:

`100.000 / 12 = 8.333,33`

Si la bonificación es únicamente un concepto promocional sin costo y la lógica definida para el caso requiere descuento total, debe registrarse como tal en la auditoría.

### Descuentos

Los descuentos individuales deben conservarse por producto. No se debe aplicar un descuento global uniforme a todos los artículos salvo que el documento indique realmente que es global y las reglas de procesamiento lo permitan.

### IVA

El IVA debe ser configurable. Si el precio recibido ya incluye IVA, primero se debe obtener la base neta:

`Base = Precio con IVA / (1 + IVA)`

Con IVA del 19 %:

`Base = Precio con IVA / 1,19`

## Coincidencias

Prioridad recomendada:

1. GTIN exacto.
2. Código/referencia exacta cuando corresponda.
3. Nombre normalizado.
4. Coincidencia aproximada/semántica.
5. Conversión de presentación/unidad.
6. Validación de confianza.

Las coincidencias de baja confianza se omiten de la plantilla principal.

## Estructura del proyecto

```text
Script-Facturas/
├── app_effi.py
├── requirements.txt
├── requirements-dev.txt
├── .env.example
├── .gitignore
├── README.md
├── CONTRIBUTING.md
├── SECURITY.md
├── docs/
│   ├── INSTALACION.md
│   ├── USO.md
│   └── ARQUITECTURA.md
├── src/
│   └── effi_processor/
├── tests/
└── salidas/
```

Los nombres internos concretos de los módulos dentro de `src/effi_processor/` deben mantenerse alineados con la implementación real.

## Calidad y pruebas

Comandos recomendados:

```bash
pytest
black . --check
flake8 .
bandit -r src app_effi.py
pip-audit
```

## Seguridad y privacidad

No subir al repositorio:

- Facturas reales.
- Catálogos reales con información confidencial.
- Credenciales.
- Tokens.
- Contraseñas.
- Archivos `.env`.
- Archivos generados con información sensible.

Usar datos sintéticos para pruebas.

## Licencia

**Pendiente de confirmar.** Definir la licencia del proyecto antes de realizar una publicación pública.
