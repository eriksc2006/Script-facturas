# Instalación — Effi Processor

**Versión:** 0.1.0

## 1. Requisitos previos

Instalar:

- Python 3.10 o superior.
- Git.
- Tesseract OCR.
- Idiomas OCR `spa` y `eng`.

La aplicación está diseñada para Windows, macOS y Linux.

## 2. Obtener el proyecto

Clonar el repositorio y entrar en su carpeta:

```bash
git clone <URL_DEL_REPOSITORIO>
cd Script-Facturas
```

## 3. Crear el entorno virtual

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

## 4. Actualizar pip

```bash
python -m pip install --upgrade pip
```

## 5. Instalar dependencias

```bash
pip install -r requirements.txt
```

Para desarrollo:

```bash
pip install -r requirements-dev.txt
```

## 6. Instalar Tesseract

### Ubuntu/Debian

```bash
sudo apt update
sudo apt install tesseract-ocr tesseract-ocr-spa tesseract-ocr-eng
```

### macOS

Instalar Tesseract mediante el gestor de paquetes disponible en el equipo y verificar que los idiomas `spa` y `eng` estén disponibles.

### Windows

Instalar Tesseract y agregar su ejecutable al PATH del sistema, o utilizar la configuración de ruta que implemente la aplicación.

Verificación:

```bash
tesseract --version
tesseract --list-langs
```

## 7. Ejecutar

Desde la raíz del proyecto:

```bash
streamlit run app_effi.py
```

## 8. Verificación inicial

Realizar una prueba con:

- Un catálogo Effi de prueba.
- Una factura sintética.
- Al menos un producto con coincidencia clara.
- Un producto que deba ser omitido por baja confianza.

Verificar que:

- La aplicación abre correctamente.
- El catálogo se carga.
- El documento se procesa.
- La hoja `Plantilla_Importacion` se genera.
- La hoja principal no contiene fórmulas.
- Las verificaciones aparecen en `Calculos_y_Verificaciones`.
- Los productos de baja confianza no se inventan ni se asignan arbitrariamente.

## 9. Errores frecuentes

### `streamlit` no se reconoce

Comprobar que el entorno virtual esté activo:

```bash
python -m streamlit run app_effi.py
```

### Tesseract no encontrado

Ejecutar:

```bash
tesseract --version
```

Si falla, instalar Tesseract y configurar el PATH del sistema.

Si la aplicación implementa una configuración específica para indicar la ruta, documentarla aquí cuando esté confirmada.

### No aparecen coincidencias

Revisar:

- GTIN.
- Nombre del producto.
- Presentación.
- Unidad.
- Cantidad.
- Umbral de coincidencia.
- Contenido del catálogo maestro.

### OCR incorrecto

Preferir, cuando sea posible:

1. Excel/CSV.
2. PDF con texto seleccionable.
3. PDF escaneado de buena calidad.
4. Imagen nítida.

Fotografías inclinadas, borrosas, arrugadas o manuscritas pueden producir errores de OCR.

## 10. Entorno de desarrollo

Instalar:

```bash
pip install -r requirements-dev.txt
```

Ejecutar pruebas:

```bash
pytest
```

Validar formato:

```bash
black . --check
```

Validar estilo:

```bash
flake8 .
```

Validar seguridad:

```bash
bandit -r src app_effi.py
pip-audit
```
