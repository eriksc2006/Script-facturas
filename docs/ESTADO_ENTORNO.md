# Estado del entorno local

Fecha: 2026-09-22

## Confirmado
- Python 3.12.10 instalado desde python.org con la opción "Add Python to PATH" marcada.
- Entorno virtual .venv creado en la raíz del proyecto y activado.
- Dependencias instaladas dentro del venv desde `requirements.txt` y `requirements-dev.txt`.
  - `streamlit`
  - `pytest`
  - `black`
  - `pytesseract`
  - paquetes adicionales del proyecto y herramientas de desarrollo.
- Sistema operativo: Windows.
- Terminal utilizada: Git Bash desde VS Code.
- Rama de trabajo: `develop_danilo`.

## Tesseract OCR verificado
- Tesseract OCR instalado con el instalador oficial de UB-Mannheim para Windows.
- Ruta por defecto detectada: `C:\Program Files\Tesseract-OCR\tesseract.exe`.
- Versión verificada: `tesseract v5.5.3.20260724`.
- Idiomas verificados disponibles: `eng` y `spa`.
- Se definió la variable de entorno `TESSERACT_CMD` para la sesión actual apuntando a `C:\Program Files\Tesseract-OCR\tesseract.exe` porque la ruta no estaba en `PATH` del sistema en esta terminal.

## Pendiente
- Ejecución de pytest desde la raíz del proyecto.
- Ejecución de Streamlit para validar arranque de la app.
