# Contribuir a Effi Processor

Gracias por contribuir al proyecto.

## 1. Antes de comenzar

Leer:

- `README.md`
- `docs/INSTALACION.md`
- `docs/USO.md`
- `docs/ARQUITECTURA.md`
- `SECURITY.md`

## 2. Preparar el entorno

Crear entorno virtual e instalar dependencias:

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
pip install -r requirements-dev.txt
```

En Windows utilizar la activación equivalente.

## 3. Cambios de código

Antes de realizar un cambio:

- Entender el flujo actual.
- Evitar duplicar lógica.
- Mantener la lógica de negocio fuera de la interfaz cuando sea posible.
- No modificar los nombres de columnas Effi sin validar el formato vigente.
- No introducir códigos Effi inventados.

## 4. Pruebas

Ejecutar:

```bash
pytest
black . --check
flake8 .
bandit -r src app_effi.py
pip-audit
```

## 5. Datos de prueba

No utilizar facturas reales ni información confidencial.

Las pruebas deben utilizar datos sintéticos o anonimizados.

## 6. Cambios en cálculos

Todo cambio que afecte:

- cantidades,
- unidades,
- conversiones,
- bonificaciones,
- descuentos,
- IVA,
- prorrateos,

debe incluir pruebas numéricas que permitan comprobar el resultado.

## 7. Matching

Toda modificación del matching debe comprobar como mínimo:

- coincidencia GTIN.
- coincidencia exacta.
- coincidencia aproximada.
- baja confianza.
- productos similares.
- diferencias de presentación.

## 8. Excel

Comprobar siempre:

- Primera hoja: `Plantilla_Importacion`.
- Primera hoja sin fórmulas.
- Valores correctos.
- Encabezados Effi exactos.
- Auditoría separada.
- Tamaño final compatible con Effi.

## 9. Pull Request

La solicitud de cambio debería incluir:

- Qué se modificó.
- Por qué se modificó.
- Cómo se probó.
- Ejemplos de entrada/salida cuando corresponda.
- Impacto sobre cálculos o matching.

## 10. Convención de commits

Se recomienda utilizar mensajes claros, por ejemplo:

```text
feat: agregar conversión de presentaciones
fix: corregir cálculo de bonificación
docs: actualizar guía de instalación
test: agregar casos de matching
refactor: separar lógica de exportación
```

Esta convención es una recomendación y no una regla obligatoria hasta que el proyecto la adopte formalmente.

## 11. Protección de datos

Nunca subir:

- `.env`
- contraseñas
- tokens
- facturas reales
- catálogos privados
- información personal o financiera

Si accidentalmente se publica información sensible, informar inmediatamente al mantenedor y seguir `SECURITY.md`.
