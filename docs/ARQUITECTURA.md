# Arquitectura — Effi Processor

**Versión:** 0.1.0

## 1. Visión general

Effi Processor separa conceptualmente la interfaz de usuario de la lógica de procesamiento.

La interfaz se encuentra en:

`app_effi.py`

La lógica de negocio se organiza dentro de:

`src/effi_processor/`

Los nombres concretos de módulos internos deben coincidir con la implementación real del repositorio.

## 2. Flujo de datos

```text
Catálogo Effi
      │
      ▼
Carga / validación
      │
      ├──────────────┐
      │              │
      ▼              ▼
Facturas        Documentos
      │              │
      └───────┬──────┘
              ▼
       Extracción
              │
              ▼
       Normalización
              │
              ▼
        Matching
              │
              ▼
 Conversiones y cálculos
              │
              ▼
        Validaciones
              │
       ┌──────┴───────┐
       ▼              ▼
 Plantilla         Auditoría
 Effi              / alertas
       │
       ▼
     XLSX
```

## 3. Capas conceptuales

### 3.1 Interfaz

`app_effi.py`

Responsabilidades:

- Mostrar la interfaz Streamlit.
- Recibir archivos.
- Mostrar opciones.
- Ejecutar el procesamiento.
- Mostrar resultados.
- Permitir descargas.

La interfaz no debería contener la lógica matemática principal.

### 3.2 Ingesta

Responsabilidades:

- Leer PDF.
- Leer imágenes.
- Leer Excel.
- Leer CSV.
- Leer TXT.
- Ejecutar OCR cuando corresponda.

Tecnologías previstas:

- PyMuPDF.
- Pillow.
- pytesseract.
- pandas.
- openpyxl.

### 3.3 Normalización

Responsabilidades:

- Normalizar nombres.
- Normalizar espacios y caracteres.
- Interpretar presentaciones.
- Identificar cantidades por presentación.
- Normalizar unidades.

### 3.4 Matching

Prioridad:

1. GTIN.
2. Código/referencia.
3. Coincidencia de nombre.
4. Fuzzy matching.
5. Validación por presentación.

Tecnología prevista:

`rapidfuzz`

El matching debe devolver, como mínimo conceptualmente:

- candidato.
- nivel de confianza.
- motivo de coincidencia.
- datos utilizados para decidir.

### 3.5 Conversión

Debe contemplar relaciones como:

- g ↔ kg.
- ml ↔ L.
- galón ↔ litros.
- unidades individuales ↔ cajas/presentaciones.

Las equivalencias deben basarse en información explícita del documento y/o catálogo.

No debe realizarse una conversión arbitraria cuando falten datos.

### 3.6 Cálculos

Incluye:

- Cantidades.
- Precio unitario.
- Descuentos.
- Bonificaciones.
- Prorrateos.
- IVA.

Las operaciones deben mantener precisión suficiente durante el cálculo y redondear conforme al formato de salida requerido.

### 3.7 Validación y auditoría

Debe registrar:

- Coincidencias.
- Confianza.
- Conversiones.
- Cálculos.
- Alertas.
- Productos omitidos.
- Origen del documento.

## 4. Exportación

La primera hoja:

`Plantilla_Importacion`

debe contener únicamente valores/texto y no fórmulas.

Los cálculos auxiliares deben permanecer fuera de esta hoja.

La hoja:

`Calculos_y_Verificaciones`

se utiliza para documentar cálculos y verificaciones cuando esté implementada.

## 5. Modelo de procesamiento

Cada artículo puede representarse conceptualmente como:

```text
Documento
 └── Línea de compra
      ├── Descripción original
      ├── Referencia
      ├── GTIN
      ├── Cantidad
      ├── Unidad
      ├── Presentación
      ├── Precio
      ├── Descuento
      ├── IVA
      └── Origen
```

Después del matching:

```text
Línea de compra
      │
      ▼
Producto Effi
      ├── ID Effi
      ├── GTIN
      ├── Presentación normalizada
      ├── Cantidad final
      ├── Precio final
      ├── Descuento
      └── Impuesto
```

## 6. Principio de seguridad del matching

El sistema debe preferir omitir un producto antes que asociarlo incorrectamente.

Regla:

```text
Confianza suficiente
       │
   ┌───┴───┐
   │       │
  Sí       No
   │       │
Exportar  Omitir
```

## 7. Rendimiento y tamaño

El archivo final debe respetar el límite de importación de Effi de 5 MB.

Para documentos grandes se recomienda:

- Procesar por lotes.
- Evitar mantener innecesariamente archivos completos en memoria.
- Registrar el origen de cada línea.
- Generar auditoría de forma controlada.

## 8. Seguridad

Nunca almacenar en Git:

- Facturas reales.
- Credenciales.
- Tokens.
- Catálogos confidenciales.
- Archivos `.env`.

Usar fixtures sintéticos para las pruebas.

## 9. Componentes tecnológicos

| Componente | Tecnología |
|---|---|
| Interfaz | Streamlit |
| Lenguaje | Python >= 3.10 |
| Datos | pandas |
| Excel | openpyxl |
| PDF | PyMuPDF |
| Imágenes | Pillow |
| OCR | Tesseract / pytesseract |
| Extracción LLM local (opcional) | Ollama (`llama3.2` u otro modelo local) |
| Matching | rapidfuzz |
| Pruebas | pytest |
| Formato | black |
| Lint | flake8 |
| Seguridad | bandit / pip-audit |

## 10. Estado de confirmación

Los módulos conceptuales anteriores describen la arquitectura prevista.

Los nombres concretos de clases, funciones y archivos internos de `src/effi_processor/` deben documentarse después de verificar el código real y no deben inventarse.
