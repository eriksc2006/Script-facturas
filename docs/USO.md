# Uso de Effi Processor

**Versión:** 0.1.0

## 1. Objetivo

Esta aplicación permite transformar facturas o listas de compra en una plantilla de importación para Effi, utilizando el catálogo maestro para identificar los productos y realizando las verificaciones necesarias antes de exportar.

## 2. Interfaz de usuario

La interfaz está construida con Streamlit y se encuentra en el punto de entrada:

`app_effi.py`

La documentación de esta interfaz se mantiene en este archivo.

## 3. Carga del catálogo Effi

El primer paso es cargar el catálogo maestro exportado desde Effi.

Formatos previstos:

- CSV.
- XLSX.

El catálogo es la fuente de referencia para:

- Identificar productos.
- Obtener el identificador Effi.
- Validar códigos.
- Comparar nombres y presentaciones.
- Resolver coincidencias.

## 4. Carga de documentos

Se pueden cargar una o varias fuentes de compra, según los formatos soportados por la implementación:

- PDF.
- Imagen.
- Excel.
- CSV.
- TXT.

Cuando se procesan varios documentos, debe conservarse el origen de cada registro para facilitar la auditoría.

## 5. Configuración

La barra lateral puede incluir parámetros como:

- Porcentaje de IVA.
- Umbral mínimo de coincidencia.
- Otras opciones disponibles en la versión instalada.

No asumir parámetros que no estén implementados en `app_effi.py`.

## 6. Vista previa

Antes de generar el Excel se debe revisar:

- Producto detectado.
- Presentación.
- Unidad.
- Cantidad.
- Precio.
- Descuento.
- IVA.
- Coincidencia encontrada.
- Confianza de la coincidencia.
- Documento de origen.

## 7. Procesamiento

El procesamiento sigue conceptualmente este flujo:

1. Lectura del documento.
2. Extracción de información.
3. Normalización.
4. Identificación del producto.
5. Conversión de unidades/presentaciones.
6. Aplicación de descuentos.
7. Tratamiento de bonificaciones.
8. Cálculo de IVA.
9. Validación.
10. Generación del Excel.
11. Registro de auditoría y alertas.

## 8. Coincidencias

La prioridad es:

1. GTIN.
2. Código/referencia cuando corresponda.
3. Nombre normalizado.
4. Coincidencia aproximada.
5. Validación por presentación/unidad.

Nunca se debe crear un código Effi por inferencia.

Si la confianza es insuficiente, el producto se omite de `Plantilla_Importacion`.

## 9. Conversiones de unidades

El sistema debe analizar la presentación completa y no únicamente el texto de la unidad.

Ejemplos:

- `1 L`.
- `1000 ml`.
- `1 galón`.
- `3,785 L`.
- `4 L`.
- `12 UNID`.
- `12 g x 24 unidades`.

Cuando una presentación es ambigua, debe utilizarse la equivalencia disponible en el catálogo y dejar constancia de la decisión en la auditoría.

## 10. Bonificaciones

Ejemplo:

Compra:

- 10 unidades pagadas.
- 2 unidades gratis.
- Total: 12 unidades.
- Valor pagado: $100.000.

Costo efectivo:

`$100.000 / 12 = $8.333,33 por unidad`

La lógica final debe reflejarse en la auditoría.

## 11. Descuentos

Los descuentos individuales deben conservarse por artículo.

No repartir un descuento global sobre productos omitidos. Si existen productos que no pueden identificarse y se omiten, los artículos válidamente exportados deben conservar sus valores calculados según las reglas aplicables a cada uno.

## 12. IVA

Si el valor recibido incluye IVA:

`Base = Precio con IVA / (1 + tasa IVA)`

Para 19 %:

`Base = Precio con IVA / 1,19`

Si el precio ya corresponde a una base neta, no se debe volver a descontar IVA.

## 13. Productos omitidos

Un producto puede omitirse cuando:

- No existe coincidencia suficientemente confiable.
- El GTIN no coincide y el nombre/presentación no permiten validar la identidad.
- La conversión de presentación no puede determinarse con seguridad.
- Existe ambigüedad entre varios productos.

Los omitidos deben quedar registrados para revisión.

## 14. Archivo de salida

La hoja principal debe llamarse exactamente:

`Plantilla_Importacion`

Debe contener solamente valores/texto planos y no fórmulas.

Encabezados confirmados:

- `Artículo (ID EFFI | Código de barras GTIN | Serie)`
- `Observación`
- `Cantidad *`
- `Precio ud. *`
- `Valor descuento total. *`
- `Código Effi Impuesto`

> El listado completo de 9 columnas debe confirmarse con la plantilla vigente de Effi antes de fijarlo en documentación o código.

La hoja secundaria de verificaciones debe llamarse:

`Calculos_y_Verificaciones`

Los nombres de hojas adicionales deben corresponder únicamente a hojas realmente implementadas.

## 15. Importación en Effi

Formatos previstos para el archivo de importación:

- XLSX.
- XLSM.
- XLS.
- XLT.

El archivo debe mantenerse dentro del límite de **5 MB** indicado para la importación.

Antes de importar:

1. Abrir el archivo.
2. Revisar `Plantilla_Importacion`.
3. Confirmar cantidades y precios.
4. Revisar descuentos.
5. Revisar IVA.
6. Revisar productos omitidos.
7. Revisar alertas.
8. Confirmar que la primera hoja contiene únicamente valores.
9. Importar en Effi.

## 16. Logs y salidas

Los archivos generados localmente deben almacenarse en `salidas/` cuando esa sea la configuración implementada.

Los logs deben facilitar la identificación de:

- Documento procesado.
- Cantidad de artículos.
- Coincidencias.
- Omisiones.
- Errores.
- Advertencias.
- Archivo generado.

## 17. Recomendación de uso

Para máxima precisión:

- Preferir Excel/CSV cuando estén disponibles.
- Utilizar documentos PDF con texto seleccionable.
- Usar imágenes nítidas.
- Revisar productos omitidos.
- No reducir el umbral de coincidencia solo para aumentar el número de resultados.
- Conservar los documentos originales como respaldo, fuera del repositorio.
