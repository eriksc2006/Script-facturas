# Seguridad

## 1. Objetivo

Effi Processor procesa documentos que pueden contener información comercial, financiera o identificable.

Por ello, la seguridad y privacidad de los archivos deben considerarse parte del diseño del proyecto.

## 2. Información que no debe publicarse

Nunca subir al repositorio:

- Facturas reales.
- Listados reales de compras.
- Catálogos confidenciales.
- Información de clientes o proveedores.
- Credenciales.
- Tokens.
- Contraseñas.
- Claves privadas.
- Archivos `.env`.

## 3. Datos de prueba

Utilizar:

- Datos sintéticos.
- Datos anonimizados.
- Fixtures creados específicamente para pruebas.

## 4. Vulnerabilidades

Si se encuentra una vulnerabilidad de seguridad, evitar publicarla inicialmente en un issue público.

Utilizar un canal privado con el mantenedor del proyecto.

**Contacto de seguridad: PENDIENTE DE CONFIRMAR antes de publicar el repositorio.**

Como alternativa, si el repositorio dispone de la función correspondiente, se recomienda habilitar el mecanismo privado de reporte de vulnerabilidades de GitHub.

## 5. Reporte

El reporte debería incluir:

- Descripción.
- Impacto.
- Pasos para reproducir.
- Versión afectada.
- Evidencia mínima necesaria.
- Propuesta de mitigación, si existe.

No adjuntar información sensible innecesaria.

## 6. Secretos

Los secretos deben mantenerse fuera del código fuente.

El archivo `.env` debe estar excluido mediante `.gitignore`.

El archivo `.env.example` debe contener únicamente nombres de variables y valores de ejemplo no sensibles.

## 7. Documentos de entrada

Los documentos cargados por usuarios pueden contener información sensible. La aplicación debe evitar generar copias innecesarias y debe almacenar salidas únicamente donde esté definido por la implementación.

## 8. Regla de integridad de datos

La aplicación no debe inventar:

- IDs Effi.
- GTIN.
- Precios.
- Cantidades.
- Impuestos.
- Productos.

Ante una duda que no pueda resolverse de forma confiable, debe registrarse una alerta y, cuando corresponda, omitirse el artículo.
ejecutalo
