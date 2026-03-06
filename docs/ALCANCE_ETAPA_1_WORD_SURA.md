# Etapa 1 - Exportacion a Word (Sura)

## Objetivo
Permitir convertir informacion de la matriz Excel cargada en documentos Word para procesos operativos de Sura, manteniendo trazabilidad basica y sin introducir autenticacion en esta etapa.

## Alcance incluido
- Exportar la hoja activa a Word (`.doc`) desde la aplicacion.
- Exportar una fila especifica de la hoja activa a Word (`.doc`).
- Mantener estructura tabular con encabezados `C1..C20`.
- Trabajar sobre la matriz normalizada `1000 x 20`.
- Soporte para archivos con multiples hojas.

## Fuera de alcance (Etapa 1)
- Motor de plantillas `.docx` con placeholders avanzados.
- Firma digital, envio automatico por correo o integracion con BPM.
- Persistencia en base de datos.
- Control de acceso/roles.
- Validaciones de negocio especificas por producto Sura.

## Flujo funcional
1. Usuario carga archivo Excel/CSV (drag and drop o explorar).
2. Usuario selecciona hoja activa.
3. Usuario edita celdas necesarias.
4. Usuario exporta:
   - Hoja activa completa a Word, o
   - Fila puntual a Word.
5. Sistema descarga archivo local.

## Reglas de datos
- Cada hoja se recorta/normaliza a 1000 filas por 20 columnas.
- Se exportan solo filas no vacias para el documento de hoja completa.
- Para exportar fila puntual, el indice debe estar entre 1 y 1000.

## Criterios de aceptacion
- La app descarga archivo Word sin errores visibles en navegador.
- Word abre el archivo y muestra tabla legible.
- Encabezados y valores coinciden con la hoja/fila seleccionada.
- Si la fila solicitada esta vacia o fuera de rango, la app muestra error claro.

## Riesgos y proxima iteracion
- `.doc` basado en HTML cumple rapido, pero no reemplaza plantillas formales.
- Proxima iteracion recomendada: plantillas `.docx` por proceso Sura y mapeo de campos de negocio (`{{campo}}`).
