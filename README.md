# Herramienta de Transformación ISIS-CSV-ISIS

Esta herramienta facilita la exportación, edición masiva (mediante hojas de cálculo tipo Excel/CSV) y re-importación segura de registros en bases de datos CDS/ISIS (utilizadas comúnmente en sistemas como ABCD). 

Está diseñada para permitir a los catalogadores y administradores realizar correcciones o actualizaciones en lote sobre la base de datos `CAT` (o cualquier otra base ISIS), minimizando los riesgos de corrupción de datos y facilitando un flujo de trabajo que incluye pruebas en un entorno clonado antes de afectar la base en producción.

## Características Principales

*   **Exportación a Excel/CSV**: Extrae los datos de la base maestra (`.mst`/`.xrf`) a un formato amigable para su edición masiva con herramientas ofimáticas comunes.
*   **Edición Desacoplada**: Permite trabajar sobre un archivo `editable` fuera de línea, a tu propio ritmo, sin bloquear la base de datos original.
*   **Generación de ISO**: Convierte las modificaciones del Excel de nuevo a un formato de intercambio estándar (ISO 2709) listo para ser ingerido por ISIS.
*   **Aplicación Segura de Cambios**: Actualiza la base de datos inyectando únicamente los registros modificados, y sugiere probar primero en un clon de la base.

## Documentación

El proyecto cuenta con documentación detallada para diferentes niveles de profundidad. Se recomienda leer los siguientes archivos antes de operar la herramienta:

*   [**PASO_A_PASO_SIMPLE.md**](PASO_A_PASO_SIMPLE.md): Una guía rápida enumerando los pasos exactos a seguir para una corrección estándar. Ideal para uso diario.
*   [**PROCEDIMIENTO.md**](PROCEDIMIENTO.md): El manual completo que detalla el funcionamiento interno, explicaciones de las opciones, solución de problemas y justificaciones de diseño.
*   [**ESPECIFICACION_CSV.md**](ESPECIFICACION_CSV.md): Documentación técnica sobre la estructura y el formato esperado de los archivos CSV generados y procesados.

## Uso Básico

La interfaz principal de la herramienta se ejecuta a través del script principal:

```bat
REPARADOR_CAT.bat
```

*(En entornos donde no se use el script .bat, se puede invocar directamente mediante `python reparador.py`)*

El flujo normal de trabajo consta de 3 fases principales que se van guiando desde el menú:
1.  **Exportar**: Opción 1 en el menú para generar los archivos Excel/CSV con la base actual.
2.  **Editar**: Modificar el archivo Excel `editable` resultante y guardar los cambios.
3.  **Generar e Inyectar**: Opciones 2 y 3 para empaquetar los cambios en formato ISO, y luego aplicarlos (idealmente primero a un clon de prueba, y finalmente a la base real).

Para un tutorial más detallado paso a paso, consulta el archivo [PASO_A_PASO_SIMPLE.md](PASO_A_PASO_SIMPLE.md).
