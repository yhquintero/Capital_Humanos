# Registro de cambios

Cambios funcionales y de operación, organizados por versión.

## Sin publicar

- Añadida comparación de dos escenarios salariales por grupo, con gráfico y descargas CSV/Excel.
- Centralizada la regla de cálculo con validación de entradas y límite técnico absoluto 6×.
- Incorporadas pruebas pytest y Streamlit AppTest de reglas, fuentes, comparación y navegación.
- Añadida CI de GitHub para pruebas Python, análisis sintáctico PowerShell y compilación de la imagen Docker.
- Activadas propuestas de mejora, plantilla de pull request y actualizaciones de dependencias mediante Dependabot.
- Publicada hoja de ruta priorizada con criterios de aceptación, límites de datos y necesidades de aprobación humana.
- Añadida trazabilidad SHA-256/versionado a ejemplos y exportaciones: hoja en Excel, página en PDF y manifiesto JSON en ZIP de CSV.
- La app compara los archivos actuales con las huellas aprobadas y advierte si encuentra cambios, faltantes o un manifiesto inconsistente.
- Corregida la carga ambigua de `generar_escala`: la interfaz carga explícitamente el motor que está en su misma carpeta y falla con un mensaje claro si son versiones incompatibles.
- Añadido un Id. correlativo antes de la hoja de origen en listados y exportaciones; los valores nulos se dejan en blanco y no se imprimen como `None`.
- Añadido un informe A4 para revisión y aprobación humana, con método, resumen por grupo, trazabilidad completa y campos manuales de decisión/firma; se marca siempre como borrador.
- En curso: diagnóstico auditable de calidad de plantilla con hallazgos por hoja/fila, conservando la fuente sin autocorrecciones.

> Esta sección documenta cambios en esta rama; la versión no se considera publicada hasta integrar y verificar el pull request.
