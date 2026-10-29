# Importar una plantilla para revisión

La opción **Plantilla por cargo → Importar otra plantilla para revisar** acepta libros `.xls` y `.xlsx` sin macros de hasta 15 MiB. El archivo se inspecciona en memoria; la aplicación no lo escribe al repositorio ni a una ruta del servidor.

1. Selecciona el archivo y revisa las comprobaciones, la vista previa de hasta 25 filas y los hallazgos por hoja/fila.
2. Si faltan encabezados de grupo o hay salarios ausentes/no numéricos/no positivos, la app no permite activar el libro.
3. Los posibles duplicados se señalan pero permanecen intactos. Las discrepancias se muestran; la herramienta no corrige valores ni elimina filas.
4. **Usar temporalmente en esta sesión** sustituye los datos activos solo en esa sesión de Streamlit. No cambia el archivo oficial ni el manifiesto aprobado.
5. La carga temporal queda identificada por SHA-256 como **NO APROBADA · SOLO SESIÓN**. Se bloquean todas las exportaciones salariales hasta volver a la plantilla oficial o hasta que una revisión institucional apruebe formalmente la fuente y su manifiesto.

El lector limita tamaño comprimido y expandido, entradas ZIP, hojas y celdas, y admite únicamente `.xls`/`.xlsx`. Una vista previa válida no confirma la autenticidad, vigencia normativa ni autorización institucional del archivo. No cargues datos personales o salariales a servicios ajenos al despliegue aprobado.
