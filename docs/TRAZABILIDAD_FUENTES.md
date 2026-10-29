# Trazabilidad e integridad de las fuentes

La app incluye una comprobación SHA-256 del archivo fuente de la plantilla, la escala histórica, la Gaceta Oficial utilizada como referencia y los módulos de cálculo/interfaz. Los hashes aprobados se guardan en `docs/source_manifest.json`; la verificación compara el contenido actual, tamaño, ruta y versión de la app con ese manifiesto.

## Qué se entrega

- Los Excel de la app y los archivos de ejemplo incluyen una hoja **Trazabilidad SHA-256** con estado, ruta, tamaño y hash aprobado/actual.
- Los PDF incluyen una página de trazabilidad con versión y huellas.
- Las descargas CSV se empaquetan como ZIP con el CSV sin modificaciones y `manifiesto_fuentes.json` como archivo acompañante.
- En **Sistema y HTTPS** se muestra el estado de cada archivo, hash esperado y hash observado. Si hay discrepancia o falta el manifiesto, la app advierte y el informe queda marcado para revisión.

## Actualizar el manifiesto

Solo después de revisar un cambio de fuente/código, confirmar su procedencia y aprobarlo, ejecuta desde la raíz:

```bash
python "tools/actualizar_manifiesto_fuentes.py"
```

Inspecciona el diff de `docs/source_manifest.json` junto al diff del archivo que cambió y registra la revisión en el pull request. La actualización del manifiesto **no valida el contenido**: simplemente acepta las huellas actuales, por lo que no debe ejecutarse para silenciar una discrepancia desconocida.

## Corrección autorizada de dos salarios (2026-10-29)

Tras revisar el CSV de incidencias, el usuario autorizó explícitamente sustituir dos valores singulares por el salario mayoritario de su grupo. Se modificaron únicamente las celdas salariales **J44** de `P-2  Aguas Claras` (grupo XII) y **K51** de `Plantilla Emp` (grupo XV). No se añadieron ni eliminaron filas; los estilos, hojas y nombres definidos del libro se conservaron. Los 183 cargos siguen presentes y la extracción ya no reporta incidencias de calidad.

- SHA-256 de la plantilla antes del cambio: `02399d1abf33cb821829df57f146e9011488ed86a11e539c664a2114c7b8b6cc`.
- SHA-256 de la plantilla después del cambio: `5542e23b705462738e66a82ed1aec8d60727c6d1352438266e620878ab63e473`.

La revisión y corrección fue autorizada por el usuario; no es una corrección automática del proceso de calidad. El manifiesto se actualizó tras verificar el cambio.

## Alcance y límites

- SHA-256 permite detectar si el archivo actual difiere de los bytes fijados en el manifiesto versionado.
- La coincidencia **no prueba autoría ni integridad del historial de Git**, no firma el informe y no verifica que una norma siga vigente o se aplique correctamente.
- No se envían archivos, hashes ni datos salariales a servicios externos.
- Para una garantía de autenticidad superior se necesita firmar releases/commits y custodiar las claves fuera del repositorio; eso requiere acordar antes quién autoriza y administra las firmas.
