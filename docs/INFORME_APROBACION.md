# Informe de metodología y solicitud de aprobación

La sección **Exportaciones** genera un PDF A4 independiente para facilitar la revisión humana del escenario activo. No reemplaza el PDF detallado de escala y cargos ni constituye una resolución, dictamen jurídico o autorización salarial.

## Contenido

1. **Resumen del escenario:** fecha, factor aplicado, cantidad de registros de cargos, versión de la aplicación y estado de integridad.
2. **Resumen salarial por grupo:** base/referencia y propuesta para los grupos de escala. No incluye nombres de cargos ni calcula el gasto total de nómina.
3. **Metodología y límites:** regla de cálculo por cargo, tope absoluto 6×, tratamiento del grupo I y de los grupos sin propuesta, supuestos e incidencias de extracción.
4. **Trazabilidad:** estado, ruta y hashes SHA-256 aprobados/actuales de las fuentes supervisadas.
5. **Registro humano:** casillas para la decisión y espacios de preparación, revisión, autorización, fecha, firma manuscrita y observaciones.

## Uso y responsabilidades

- El PDF se rotula **BORRADOR / PENDIENTE DE REVISIÓN Y APROBACIÓN**. Esa marca también aparece en sus pies de página.
- La persona responsable debe confirmar fuentes, normativa vigente, supuestos y autoridad de quienes completan el registro antes de dar uso institucional al documento.
- Las casillas y firmas están en blanco: la app no inventa nombres, credenciales, firmas digitales ni un estado de aprobación.
- Una fuente discrepante o ausente se indica como **REVISIÓN DE FUENTES REQUERIDA**; el informe no debe distribuirse como aprobado hasta resolverla.
- SHA-256 compara bytes con el manifiesto versionado. No certifica la autoría, vigencia legal ni validez sustantiva de una fuente.
- No se estima un presupuesto de nómina porque no hay un conteo autorizado de plazas ocupadas/vacantes ni componentes validados.

## Verificación del documento

Las pruebas automatizadas comprueban que el PDF contiene metodología, hashes completos, advertencias, decisiones/firmas manuales y que la app muestra una descarga separada para este borrador. La revisión visual de impresión sigue siendo necesaria antes de adoptar un formato institucional definitivo.
