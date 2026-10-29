# Hoja de ruta de Capital Humano

**Propósito:** priorizar mejoras que hagan el análisis salarial más confiable, auditable, seguro y fácil de usar. Inspirada en el modelo de supervisión por etapas del repositorio [yhquintero/inventario](https://github.com/yhquintero/inventario), adaptada al contexto salarial y a la sensibilidad de los datos de personal.

> **Alcance realista:** una hoja de ruta no garantiza un puesto de “Top 1”. Ese reconocimiento depende de adopción, resultados medibles, evaluaciones independientes y necesidades de los usuarios. No se publican datos salariales ni se envían sugerencias fuera de la app. Cada mejora se considera terminada solo después de pasar su criterio de aceptación.

## Estado actual y decisiones de producto

La aplicación es una herramienta de modelación y apoyo basada en normativa y archivos de origen cubanos. Mantiene el límite absoluto de 6×, el mínimo del grupo I y la exportación imprimible. No hay un censo de plazas ocupadas/vacantes, así que **no se estimará el costo total de nómina** a partir de los salarios por grupo. Las funciones multiusuario, presupuestos y datos personales requieren primero controles de acceso, autorización del responsable y diseño de privacidad.

## Propuesta priorizada

La prioridad indica urgencia relativa, no una fecha comprometida. Alcance y esfuerzo se estiman antes de iniciar cada etapa; se trabaja una mejora a la vez.

**Estados:** *Integrada* significa fusionada en `main` con la verificación indicada aprobada; *En curso* identifica una mejora autorizada aún no integrada; *Parcial* identifica solo los componentes ya disponibles; *Propuesta* no autoriza por sí sola el inicio; *Bloqueada* requiere resolver primero su condición indicada.

| ID | Prioridad | Mejora propuesta | Beneficio | Criterio de aceptación / verificación | Estado |
|---|---|---|---|---|---|
| CH-01 | P0 · inmediata | Pruebas automatizadas de reglas salariales y flujo UI en GitHub Actions | Evita regresiones del tope 6× y del mínimo legal | CI pasa pruebas unitarias y recorre cada sección; falla si se rompe la regla | Integrada en main vía PR #5; CI aprobada |
| CH-02 | P0 · inmediata | Comparar dos escenarios (1×–6×) por grupo, con CSV/Excel | Permite evaluar alternativas antes de aprobar un factor | 32 grupos visibles, grupos sin propuesta sin cambios, factores acotados a 6× y exportaciones verificables | Integrada en main vía PR #5 |
| CH-03 | P0 · inmediata | Plantilla de pull request, formularios de errores/ideas y Dependabot | Mejora trazabilidad y revisión de código/dependencias | Solicitudes incluyen impacto, pruebas y estado; actualizaciones de dependencias llegan como PR | Integrada en main vía PR #5 |
| CH-04 | P1 · alta | Validación de integridad de las fuentes y huellas SHA-256 en exportaciones | Identifica el origen exacto de cada informe | PDF/XLSX incluyen manifiesto, ZIP de CSV lleva JSON lateral; comparación contra hashes aprobados y alerta ante cambios | Integrada en main vía PR #12; CI aprobada |
| CH-05 | P1 · alta | Exportación de informe de aprobación con metodología, versión y espacios de firma | Facilita revisión formal sin presentar el modelo como decisión oficial | PDF A4 identifica escenario, supuestos, fuentes/hash completos, incidencias y campos manuales para decisión; permanece marcado como borrador | Integrada en main vía PR #14; CI aprobada |
| CH-06 | P1 · alta | Usabilidad y accesibilidad con objetivo WCAG 2.2 AA (teclado, contraste, reflujo y lector de pantalla) | Mejora el uso en distintas capacidades y dispositivos | Definir flujos representativos; análisis automático sin hallazgos graves/críticos pendientes; comprobar teclado sin trampas y foco visible, contraste mínimo AA, reflujo a 320 px y etiquetas/errores; probar flujos manualmente con lector de pantalla; registrar herramientas, resultados y excepciones sin afirmar certificación | En curso, autorizada; estilos y prueba automatizada iniciales; falta auditoría en navegador, teclado, reflujo y lector de pantalla |
| CH-07 | P1 · alta | Presupuesto de nómina con plazas ocupadas/vacantes y componentes definidos | Permite calcular un impacto fiscal real, no una suma engañosa de escalas | Fuente autorizada de plazas, reglas de contratación y fórmulas revisadas por Finanzas antes de publicar totales | Bloqueada por datos y aprobación del responsable |
| CH-08 | P1 · alta | Acceso individual, roles y registro de acciones | Requisito para colaboración sobre información salarial sensible | Pruebas de autorización por rol, sesiones protegidas, bitácora y revisión de seguridad antes de producción | Propuesta; exige diseño de servidor |
| CH-09 | P2 · media | Importar nueva plantilla con vista previa y validación de columnas | Reduce el trabajo manual ante cambios de archivo | Detecta formato inválido, duplicados y salarios no positivos; nunca reemplaza fuentes sin confirmación | Propuesta |
| CH-10 | P2 · media | Guardar, nombrar y volver a abrir escenarios de forma segura | Evita rehacer análisis | Persistencia opt-in, aislamiento por usuario, exportación/restauración y pruebas de privacidad | Propuesta; requiere persistencia segura |
| CH-11 | P2 · media | Interfaz ES/EN y formatos regionales sin traducir ni cambiar la normativa fuente | Facilita colaboración internacional y mantiene el contexto cubano | Textos traducidos revisados; CUP, separadores y referencias legales inequívocos | Propuesta |
| CH-12 | P2 · media | Operación y recuperación: healthcheck, monitorización y runbook de incidentes | Ayuda al responsable de despliegue a detectar y resolver fallos | Distinguir healthcheck del contenedor de la disponibilidad externa; probar alertas sin datos personales; documentar diagnóstico, recuperación y simulacro de indisponibilidad | Parcial: hay healthchecks en Dockerfile y Compose y renovación TLS documentada en [HTTPS.md](HTTPS.md); monitorización externa, alertas y runbook de incidentes siguen pendientes |
| CH-13 | P3 · evolución | API y modelo de datos con versionado para integrar fuentes autorizadas | Abre interoperabilidad sin acoplar la interfaz a archivos locales | Contrato documentado, control de acceso y pruebas de compatibilidad hacia atrás | Propuesta; requiere arquitectura y aprobación |
| CH-14 | P3 · evolución | Revisión independiente de seguridad, accesibilidad y exactitud normativa | Aporta evidencia externa sobre controles ya implementados, sin sustituir las pruebas internas de cada mejora | Acordar alcance y especialista; registrar hallazgos con responsable y fecha; verificar correcciones y retest; publicar límites del examen sin afirmar certificación no obtenida | Propuesta; requiere especialistas y aprobación del alcance |
| CH-15 | P3 · evolución | Firma digital de releases e informes con clave institucional protegida externamente | Aporta autenticidad y detección de sustitución del artefacto | Verificar una firma con clave pública publicada; custodia/rotación de la clave definida por el titular | Propuesta; requiere gobierno de claves |
| CH-16 | P1 · alta | Informe de calidad de la plantilla de origen | Hace visibles errores y casos ambiguos antes de revisar o usar la propuesta | Detecta salarios vacíos/no válidos, cargos sin descripción, posibles duplicados y discrepancias de grupo; cada hallazgo incluye hoja/fila; tabla y descarga rastreables; no corrige ni elimina filas | Integrada en main vía PR #17; CI aprobada |
| CH-17 | P2 · media | Comparación A/B por cargo con diferencias nominales y relativas | Permite revisar el efecto de factores alternativos fila por fila | Cada cargo conserva ID y origen; muestra CUP y porcentaje; topes de 6×; CSV/XLSX con trazabilidad; no estima costo de nómina | Integrada en main vía PR #18; CI aprobada |
| CH-18 | P1 · alta | Preflight de integridad antes de exportar | Evita distribuir archivos cuando conteos, IDs, factor, topes o fuentes no pasan controles | Tabla de validaciones; las descargas se bloquean ante errores; no repara datos ni altera fuentes | Integrada en main vía PR #18; CI aprobada |
| CH-19 | P2 · media | Revisión humana de incidencias de calidad | Documenta estado y contexto de revisión sin cambiar la fuente | Estados y notas editables durante la sesión; la exportación conserva anotaciones; los campos de origen permanecen inmutables | Integrada en main vía PR #18; CI aprobada |
| CH-20 | P2 · media | Paquete ZIP integral de escenario | Reduce pasos y mantiene juntos los documentos y su contexto | Un ZIP incluye Excel, PDF, borrador de aprobación, parámetros, incidencias anotadas, manifiesto de fuentes y hashes SHA-256 verificables | Integrada en main vía PR #18; CI aprobada |
| CH-21 | P1 · alta | Verificador offline de integridad del paquete | Permite revisar un ZIP descargado sin abrir ni extraer datos salariales | Comprueba estructura requerida, nombres seguros, tamaños y SHA-256 de cada archivo; detecta alteraciones, faltantes y extras; no imprime contenidos ni afirma firma/autenticidad | Integrada en main vía PR #19; CI aprobada |

## Secuencia recomendada y puertas de avance

1. **P0 — completada:** CI, comparación de escenarios y flujo de contribuciones están integrados; no reabrirlos salvo regresión o cambio de requisito.
2. **P1 — informes, completada:** trazabilidad y documento de aprobación están integrados. **CH-06 está en curso y autorizada**; se definieron cinco flujos representativos y se añadió una base inicial de contraste, foco visible y reflujo. La auditoría de navegador, teclado y lector de pantalla sigue pendiente.
3. **P1 — calidad de fuente, integrada:** CH-16 está en `main` vía PR #17; los hallazgos siguen siendo informativos y no modifican la plantilla.
4. **CH-17–CH-20 — ciclo de auditoría y exportación, integrado:** PR #18 está en `main`; mantiene trazabilidad por cargo, preflight que bloquea errores, revisión manual de incidencias y un ZIP con hashes de salida.
5. **CH-21 — verificación offline, integrada vía PR #19:** herramienta local de solo lectura que detecta faltantes, extras, rutas inseguras y discrepancias SHA-256 sin extraer ni mostrar el contenido salarial.
6. **P1 — dependencias externas:** CH-07 no comienza sin fuente autorizada de plazas y aprobación de Finanzas/responsable. CH-08 requiere diseño de arquitectura, roles, protección de sesiones y revisión de seguridad antes de tratar datos salariales en colaboración.
7. **P2 — operación y uso:** priorizar importación o recuperación solo con responsable, necesidad concreta y criterios de aceptación acordados; CH-12 debe cerrar primero sus brechas de monitorización/runbook.
8. **P3 — evolución controlada:** API, revisiones independientes y firma digital requieren alcance aprobado, especialistas o gobierno de claves, según corresponda.

No se inicia una propuesta siguiente mientras otra esté en implementación; cada etapa se cierra con PR integrado y verificación aprobada.

## Cómo proponer o cerrar una mejora

1. Abrir un issue de GitHub con problema, usuario afectado y resultado esperado (sin adjuntar datos personales/salariales).
2. Acordar prioridad, alcance, esfuerzo, responsable, criterio de aceptación y tratamiento de privacidad; una propuesta en la tabla no equivale a autorización para implementarla.
3. Trabajar una mejora a la vez y entregar un pull request acotado, con pruebas y documentación.
4. Marcar “Integrada” solo cuando el criterio está probado, CI pasa y el cambio está en `main`; documentar explícitamente cualquier parte que quede parcial o bloqueada.

## Prácticas internacionales adoptadas / previstas

- GitHub Actions para CI; Dependabot para actualizaciones revisables.
- Pruebas automatizadas de lógica y de interfaz, HTTPS/TLS y contenedores no-root.
- Trazabilidad, mínima exposición de puertos, documentación de privacidad y aprobación humana.
- Accesibilidad WCAG como objetivo por validar, no como certificación autodeclarada.
- Nunca usar “Top 1”, “certificado” o “cumple legalmente” como afirmación sin evidencia independiente.
