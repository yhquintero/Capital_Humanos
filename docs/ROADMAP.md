# Hoja de ruta de Capital Humano

**Propósito:** priorizar mejoras que hagan el análisis salarial más confiable, auditable, seguro y fácil de usar. Inspirada en el modelo de supervisión por etapas del repositorio [yhquintero/inventario](https://github.com/yhquintero/inventario), adaptada al contexto salarial y a la sensibilidad de los datos de personal.

> **Alcance realista:** una hoja de ruta no garantiza un puesto de “Top 1”. Ese reconocimiento depende de adopción, resultados medibles, evaluaciones independientes y necesidades de los usuarios. No se publican datos salariales ni se envían sugerencias fuera de la app. Cada mejora se considera terminada solo después de pasar su criterio de aceptación.

## Estado actual y decisiones de producto

La aplicación es una herramienta de modelación y apoyo basada en normativa y archivos de origen cubanos. Mantiene el límite absoluto de 6×, el mínimo del grupo I y la exportación imprimible. No hay un censo de plazas ocupadas/vacantes, así que **no se estimará el costo total de nómina** a partir de los salarios por grupo. Las funciones multiusuario, presupuestos y datos personales requieren primero controles de acceso, autorización del responsable y diseño de privacidad.

## Propuesta priorizada

Escala de impacto/esfuerzo: **Alto / Medio / Bajo**. Los esfuerzos son aproximados y se refinan antes de iniciar cada etapa.

| ID | Prioridad | Mejora propuesta | Beneficio | Criterio de aceptación / verificación | Estado |
|---|---|---|---|---|---|
| CH-01 | P0 · inmediata | Pruebas automatizadas de reglas salariales y flujo UI en GitHub Actions | Evita regresiones del tope 6× y del mínimo legal | CI pasa pruebas unitarias y recorre cada sección; falla si se rompe la regla | Integrada en main vía PR #5; CI aprobada |
| CH-02 | P0 · inmediata | Comparar dos escenarios (1×–6×) por grupo, con CSV/Excel | Permite evaluar alternativas antes de aprobar un factor | 32 grupos visibles, grupos sin propuesta sin cambios, factores acotados a 6× y exportaciones verificables | Integrada en main vía PR #5 |
| CH-03 | P0 · inmediata | Plantilla de pull request, formularios de errores/ideas y Dependabot | Mejora trazabilidad y revisión de código/dependencias | Solicitudes incluyen impacto, pruebas y estado; actualizaciones de dependencias llegan como PR | Configurada en main vía PR #5 |
| CH-04 | P1 · alta | Validación de integridad de las fuentes y huellas SHA-256 en exportaciones | Identifica el origen exacto de cada informe | PDF/XLSX incluyen manifiesto, ZIP de CSV lleva JSON lateral; comparación contra hashes aprobados y alerta ante cambios | Integrada en main vía PR #12; CI aprobada |
| CH-05 | P1 · alta | Exportación de informe de aprobación con metodología, versión y espacios de firma | Facilita revisión formal sin presentar el modelo como decisión oficial | PDF A4 identifica escenario, supuestos, fuentes/hash completos, incidencias y campos manuales para decisión; permanece marcado como borrador | Integrada en main vía PR #14; CI aprobada |
| CH-06 | P1 · alta | Usabilidad y accesibilidad WCAG 2.2 AA (teclado, contraste, lector de pantalla) | Mejora uso en distintas capacidades y dispositivos | Auditoría automática complementada por lista de verificación manual en navegador/lector | Propuesta |
| CH-07 | P1 · alta | Presupuesto de nómina con plazas ocupadas/vacantes y componentes definidos | Permite calcular un impacto fiscal real, no una suma engañosa de escalas | Fuente autorizada de plazas, reglas de contratación y fórmulas revisadas por Finanzas antes de publicar totales | Bloqueada por datos y aprobación del responsable |
| CH-08 | P1 · alta | Acceso individual, roles y registro de acciones | Requisito para colaboración sobre información salarial sensible | Pruebas de autorización por rol, sesiones protegidas, bitácora y revisión de seguridad antes de producción | Propuesta; exige diseño de servidor |
| CH-09 | P2 · media | Importar nueva plantilla con vista previa y validación de columnas | Reduce el trabajo manual ante cambios de archivo | Detecta formato inválido, duplicados y salarios no positivos; nunca reemplaza fuentes sin confirmación | Propuesta |
| CH-10 | P2 · media | Guardar, nombrar y volver a abrir escenarios de forma segura | Evita rehacer análisis | Persistencia opt-in, aislamiento por usuario, exportación/restauración y pruebas de privacidad | Propuesta; requiere persistencia segura |
| CH-11 | P2 · media | Interfaz ES/EN y formatos regionales sin traducir ni cambiar la normativa fuente | Facilita colaboración internacional y mantiene el contexto cubano | Textos traducidos revisados; CUP, separadores y referencias legales inequívocos | Propuesta |
| CH-12 | P2 · media | Estado operativo: healthcheck, disponibilidad y runbook de incidentes | Ayuda al responsable de despliegue a detectar fallos | Monitorización sin datos personales, alertas probadas y pasos de recuperación documentados | Parcial; healthcheck ya existe |
| CH-13 | P3 · evolución | API y modelo de datos con versionado para integrar fuentes autorizadas | Abre interoperabilidad sin acoplar la interfaz a archivos locales | Contrato documentado, control de acceso y pruebas de compatibilidad hacia atrás | Propuesta; requiere arquitectura y aprobación |
| CH-14 | P3 · evolución | Revisiones externas periódicas de seguridad, accesibilidad y exactitud normativa | Aporta evidencia independiente | Hallazgos registrados con responsable, fecha y prueba de cierre | Propuesta; requiere especialistas externos |
| CH-15 | P3 · evolución | Firma digital de releases e informes con clave institucional protegida externamente | Aporta autenticidad y detección de sustitución del artefacto | Verificar una firma con clave pública publicada; custodia/rotación de la clave definida por el titular | Propuesta; requiere gobierno de claves |

## Secuencia recomendada

1. **Asegurar calidad y visibilidad (P0):** CI, comparación de escenarios y flujo de contribuciones en GitHub.
2. **Fortalecer informes (P1):** trazabilidad de fuentes y documento de aprobación; medir accesibilidad manualmente.
3. **Acordar datos y responsabilidades (P1):** solo después de obtener una fuente autorizada de plazas y decisiones sobre acceso individual.
4. **Evolucionar con aprobación (P2–P3):** importación, persistencia, idiomas e integraciones en propuestas pequeñas revisables.

## Cómo proponer o cerrar una mejora

1. Abrir un issue de GitHub con problema, usuario afectado y resultado esperado (sin adjuntar datos personales/salariales).
2. Acordar prioridad, esfuerzo, responsable, criterio de aceptación y tratamiento de privacidad.
3. Entregar en un pull request pequeño con pruebas y documentación.
4. Marcar “Hecha” solo cuando el criterio está probado y el cambio está integrado.

## Prácticas internacionales adoptadas / previstas

- GitHub Actions para CI; Dependabot para actualizaciones revisables.
- Pruebas automatizadas de lógica y de interfaz, HTTPS/TLS y contenedores no-root.
- Trazabilidad, mínima exposición de puertos, documentación de privacidad y aprobación humana.
- Accesibilidad WCAG como objetivo por validar, no como certificación autodeclarada.
- Nunca usar “Top 1”, “certificado” o “cumple legalmente” como afirmación sin evidencia independiente.
