# Mejoras propuestas — tablero de seguimiento

Lista operativa vinculada a la hoja de ruta de [Capital Humano](ROADMAP.md). Cada renglón debe tener un criterio verificable antes de comenzar. Una propuesta no se describe como realizada hasta que la verificación correspondiente pasa.

| ID | Mejora | Prioridad | Criterio de cierre | Verificación | Estado |
|---|---|---|---|---|---|
| CH-01 | Pruebas automáticas de reglas salariales y navegación | P0 | No exceder 6×; mínimo I $3,210; siete secciones y comparador sin excepción | pytest + Streamlit AppTest en GitHub Actions | Integrada en main por PR #5; CI aprobada |
| CH-02 | Comparador por grupos y exportación CSV/XLSX | P0 | Mostrar A/B, cambio nominal/relativo, conservar grupos sin propuesta y advertir que no es presupuesto de nómina | pruebas de datos, UI y lectura del Excel | Integrado en main por PR #5 |
| CH-03 | Mejorar contribución y revisión en GitHub | P0 | Plantillas de issue/PR y Dependabot activas para Python y Actions | revisar archivos y PR de actualización | Configurado en main por PR #5 |
| CH-04 | Huella de fuentes en cada exportación | P1 | PDF/XLSX incluyen SHA-256; ZIP del CSV incluye JSON lateral; cambios contra manifiesto aprobado disparan advertencia | pruebas contra hashes reales, fuentes modificadas, lectura de Excel/ZIP/PDF y build Docker | Integrada en main por PR #12; CI aprobada |
| CH-05 | Informe imprimible de metodología y aprobación | P1 | Identificar escenario, fuentes y hashes completos, supuestos, incidencias y espacios de revisión/firma; permanecer como borrador y no aparentar aprobación | extracción y lectura del PDF, verificación de contenido y prueba AppTest de descarga | Integrada en main por PR #14; CI aprobada |
| CH-06 | Accesibilidad y uso móvil | P1 | navegación por teclado, contraste, etiquetas y diseño adaptable | auditoría automatizada + inspección manual | Pendiente |
| CH-07 | Presupuesto por plazas reales | P1 | Solo mostrar totales tras importar un recuento de plazas autorizado y validar fórmulas con Finanzas | casos con plazas ocupadas/vacantes conocidas | Bloqueado por fuente de datos |
| CH-08 | Usuarios, roles y bitácora | P1 | aislar permisos y probar accesos denegados; revisar seguridad antes de despliegue | pruebas autorización y revisión externa | Pendiente; rediseño servidor |
| CH-09 | Carga guiada de plantilla actualizada | P2 | vista previa, validación, manejo de discrepancias y confirmación antes de usar | libro válido, inválido y con duplicados | Pendiente |
| CH-10 | Escenarios con nombre y persistencia | P2 | persistencia voluntaria con controles de acceso, exportar y recuperar sin pérdida | pruebas de recuperación y privacidad | Pendiente |
| CH-11 | Idiomas ES/EN y formato internacional | P2 | interfaz traducida y referencias legales originales intactas | inspección de textos/números en ambos idiomas | Pendiente |
| CH-12 | Operación y recuperación del servidor | P2 | instrucciones para healthcheck, logs, renovación TLS y recuperación | simulación documentada de indisponibilidad | Parcial |
| CH-13 | API versionada e integraciones | P3 | contrato estable con control de acceso y compatibilidad | pruebas de contrato y autorización | Pendiente |
| CH-14 | Revisión independiente de seguridad y accesibilidad | P3 | registrar hallazgos, responsable y prueba de remediación | evaluación externa fechada | Pendiente |
| CH-15 | Firma digital de releases e informes | P3 | Verificar firma con clave pública publicada; custodiar y rotar la clave privada fuera del repo | verificación criptográfica reproducible por una segunda persona | Pendiente; requiere gobierno de claves |

## Regla de aceptación

- Toda propuesta de cálculo conserva el límite absoluto 6× y prueba el mínimo del grupo I.
- Totales de nómina no se calculan sin cantidad de plazas aprobada y documentada.
- Los archivos salariales no deben adjuntarse a issues públicos ni enviarse a servicios externos.
- El responsable humano verifica las fuentes normativas y aprueba cualquier uso oficial.
