# Diseño propuesto de roles

**Estado: borrador de diseño; no está implementado ni protege información todavía.** La aplicación actual no identifica usuarios individualmente. No se debe interpretar que esconder secciones en el cliente constituya autorización.

## Perfiles iniciales

| Perfil | Permisos propuestos | Límites propuestos |
|---|---|---|
| **Administrador de sistema** | Gestionar cuentas, asignaciones de rol y configuración técnica; consultar metadatos de auditoría | No obtiene acceso al detalle salarial por defecto; cualquier acceso excepcional debe asignarse y registrarse explícitamente |
| **Analista de Recursos Humanos** | Consultar fuentes y detalle por cargo; preparar y comparar escenarios; revisar la vista previa de una plantilla temporal; generar informes cuando pase el preflight | No reemplaza ni aprueba por sí mismo la fuente oficial; la importación temporal no modifica datos ni elimina el bloqueo actual de descargas |
| **Revisor de Finanzas** | Consultar escenarios y comparaciones; revisar y descargar los informes autorizados | Solo lectura; no importa plantillas, altera fuentes ni administra cuentas |
| **Consulta / Auditoría** | Consultar estado de integridad, huellas de fuentes y metadatos de revisión autorizados | Sin detalle salarial por defecto, modificaciones, importaciones ni exportaciones; ampliar el acceso requiere aprobación específica |

Esta matriz recoge los perfiles iniciales solicitados, pero **los permisos aún requieren revisión y aprobación del responsable institucional**. El nombre de un rol no debe sustituir una asignación individual, necesidad de acceso ni segregación de funciones.

## Requisitos antes de activar roles

1. Elegir y documentar el proveedor institucional de identidad (por ejemplo, SSO/OIDC o un gateway autenticado administrado por la institución). La selección sigue pendiente.
2. Aprobar la matriz detallada de operaciones y datos permitidos para cada rol, incluidos acceso a salarios individuales, vistas, descargas, importación y aprobación de fuentes.
3. Hacer cumplir la autorización en el servidor y en cada operación/descarga; ocultar controles en la interfaz no es suficiente.
4. Usar identidad individual, sesiones protegidas, expiración/cierre de sesión, asignación de privilegio mínimo y una bitácora que no copie datos salariales.
5. Denegar por defecto cuando no se pueda validar la identidad o el permiso; añadir pruebas positivas y negativas por rol y revisar el diseño de seguridad antes de producción.

No se deben introducir contraseñas compartidas, una lista local de usuarios sin protección, ni un selector de rol visible como sustituto de autenticación. No compartas contraseñas, tokens, claves privadas u otros secretos en issues, pull requests o conversaciones.

## Estado actual

Los roles descritos arriba **no existen en la aplicación**. La autenticación compartida de un reverse proxy, si se configura, no establece por sí sola roles individuales. CH-08 sigue bloqueada hasta elegir proveedor de identidad, aprobar permisos y diseñar la autorización de servidor.
