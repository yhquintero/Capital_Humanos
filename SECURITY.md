# Seguridad

## Versiones con soporte

Se da seguimiento a la rama `main`. No hay una política de versiones publicadas con soporte independiente; mantén una copia local actualizada y revisa los avisos del repositorio antes de desplegar.

## Reportar una vulnerabilidad

**No publiques detalles de una vulnerabilidad en un issue, discusión o pull request público.** Usa los [avisos de seguridad privados de GitHub](https://github.com/yhquintero/Capital_Humanos/security/advisories/new) para contactar al mantenedor de forma confidencial. Incluye pasos mínimos para reproducir el problema, el impacto y una mitigación posible; no adjuntes datos personales, plantillas reales ni exportaciones confidenciales.

Si la función de aviso privado no está disponible, abre un issue público solicitando un canal privado, sin describir el problema técnico.

No existe un plazo de respuesta o una garantía de corrección. Evita pruebas que accedan, alteren o exfiltren datos de terceros. No sometas servicios en producción a pruebas sin autorización expresa.

## Consideraciones de datos

No despliegues la aplicación públicamente sin controles de acceso apropiados. Los archivos cargados se analizan en memoria según la implementación actual, pero los operadores deben validar configuración, registros, copias de seguridad y controles del entorno antes de procesar información salarial. Consulta [`docs/HTTPS.md`](docs/HTTPS.md) para las pautas de despliegue.

---

# Security

## Supported versions

The `main` branch is monitored. There is no separate support policy for released versions; keep local deployments updated and review repository advisories before deployment.

## Reporting a vulnerability

**Do not disclose vulnerability details in a public issue, discussion, or pull request.** Use [GitHub private security advisories](https://github.com/yhquintero/Capital_Humanos/security/advisories/new) to contact the maintainer confidentially. Include minimal reproduction steps, impact, and a possible mitigation; do not attach personal data, real templates, or confidential exports.

If private advisories are unavailable, open a public issue requesting a private contact channel without describing the technical issue.

There is no response-time or remediation guarantee. Avoid testing that accesses, changes, or exfiltrates third-party data. Do not test production services without explicit authorization.

## Data considerations

Do not expose the application publicly without appropriate access controls. Uploaded files are analyzed in memory according to the current implementation, but operators must validate configuration, logs, backups, and environmental controls before processing salary information. See [`docs/HTTPS.md`](docs/HTTPS.md) for deployment guidance.
