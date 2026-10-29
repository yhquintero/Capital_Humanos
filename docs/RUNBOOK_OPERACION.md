# Runbook de disponibilidad y recuperación

Este documento distingue dos señales independientes:

- **Salud del contenedor:** Docker Compose consulta `http://127.0.0.1:8501/_stcore/health` dentro del contenedor. Solo confirma que el proceso Streamlit responde desde el mismo host.
- **Disponibilidad externa:** `tools/monitorar_disponibilidad.py` solicita por HTTPS `https://DOMINIO/healthz` desde una máquina externa. Caddy reenvía esa ruta al endpoint de salud de Streamlit. La ruta pública solo devuelve el estado genérico `ok`; no requiere credenciales ni expone páginas o datos.

## Instalar una comprobación externa

Desde un host de monitorización con Python 3.10+ y acceso al dominio:

```bash
python3 tools/monitorar_disponibilidad.py --url https://TU_DOMINIO/healthz
```

Una respuesta saludable produce una línea JSON con estado `UP`, código HTTP y latencia, y termina con código 0. Un servicio caído, respuesta incorrecta, error HTTP o timeout produce `DOWN` y código 1. La herramienta no sigue redirecciones, no acepta credenciales ni parámetros en la URL y no lee ni escribe el cuerpo en logs.

Ejemplo de comprobación por minuto mediante cron; ajusta las rutas al despliegue:

```cron
* * * * * cd /opt/Capital_Humanos && /usr/bin/python3 tools/monitorar_disponibilidad.py --url https://TU_DOMINIO/healthz >> /var/log/capital-humano-health.jsonl 2>&1 || /usr/bin/logger -p daemon.alert -t capital-humano-health 'La sonda externa falló; revisar runbook'
```

El código de salida y el registro `logger` permiten conectar la sonda a la plataforma de alertas que ya use la institución. No se envían salarios, filas, contraseñas ni cuerpos HTTP a un proveedor. No añadas secretos ni datos de personal a los argumentos de cron o a este repositorio.

## Diagnóstico y recuperación

1. Desde el host de despliegue, ejecuta `docker compose --env-file deploy/.env -f deploy/compose.yaml ps` y comprueba la salud del servicio `app` y el estado de `caddy`.
2. Si `app` está unhealthy, revisa `docker compose --env-file deploy/.env -f deploy/compose.yaml logs --tail=100 app`. Comprueba espacio, memoria y que el proceso escucha en el puerto interno 8501.
3. Si el contenedor está saludable pero falla `/healthz`, revisa Caddy con `docker compose --env-file deploy/.env -f deploy/compose.yaml logs --tail=100 caddy`; comprueba DNS y conectividad TCP 443 desde el host externo.
4. Si hay error TLS, revisa la fecha del sistema y los logs de Caddy. No borres `caddy_data` ni claves para intentar recuperarlo.
5. Después de corregir la causa, reinicia solo el servicio afectado con `docker compose ... restart app` o `restart caddy`, y ejecuta de nuevo la sonda desde fuera del servidor.
6. Registra hora UTC, servicio, síntomas no sensibles, causa y acción. Nunca copies registros salariales, respuestas del informe ni credenciales al parte del incidente.

## Simulacro seguro

Las pruebas `tests/test_monitorar_disponibilidad.py` levantan un servidor local temporal que responde `200/ok`, `503` y redirección. Verifican estados/alertas por código de salida y que la respuesta HTTP no aparezca en logs. No contactan producción ni usan datos personales.

La integración con un destino de paging concreto, horarios, responsables y retención de alertas debe configurarse en la infraestructura institucional; este repositorio no elige proveedor ni guarda credenciales.
