# HTTPS, certificados y acceso seguro

La aplicación se puede ejecutar de dos maneras. **Para que el navegador valide un certificado como perteneciente a un sitio público, hace falta un dominio propio** que apunte al servidor. No se puede emitir un certificado público confiable para `localhost` o para una IP privada.

## Opción recomendada para un servidor público: Caddy + Let's Encrypt

El despliegue de `deploy/compose.yaml` pone Caddy delante de Streamlit. Caddy solicita y renueva automáticamente un certificado TLS válido de Let's Encrypt/ZeroSSL, mantiene la clave privada en volúmenes locales y añade autenticación por contraseña delante de la aplicación. El contenedor Streamlit no publica su puerto al exterior; solo Caddy expone 80/443.

### Requisitos previos

- Un dominio bajo tu control; crea un registro DNS `A` (y `AAAA` solo si el servidor tiene IPv6 pública funcional) apuntando a la IP pública.
- Enrutamiento/NAT correcto y los puertos TCP **80 y 443** accesibles desde Internet.
- Docker Engine con Docker Compose.

### Puesta en marcha

Desde la raíz del repositorio:

```bash
cp deploy/.env.example deploy/.env
```

Edita `deploy/.env`: establece `DOMINIO` y un usuario. Genera un hash de contraseña con Caddy:

```bash
docker run --rm caddy:2-alpine caddy hash-password --plaintext 'una-frase-secreta-larga-y-unica'
```

Pega el resultado en `APP_PASSWORD_HASH` de `deploy/.env`, conservando comillas simples porque el hash contiene `$`. No pongas la contraseña sin hash, no compartas el archivo `.env` y no lo añadas a Git.

```bash
docker compose --env-file deploy/.env -f deploy/compose.yaml up --build -d
```

Abre `https://TU_DOMINIO`. El navegador debe mostrar un certificado válido cuyo emisor sea una autoridad pública confiable. Caddy renueva el certificado mientras conserve el volumen `caddy_data`; **no borres ese volumen** durante la operación habitual.

Administración:

```bash
docker compose --env-file deploy/.env -f deploy/compose.yaml logs -f caddy
# Detener sin borrar los datos/certificados de Caddy:
docker compose --env-file deploy/.env -f deploy/compose.yaml down
```

No abras directamente el puerto 8501 en el router/firewall. Solo se deben exponer 80/443 del proxy Caddy.

## Opción para uso local: certificado confiable con mkcert

`mkcert` crea una autoridad certificadora local y agrega su raíz de confianza **a este equipo**. Es apropiado para desarrollo y una red privada controlada, no para publicar el sitio en Internet. La aplicación arranca enlazada a `127.0.0.1` por defecto.

### Windows

1. Instala mkcert: `winget install FiloSottile.mkcert`.
2. Desde PowerShell, en la raíz del repositorio:

   ```powershell
   .\TES\Salariales\Nueva Escala 3210\run_https.ps1
   ```

3. Abre `https://localhost:8501`.

Para probar en una LAN privada, opcionalmente ejecuta el script con `-EnRedLocal`. Esto enlaza el servicio a las interfaces de red y muestra las direcciones; configura el Firewall para permitir el puerto solo en perfil de red **Privado**.

### Linux / macOS

Instala mkcert desde [su página oficial](https://github.com/FiloSottile/mkcert), y luego:

```bash
chmod +x "TES/Salariales/Nueva Escala 3210/run_https.sh"
"TES/Salariales/Nueva Escala 3210/run_https.sh"
```

Para uso explícito en la LAN privada:

```bash
"TES/Salariales/Nueva Escala 3210/run_https.sh" --lan
```

Para que otro dispositivo valide ese certificado local debe instalar la raíz pública de mkcert como autoridad de confianza. `mkcert -CAROOT` indica su ubicación: comparte únicamente `rootCA.pem`, nunca `rootCA-key.pem`. La llave de la CA y la llave TLS local permanecen fuera del repositorio, en el equipo que generó el certificado.

## Qué protege y qué no

- TLS cifra el tráfico y permite al navegador comprobar que el certificado corresponde al dominio indicado.
- La autenticación de Caddy solicita usuario y contraseña antes de abrir la aplicación pública. Usa una contraseña larga, exclusiva y no reutilizada.
- El certificado no cifra los archivos fuente ni concede control de acceso a los documentos almacenados en GitHub. Revisa quién tiene permiso al repositorio y al servidor antes de incluir información personal o confidencial.
- La autenticación Basic configurada para el despliegue es compartida por la instalación; no proporciona usuarios individuales ni auditoría por persona.
- Los certificados locales de mkcert son confiables solo en los dispositivos en los que se instale explícitamente su raíz.
- No subas certificados, claves privadas, contraseñas ni `deploy/.env`. Las rutas locales están excluidas mediante `.gitignore`.
