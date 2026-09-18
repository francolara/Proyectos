# Despliegue de producción restringido

Estos archivos preparan un despliegue manual desde GitHub Actions hacia Contabo. No despliegan ni modifican el VPS por sí solos.

## Archivos

- `ghcr.override.yml`: reemplaza exclusivamente las imágenes de `fralsecont`, `fralsetech` y `zonadeportiva` por las publicadas en GHCR.
- `fralse-deploy`: realiza el despliegue restringido en el VPS, incluidos bloqueo, validación, health checks y rollback.
- `fralse-deploy-ssh`: comando forzado de SSH; solo permite `deploy <tag> <target>`.
- `install-deploy.sh`: instala los tres archivos en el VPS, la regla sudoers mínima y la entrada SSH restringida.
- `../.github/workflows/deploy-production.yml`: workflow manual de GitHub Actions.

## Instalación inicial en Contabo

Copie la carpeta `deploy/` al VPS por un canal administrativo ya autorizado con el usuario personal `franco`. **No copie ni revele la clave privada** `contabo_github_actions`. Solo transfiera su archivo público, por ejemplo:

```bash
sudo bash ./install-deploy.sh --public-key-file /ruta/contabo_github_actions.pub
```

El instalador exige root, crea —si todavía no existe— la cuenta exclusiva y sin contraseña utilizable `fralse-deploy`, y puede ejecutarse de nuevo. No permite elegir otro usuario ni utilizar `root`; tampoco agrega la cuenta al grupo `docker` ni a grupos administrativos. Antes de modificar el VPS, prepara y valida los archivos; después instala cada destino de forma atómica, respalda los existentes y restaura el estado previo si se produce un error.

Instala el override en `/srv/fralse/compose/ghcr.override.yml`, deja los ejecutables bajo propiedad de root y agrega una sola entrada a `~fralse-deploy/.ssh/authorized_keys`. Esa entrada bloquea terminales, PTY, reenvíos de puertos, agente y X11; no da acceso a comandos arbitrarios.

La única regla sudoers creada permite a `fralse-deploy` ejecutar `/usr/local/sbin/fralse-deploy` como root sin contraseña. Se valida primero en un archivo temporal mediante `visudo -cf`, se instala atómicamente y se valida de nuevo. No concede privilegios generales y no modifica SQL Server, Nginx, Cloudflare, volúmenes ni archivos `.env`.

## Secrets de GitHub

En **Settings → Secrets and variables → Actions**, cree estos cuatro secretos:

- `CONTABO_HOST`: nombre DNS o IP del VPS.
- `CONTABO_USER`: debe tener exactamente el valor `fralse-deploy`.
- `CONTABO_SSH_PRIVATE_KEY`: contenido completo de la clave privada dedicada. Nunca lo registre en el repositorio ni lo imprima.
- `CONTABO_KNOWN_HOSTS`: salida de confianza ya verificada para el host, por ejemplo obtenida por un administrador con `ssh-keyscan` y comprobada contra la huella del VPS.

## Ejecución

Abra **Actions → Deploy production → Run workflow**. Indique preferentemente un SHA completo de 40 caracteres hexadecimales en minúsculas: es trazable e inmutable. `main` continúa permitido para una ejecución manual rápida. Seleccione el destino:

- Una aplicación: `fralsecont`, `fralsetech` o `zonadeportiva`.
- Las tres: `all`.

El workflow solo se inicia manualmente. Después de validar sus parámetros, abre SSH con verificación estricta de host y ejecuta exclusivamente `deploy <tag> <target>`.

## Validación y rollback

El servidor bloquea despliegues simultáneos con `flock`, valida la configuración de Compose y solo descarga/recrea los servicios elegidos con `--no-deps`. Nunca usa `docker compose down`.

Antes de actualizar, obtiene del contenedor en ejecución el ID inmutable `sha256:...` de cada servicio elegido, incluso si su origen previo era `fralse/...`. Tras recrear, comprueba `/healthz/live` y `/healthz/ready` en los puertos locales 5103, 5102 y 5101, con 30 intentos separados por 2 segundos. Si falla la recreación o un health check, restaura con esos IDs y levanta únicamente los servicios seleccionados.

Si la descarga falla, no se recrea ningún contenedor. El script verifica que los contenedores conservan sus IDs previos; solo si detecta una modificación inesperada ejecuta un rollback preventivo. En ambos casos el job termina con error y producción queda intacta o restaurada.

## Retirar el acceso

Como administrador del VPS, retire únicamente la línea cuyo `command="/usr/local/sbin/fralse-deploy-ssh"` aparece en `~fralse-deploy/.ssh/authorized_keys` y conserve un respaldo de ese archivo. Después valide y retire `/etc/sudoers.d/fralse-deploy`, y elimine `/usr/local/sbin/fralse-deploy`, `/usr/local/sbin/fralse-deploy-ssh` y, si ya no se necesita, `/srv/fralse/compose/ghcr.override.yml`.

Opcionalmente, una vez retirados la clave y sudoers, elimine la cuenta exclusiva `fralse-deploy`. Esta retirada no debe tocar el usuario `franco`, Docker, contenedores, imágenes, SQL Server, volúmenes ni archivos `.env`. Finalmente elimine o rote los cuatro secrets de GitHub.
