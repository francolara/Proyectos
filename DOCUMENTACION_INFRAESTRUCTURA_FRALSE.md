# Infraestructura, despliegue y operación de FRALSE en Contabo

> Documento de contexto para administradores y asistentes de desarrollo (incluido Codex).
>
> **Última consolidación:** 2026-10-06  
> **Zona horaria operativa:** `America/Lima` (UTC-5)  
> **Entorno:** Producción  
> **Regla:** este documento no debe contener contraseñas, claves privadas, tokens ni valores de archivos `.env`.
> **Alcance:** esta documentación técnica se aplica únicamente a los proyectos ubicados dentro de la carpeta raíz `.NET2026`. No debe utilizarse como configuración de infraestructura para proyectos alojados fuera de esa carpeta.

---

## 1. Resumen ejecutivo

FRALSE opera tres aplicaciones web ASP.NET Core y un SQL Server en un VPS Ubuntu de Contabo. Las aplicaciones se ejecutan como contenedores Docker coordinados por Docker Compose. Nginx recibe HTTPS y reenvía cada dominio a un puerto publicado únicamente en el loopback del VPS. El tráfico web público debe atravesar Cloudflare. SQL Server no se publica directamente en Internet; las aplicaciones acceden por la red privada de Docker y la administración remota se realiza mediante un túnel SSH.

Los respaldos de SQL Server se ejecutan diariamente a las **03:00 hora de Perú**, se verifican, comprimen y cifran con `age` antes de enviarse a un bucket privado de Cloudflare R2.

### Inventario conocido

| Componente | Nombre o ubicación | Función |
|---|---|---|
| VPS | Contabo, Ubuntu, IP `13.140.42.106` | Host de producción |
| Usuario administrativo | `franco` | Administración mediante SSH con clave |
| Aplicación Zona Deportiva | Contenedor `fralse-zonadeportiva` | Sitio `lazonadeportiva.com` |
| Aplicación Fralse Tech | Contenedor `fralse-fralsetech` | Sitio `fralsetech.com` |
| Aplicación FralseCont | Contenedor `fralse-fralsecont` | Sitio `conta.fralsetech.com` |
| Base de datos | Contenedor `fralse-sqlserver` | SQL Server, puerto interno `1433` |
| Orquestación | `/srv/fralse/compose/apps.compose.yml` | Configuración Docker Compose |
| Proxy inverso | Nginx en el host | TLS de origen y enrutamiento |
| Backup SQL | `/usr/local/sbin/fralse-sql-backup` | Backup, verificación, cifrado y carga |
| Timer del backup | `fralse-sql-backup.timer` | Ejecución diaria a las 03:00 Perú |
| Destino del backup | R2 `fralse-sql-backups` | Almacenamiento privado cifrado |

> Los nombres exactos de los **servicios Compose**, repositorios, contextos de compilación y archivos `.csproj` deben verificarse en el repositorio antes de modificar un workflow o ejecutar un despliegue.

---

## 2. Arquitectura de conexiones

### 2.1 Tráfico web

```text
Navegador
  -> Cloudflare (DNS proxied, TLS y protección)
  -> Firewall de Contabo
  -> UFW del VPS (443 solo desde rangos de Cloudflare)
  -> Nginx (certificado Cloudflare Origin CA)
  -> 127.0.0.1:5101, 5102 o 5103
  -> contenedor ASP.NET Core:8080
```

| Dominio | Puerto del host | Puerto del contenedor | Contenedor |
|---|---:|---:|---|
| `lazonadeportiva.com` | `127.0.0.1:5101` | `8080` | `fralse-zonadeportiva` |
| `fralsetech.com` | `127.0.0.1:5102` | `8080` | `fralse-fralsetech` |
| `conta.fralsetech.com` | `127.0.0.1:5103` | `8080` | `fralse-fralsecont` |

Reglas esenciales:

- Los puertos `5101`, `5102` y `5103` deben enlazarse a `127.0.0.1`, nunca a `0.0.0.0`.
- Los registros DNS web deben conservar el proxy naranja de Cloudflare.
- Cloudflare debe usar SSL/TLS **Full (strict)**.
- Nginx utiliza certificados Cloudflare Origin CA.
- El puerto 80 del origen no es necesario si la redirección HTTP a HTTPS se hace en Cloudflare.

### 2.2 Aplicaciones hacia SQL Server

```text
Contenedor ASP.NET Core
  -> red privada Docker
  -> fralse-sqlserver:1433
  -> base de datos asignada
```

- Las aplicaciones deben usar el nombre DNS del servicio o contenedor en Docker, no la IP pública del VPS.
- La cadena de conexión se inyecta como variable de entorno o secreto; no se guarda en Git.
- Se recomienda un usuario SQL de permisos limitados por aplicación. No usar `sa` desde las aplicaciones.
- Las bases conocidas son `FralseCont` y `ZonaDeportiva`.

### 2.3 Administración de SQL desde Windows

```text
SSMS / aplicación local
  -> 127.0.0.1:11433
  -> túnel SSH por VPS:22
  -> endpoint SQL privado:1433
  -> fralse-sqlserver
```

En SSMS:

- Servidor: `127.0.0.1,11433`
- Encrypt: `Mandatory`
- Trust server certificate: activado, mientras SQL utilice su certificado interno actual

#### Configuración recomendada y estable

Publicar SQL únicamente en loopback del VPS desde Compose:

```yaml
services:
  sqlserver:
    # El nombre real del servicio debe confirmarse.
    ports:
      - "127.0.0.1:1433:1433"
```

Después el túnel no depende de una IP interna cambiante:

```bat
ssh.exe -i "%USERPROFILE%\.ssh\contabo_vps_ed25519" ^
  -N -T ^
  -o ExitOnForwardFailure=yes ^
  -o ServerAliveInterval=30 ^
  -o ServerAliveCountMax=3 ^
  -L 127.0.0.1:11433:127.0.0.1:1433 ^
  franco@13.140.42.106
```

#### Configuración histórica

El BAT existente apunta a `172.18.0.2:1433`. Esa IP de Docker puede cambiar al recrear el contenedor o la red. Si todavía no se ha aplicado el mapeo estable, consultar la IP actual:

```bash
sudo docker inspect fralse-sqlserver \
  --format '{{range .NetworkSettings.Networks}}{{.IPAddress}}{{println}}{{end}}'
```

Si el túnel informa `Connection refused`, comprobar primero:

```bash
sudo docker ps --format 'table {{.Names}}\t{{.Status}}\t{{.Ports}}'
sudo docker logs --tail 100 fralse-sqlserver
sudo docker inspect fralse-sqlserver --format '{{json .NetworkSettings.Networks}}'
```

---

## 3. Capas de seguridad

| Capa | Responsabilidad | Estado esperado |
|---|---|---|
| Cloudflare | DNS proxied, TLS público y filtrado web | Proxy naranja; Full (strict) |
| Firewall Contabo | Filtrado antes de llegar al VPS | Permitir 22 y 443; bloquear el resto |
| UFW | Segundo filtro en Ubuntu | 443 solo desde Cloudflare; 22 para administración |
| SSH | Administración cifrada | Solo clave pública, root y contraseña deshabilitados |
| Fail2ban | Bloquear intentos abusivos de SSH | Servicio y cárcel `sshd` activos |
| Nginx | TLS de origen y reverse proxy | Configuración válida; apps solo en loopback |
| Docker | Aislamiento de servicios | SQL privado; redes separadas cuando corresponda |
| Backups | Recuperación ante pérdida o corrupción | R2 privado, cifrado `age`, SHA-256 y prueba de restauración |

### 3.1 Puertos esperados

| Puerto | Exposición | Uso |
|---:|---|---|
| `22/tcp` | Público y filtrado | SSH con claves + Fail2ban |
| `443/tcp` | Público, solo Cloudflare en UFW | HTTPS hacia Nginx |
| `80/tcp` | Cerrado en origen | Redirección realizada por Cloudflare |
| `1433/tcp` | Privado o solo `127.0.0.1` | SQL Server y túnel SSH |
| `5101–5103/tcp` | Solo `127.0.0.1` | Nginx hacia aplicaciones |

### 3.2 SSH endurecido

Estado validado:

```text
PermitRootLogin no
PubkeyAuthentication yes
PasswordAuthentication no
```

Comprobación:

```bash
sudo sshd -T | grep -E 'permitrootlogin|passwordauthentication|pubkeyauthentication|allowusers'
systemctl is-enabled ssh
systemctl is-active ssh
sudo fail2ban-client status sshd
```

Acceso desde PowerShell:

```powershell
ssh -i "$env:USERPROFILE\.ssh\contabo_vps_ed25519" franco@13.140.42.106
```

Buenas prácticas:

- Una pareja de claves por computadora.
- La clave privada nunca sale de la computadora que la creó.
- Solo la clave pública se agrega a `/home/franco/.ssh/authorized_keys`.
- Antes de retirar una clave antigua, probar la nueva en una segunda sesión abierta.
- Si una PC se pierde, eliminar únicamente su línea pública de `authorized_keys`.

### 3.3 Revisión rápida de seguridad

```bash
sudo ufw status verbose
sudo ss -tulpn
sudo fail2ban-client status sshd
sudo nginx -t
sudo systemctl --failed --no-pager
sudo docker ps --format 'table {{.Names}}\t{{.Image}}\t{{.Status}}\t{{.Ports}}'
```

---

## 4. Estructura recomendada del repositorio

La estructura real puede variar. Esta es la organización objetivo para que código, infraestructura y documentación sean fáciles de mantener:

```text
repositorio/
├── .github/
│   └── workflows/
│       ├── build-images.yml
│       └── deploy-production.yml
├── src/
│   ├── FralseTech/
│   │   ├── FralseTech.csproj
│   │   └── Dockerfile
│   ├── FralseCont/
│   │   ├── FralseCont.csproj
│   │   └── Dockerfile
│   └── ZonaDeportiva/
│       ├── ZonaDeportiva.csproj
│       └── Dockerfile
├── deploy/
│   ├── apps.compose.yml.example
│   └── env.example
├── scripts/
│   ├── deploy.sh
│   └── health-check.sh
├── .dockerignore
├── .gitignore
└── DOCUMENTACION_INFRAESTRUCTURA_FRALSE.md
```

Nunca subir:

```gitignore
.env
.env.*
!.env.example
*.key
*.pem
*.pfx
*.p12
secrets/
backup/
*.bak
*.age
rclone.conf
```

---

## 5. Dockerfile por proyecto ASP.NET Core

Cada aplicación debe tener su propio Dockerfile. La receta habitual usa dos etapas: una SDK para compilar y una runtime para ejecutar.

### 5.1 Plantilla base

```dockerfile
# syntax=docker/dockerfile:1

FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS runtime
WORKDIR /app
EXPOSE 8080
ENV ASPNETCORE_URLS=http://+:8080

FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
ARG BUILD_CONFIGURATION=Release
WORKDIR /src

# Ajustar las rutas al proyecto real.
COPY ["src/PROYECTO/PROYECTO.csproj", "src/PROYECTO/"]
RUN dotnet restore "src/PROYECTO/PROYECTO.csproj"

COPY . .
RUN dotnet publish "src/PROYECTO/PROYECTO.csproj" \
    -c "$BUILD_CONFIGURATION" \
    -o /app/publish \
    --no-restore \
    /p:UseAppHost=false

FROM runtime AS final
WORKDIR /app
COPY --from=build /app/publish .
ENTRYPOINT ["dotnet", "PROYECTO.dll"]
```

Crear tres copias ajustando `PROYECTO`:

| Aplicación | Proyecto/DLL esperado | Imagen GHCR sugerida |
|---|---|---|
| Fralse Tech | Verificar `.csproj` y DLL reales | `ghcr.io/<owner>/fralsetech` |
| FralseCont | Verificar `.csproj` y DLL reales | `ghcr.io/<owner>/fralsecont` |
| Zona Deportiva | Verificar `.csproj` y DLL reales | `ghcr.io/<owner>/zonadeportiva` |

Si los proyectos comparten librerías, copiar primero todos los `.csproj` necesarios antes de `dotnet restore` para aprovechar la caché de Docker.

### 5.2 `.dockerignore`

```dockerignore
**/bin/
**/obj/
.git/
.github/
.vs/
.vscode/
.env
.env.*
secrets/
**/*.user
**/*.suo
**/*.pfx
**/*.key
**/*.pem
```

### 5.3 Qué es cada elemento

| Elemento | Qué hace |
|---|---|
| Dockerfile | Receta que compila una aplicación y construye una imagen |
| Imagen | Paquete inmutable con aplicación y dependencias |
| Contenedor | Instancia en ejecución de una imagen |
| Compose | Declara cómo se ejecutan y conectan todos los contenedores |
| Volumen | Datos persistentes independientes del contenedor |

---

## 6. Docker Compose de producción

El archivo operativo está en:

```text
/srv/fralse/compose/apps.compose.yml
```

Modelo de referencia, no sustituto automático del archivo real:

```yaml
name: fralse

services:
  zonadeportiva:
    image: ghcr.io/<owner>/zonadeportiva:${IMAGE_TAG:-main}
    container_name: fralse-zonadeportiva
    restart: unless-stopped
    env_file:
      - /srv/fralse/secrets/zonadeportiva.env
    ports:
      - "127.0.0.1:5101:8080"
    networks:
      - backend
      - egress
    depends_on:
      sqlserver:
        condition: service_healthy

  fralsetech:
    image: ghcr.io/<owner>/fralsetech:${IMAGE_TAG:-main}
    container_name: fralse-fralsetech
    restart: unless-stopped
    env_file:
      - /srv/fralse/secrets/fralsetech.env
    ports:
      - "127.0.0.1:5102:8080"
    networks:
      - egress

  fralsecont:
    image: ghcr.io/<owner>/fralsecont:${IMAGE_TAG:-main}
    container_name: fralse-fralsecont
    restart: unless-stopped
    env_file:
      - /srv/fralse/secrets/fralsecont.env
    ports:
      - "127.0.0.1:5103:8080"
    networks:
      - backend
      - egress
    depends_on:
      sqlserver:
        condition: service_healthy

  sqlserver:
    image: mcr.microsoft.com/mssql/server:2022-latest
    container_name: fralse-sqlserver
    restart: unless-stopped
    env_file:
      - /srv/fralse/secrets/sqlserver.env
    ports:
      - "127.0.0.1:1433:1433"
    volumes:
      - sqlserver-data:/var/opt/mssql
      - /srv/backups/sqlserver:/var/opt/mssql/backups
    networks:
      - backend
      - egress
    healthcheck:
      test: ["CMD-SHELL", "/opt/mssql-tools18/bin/sqlcmd -S localhost -U sa -P \"$$MSSQL_SA_PASSWORD\" -C -Q 'SELECT 1' || exit 1"]
      interval: 30s
      timeout: 10s
      retries: 10
      start_period: 60s

volumes:
  sqlserver-data:

networks:
  backend:
    internal: true
  egress:
```

### Precauciones

- Confirmar el nombre real del volumen SQL con `docker inspect` antes de editar.
- No ejecutar `docker compose down -v`; `-v` puede eliminar volúmenes.
- No publicar SQL como `0.0.0.0:1433:1433`.
- No copiar secretos dentro de imágenes.
- Validar siempre antes de aplicar:

```bash
sudo docker compose -f /srv/fralse/compose/apps.compose.yml config -q
sudo docker compose -f /srv/fralse/compose/apps.compose.yml config --services
```

### Descubrir la configuración real

```bash
sudo docker inspect fralse-sqlserver \
  --format 'Directorio={{index .Config.Labels "com.docker.compose.project.working_dir"}} Archivo={{index .Config.Labels "com.docker.compose.project.config_files"}}'

sudo docker inspect fralse-sqlserver \
  --format '{{range .Mounts}}Tipo={{.Type}} Nombre={{.Name}} Origen={{.Source}} Destino={{.Destination}}{{println}}{{end}}'

sudo docker network ls
sudo docker volume ls
```

---

## 7. Variables de entorno y secretos

### 7.1 Ubicación

| Lugar | Contenido | ¿Git? |
|---|---|---|
| `apps.compose.yml` | Estructura, nombres de variables, puertos y redes | Sí, si no contiene valores sensibles |
| `/srv/fralse/secrets/*.env` | Valores reales de producción | No |
| GitHub Actions Secrets | Claves para CI/CD | No; GitHub las almacena cifradas |
| Variables del contenedor | Copia cargada al crear el contenedor | No es fuente editable |

Ejemplo sin valores reales:

```dotenv
ASPNETCORE_ENVIRONMENT=Production
ConnectionStrings__DefaultConnection=Server=fralse-sqlserver,1433;Database=<DB>;User Id=<USUARIO>;Password=<SECRETO>;Encrypt=True;TrustServerCertificate=True
Mail__Host=<HOST>
Mail__User=<USUARIO>
Mail__Password=<SECRETO>
```

### 7.2 Ver nombres sin revelar valores

```bash
sudo docker inspect -f '{{range .Config.Env}}{{println .}}{{end}}' NOMBRE_CONTENEDOR \
  | sed 's/=.*$/=<oculto>/'
```

### 7.3 Modificar una variable

```bash
cd /srv/fralse/compose
sudo cp -a /srv/fralse/secrets/APP.env "/srv/fralse/secrets/APP.env.bak.$(date +%Y%m%d-%H%M%S)"
sudo nano /srv/fralse/secrets/APP.env
sudo docker compose -f apps.compose.yml config -q
sudo docker compose -f apps.compose.yml up -d --no-deps --force-recreate SERVICIO
sudo docker compose -f apps.compose.yml ps
sudo docker compose -f apps.compose.yml logs --tail 100 SERVICIO
```

No pegar en chats ni capturas la salida de `docker compose config` porque puede mostrar secretos resueltos.

---

## 8. GitHub Actions y GHCR

### 8.1 Flujo normal

```text
Cambios locales
  -> rama y Pull Request
  -> merge a main
  -> workflow Build and publish container images
  -> imágenes versionadas en GHCR
  -> workflow manual Deploy production
  -> SSH al VPS
  -> docker compose pull
  -> recrear solo servicios seleccionados
  -> health checks y validación pública
```

El build automático y el despliegue manual son dos etapas separadas. No desplegar hasta que el build del mismo commit esté en verde.

### 8.2 Secretos de GitHub esperados

Los nombres reales deben verificarse en `Settings -> Secrets and variables -> Actions`:

| Secreto | Uso |
|---|---|
| `VPS_HOST` | IP o host del VPS |
| `VPS_USER` | Usuario de despliegue, idealmente `fralse-deploy` |
| `VPS_SSH_KEY` | Clave privada exclusiva del robot de despliegue |
| `VPS_KNOWN_HOSTS` | Huella SSH verificada del VPS |
| `GHCR_USER` | Usuario técnico para descargar imágenes, si es necesario |
| `GHCR_TOKEN` | Token mínimo `read:packages`, si las imágenes son privadas |

`GITHUB_TOKEN` se genera automáticamente para el workflow de construcción. No reutilizar la clave privada personal de `franco` como clave del robot.

### 8.3 Workflow de construcción sugerido

Archivo: `.github/workflows/build-images.yml`

```yaml
name: Build and publish container images

on:
  push:
    branches: [main]
  workflow_dispatch:

permissions:
  contents: read
  packages: write

jobs:
  build:
    runs-on: ubuntu-latest
    strategy:
      fail-fast: false
      matrix:
        include:
          - image: fralsetech
            context: .
            dockerfile: src/FralseTech/Dockerfile
          - image: fralsecont
            context: .
            dockerfile: src/FralseCont/Dockerfile
          - image: zonadeportiva
            context: .
            dockerfile: src/ZonaDeportiva/Dockerfile

    steps:
      - uses: actions/checkout@v4

      - uses: docker/setup-buildx-action@v3

      - uses: docker/login-action@v3
        with:
          registry: ghcr.io
          username: ${{ github.actor }}
          password: ${{ secrets.GITHUB_TOKEN }}

      - id: meta
        uses: docker/metadata-action@v5
        with:
          images: ghcr.io/${{ github.repository_owner }}/${{ matrix.image }}
          tags: |
            type=raw,value=main
            type=sha,format=long

      - uses: docker/build-push-action@v6
        with:
          context: ${{ matrix.context }}
          file: ${{ matrix.dockerfile }}
          push: true
          platforms: linux/amd64
          tags: ${{ steps.meta.outputs.tags }}
          labels: ${{ steps.meta.outputs.labels }}
          cache-from: type=gha
          cache-to: type=gha,mode=max
```

### 8.4 Workflow de despliegue sugerido

Archivo: `.github/workflows/deploy-production.yml`

```yaml
name: Deploy production

on:
  workflow_dispatch:
    inputs:
      target:
        description: Servicio a desplegar
        required: true
        type: choice
        options:
          - all-apps
          - fralsetech
          - fralsecont
          - zonadeportiva
      image_tag:
        description: Etiqueta GHCR, normalmente main o sha-<SHA_COMPLETO>
        required: true
        default: main

concurrency:
  group: production-deploy
  cancel-in-progress: false

jobs:
  deploy:
    runs-on: ubuntu-latest
    environment: production
    steps:
      - name: Preparar SSH
        shell: bash
        run: |
          install -d -m 700 ~/.ssh
          printf '%s\n' "${{ secrets.VPS_SSH_KEY }}" > ~/.ssh/deploy_key
          chmod 600 ~/.ssh/deploy_key
          printf '%s\n' "${{ secrets.VPS_KNOWN_HOSTS }}" > ~/.ssh/known_hosts

      - name: Desplegar servicio seleccionado
        env:
          TARGET: ${{ inputs.target }}
          IMAGE_TAG: ${{ inputs.image_tag }}
        run: |
          case "$TARGET" in
            all-apps) SERVICES="fralsetech fralsecont zonadeportiva" ;;
            fralsetech|fralsecont|zonadeportiva) SERVICES="$TARGET" ;;
            *) echo "Destino inválido" >&2; exit 2 ;;
          esac

          ssh -i ~/.ssh/deploy_key \
            "${{ secrets.VPS_USER }}@${{ secrets.VPS_HOST }}" \
            "IMAGE_TAG='$IMAGE_TAG' /usr/local/sbin/fralse-deploy $SERVICES"
```

El script `/usr/local/sbin/fralse-deploy` debe aplicar una lista cerrada de servicios, validar Compose, descargar imágenes, recrear solo las aplicaciones solicitadas y comprobar salud. No debe aceptar comandos arbitrarios ni recrear SQL Server durante un despliegue normal.

### 8.5 Proceso operativo de publicación

1. Probar localmente y revisar que no haya secretos.
2. Crear rama, commit y Pull Request.
3. Integrar a `main` después de revisión.
4. Esperar `Build and publish container images` en verde.
5. Abrir `Deploy production -> Run workflow`.
6. Elegir solo la aplicación modificada, salvo despliegue conjunto justificado.
7. Confirmar el commit o etiqueta que se va a publicar.
8. Esperar resultado verde.
9. Validar página, función modificada, health checks y logs.
10. Conservar la versión anterior hasta cerrar la validación.

Para Zona Deportiva, `/healthz/live` confirma que el proceso responde y `/healthz/ready` confirma además la conexión SQL y el contrato canónico `CANONICO_MAESTROS_V2` (moneda, comprobantes, deporte y suelo). Antes de publicar una imagen que incluya esta validación se debe desplegar `dbo.Sp_Sistema_ValidarContratoCanonico`; de lo contrario readiness permanecerá no saludable. Las conciliaciones históricas completas se ejecutan mediante los scripts de Fase 5 y Fase 6, no dentro de cada sonda HTTP. La migración de espacios se despliega en dos pasos: Fase 7A aditiva antes de la aplicación y Fase 7B destructiva solo después de pruebas y con respaldo restaurable verificado.

---

## 9. Despliegue manual controlado

Usar cuando GitHub Actions no esté disponible o para recuperación.

```bash
COMPOSE=/srv/fralse/compose/apps.compose.yml

sudo docker compose -f "$COMPOSE" ps
sudo docker compose -f "$COMPOSE" config -q
sudo docker compose -f "$COMPOSE" pull SERVICIO
sudo docker compose -f "$COMPOSE" up -d --no-deps SERVICIO
sudo docker compose -f "$COMPOSE" ps SERVICIO
sudo docker compose -f "$COMPOSE" logs --tail 100 SERVICIO
```

Validar públicamente:

```bash
curl -fsS -o /dev/null -w '%{http_code}\n' https://fralsetech.com/
curl -fsS -o /dev/null -w '%{http_code}\n' https://conta.fralsetech.com/
curl -fsS -o /dev/null -w '%{http_code}\n' https://lazonadeportiva.com/
```

Un `301` puede ser normal si el dominio redirige a su URL canónica. Seguir la redirección con `curl -IL` para verificar el destino.

No usar durante un despliegue normal:

```text
docker compose down
docker compose down -v
docker system prune -a
docker volume prune
```

---

## 10. SQL Server: persistencia y mantenimiento

### 10.1 Modelo

- Motor: contenedor `fralse-sqlserver`.
- Puerto: `1433` interno; no público.
- Datos: volumen Docker montado en `/var/opt/mssql`.
- Backups temporales: `/srv/backups/sqlserver` en el host, montado en `/var/opt/mssql/backups` dentro del contenedor.
- Reinicio: política `unless-stopped`.

Inventario:

```bash
sudo docker inspect fralse-sqlserver \
  --format '{{range .Mounts}}Tipo={{.Type}} Nombre={{.Name}} Origen={{.Source}} Destino={{.Destination}}{{println}}{{end}}'

sudo docker inspect fralse-sqlserver \
  --format 'Reinicio={{.HostConfig.RestartPolicy.Name}} Estado={{.State.Status}} Salud={{if .State.Health}}{{.State.Health.Status}}{{else}}sin-healthcheck{{end}}'
```

### 10.2 Restauración y cambios

- Antes de una migración de esquema, generar un backup adicional.
- Probar restauraciones fuera de producción.
- No reemplazar el volumen para actualizar la imagen de SQL Server.
- Confirmar compatibilidad y backup antes de cambiar la versión mayor de SQL Server.

---

## 11. Backups SQL cifrados en Cloudflare R2

### 11.1 Flujo

```text
systemd timer 03:00 America/Lima
  -> BACKUP DATABASE con COPY_ONLY y CHECKSUM
  -> RESTORE VERIFYONLY
  -> gzip
  -> cifrado age con clave pública
  -> SHA-256
  -> carga por rclone a R2
  -> comparación remota
  -> retención de 7 copias por base
```

Configuración conocida:

| Elemento | Valor |
|---|---|
| Script | `/usr/local/sbin/fralse-sql-backup` |
| Contenedor | `fralse-sqlserver` |
| Bases | `FralseCont`, `ZonaDeportiva` |
| Área temporal | `/srv/backups/sqlserver` |
| Destino R2 | `r2:fralse-sql-backups` |
| Recipient público age | `/etc/fralse-backup/age-recipient.txt` |
| Configuración rclone | `/root/.config/rclone/rclone.conf` |
| Retención | 7 copias por base |

La clave privada `age` debe permanecer fuera del VPS, con copia segura. En el VPS solo se necesita la clave pública para cifrar.

### 11.2 Revisión manual

```bash
sudo systemctl list-timers fralse-sql-backup.timer --no-pager
sudo systemctl status fralse-sql-backup.service --no-pager
sudo journalctl -u fralse-sql-backup.service -n 50 --no-pager
sudo /usr/local/sbin/fralse-sql-backup
```

### 11.3 Agregar una base al backup

1. Confirmar que la base existe y su nombre exacto.
2. Crear copia del script.
3. Editar únicamente el arreglo `DATABASES`.
4. Validar sintaxis.
5. Ejecutar prueba manual.
6. Confirmar `.age` y `.sha256` en la carpeta R2 de la nueva base.
7. Descargar, descifrar y restaurar una copia de prueba.

```bash
sudo cp -a /usr/local/sbin/fralse-sql-backup \
  "/usr/local/sbin/fralse-sql-backup.bak.$(date +%Y%m%d-%H%M%S)"
sudo nano /usr/local/sbin/fralse-sql-backup
sudo bash -n /usr/local/sbin/fralse-sql-backup && echo 'SINTAXIS CORRECTA'
sudo /usr/local/sbin/fralse-sql-backup
```

---

## 12. Mantenimiento del VPS

### Semanal

```bash
sudo systemctl --failed --no-pager
sudo docker ps --format 'table {{.Names}}\t{{.Image}}\t{{.Status}}\t{{.Ports}}'
df -h
free -h
sudo journalctl -p warning --since '7 days ago' --no-pager
sudo journalctl -u fralse-sql-backup.service --since '7 days ago' --no-pager
```

### Mensual

```bash
sudo apt update
apt list --upgradable
sudo ufw status verbose
sudo fail2ban-client status sshd
sudo nginx -t
sudo docker system df
sudo docker image ls
```

Aplicar actualizaciones:

```bash
sudo apt update
sudo apt upgrade
```

Después:

```bash
if [ -f /var/run/reboot-required ]; then
  cat /var/run/reboot-required
else
  echo 'NO SE REQUIERE REINICIO'
fi
```

### Reinicio controlado

Antes:

```bash
sudo docker ps --format 'table {{.Names}}\t{{.Status}}\t{{.Ports}}'
for s in ssh docker nginx fralse-sql-backup.timer; do
  printf '%-28s' "$s"
  systemctl is-active "$s"
done
sudo systemctl --failed --no-pager
```

Reiniciar:

```bash
sudo reboot
```

Después de reconectar:

```bash
uptime -p
for s in ssh docker nginx fralse-sql-backup.timer; do
  printf '%-28s' "$s"
  systemctl is-active "$s"
done
sudo docker ps --format 'table {{.Names}}\t{{.Status}}\t{{.Ports}}'
sudo systemctl --failed --no-pager
sudo systemctl list-timers fralse-sql-backup.timer --no-pager
```

Verificar también los tres dominios y el túnel SQL.

---

## 13. Diagnóstico rápido

### Un sitio no abre

```bash
sudo nginx -t
sudo systemctl status nginx --no-pager
sudo docker ps
sudo docker logs --tail 100 NOMBRE_CONTENEDOR
curl -I https://DOMINIO
```

Revisar, en orden: DNS/proxy Cloudflare, certificado, firewall, Nginx, puerto loopback, contenedor y logs.

### Un contenedor reinicia continuamente

```bash
sudo docker ps -a
sudo docker inspect NOMBRE_CONTENEDOR --format '{{json .State}}'
sudo docker logs --tail 200 NOMBRE_CONTENEDOR
```

No borrar ni recrear volúmenes para “probar”. Primero identificar la causa.

### El túnel SSH abre, pero SQL rechaza la conexión

```bash
sudo docker ps --filter name=fralse-sqlserver
sudo docker logs --tail 100 fralse-sqlserver
sudo docker inspect fralse-sqlserver \
  --format '{{range .NetworkSettings.Networks}}{{.IPAddress}}{{println}}{{end}}'
sudo ss -ltnp | grep 1433 || true
```

Si el BAT usa una IP `172.x.x.x`, comprobar si cambió después del reinicio. La corrección permanente es el binding `127.0.0.1:1433:1433` más el túnel hacia `127.0.0.1:1433`.

### El backup no corrió

```bash
sudo systemctl status fralse-sql-backup.timer --no-pager
sudo systemctl status fralse-sql-backup.service --no-pager
sudo journalctl -u fralse-sql-backup.service --since '24 hours ago' --no-pager
sudo systemctl list-timers fralse-sql-backup.timer --no-pager
```

No dar por válido un backup solo porque existe en R2: debe tener su `.sha256`, haber sido verificado y, periódicamente, restaurado en un entorno seguro.

---

## 14. Alta de una nueva aplicación

1. Crear el proyecto y sus pruebas.
2. Agregar Dockerfile multietapa y `.dockerignore`.
3. Definir health checks `/healthz/live` y `/healthz/ready` si la aplicación los soporta.
4. Agregar la aplicación a la matriz de `build-images.yml`.
5. Publicar imagen en GHCR con etiqueta de commit.
6. Crear archivo env de producción fuera de Git.
7. Agregar servicio a Compose con puerto solo loopback.
8. Conectar solo las redes necesarias.
9. Agregar Nginx y certificado Origin CA.
10. Crear DNS proxied en Cloudflare.
11. Validar Compose y Nginx.
12. Desplegar, revisar logs y probar por Cloudflare.
13. Añadir el servicio a `deploy-production.yml` y al script de lista cerrada.
14. Documentar dominio, puerto, base, variables, propietario y rollback.

Si necesita SQL, crear una credencial limitada y agregar la base al proceso de backup.

---

## 15. Rollback

La unidad de rollback de una aplicación es la imagen Docker correspondiente a un commit anterior con build exitoso.

Proceso:

1. Identificar el SHA correcto.
2. Confirmar que la imagen existe en GHCR.
3. Ejecutar `Deploy production` con esa etiqueta y solo la aplicación afectada.
4. Validar salud, logs y comportamiento público.
5. No cambiar ni eliminar el volumen SQL.

Si hubo una migración incompatible de base de datos, el rollback requiere un plan específico de datos; no restaurar producción impulsivamente sobre datos recientes.

---

## 16. Reglas para Codex y otros asistentes

Al trabajar con este repositorio:

1. Aplicar este documento únicamente si el proyecto se encuentra dentro de la carpeta raíz `.NET2026`.
2. Para proyectos fuera de `.NET2026`, buscar su propia documentación y no asumir que utilizan este VPS, Docker Compose, Cloudflare, SQL Server ni los workflows descritos aquí.
3. Leer primero este documento, `apps.compose.yml`, los Dockerfiles y los workflows relevantes.
4. Leer también el archivo Markdown funcional propio del proyecto seleccionado.
5. No asumir nombres de servicios, rutas, volúmenes o secretos; confirmarlos en archivos o con comandos de solo lectura.
6. Nunca mostrar valores de `.env`, claves privadas, tokens, contraseñas o cadenas de conexión completas.
7. No modificar el VPS cuando la solicitud sea solo explicar, revisar o diagnosticar.
8. Antes de cambios de infraestructura, presentar el impacto y el rollback.
9. Recrear solo el servicio afectado.
10. No usar `docker compose down -v`, `docker volume prune`, `docker system prune -a` ni publicar SQL en `0.0.0.0`.
11. Antes de tocar SQL, confirmar el volumen y el último backup verificable.
12. Después de un despliegue, verificar contenedor, logs, health checks y URL pública.
13. Mantener Cloudflare obligatorio para la web y SSH con claves para administración.

### Contexto corto para iniciar una tarea con Codex

```text
Si el proyecto está dentro de .NET2026, lee
DOCUMENTACION_INFRAESTRUCTURA_FRALSE.md antes de actuar.
Si está fuera de .NET2026, no apliques esta infraestructura automáticamente.
La producción está en un VPS Ubuntu de Contabo y usa Cloudflare, Nginx,
Docker Compose y SQL Server en Docker. Las apps son Fralse Tech,
FralseCont y Zona Deportiva. No reveles secretos ni modifiques volúmenes SQL.
Primero inspecciona los archivos reales y confirma cualquier dato marcado
como pendiente. Para despliegues, cambia solo el servicio solicitado, conserva
rollback y valida estado, logs, health checks y URL pública.
```

### Información que Codex debe pedir si falta

- Repositorio y rama objetivo.
- Aplicación o servicio exacto.
- Resultado esperado.
- Si se autoriza solo diagnóstico o también implementación/despliegue.
- SHA o versión a publicar.
- Si hay migración de base de datos.
- Salida de comandos de solo lectura cuando el problema está en el VPS.

---

## 17. Datos pendientes de confirmar en el repositorio

Marcar estos puntos cuando se valide la configuración fuente:

- [ ] Rutas exactas de los tres `.csproj` y DLL.
- [ ] Nombres exactos de los servicios de `apps.compose.yml`.
- [ ] Nombre real del volumen persistente de SQL Server.
- [ ] Redes Docker reales y qué aplicación pertenece a cada red.
- [ ] Nombres exactos de secretos de GitHub Actions.
- [ ] Nombre y contenido revisado del script `/usr/local/sbin/fralse-deploy`.
- [ ] Endpoints reales de health check por aplicación.
- [ ] Confirmar que SQL está enlazado a `127.0.0.1:1433` y actualizar el BAT.
- [ ] Política de retención y protección de paquetes en GHCR.
- [ ] Procedimiento probado de rollback con una etiqueta SHA.

Cuando se confirme un dato, reemplazar el marcador correspondiente y actualizar la fecha de consolidación.

---

## 18. Comandos de inventario para actualizar esta documentación

Estos comandos no deberían revelar secretos, salvo donde se advierte:

```bash
sudo docker ps --format 'table {{.Names}}\t{{.Image}}\t{{.Status}}\t{{.Ports}}'
sudo docker image ls
sudo docker volume ls
sudo docker network ls
sudo docker compose -f /srv/fralse/compose/apps.compose.yml config --services
sudo docker inspect fralse-sqlserver \
  --format '{{range .Mounts}}Tipo={{.Type}} Nombre={{.Name}} Origen={{.Source}} Destino={{.Destination}}{{println}}{{end}}'
sudo systemctl list-timers fralse-sql-backup.timer --no-pager
sudo ufw status verbose
sudo ss -tulpn
```

No compartir la salida completa de:

```text
docker compose config
docker inspect ... .Config.Env
cat /srv/fralse/secrets/*.env
cat /root/.config/rclone/rclone.conf
```

porque puede contener credenciales.
