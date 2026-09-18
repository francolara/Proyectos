# Automatizacion de cancelacion de reservas en Ubuntu

Este paquete esta preparado para ejecutar, cada 10 minutos, exclusivamente el endpoint interno de Zona Deportiva. No debe usarse ni el dominio publico ni Cloudflare.

## Archivos

- `zona-deportiva-autocancel`: llama por `POST` a `http://127.0.0.1:5101/jobs/reservas/autocancelar-no-confirmadas` y transmite el token solo por el encabezado `X-Jobs-Token`.
- `zona-deportiva-autocancel.service`: limita cada ejecucion a dos minutos. El script usa `flock` para impedir instancias simultaneas y registra una coincidencia como `omitido_ya_en_ejecucion`.
- `zona-deportiva-autocancel.timer`: espera 10 minutos desde el final de la ejecucion anterior. Esto evita acumulaciones cuando una llamada tarda mas de lo normal.

El script escribe en el journal la fecha en `America/Lima`, duracion, HTTP y estado. No muestra el token ni el cuerpo de respuesta. No conecta a SQL Server ni contiene credenciales SQL.

## Instalacion pendiente de autorizacion

Copiar los archivos con propietario `root:root`:

```bash
install -o root -g root -m 0750 zona-deportiva-autocancel /usr/local/sbin/zona-deportiva-autocancel
install -o root -g root -m 0644 zona-deportiva-autocancel.service /etc/systemd/system/zona-deportiva-autocancel.service
install -o root -g root -m 0644 zona-deportiva-autocancel.timer /etc/systemd/system/zona-deportiva-autocancel.timer
install -d -o root -g root -m 0700 /srv/fralse/secrets
install -o root -g root -m 0600 /dev/null /srv/fralse/secrets/zona-deportiva-jobs-token
```

El ultimo archivo debe recibir el token existente de `Jobs__Token` mediante un metodo seguro en el servidor. No se debe pegar ni registrar el token en comandos, historial, unidades systemd o logs.

Despues de la autorizacion se validaran las unidades con `systemd-analyze verify`, se recargara systemd y se habilitara el timer. Ninguno de esos pasos ha sido realizado.

## Verificacion posterior

```bash
systemctl list-timers zona-deportiva-autocancel.timer
journalctl -u zona-deportiva-autocancel.service --since today
```
