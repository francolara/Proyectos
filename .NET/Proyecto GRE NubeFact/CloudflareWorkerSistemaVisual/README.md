# Cloudflare Worker de SistemaVisual

Worker HTTP de solo lectura que lista todos los objetos bajo `PROVEEPERU/Sistema Visual/`. El código solo expone `GET /api/archivos`; no implementa operaciones para subir, reemplazar o eliminar objetos.

## Requisitos

- Node.js y npm.
- Una cuenta Cloudflare con el bucket R2 `sistemavisual-actualizaciones` ya creado.
- Los archivos publicados bajo el prefijo `PROVEEPERU/Sistema Visual/`.

## Publicación exacta

Abra PowerShell en esta carpeta y ejecute:

```powershell
npm install
npx wrangler login
npm run check
npm run deploy
```

`wrangler.toml` declara el binding:

```toml
[[r2_buckets]]
binding = "SISTEMA_VISUAL_BUCKET"
bucket_name = "sistemavisual-actualizaciones"
```

El Worker configurado utiliza esta URL:

```text
https://sistema-visual-api.larasoft-dev.workers.dev
```

El endpoint que debe probar es:

```text
https://sistema-visual-api.larasoft-dev.workers.dev/api/archivos
```

Puede probarlo desde PowerShell:

```powershell
Invoke-RestMethod -Method Get -Uri "https://sistema-visual-api.larasoft-dev.workers.dev/api/archivos"
```

La respuesta debe contener `prefijo`, `fechaConsulta` y `archivos`. Si existen más de 1,000 objetos, el Worker continúa solicitando páginas mediante el cursor de R2 hasta completar el listado.

## Configurar el actualizador

Edite `actualizador.config.json` que se distribuye junto a `SistemaVisual.exe` y coloque la URL completa:

```json
"UrlWorker": "https://sistema-visual-api.larasoft-dev.workers.dev/api/archivos"
```

Mantenga:

```json
"DominioDescargas": "https://actualizaciones.fralsetech.com",
"NombreBucket": "sistemavisual-actualizaciones",
"PrefijoRemoto": "PROVEEPERU/Sistema Visual/"
```

No agregue tokens ni claves R2 a `actualizador.config.json`. El binding se resuelve dentro de Cloudflare y las computadoras cliente solo reciben el listado público.

## Prueba de descarga

Tome la propiedad `url` de cualquier objeto devuelto y compruebe que responde correctamente:

```powershell
Invoke-WebRequest -Method Head -Uri "https://actualizaciones.fralsetech.com/PROVEEPERU/Sistema%20Visual/Ventas.exe"
```

Publique primero los archivos en R2. No se requiere manifiesto, ZIP ni número de versión.
