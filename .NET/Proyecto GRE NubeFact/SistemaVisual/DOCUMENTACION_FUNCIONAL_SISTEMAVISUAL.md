# Documentación funcional de SistemaVisual

## 1. Objetivo

SistemaVisual es el instalador y actualizador del sistema de ventas. Su responsabilidad es preparar una computadora cliente, sincronizar los archivos publicados en Cloudflare R2 y finalmente ejecutar:

`C:\Sistema Visual\Ventas.exe`

La solución se divide en tres componentes:

1. **Instalador Inno Setup:** instala el actualizador y sus prerrequisitos.
2. **SistemaVisual:** aplicación WinForms que instala o actualiza los archivos del sistema.
3. **CloudflareWorkerSistemaVisual:** API de solo lectura que enumera los objetos publicados en R2.

## 2. Ubicación de los componentes

- Actualizador: `Proyecto GRE NubeFact\SistemaVisual\`
- Worker: `Proyecto GRE NubeFact\CloudflareWorkerSistemaVisual\`
- Script del instalador: `SistemaVisual\InstaladorSistemaVisual.iss`
- Instalador generado: `SistemaVisual\Salida\InstaladorSistemaVisual.exe`
- Configuración incluida: `SistemaVisual\actualizador.config.json`
- Prueba de integración: `SistemaVisual\Pruebas\PruebaIntegracionLocal.ps1`
- Comando unificado de compilación: `SistemaVisual\GenerarInstalador.ps1`

## 3. Infraestructura configurada

- Worker: `https://sistema-visual-api.larasoft-dev.workers.dev/api/archivos`
- Dominio público de descargas: `https://actualizaciones.fralsetech.com`
- Bucket R2: `sistemavisual-actualizaciones`
- Prefijo remoto: `PROVEEPERU/Sistema Visual/`
- Directorio final del sistema: `C:\Sistema Visual`
- Ejecutable principal: `Ventas.exe`

Las computadoras cliente no reciben credenciales de Cloudflare. El binding de R2 se resuelve dentro del Worker y el contenido se descarga mediante el dominio público.

## 4. Funcionalidad del instalador

El instalador generado con Inno Setup:

1. Solicita permisos de administrador.
2. Comprueba si existe .NET Framework 4.8 o superior.
3. Si falta .NET Framework 4.8, lo instala desde el redistribuible oficial incluido; no necesita Internet para este prerrequisito.
4. Si .NET requiere reiniciar Windows, marca el reinicio y no intenta abrir prematuramente el actualizador.
5. Instala `SistemaVisual.exe`, su configuración y dependencias en `{autopf}\FRALSETECH\SistemaVisual`.
6. Crea accesos directos en el menú Inicio y en el escritorio.
7. Al finalizar puede ejecutar SistemaVisual con los permisos administrativos del instalador.

El redistribuible requerido debe existir en:

`Prerequisitos\ndp48-x86-x64-allos-enu.exe`

La carpeta `Prerequisitos/` se mantiene fuera de Git por el tamaño del redistribuible. Cada computadora que compile el instalador debe descargar previamente el instalador oficial sin conexión de .NET Framework 4.8 desde Microsoft y conservarlo con ese nombre y ubicación. `GenerarInstalador.ps1` detiene el proceso si el archivo no está disponible.

## 5. Funcionamiento del Worker

El Worker acepta únicamente:

`GET /api/archivos`

Para cualquier otra ruta responde 404 y para cualquier otro método responde 405. No permite subir, reemplazar ni eliminar objetos.

El proceso del Worker es:

1. Lee `PREFIJO_REMOTO` y `DOMINIO_PUBLICO` desde `wrangler.toml`.
2. Lista el binding `SISTEMA_VISUAL_BUCKET` usando el prefijo configurado.
3. Procesa páginas de hasta 1.000 objetos y continúa mediante el cursor de R2.
4. Ignora marcadores de carpetas y conserva solo objetos que representan archivos.
5. Convierte cada clave en una ruta relativa respecto del prefijo.
6. Devuelve los archivos ordenados por ruta.
7. Deshabilita la caché de la respuesta para mostrar el estado actual de R2.

Cada archivo devuelto contiene:

- `ruta`: ruta relativa que se reproducirá bajo `C:\Sistema Visual`.
- `clave`: clave completa dentro de R2.
- `tamano`: tamaño remoto en bytes.
- `etag`: identificador del objeto remoto.
- `fechaModificacion`: fecha de carga informada por R2.
- `url`: dirección pública codificada del objeto.

La respuesta general contiene `prefijo`, `fechaConsulta` y `archivos`.

## 6. Reglas para publicar archivos

R2 representa la versión publicada del sistema. Para publicar:

1. Compilar o preparar los archivos definitivos del sistema de ventas.
2. Subirlos debajo de `PROVEEPERU/Sistema Visual/`.
3. Conservar exactamente la estructura relativa deseada en el cliente.
4. Reemplazar en R2 los archivos modificados y agregar los nuevos.
5. No es necesario crear ZIP, manifiesto ni número de versión.
6. Consultar el Worker y comprobar que los objetos aparecen con sus rutas y tamaños correctos.
7. Probar la URL pública de al menos un archivo antes de distribuir la actualización.

Ejemplos de correspondencia:

| R2 | Computadora cliente |
| --- | --- |
| `PROVEEPERU/Sistema Visual/Ventas.exe` | `C:\Sistema Visual\Ventas.exe` |
| `PROVEEPERU/Sistema Visual/Conexion.ini` | `C:\Sistema Visual\Conexion.ini` |
| `PROVEEPERU/Sistema Visual/Temporales/Plantillas/base.txt` | `C:\Sistema Visual\Temporales\Plantillas\base.txt` |

No deben publicarse las carpetas técnicas `Respaldo/` ni `.actualizador/`. Si aparecen accidentalmente, SistemaVisual las omite.

## 7. Determinación del modo de operación

### 7.1 Instalación completa

Se activa cuando no existe `C:\Sistema Visual` o existe el marcador:

`C:\Sistema Visual\.actualizador\instalacion.pendiente`

El marcador permite reanudar una instalación interrumpida. Se elimina únicamente después de comprobar que el proceso terminó y que `Ventas.exe` existe.

### 7.2 Actualización

Se activa cuando `C:\Sistema Visual` ya existe y no tiene un marcador de instalación pendiente.

## 8. Reglas de instalación completa

Durante una instalación completa:

1. Se crean `C:\Sistema Visual`, `.actualizador`, `.actualizador\temporal`, `Respaldo`, `Respaldo\Logs` y `Temporales`.
2. Se consulta el listado completo del Worker.
3. Se descargan todos los archivos válidos publicados bajo el prefijo remoto.
4. Se preserva exactamente la estructura de carpetas y subcarpetas.
5. Se incluyen `Ventas.exe`, `Conexion.ini`, `Temporales/` y cualquier extensión de archivo.
6. No se crean respaldos porque no existe una versión anterior válida.
7. El estado se guarda progresivamente en `.actualizador\estado.json`.
8. Si el proceso se interrumpe, los archivos completados se conservan y la próxima ejecución continúa en modo instalación.
9. Antes de finalizar se comprueba que exista `Ventas.exe`.
10. Solo después de completar correctamente se elimina el marcador y se abre `Ventas.exe`.

## 9. Reglas de actualización

Durante una actualización:

1. Se consulta el Worker y se valida el prefijo recibido.
2. Cada archivo remoto se clasifica como nuevo, actualizado, sin cambios u omitido.
3. Un archivo inexistente se descarga como nuevo.
4. Un archivo modificado se respalda antes de reemplazarlo.
5. Un archivo sin cambios se conserva.
6. Los archivos locales que ya no aparezcan en R2 no se eliminan.
7. `Conexion.ini` se procesa igual que los demás archivos: si cambió, se respalda y reemplaza; si falta, se descarga.
8. Todos los archivos dentro de `Temporales/` se procesan con las mismas reglas.
9. La carpeta `Temporales` nunca se vacía automáticamente.
10. `Respaldo/` y `.actualizador/` siempre se omiten por ser carpetas locales protegidas.

Los respaldos usan una carpeta con fecha y hora, conservando la ruta relativa original:

`C:\Sistema Visual\Respaldo\yyyy-MM-dd_HHmmss\`

Ejemplo:

`C:\Sistema Visual\Respaldo\2026-07-21_163000\Temporales\Plantillas\plantilla.txt`

## 10. Comparación y estado local

El archivo `.actualizador\estado.json` registra por archivo:

- ETag remoto.
- Tamaño remoto.
- Fecha de modificación remota.
- SHA-256 calculado después de descargar.
- Fecha de instalación local.

La comparación sigue estas reglas:

1. Si el archivo local no existe, es nuevo.
2. Si no existe estado previo, se considera actualizado para resincronizarlo de forma segura.
3. Si cambió ETag o tamaño, se considera actualizado.
4. Si coinciden ETag y tamaño, se calcula SHA-256 del archivo local y se compara con el estado guardado.
5. Si `estado.json` está dañado, se conserva una copia `.danado.<fecha>.json` y se reconstruye el estado durante la resincronización.

El estado se escribe de forma atómica después de completar cada archivo.

## 11. Descarga y reemplazo seguro

- Se usa TLS 1.2 o superior.
- La configuración predeterminada usa un timeout de 60 segundos y hasta 3 intentos.
- Los reintentos esperan progresivamente antes de repetirse.
- Cada descarga se escribe primero en `.actualizador\temporal`.
- Se valida que la cantidad de bytes descargados coincida con el tamaño anunciado.
- Se calcula SHA-256 antes de aplicar el archivo.
- Los reemplazos usan una operación segura del sistema de archivos.
- Se verifica espacio disponible; en actualización se reserva aproximadamente el doble por los respaldos.
- Si `Ventas.exe` está abierto y debe modificarse, se solicita un cierre normal. Nunca se termina el proceso a la fuerza.

## 12. Funcionamiento sin conexión

Si no se puede consultar el Worker:

- Si existe una instalación completa, no hay marcador pendiente y existe `C:\Sistema Visual\Ventas.exe`, se registra el problema y se abre la versión local sin actualizar.
- Si es una instalación nueva o incompleta, no se abre el sistema y se permite reintentar.

Si la consulta al Worker sí terminó pero la conexión se pierde durante una descarga, la operación se detiene, conserva los archivos ya completados y permite reintentar. En ese caso no se abre automáticamente el sistema, porque podría existir una actualización parcialmente aplicada.

## 13. Seguridad de rutas

SistemaVisual rechaza:

- Rutas absolutas recibidas desde el Worker.
- Segmentos `.` o `..`.
- Nombres con caracteres inválidos, espacios o puntos finales.
- Rutas que intenten salir de `C:\Sistema Visual`.
- Rutas duplicadas ignorando mayúsculas y minúsculas.
- Directorios intermedios que sean enlaces o puntos de unión.
- Claves remotas que no correspondan exactamente al prefijo y ruta informados.
- Intentos de reemplazar el propio `SistemaVisual.exe` mientras está ejecutándose.

`actualizador.config.json` solo acepta URLs HTTPS, excepto HTTP hacia loopback para pruebas locales, y no permite credenciales dentro de las URLs.

## 14. Registros y diagnóstico

El log se guarda en:

`C:\Sistema Visual\Respaldo\Logs\actualizador.log`

Registra consultas, reintentos, archivos sin cambios, omisiones, respaldos, reemplazos, errores, funcionamiento sin conexión y apertura de `Ventas.exe`.

## 15. Regla obligatoria para cualquier cambio

Después de cualquier modificación relacionada con SistemaVisual, su configuración, el Worker, las pruebas o esta documentación, se debe regenerar el instalador. El trabajo no se considera terminado mientras no se ejecute correctamente:

```powershell
.\GenerarInstalador.ps1
```

El script realiza obligatoriamente:

1. Compilación de `SistemaVisual.csproj` en `Release`.
2. Ejecución de `Pruebas\PruebaIntegracionLocal.ps1` sobre la salida `Release`.
3. Compilación de `InstaladorSistemaVisual.iss` con Inno Setup.
4. Comprobación de que se generó `Salida\InstaladorSistemaVisual.exe`.
5. Presentación del tamaño y SHA-256 del instalador final.

Si falla cualquiera de esos pasos, no debe entregarse ni publicarse el instalador.

## 16. Lista mínima de validación

- Instalación nueva con Internet.
- Reanudación de instalación interrumpida.
- Actualización con archivos nuevos y modificados.
- Respaldo de `Ventas.exe`, `Conexion.ini` y archivos de `Temporales`.
- Conservación de archivos locales ausentes en R2.
- Inicio sin conexión con instalación completa.
- Bloqueo sin conexión con instalación incompleta.
- Estado dañado y reconstrucción.
- Rechazo de rutas remotas inseguras.
- Instalación automática de .NET Framework 4.8 cuando falte.
- Apertura final sin error 740.

