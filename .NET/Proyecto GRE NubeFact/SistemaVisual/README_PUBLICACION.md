# SistemaVisual

Instalador y actualizador WinForms para .NET Framework 4.8, compatible con Visual Studio 2019.

## Configuración

El archivo `actualizador.config.json` se copia junto a `SistemaVisual.exe` y utiliza `https://sistema-visual-api.larasoft-dev.workers.dev/api/archivos`. El dominio de descarga es `https://actualizaciones.fralsetech.com`. No coloque credenciales, tokens ni secretos.

La instalación local siempre se administra en `C:\Sistema Visual`. Si la carpeta no existe, se crea `.actualizador`, `.actualizador\temporal`, `Respaldo` y `Respaldo\Logs`, y se instala todo el prefijo remoto. Un marcador interno permite reanudar una instalación interrumpida.

Si la carpeta ya existe, se comparan ETag, tamaño y SHA-256 registrado en `.actualizador\estado.json`. Los archivos modificados se respaldan en `Respaldo\yyyy-MM-dd_HHmmss`. Nunca se eliminan archivos locales ausentes en R2.

`Conexion.ini` y todos los archivos de `Temporales` se comparan y actualizan igual que los demás archivos. Si existen y cambiaron, se respaldan antes de reemplazarse; si faltan, se descargan como archivos nuevos. `Temporales` nunca se vacía y sus archivos locales ausentes en R2 no se eliminan.

Si el Worker no está disponible y ya existe una instalación completa con `C:\Sistema Visual\Ventas.exe`, el actualizador registra la incidencia y abre la versión local sin actualizar. En una instalación nueva o marcada como incompleta, la falta de conexión bloquea la ejecución y permite reintentar; nunca se abre un sistema incompleto.

## Compilación y distribución

Compile `SistemaVisual` en `Release`. El instalador de Inno Setup toma los archivos de `bin\Release\net48` e incorpora el redistribuible oficial sin conexión ubicado en `Prerequisitos\ndp48-x86-x64-allos-enu.exe`.

Al ejecutar `InstaladorSistemaVisual.exe`, se comprueba primero si el equipo tiene .NET Framework 4.8 o superior. Si falta, se instala automáticamente desde el paquete incorporado, sin requerir Internet. Cuando .NET solicita reiniciar Windows, el instalador termina correctamente, solicita el reinicio y no abre `SistemaVisual.exe` hasta una ejecución posterior.

El contenido de `bin\Release\net48` debe incluir:

- `SistemaVisual.exe`
- `Newtonsoft.Json.dll`
- `actualizador.config.json`

La aplicación solicita permisos de administrador mediante `app.manifest`.

Compile `InstaladorSistemaVisual.iss` con Inno Setup para generar `Salida\InstaladorSistemaVisual.exe`. Ese único ejecutable contiene SistemaVisual y el prerrequisito de .NET Framework 4.8 necesario para instalar en otras máquinas.

## Prueba aislada

Para no afectar datos reales, cambie temporalmente `CarpetaLocal` a una carpeta de pruebas y use un Worker/bucket de pruebas. El código permite HTTP únicamente cuando el destino es loopback, facilitando pruebas locales controladas.

Compruebe instalación, reanudación, archivos nuevos, reemplazos, `Conexion.ini`, contenido bajo `Temporales`, indisponibilidad de red, estado dañado y rutas inválidas antes de pasar a producción.
