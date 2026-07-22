param(
    [string]$Configuracion = "Debug",
    [switch]$ConservarTemporal
)

$ErrorActionPreference = "Stop"
$proyecto = Split-Path -Parent $PSScriptRoot
$salida = Join-Path $proyecto "bin\$Configuracion\net48"
[Reflection.Assembly]::LoadFrom((Join-Path $salida "Newtonsoft.Json.dll")) | Out-Null
[Reflection.Assembly]::LoadFrom((Join-Path $salida "SistemaVisual.exe")) | Out-Null

$raizPrueba = Join-Path ([IO.Path]::GetTempPath()) ("SistemaVisualIntegracion-" + [Guid]::NewGuid().ToString("N"))
$raizServidor = Join-Path $raizPrueba "servidor"
$raizCliente = Join-Path $raizPrueba "cliente"
$procesoServidor = $null

function Escribir-Remoto([string]$ruta, [string]$contenido) {
    $destino = Join-Path $raizServidor ($ruta.Replace('/', '\'))
    [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($destino)) | Out-Null
    [IO.File]::WriteAllBytes($destino, [Text.UTF8Encoding]::new($false).GetBytes($contenido))
}

function Crear-Objeto([string]$ruta, [string]$etag, [string]$contenido) {
    [pscustomobject]@{
        ruta = $ruta
        clave = "PROVEEPERU/Sistema Visual/$ruta"
        tamano = [Text.UTF8Encoding]::new($false).GetByteCount($contenido)
        etag = $etag
        fechaModificacion = "2026-07-21T16:00:00Z"
        url = "http://127.0.0.1:$puerto/PROVEEPERU/Sistema%20Visual/$ruta"
    }
}

function Publicar-Listado([object[]]$archivos) {
    $directorioApi = Join-Path $raizServidor "api"
    [IO.Directory]::CreateDirectory($directorioApi) | Out-Null
    $respuesta = [pscustomobject]@{
        prefijo = "PROVEEPERU/Sistema Visual/"
        fechaConsulta = "2026-07-21T16:00:00Z"
        archivos = $archivos
    }
    [IO.File]::WriteAllText(
        (Join-Path $directorioApi "archivos"),
        ($respuesta | ConvertTo-Json -Depth 6),
        [Text.UTF8Encoding]::new($false))
}

function Nueva-Configuracion([string]$urlWorker) {
    $config = [SistemaVisual.Models.UpdateConfig]::new()
    $config.UrlWorker = $urlWorker
    $config.DominioDescargas = "http://127.0.0.1:$puerto"
    $config.NombreBucket = "sistemavisual-actualizaciones"
    $config.PrefijoRemoto = "PROVEEPERU/Sistema Visual/"
    $config.CarpetaLocal = $raizCliente
    $config.EjecutablePrincipal = "Ventas.exe"
    $config.CarpetaTemporal = ".actualizador\temporal"
    $config.CarpetaRespaldo = "Respaldo"
    $config.CarpetasExcluidas = [Collections.Generic.List[string]]::new()
    $config.CarpetasExcluidas.Add("Respaldo")
    $config.CarpetasExcluidas.Add(".actualizador")
    $config.TimeoutSegundos = 10
    $config.MaximoReintentos = 1
    $config.EsperaCierreVentasSegundos = 1
    return $config
}

function Ejecutar-Actualizador($config) {
    $rutas = [SistemaVisual.Services.AppPaths]::new($config)
    $log = [SistemaVisual.Services.LogService]::new($rutas.ArchivoLog)
    $servicio = [SistemaVisual.Services.UpdateService]::new($config, $rutas, $log)
    return $servicio.EjecutarAsync($null, $null, [Threading.CancellationToken]::None).GetAwaiter().GetResult()
}

try {
    [IO.Directory]::CreateDirectory($raizServidor) | Out-Null
    $listener = [Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback, 0)
    $listener.Start()
    $puerto = ([Net.IPEndPoint]$listener.LocalEndpoint).Port
    $listener.Stop()

    $ventas1 = "ventas-version-1"
    $conexion1 = "servidor=instalacion"
    $reporte1 = "reporte-version-1"
    $temporal1 = "plantilla-version-1"
    Escribir-Remoto "PROVEEPERU/Sistema Visual/Ventas.exe" $ventas1
    Escribir-Remoto "PROVEEPERU/Sistema Visual/Reportes/ReporteVentas.rpt" $reporte1
    Escribir-Remoto "PROVEEPERU/Sistema Visual/Temporales/Plantillas/plantilla.txt" $temporal1
    $listado1 = @(
        (Crear-Objeto "Ventas.exe" "etag-v1" $ventas1),
        (Crear-Objeto "Conexion.ini" "etag-c1" $conexion1),
        (Crear-Objeto "Reportes/ReporteVentas.rpt" "etag-r1" $reporte1),
        (Crear-Objeto "Temporales/Plantillas/plantilla.txt" "etag-t1" $temporal1)
    )
    Publicar-Listado $listado1

    $procesoServidor = Start-Process -FilePath "python" -ArgumentList @("-m", "http.server", $puerto, "--bind", "127.0.0.1") -WorkingDirectory $raizServidor -WindowStyle Hidden -PassThru
    $urlWorker = "http://127.0.0.1:$puerto/api/archivos"
    $servidorListo = $false
    for ($intento = 0; $intento -lt 25; $intento++) {
        try { Invoke-WebRequest -UseBasicParsing -Uri $urlWorker | Out-Null; $servidorListo = $true; break }
        catch { Start-Sleep -Milliseconds 200 }
    }
    if (-not $servidorListo) { throw "No se pudo iniciar el servidor HTTP local de pruebas." }

    $config = Nueva-Configuracion $urlWorker
    $instalacionInterrumpida = $false
    try { Ejecutar-Actualizador $config | Out-Null }
    catch { $instalacionInterrumpida = $true }
    $marcadorConservado = Test-Path (Join-Path $raizCliente ".actualizador\instalacion.pendiente")
    $ventasConservado = Test-Path (Join-Path $raizCliente "Ventas.exe")

    Escribir-Remoto "PROVEEPERU/Sistema Visual/Conexion.ini" $conexion1
    $resultadoInstalacion = Ejecutar-Actualizador $config
    $instalacionReanudada = $resultadoInstalacion.ModoInstalacion -and -not (Test-Path (Join-Path $raizCliente ".actualizador\instalacion.pendiente"))
    $sinRespaldosInstalacion = @(
        Get-ChildItem -LiteralPath (Join-Path $raizCliente "Respaldo") -Directory |
            Where-Object { $_.Name -ne "Logs" }
    ).Count -eq 0

    [IO.File]::WriteAllText((Join-Path $raizCliente "Conexion.ini"), "conexion-particular-cliente")
    [IO.Directory]::CreateDirectory((Join-Path $raizCliente "Temporales")) | Out-Null
    [IO.File]::WriteAllText((Join-Path $raizCliente "Temporales\solo-local.tmp"), "no eliminar")

    $ventas2 = "ventas-version-2"
    $conexion2 = "servidor=nuevo-no-aplicar"
    $temporal2 = "plantilla-version-2"
    $logo1 = "logo-nuevo"
    Escribir-Remoto "PROVEEPERU/Sistema Visual/Ventas.exe" $ventas2
    Escribir-Remoto "PROVEEPERU/Sistema Visual/Conexion.ini" $conexion2
    Escribir-Remoto "PROVEEPERU/Sistema Visual/Temporales/Plantillas/plantilla.txt" $temporal2
    Escribir-Remoto "PROVEEPERU/Sistema Visual/Logos/logo.png" $logo1
    $listado2 = @(
        (Crear-Objeto "Ventas.exe" "etag-v2" $ventas2),
        (Crear-Objeto "Conexion.ini" "etag-c2" $conexion2),
        (Crear-Objeto "Reportes/ReporteVentas.rpt" "etag-r1" $reporte1),
        (Crear-Objeto "Temporales/Plantillas/plantilla.txt" "etag-t2" $temporal2),
        (Crear-Objeto "Logos/logo.png" "etag-l1" $logo1)
    )
    Publicar-Listado $listado2

    $resultadoActualizacion = Ejecutar-Actualizador $config
    $conexionActualizada = [IO.File]::ReadAllText((Join-Path $raizCliente "Conexion.ini")) -eq $conexion2
    $temporalActualizado = [IO.File]::ReadAllText((Join-Path $raizCliente "Temporales\Plantillas\plantilla.txt")) -eq $temporal2
    $localNoEliminado = Test-Path (Join-Path $raizCliente "Temporales\solo-local.tmp")
    $carpetaRespaldo = Get-ChildItem -LiteralPath (Join-Path $raizCliente "Respaldo") -Directory |
        Where-Object { $_.Name -ne "Logs" } | Select-Object -First 1
    $respaldoVentas = Test-Path (Join-Path $carpetaRespaldo.FullName "Ventas.exe")
    $respaldoTemporal = Test-Path (Join-Path $carpetaRespaldo.FullName "Temporales\Plantillas\plantilla.txt")
    $conexionRespaldada = Test-Path (Join-Path $carpetaRespaldo.FullName "Conexion.ini")

    $resultadoSinCambios = Ejecutar-Actualizador $config
    $cantidadRespaldosAntes = @(Get-ChildItem -LiteralPath (Join-Path $raizCliente "Respaldo") -Directory | Where-Object { $_.Name -ne "Logs" }).Count
    Remove-Item -LiteralPath (Join-Path $raizCliente "Conexion.ini")
    $resultadoConexionFaltante = Ejecutar-Actualizador $config
    $cantidadRespaldosDespues = @(Get-ChildItem -LiteralPath (Join-Path $raizCliente "Respaldo") -Directory | Where-Object { $_.Name -ne "Logs" }).Count
    $conexionDescargadaComoNueva = [IO.File]::ReadAllText((Join-Path $raizCliente "Conexion.ini")) -eq $conexion2

    $configNoDisponible = Nueva-Configuracion "http://127.0.0.1:1/api/archivos"
    $resultadoSinConexion = Ejecutar-Actualizador $configNoDisponible
    $modoSinConexionPermiteAbrir =
        $resultadoSinConexion.SinConexion -and
        $resultadoSinConexion.SinCambios -and
        (Test-Path (Join-Path $raizCliente "Ventas.exe"))

    $marcadorInstalacion = Join-Path $raizCliente ".actualizador\instalacion.pendiente"
    [IO.File]::WriteAllText($marcadorInstalacion, "pendiente")
    $instalacionIncompletaSinConexionBloqueada = $false
    try { Ejecutar-Actualizador $configNoDisponible | Out-Null }
    catch {
        $instalacionIncompletaSinConexionBloqueada =
            $_.Exception.Message -like "*servidor de actualizaciones*"
    }
    Remove-Item -LiteralPath $marcadorInstalacion

    $resultados = [pscustomobject]@{
        InstalacionInterrumpidaDetectada = $instalacionInterrumpida
        MarcadorPermiteReanudar = $marcadorConservado
        ArchivoValidadoConservado = $ventasConservado
        InstalacionReanudada = $instalacionReanudada
        InstalacionSinRespaldos = $sinRespaldosInstalacion
        ConexionModificadaActualizada = $conexionActualizada
        TemporalActualizado = $temporalActualizado
        ArchivoLocalAusenteEnR2Conservado = $localNoEliminado
        VentasRespaldado = $respaldoVentas
        TemporalRespaldado = $respaldoTemporal
        ConexionAnteriorRespaldada = $conexionRespaldada
        ArchivoNuevoLogos = ($resultadoActualizacion.Nuevos -eq 1)
        ActualizadosEsperados = ($resultadoActualizacion.Actualizados -eq 3)
        SinCambiosDetectado = $resultadoSinCambios.SinCambios
        ConexionFaltanteDescargada = $conexionDescargadaComoNueva
        ConexionNuevaSinRespaldoAdicional = ($cantidadRespaldosAntes -eq $cantidadRespaldosDespues)
        ModoSinConexionPermiteAbrir = $modoSinConexionPermiteAbrir
        InstalacionIncompletaSinConexionBloqueada = $instalacionIncompletaSinConexionBloqueada
        EstadoJsonCreado = Test-Path (Join-Path $raizCliente ".actualizador\estado.json")
        LogCreado = Test-Path (Join-Path $raizCliente "Respaldo\Logs\actualizador.log")
    }

    $resultados | Format-List
    $pruebasFallidas = @(
        $resultados.PSObject.Properties |
            Where-Object { $_.Value -ne $true } |
            Select-Object -ExpandProperty Name
    )
    if ($pruebasFallidas.Count -gt 0) {
        throw "Fallaron las pruebas de integracion: $($pruebasFallidas -join ', ')"
    }
}
finally {
    if ($procesoServidor -and -not $procesoServidor.HasExited) {
        Stop-Process -Id $procesoServidor.Id -Force
    }
    $baseTemporal = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
    $resuelta = [IO.Path]::GetFullPath($raizPrueba)
    if (-not $ConservarTemporal -and
        $resuelta.StartsWith($baseTemporal, [StringComparison]::OrdinalIgnoreCase) -and
        (Split-Path $resuelta -Leaf).StartsWith("SistemaVisualIntegracion-")) {
        Remove-Item -LiteralPath $resuelta -Recurse -Force
    }
    elseif ($ConservarTemporal) {
        Write-Host "Directorio temporal conservado: $resuelta"
    }
}
