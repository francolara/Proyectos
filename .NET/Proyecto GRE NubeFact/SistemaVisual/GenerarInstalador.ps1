$ErrorActionPreference = "Stop"

$proyecto = $PSScriptRoot
$archivoProyecto = Join-Path $proyecto "SistemaVisual.csproj"
$pruebaIntegracion = Join-Path $proyecto "Pruebas\PruebaIntegracionLocal.ps1"
$scriptInno = Join-Path $proyecto "InstaladorSistemaVisual.iss"
$prerrequisito = Join-Path $proyecto "Prerequisitos\ndp48-x86-x64-allos-enu.exe"
$compiladorInno = Join-Path ${env:ProgramFiles(x86)} "Inno Setup 6\ISCC.exe"
$instalador = Join-Path $proyecto "Salida\InstaladorSistemaVisual.exe"

foreach ($archivoRequerido in @($archivoProyecto, $pruebaIntegracion, $scriptInno, $prerrequisito, $compiladorInno)) {
    if (-not (Test-Path -LiteralPath $archivoRequerido -PathType Leaf)) {
        throw "No se encontro el archivo requerido: $archivoRequerido"
    }
}

Push-Location $proyecto
try {
    Write-Host "Compilando SistemaVisual en Release..."
    & dotnet build $archivoProyecto -c Release
    if ($LASTEXITCODE -ne 0) {
        throw "La compilacion Release de SistemaVisual fallo."
    }

    Write-Host "Ejecutando la prueba de integracion..."
    & $pruebaIntegracion -Configuracion Release
    if ($LASTEXITCODE -ne 0) {
        throw "La prueba de integracion fallo."
    }

    Write-Host "Generando el instalador con Inno Setup..."
    & $compiladorInno $scriptInno
    if ($LASTEXITCODE -ne 0) {
        throw "La compilacion del instalador fallo."
    }

    if (-not (Test-Path -LiteralPath $instalador -PathType Leaf)) {
        throw "Inno Setup termino sin generar el instalador esperado."
    }

    $archivoSalida = Get-Item -LiteralPath $instalador
    $hash = Get-FileHash -Algorithm SHA256 -LiteralPath $instalador
    [PSCustomObject]@{
        Instalador = $archivoSalida.FullName
        Bytes = $archivoSalida.Length
        SHA256 = $hash.Hash
        Generado = $archivoSalida.LastWriteTime
    } | Format-List
}
finally {
    Pop-Location
}

