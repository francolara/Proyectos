param(
    [string]$RootPath = $PSScriptRoot,
    [string]$OutputFileName = "99_SP_Finales.sql"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-FileOrder {
    param([string]$FileName)
    if ($FileName -match "^(?<n>\d+)") {
        return [int]$Matches["n"]
    }
    return 9999
}

function New-Utf8NoBomEncoding {
    return [System.Text.UTF8Encoding]::new($false)
}

$outputPath = Join-Path $RootPath $OutputFileName

$sqlFiles = Get-ChildItem -Path $RootPath -Filter "*.sql" -File |
    Where-Object { $_.Name -ne $OutputFileName } |
    Sort-Object @{ Expression = { Get-FileOrder $_.Name } }, Name

$procedures = @{}

foreach ($file in $sqlFiles) {
    $lines = Get-Content -LiteralPath $file.FullName | ForEach-Object { $_.TrimEnd("`r") }
    $i = 0

    while ($i -lt $lines.Count) {
        $line = $lines[$i]
        if ($line -match "^\s*CREATE\s+OR\s+ALTER\s+PROCEDURE\s+dbo\.([A-Za-z0-9_]+)\b") {
            $spName = $Matches[1]
            $start = $i
            $j = $i + 1

            while ($j -lt $lines.Count -and $lines[$j] -notmatch "^\s*GO\s*$") {
                $j++
            }

            if ($j -lt $lines.Count) {
                $end = $j
            }
            else {
                $end = $lines.Count - 1
            }

            $block = ($lines[$start..$end] -join "`n").TrimEnd()
            if (-not $block.EndsWith("`n")) {
                $block += "`n"
            }

            $procedures[$spName] = [PSCustomObject]@{
                SpName     = $spName
                SourceFile = $file.Name
                FileOrder  = Get-FileOrder $file.Name
                Line       = $start + 1
                Block      = $block
            }

            $i = $end + 1
            continue
        }

        $i++
    }
}

if ($procedures.Count -eq 0) {
    throw "No se encontraron definiciones CREATE OR ALTER PROCEDURE en $RootPath."
}

$today = Get-Date -Format "dd/MM/yyyy"
$header = @"
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   $today
-- Description:   Consolidado final de stored procedures (ultima version efectiva por nombre) generado automaticamente.
-- Firma:         Codex - $today | Script final para evitar sobreescritura por orden de despliegue.
-- =============================================
-- REGLA DE USO:
-- 1) Ejecutar primero los scripts estructurales y funcionales (00..32).
-- 2) Ejecutar este archivo al final.
-- 3) Regenerar este archivo con Generate-99_SP_Finales.ps1 cada vez que cambie un SP.

USE [DbSportCenter];
GO

"@

$winners = $procedures.Values |
    Sort-Object FileOrder, SourceFile, Line, SpName

$sb = [System.Text.StringBuilder]::new()
[void]$sb.Append($header)

foreach ($proc in $winners) {
    [void]$sb.AppendLine("-- SOURCE: $($proc.SourceFile) (linea $($proc.Line))")
    [void]$sb.Append($proc.Block)
    [void]$sb.AppendLine()
}

$content = $sb.ToString().Replace("`r`n", "`n").Replace("`r", "")
[System.IO.File]::WriteAllText($outputPath, $content, (New-Utf8NoBomEncoding))

Write-Host "Generado: $outputPath"
Write-Host "Total SP consolidados: $($winners.Count)"
