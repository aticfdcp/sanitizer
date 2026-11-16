# ============ CONFIGURACION INTERACTIVA ============
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host "   ATIC WORD SANITIZER - MODO UNIDAD COMPLETA" -ForegroundColor Cyan
Write-Host "===============================================" -ForegroundColor Cyan

$unidadOrigen = Read-Host "Ingrese la letra de la unidad o ruta base a revisar (Ej: D: o D:\MiCarpeta)"
if (-not (Test-Path -LiteralPath $unidadOrigen -PathType Container)) {
    throw "La unidad o ruta especificada no existe o no es valida."
}

# Limpiar formato de la unidad/ruta base
$sourceFull = (Get-Item -LiteralPath $unidadOrigen).FullName.TrimEnd('\') + '\'

# Definir la carpeta destino global (ej: D:\UnidadD_limpio o seleccionada por el usuario)
$padreOrigen = Split-Path $sourceFull -Parent
$nombreUnidad = (Get-Item -LiteralPath $sourceFull).Name
if ([string]::IsNullOrWhiteSpace($nombreUnidad)) { $nombreUnidad = "Raiz_" + $sourceFull.Substring(0,1) }
$sugerenciaDestino = Join-Path $padreOrigen ($nombreUnidad + "_limpio")

Write-Host "Carpeta de destino general sugerida: $sugerenciaDestino" -ForegroundColor Yellow
$customDestino = Read-Host "Presione ENTER para usar esta sugerencia o escriba otra ruta para el destino"
$carpetaDestino = if ([string]::IsNullOrWhiteSpace($customDestino)) { $sugerenciaDestino } else { $customDestino }

New-Item -ItemType Directory -Force -Path $carpetaDestino | Out-Null
Write-Host "Carpeta de destino configurada en: $carpetaDestino" -ForegroundColor Green

$extensiones = @('.doc', '.docm', '.docx')
$timeoutSeg  = 120
$testPrimero = $false      
# ===================================================

$worker = Join-Path $PSScriptRoot "_worker.ps1"
if (-not (Test-Path -LiteralPath $worker)) { throw "Falta el archivo _worker.ps1 junto a este script principal." }

Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
Write-Host "`nBuscando archivos compatibles en $sourceFull (esto puede tardar dependiendo del tamaño)..." -ForegroundColor Cyan

$archivos = Get-ChildItem -LiteralPath $sourceFull -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object { $extensiones -contains $_.Extension.ToLower() -and $_.FullName -notlike "$carpetaDestino*" }

Write-Host "Archivos encontrados para limpiar: $($archivos.Count) | Timeout por archivo: ${timeoutSeg}s`n" -ForegroundColor Cyan

$resultados = @()
$n = 0

foreach ($a in $archivos) {
    $n++
    Write-Host "[$n/$($archivos.Count)] $($a.Name)" -ForegroundColor Cyan

    $src = $a.FullName
    
    # Calcular la ruta relativa respecto a la unidad origen para replicar carpetas limpias
    $relPath = $src.Substring($sourceFull.Length)
    $dst = Join-Path $carpetaDestino $relPath
    $dst = [System.IO.Path]::ChangeExtension($dst, ".docx")

    # Asegurar que la subcarpeta de destino exista
    New-Item -ItemType Directory -Force -Path ([System.IO.Path]::GetDirectoryName($dst)) | Out-Null

    $stdout = [System.IO.Path]::GetTempFileName()
    $stderr = [System.IO.Path]::GetTempFileName()

    # -STA es imprescindible para Word COM en proceso separado
    $argsStr = "-NoProfile -ExecutionPolicy Bypass -STA -File `"$worker`" -Src `"$src`" -Dst `"$dst`""
    $p = Start-Process powershell.exe -ArgumentList $argsStr `
            -PassThru -WindowStyle Hidden `
            -RedirectStandardOutput $stdout -RedirectStandardError $stderr

    $termino = $p.WaitForExit($timeoutSeg * 1000)
    $out = if (Test-Path -LiteralPath $stdout) { Get-Content -LiteralPath $stdout -Raw } else { "" }
    $err = if (Test-Path -LiteralPath $stderr) { Get-Content -LiteralPath $stderr -Raw } else { "" }

    if (-not $termino) {
        try { $p.Kill() } catch {}
        Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
        Write-Host "    TIMEOUT (${timeoutSeg}s)" -ForegroundColor Yellow
        $resultados += [pscustomobject]@{ Archivo=$a.Name; Estado="TIMEOUT"; Detalle=$src }
    }
    elseif ($out -match "OK") {
        Write-Host "    GUARDADO: $dst" -ForegroundColor Green
        $resultados += [pscustomobject]@{ Archivo=$a.Name; Estado="OK"; Detalle=$dst }
    }
    else {
        $msg = ($out + $err).Trim()
        Write-Host "    FALLO: $msg" -ForegroundColor Red
        $resultados += [pscustomobject]@{ Archivo=$a.Name; Estado="FAIL"; Detalle=$msg }
    }

    Remove-Item -LiteralPath $stdout, $stderr -Force -ErrorAction SilentlyContinue
    Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

    if ($testPrimero) {
        Write-Host "`nModo prueba: detenido tras el primer archivo." -ForegroundColor Yellow
        break
    }
}

Write-Host "`n============== RESUMEN ==============" -ForegroundColor Cyan
$resultados | Format-Table -AutoSize
$csv = Join-Path $carpetaDestino ("reporte_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".csv")
$resultados | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8
Write-Host "Reporte guardado en: $csv" -ForegroundColor Cyan