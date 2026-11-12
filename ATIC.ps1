# ============ CONFIGURACION INTERACTIVA ============
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host "     ATIC WORD SANITIZER - CONFIGURACION" -ForegroundColor Cyan
Write-Host "===============================================" -ForegroundColor Cyan

$carpetaOrigen = Read-Host "Ingrese la ruta de la carpeta de ORIGEN (donde estan los archivos)"
if (-not (Test-Path -LiteralPath $carpetaOrigen -PathType Container)) {
    throw "La carpeta de origen especificada no existe o no es una ruta valida."
}

# Obtener automáticamente el nombre de la carpeta de origen para crear la de destino por defecto
$nombreOrigen = Split-Path $carpetaOrigen -Leaf
$padreOrigen = Split-Path $carpetaOrigen -Parent
$sugerenciaDestino = Join-Path $padreOrigen ($nombreOrigen + "_limpio")

Write-Host "Carpeta de destino sugerida: $sugerenciaDestino" -ForegroundColor Yellow
$customDestino = Read-Host "Presione ENTER para usar esta sugerencia o escriba otra ruta"

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
Get-ChildItem -LiteralPath $carpetaOrigen -Recurse -File |
    Where-Object { $extensiones -contains $_.Extension.ToLower() } |
    Unblock-File -ErrorAction SilentlyContinue

$archivos = Get-ChildItem -LiteralPath $carpetaOrigen -Recurse -File |
            Where-Object { $extensiones -contains $_.Extension.ToLower() }

Write-Host "`nArchivos encontrados: $($archivos.Count) | Timeout por archivo: ${timeoutSeg}s`n" -ForegroundColor Cyan

$resultados = @()
$n = 0

foreach ($a in $archivos) {
    $n++
    Write-Host "[$n/$($archivos.Count)] $($a.Name)" -ForegroundColor Cyan

    $src = $a.FullName
    
    # Mantener la estructura de subcarpetas en el destino si existen
    $relPath = $src.Substring((Get-Item -LiteralPath $carpetaOrigen).FullName.Length)
    $dst = Join-Path $carpetaDestino $relPath
    $dst = [System.IO.Path]::ChangeExtension($dst, ".docx")

    if ($src -ieq $dst) {
        $dir  = [System.IO.Path]::GetDirectoryName($src)
        $base = [System.IO.Path]::GetFileNameWithoutExtension($src)
        $dst  = Join-Path $dir ($base + "_limpio.docx")
    }

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
        $resultados += [pscustomobject]@{ Archivo=$a.Name; Estado="TIMEOUT"; Detalle="" }
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