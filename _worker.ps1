param([string]$Src, [string]$Dst)
$ErrorActionPreference = 'Stop'
$w = $null; $doc = $null; $nuevo = $null
$step = "init"
try {
    $step = "create-word"
    $w = New-Object -ComObject Word.Application
    $w.Visible            = $false
    $w.DisplayAlerts      = 0
    $w.AutomationSecurity = 3

    # EXACTAMENTE los 4 parametros que probamos y funcionaron
    $step = "open"
    $doc = $w.Documents.Open($Src, $false, $true, $false)

    $step = "add"
    $nuevo = $w.Documents.Add()

    $step = "copy-content"
    $nuevo.Content.FormattedText = $doc.Content.FormattedText

    $step = "copy-headers-footers"
    $numSec = [Math]::Min($doc.Sections.Count, $nuevo.Sections.Count)
    for ($s = 1; $s -le $numSec; $s++) {
        foreach ($h in 1,2,3) {
            try {
                $r = $doc.Sections.Item($s).Headers.Item($h).Range
                if ($r.Text.Trim().Length -gt 0) {
                    $nuevo.Sections.Item($s).Headers.Item($h).Range.FormattedText = $r.FormattedText
                }
            } catch {}
        }
        foreach ($f in 1,2,3) {
            try {
                $r = $doc.Sections.Item($s).Footers.Item($f).Range
                if ($r.Text.Trim().Length -gt 0) {
                    $nuevo.Sections.Item($s).Footers.Item($f).Range.FormattedText = $r.FormattedText
                }
            } catch {}
        }
        try {
            $sp = $doc.Sections.Item($s).PageSetup
            $dp = $nuevo.Sections.Item($s).PageSetup
            $dp.PaperSize    = $sp.PaperSize
            $dp.Orientation  = $sp.Orientation
            $dp.TopMargin    = $sp.TopMargin
            $dp.BottomMargin = $sp.BottomMargin
            $dp.LeftMargin   = $sp.LeftMargin
            $dp.RightMargin  = $sp.RightMargin
        } catch {}
    }

    $step = "save"
    if (Test-Path -LiteralPath $Dst) { Remove-Item -LiteralPath $Dst -Force }
    $nuevo.SaveAs([ref]$Dst, [ref]16)
    Unblock-File -LiteralPath $Dst -ErrorAction SilentlyContinue

    Write-Output "OK"
}
catch {
    Write-Output "ERROR [$step]: $($_.Exception.Message)"
}
finally {
    # Quit ANTES de liberar (el bug anterior)
    if ($w) { try { $w.Quit(0) } catch {} }
    foreach ($o in @($nuevo, $doc, $w)) {
        if ($null -ne $o) {
            try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($o) | Out-Null } catch {}
        }
    }
    Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}