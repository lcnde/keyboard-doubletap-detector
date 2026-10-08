<#
    Attiva o disattiva l'avvio automatico di Double Tap Detector all'accesso a Windows.
    Crea (o rimuove) un collegamento nella cartella "Esecuzione automatica" dell'utente:
    niente registro di sistema, niente privilegi di amministratore.

      powershell -ExecutionPolicy Bypass -File avvio-automatico.ps1            attiva
      powershell -ExecutionPolicy Bypass -File avvio-automatico.ps1 -Rimuovi   disattiva

    Il collegamento contiene il percorso attuale della cartella: se la sposti, rilancia questo script.
#>
param([switch]$Rimuovi)

$collegamento = Join-Path ([Environment]::GetFolderPath('Startup')) 'Double Tap Detector.lnk'

if ($Rimuovi) {
    if (Test-Path $collegamento) { Remove-Item $collegamento }
    Write-Host 'Avvio automatico disattivato.'
    return
}

$shell = New-Object -ComObject WScript.Shell
$lnk = $shell.CreateShortcut($collegamento)
$lnk.TargetPath = Join-Path $env:SystemRoot 'System32\conhost.exe'
$lnk.Arguments = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "' + (Join-Path $PSScriptRoot 'double-tap-detector.ps1') + '"'
$lnk.WorkingDirectory = $PSScriptRoot
$lnk.WindowStyle = 7   # ridotta a icona
$lnk.IconLocation = (Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe') + ',0'
$lnk.Description = 'Double Tap Detector'
$lnk.Save()
Write-Host "Avvio automatico attivato: $collegamento"
