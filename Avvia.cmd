@echo off
rem Avvia Double Tap Detector in una finestra ridotta a icona (resta visibile nella barra delle applicazioni).
start "" /min "%SystemRoot%\System32\conhost.exe" powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0double-tap-detector.ps1"
