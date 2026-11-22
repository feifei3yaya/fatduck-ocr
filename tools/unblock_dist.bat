@echo off
rem Removes Windows "downloaded from internet" marks (Zone.Identifier)
rem so SmartScreen does not warn on first run.
rem NOTE: this file must stay 100% ASCII (comments included).
rem cmd.exe misparses double-byte characters in batch files, so the
rem Chinese messages live in unblock_files.ps1 (UTF-8 BOM), which
rem PowerShell handles correctly.
setlocal
cd /d %~dp0
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0unblock_files.ps1"
endlocal
