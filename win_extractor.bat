echo off

REM Switch to current directory
cd /d %~dp0

SET filename=%~pnx1
SET linux_filename=%filename:\=/%
SET filepath=/mnt/c%linux_filename%
echo %filepath%

SET /P speed="Speed? (just hit enter for 1.0): "

IF "%speed%" == "" (
powershell.exe wsl.exe python3 extractor.py %filepath%
) ELSE (
powershell.exe wsl.exe python3 extractor.py %filepath% --speed %speed%
)

powershell.exe sleep 5