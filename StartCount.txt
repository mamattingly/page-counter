@echo off
:: Save the current directory
set "currentDir=%cd%"

:: Start PowerShell and run the script
powershell -NoProfile -ExecutionPolicy RemoteSigned -Command "& {cd '%currentDir%'; .\DocumentPageCount.ps1; Pause}"
