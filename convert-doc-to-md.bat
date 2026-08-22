@echo off
setlocal
chcp 65001 >nul

set "SCRIPT=%~dp0tools\doc-to-docx\convert-doc-to-md.ps1"
set "TARGET=%~1"

if not defined TARGET set "TARGET=%~dp0inbox"

if not exist "%SCRIPT%" (
    echo DOC to Markdown converter not found: %SCRIPT%
    exit /b 2
)

if not exist "%TARGET%" (
    echo Input path not found: %TARGET%
    exit /b 3
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -Path "%TARGET%" -Recurse
set "EXITCODE=%ERRORLEVEL%"

echo.
if not "%EXITCODE%"=="0" (
    echo DOC to Markdown conversion failed. Exit code: %EXITCODE%
    exit /b %EXITCODE%
)

echo DOC to Markdown conversion completed.
exit /b 0
