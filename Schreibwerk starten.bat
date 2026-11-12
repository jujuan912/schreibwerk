@echo off
rem Startet Schreibwerk als eigenes App-Fenster (Brave, sonst Edge/Chrome, sonst Standardbrowser)
set "APP=%~dp0index.html"
set "URL=file:///%APP%"
set "BRAVE=%ProgramFiles%\BraveSoftware\Brave-Browser\Application\brave.exe"
if not exist "%BRAVE%" set "BRAVE=%LocalAppData%\BraveSoftware\Brave-Browser\Application\brave.exe"
if exist "%BRAVE%" (
  start "" "%BRAVE%" --app="%URL%"
  exit /b
)
set "EDGE=%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe"
if exist "%EDGE%" (
  start "" "%EDGE%" --app="%URL%"
  exit /b
)
start "" "%APP%"
