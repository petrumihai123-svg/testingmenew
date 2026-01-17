@echo off
setlocal

REM Builds a self-contained, single-file portable EXE in .\publish\win-x64

set CONFIG=Release
set RID=win-x64
set OUTDIR=%~dp0publish\%RID%

if not exist "%OUTDIR%" mkdir "%OUTDIR%"

dotnet publish -c %CONFIG% -r %RID% --self-contained true ^
  /p:PublishSingleFile=true ^
  /p:IncludeNativeLibrariesForSelfExtract=true ^
  /p:PublishTrimmed=false ^
  -o "%OUTDIR%"

if errorlevel 1 (
  echo.
  echo Publish FAILED.
  exit /b 1
)

echo.
echo Publish OK: %OUTDIR%
endlocal
