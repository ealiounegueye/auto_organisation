@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

where dotnet >nul 2>&1
if errorlevel 1 (
  echo.
  echo Installez le SDK .NET 8 sur CE PC admin seulement :
  echo https://dotnet.microsoft.com/download/dotnet/8.0
  echo.
  pause
  exit /b 1
)

echo Creation de l'application desktop OutlookOrganizer.exe ...
dotnet publish "OutlookOrganizer\OutlookOrganizer.csproj" -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:EnableCompressionInSingleFile=true -o "dist"
if errorlevel 1 (
  echo.
  echo La publication a echoue.
  pause
  exit /b 1
)

echo.
echo Envoyez UNIQUEMENT ce fichier aux collaborateurs :
echo   %cd%\dist\OutlookOrganizer.exe
echo.
echo Ils double-cliquent. L'application se connecte et organise la boite toute seule.
echo.
pause
