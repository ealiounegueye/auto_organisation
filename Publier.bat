@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

where dotnet >nul 2>&1
if errorlevel 1 (
  echo.
  echo Le SDK .NET 8 n'est pas installe sur CE PC ^(seulement le PC admin en a besoin^).
  echo Telechargez-le ici : https://dotnet.microsoft.com/download/dotnet/8.0
  echo Choisissez : SDK 8.0 pour Windows x64.
  echo.
  pause
  exit /b 1
)

echo Publication de OutlookOrganizer.exe ^(un fichier, sans Visual Studio^)...
dotnet publish "OutlookOrganizer\OutlookOrganizer.csproj" -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:EnableCompressionInSingleFile=true -o "dist"
if errorlevel 1 (
  echo.
  echo La publication a echoue.
  pause
  exit /b 1
)

echo.
echo Pret. Donnez aux collaborateurs le dossier dist :
echo   %cd%\dist\OutlookOrganizer.exe
echo   %cd%\dist\appsettings.json
echo.
echo Ils n'ont pas besoin de Visual Studio ni de .NET.
echo Double-clic sur OutlookOrganizer.exe.
echo.
pause
