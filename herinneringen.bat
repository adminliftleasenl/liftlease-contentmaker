@echo off
chcp 65001 >nul
set "PYTHONUTF8=1"
cd /d "%~dp0"
title Liftlease - Content-herinneringen

rem Bepaal het Python-commando (python of py)
set "PY=python"
where python >nul 2>nul || set "PY=py"

rem Controleer of Python aanwezig is
%PY% --version >nul 2>nul
if errorlevel 1 (
  echo.
  echo Python is niet gevonden op deze pc.
  echo Installeer Python via https://www.python.org/downloads/ en vink "Add Python to PATH" aan.
  echo.
  pause
  exit /b
)

rem Eenmalig: zorg dat de benodigde pakketten aanwezig zijn (openpyxl + pywin32 voor Outlook)
%PY% -c "import openpyxl, win32com.client" 1>nul 2>nul
if errorlevel 1 (
  echo.
  echo Eenmalige installatie van benodigde onderdelen ^(openpyxl, pywin32^)...
  echo Even geduld, dit kan een minuut duren.
  echo.
  %PY% -m pip install openpyxl pywin32
  echo.
  echo Installatie klaar.
  timeout /t 2 >nul
)

:menu
cls
echo ==================================================
echo    Liftlease - Content-herinneringen
echo ==================================================
echo.
echo    [1] Klaarzetten in Outlook (deze week)
echo        - opent de mails; JIJ drukt zelf op Verzenden
echo    [2] Alleen tonen / testen (deze week, verstuurt niets)
echo    [3] Test naar MIJN eigen adres (klaargezet in Outlook)
echo    [4] Voorbeeld-test januari 2026 (alleen tonen - om de AI te checken)
echo    [5] Sluiten
echo.
set "keuze="
set /p "keuze=Maak een keuze (1-5): "

if "%keuze%"=="1" goto klaar
if "%keuze%"=="2" goto tonen
if "%keuze%"=="3" goto eigen
if "%keuze%"=="4" goto voorbeeld
if "%keuze%"=="5" exit /b
goto menu

:klaar
echo.
echo De mails worden klaargezet in Outlook. Er wordt NIETS automatisch verzonden.
echo.
%PY% stuur_herinneringen.py --via outlook --klaarzetten
goto einde

:tonen
echo.
%PY% stuur_herinneringen.py --dry-run
goto einde

:eigen
echo.
set "adres="
set /p "adres=Typ je eigen e-mailadres: "
echo.
%PY% stuur_herinneringen.py --via outlook --klaarzetten --test "%adres%" --datum 2026-01-05 --dagen 12
goto einde

:voorbeeld
echo.
%PY% stuur_herinneringen.py --dry-run --datum 2026-01-05 --dagen 12
goto einde

:einde
echo.
echo Klaar. Druk op een toets om dit venster te sluiten.
pause >nul
