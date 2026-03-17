@echo off
title Installation Script all in One
color 0A
setlocal enabledelayedexpansion

:: =========================
:: Main Header
:: =========================
cls
echo.
echo ======================================================
echo           Installation Script all in One
echo ======================================================
echo.
echo Disable Windows Defender before proceed.
echo.
pause

:: Create temp working folder
set "WORKDIR=%TEMP%\all_in_one_install"
if not exist "%WORKDIR%" mkdir "%WORKDIR%"

:: =========================
:: Progress Bar Function
:: =========================
goto :main

:progress
setlocal EnableDelayedExpansion
set "STEP=%~1"
set "TEXT=%~2"

cls
echo.
echo ======================================================
echo           Installation Script all in One
echo ======================================================
echo.
echo Current Step: !STEP!/6
echo !TEXT!
echo.

set "BAR="
for /L %%A in (1,1,%STEP%) do set "BAR=!BAR!##"
for /L %%A in (%STEP%,1,5) do set "BAR=!BAR!.."
echo Progress: [!BAR!]
echo.
endlocal
exit /b

:main

:: ---------------------------------------------------------
:: 1. Install Ninite
:: ---------------------------------------------------------
call :progress 1 "Downloading and installing Ninite..."
set "NINITE_URL=https://ninite.com/7zip-chrome-vlc-winrar-zoom/ninite.exe"
set "NINITE_FILE=%WORKDIR%\ninite.exe"

powershell -Command "try { Invoke-WebRequest -Uri '%NINITE_URL%' -OutFile '%NINITE_FILE%' -UseBasicParsing } catch { exit 1 }"
if errorlevel 1 (
    echo Failed to download Ninite.
    pause
) else (
    start /wait "" "%NINITE_FILE%"
    echo Ninite installation finished.
    timeout /t 2 >nul
)

:: ---------------------------------------------------------
:: 2. Install Office
:: ---------------------------------------------------------
call :progress 2 "Downloading and installing Microsoft Office..."
set "OFFICE_URL=https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=O365EduCloudRetail&platform=x64&language=en-us&version=O16GA"
set "OFFICE_FILE=%WORKDIR%\OfficeSetup.exe"

powershell -Command "try { Invoke-WebRequest -Uri '%OFFICE_URL%' -OutFile '%OFFICE_FILE%' -UseBasicParsing } catch { exit 1 }"
if errorlevel 1 (
    echo Failed to download Office installer.
    pause
) else (
    start /wait "" "%OFFICE_FILE%"
    echo Office installation finished.
    timeout /t 2 >nul
)

:: ---------------------------------------------------------
:: 3. Google Sheets (Serial Number)
:: ---------------------------------------------------------
call :progress 3 "Opening Google Sheets for Laptop Serial Number..."
start "" "https://docs.google.com/spreadsheets/d/12d8Mx9k8LqKlMekrAp3OFEqPkW38UNwj2SbmJMH0goQ/edit?gid=473635777#gid=473635777"

:ASK_SERIAL
echo.
set /p SERIALDONE=You entered serial number in Google Sheets? (Y/N): 
if /I "%SERIALDONE%"=="Y" goto STEP4
if /I "%SERIALDONE%"=="N" (
    echo Please enter the Serial Number in Google Sheets.
    pause
    goto STEP4
)
echo Invalid input. Please type Y or N.
goto ASK_SERIAL

:: ---------------------------------------------------------
:: 4. Wallpaper
:: ---------------------------------------------------------
:STEP4
call :progress 4 "Opening DILG wallpaper link..."
start "" "https://bit.ly/r10intensitywp"

:ASK_WALL
echo.
set /p WALLDONE=You apply the wallpaper? (Y/N): 
if /I "%WALLDONE%"=="Y" goto STEP5
if /I "%WALLDONE%"=="N" (
    echo Please apply the image as wallpaper.
    pause
    goto STEP5
)
echo Invalid input. Please type Y or N.
goto ASK_WALL

:: ---------------------------------------------------------
:: 5. Download Nitro Pro 13
:: ---------------------------------------------------------
:STEP5
call :progress 5 "Downloading Nitro Pro 13..."
set "NITRO_URL=http://bit.ly/nitropro13exe"
set "NITRO_FILE=%WORKDIR%\nitropro13.exe"

powershell -Command "try { Invoke-WebRequest -Uri '%NITRO_URL%' -OutFile '%NITRO_FILE%' -UseBasicParsing } catch { exit 1 }"

if exist "%NITRO_FILE%" (
    echo Download completed.
    echo File saved to: %NITRO_FILE%
    echo Opening download folder...
    explorer "%WORKDIR%"
) else (
    echo Failed to download Nitro Pro.
)

echo.
pause
goto STEP6

:: ---------------------------------------------------------
:: 6. Download and Run Activate.cmd
:: ---------------------------------------------------------
:STEP6
call :progress 6 "Downloading and running Activate.cmd..."

set "ACTIVATE_URL=https://raw.githubusercontent.com/iamrudyard/rx-script/main/Activate.cmd"
set "ACTIVATE_FILE=%WORKDIR%\Activate.cmd"

echo Downloading Activate.cmd...
powershell -Command "Invoke-WebRequest -Uri '%ACTIVATE_URL%' -OutFile '%ACTIVATE_FILE%' -UseBasicParsing"

if exist "%ACTIVATE_FILE%" (
    echo Download complete.
    echo Running as Administrator...
    
    powershell -Command "Start-Process '%ACTIVATE_FILE%' -Verb RunAs"
) else (
    echo Failed to download Activate.cmd.
    pause
    goto FINISH
)

echo.
echo Instructions:
echo - Activate Windows: Select 1
echo - Activate Microsoft Office: Select 2 then 1
echo.

pause
goto FINISH
)

echo.
echo Instructions:
echo - Activate Windows: Select 1
echo - Activate Microsoft Office: Select 2 then 1
echo.
pause

:: ---------------------------------------------------------
:: Finish
:: ---------------------------------------------------------
:FINISH
cls
echo.
echo ======================================================
echo         All job done successfully.
echo ======================================================
echo.
pause

endlocal
exit /b
