@echo off
setlocal

set APPNAME=BOHO_ImportBOM_THDM
set VERSION=2.1

echo.
echo ============================================
echo   BUILD - BOHO IMPORT BOM/THDM v%VERSION%
echo ============================================
echo.
set ISCC="C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if not exist %ISCC% set ISCC="C:\Program Files\Inno Setup 6\ISCC.exe"
if not exist %ISCC% set ISCC="C:\Users\%USERNAME%\AppData\Local\Programs\Inno Setup 6\ISCC.exe"

:: --- Kiem tra PyInstaller ---
echo [Check] PyInstaller...
python -m PyInstaller --version >nul 2>&1
if errorlevel 1 (
    echo [LOI] Chua cai PyInstaller. Chay: pip install pyinstaller
    pause
    exit /b 1
)
echo [OK] PyInstaller san sang.

:: --- Kiem tra Inno Setup ---
echo [Check] Inno Setup 6...
if not exist %ISCC% (
    echo [LOI] Khong tim thay Inno Setup 6.
    echo       Tai ve: https://jrsoftware.org/isdl.php
    echo       (Chi can cho ban INSTALLER, ban PORTABLE van build duoc)
    echo.
)

:: --- Kiem tra ODBC Driver file ---
echo [Check] redist\msodbcsql.msi...
if not exist "redist\msodbcsql.msi" (
    echo [CANH BAO] Chua co file redist\msodbcsql.msi
    echo            Installer se KHONG tu dong cai ODBC Driver.
    echo.
) else (
    echo [OK] ODBC Driver file san sang.
)

:: ============================================================
:: BUOC 1: PyInstaller — build file .exe
:: ============================================================
echo.
echo --------------------------------------------
echo  [1/3] Build exe voi PyInstaller...
echo --------------------------------------------

if exist "build" rmdir /s /q "build"
if exist "dist\BOM_Tool.exe" del /f /q "dist\BOM_Tool.exe"

python -m PyInstaller --onefile --noconsole ^
    --name BOM_Tool ^
    --icon=icon.ico ^
    --add-data "icon.ico;." ^
    --add-data "config;config" ^
    --collect-all customtkinter ^
    --hidden-import pyodbc ^
    --hidden-import pandas ^
    --hidden-import openpyxl ^
    --hidden-import openpyxl.cell._writer ^
    --hidden-import rapidfuzz ^
    --hidden-import msoffcrypto ^
    --hidden-import tksheet ^
    main.py

if errorlevel 1 (
    echo [LOI] PyInstaller build that bai!
    pause
    exit /b 1
)

if not exist "dist\BOM_Tool.exe" (
    echo [LOI] Khong tim thay dist\BOM_Tool.exe!
    pause
    exit /b 1
)
echo [OK] dist\BOM_Tool.exe tao thanh cong.

if not exist "installer_output" mkdir "installer_output"

:: ============================================================
:: BUOC 2: Ban PORTABLE — giai nen la chay, khong can cai dat
:: ============================================================
echo.
echo --------------------------------------------
echo  [2/3] Tao ban PORTABLE (ZIP)...
echo --------------------------------------------

set PORTABLE_DIR=installer_output\%APPNAME%_v%VERSION%_Portable
set PORTABLE_ZIP=installer_output\%APPNAME%_v%VERSION%_Portable.zip

if exist "%PORTABLE_DIR%" rmdir /s /q "%PORTABLE_DIR%"
mkdir "%PORTABLE_DIR%"
mkdir "%PORTABLE_DIR%\config"

:: Copy exe + config
copy /y "dist\BOM_Tool.exe"        "%PORTABLE_DIR%\BOHO_ImportBOM_THDM.exe" >nul
copy /y "config\CK_Mapping_v5.xlsx" "%PORTABLE_DIR%\config\"            >nul
copy /y "config\db_config.json"     "%PORTABLE_DIR%\config\"            >nul

:: Nen thanh ZIP — retry neu AV/file dang bi khoa.
:: Compress-Archive bao loi "being used by another process" o dang
:: NON-TERMINATING (Write-Error tu ben trong), nen try/catch KHONG bat duoc
:: neu thieu -ErrorAction Stop — retry se im lang khong chay, loi se van
:: hien ra ngoai nhu chua he co try/catch. Bat buoc phai co Stop o day.
if exist "%PORTABLE_ZIP%" del /f /q "%PORTABLE_ZIP%"
powershell -NoProfile -Command ^
  "$src='%PORTABLE_DIR%'; $dst='%PORTABLE_ZIP%'; $ok=$false;" ^
  "for($i=0;$i-lt8;$i++){" ^
  "  try{ Compress-Archive -Path \"$src\*\" -DestinationPath $dst -Force -ErrorAction Stop; $ok=$true; break }" ^
  "  catch{ Write-Host \"[Retry $($i+1)/8] File dang bi khoa (AV/Explorer/app khac dang mo)...\"; Start-Sleep 3 }" ^
  "}" ^
  "if(-not $ok){ Write-Error 'ZIP that bai sau 8 lan thu — dong file trong config (vd db_config.json dang mo o Notepad/VSCode), tat tam AV, hoac dong BOM_Tool.exe dang chay tu portable roi build lai.'; exit 1 }"

if errorlevel 1 (
    echo [CANH BAO] Nen ZIP Portable that bai.
) else (
    echo [OK] Ban Portable: %PORTABLE_ZIP%
)

:: ============================================================
:: BUOC 3: Ban INSTALLER — Setup.exe (can Inno Setup)
:: ============================================================
echo.
echo --------------------------------------------
echo  [3/3] Tao ban INSTALLER (Setup.exe)...
echo --------------------------------------------

if not exist %ISCC% (
    echo [BO QUA] Khong tim thay Inno Setup 6 — chi co ban Portable.
    goto :done
)

%ISCC% "setup.iss"

if errorlevel 1 (
    echo [LOI] Inno Setup build that bai!
    goto :done
)
echo [OK] Ban Installer: installer_output\%APPNAME%_Setup_v%VERSION%.exe

:done
echo.
echo ============================================
echo  HOAN THANH!
echo.
echo  Ban PORTABLE  : installer_output\%APPNAME%_v%VERSION%_Portable.zip
echo    -^> Giai nen, chay BOHO_ImportBOM_THDM.exe truc tiep
echo.
echo  Ban INSTALLER : installer_output\%APPNAME%_Setup_v%VERSION%.exe
echo    -^> Chay Setup.exe -^> Next -^> Finish
echo    -^> Tu dong cai ODBC Driver (neu co file redist)
echo ============================================
echo.
pause
