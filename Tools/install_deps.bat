@echo off
echo ============================================
echo  BOHO Import BOM-THDM - Cai dat thu vien
echo ============================================
echo.

:: Kiem tra Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [LOI] Chua cai Python. Tai tai: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [OK] Python da co:
python --version
echo.

:: Upgrade pip truoc
echo [1/7] Nang cap pip...
python -m pip install --upgrade pip

echo.
echo [2/7] Cai pandas...
pip install pandas

echo.
echo [3/7] Cai customtkinter...
pip install customtkinter

echo.
echo [4/7] Cai openpyxl...
pip install openpyxl

echo.
echo [5/7] Cai pyodbc...
pip install pyodbc

echo.
echo [6/7] Cai tksheet (bang du lieu kieu Excel)...
pip install tksheet

echo.
echo [7/7] Cai rapidfuzz (goi y ma gan dung - fuzzy lookup)...
pip install rapidfuzz

echo.
echo [Optional] Cai msoffcrypto-tool (mo Excel co mat khau)...
pip install msoffcrypto-tool

echo.
echo ============================================
echo  Kiem tra ket qua:
echo ============================================
python -c "import pandas; print('pandas:', pandas.__version__)"
python -c "import customtkinter; print('customtkinter:', customtkinter.__version__)"
python -c "import openpyxl; print('openpyxl:', openpyxl.__version__)"
python -c "import pyodbc; print('pyodbc:', pyodbc.version)"
python -c "import rapidfuzz; print('rapidfuzz:', rapidfuzz.__version__)"
python -c "import tksheet; print('tksheet:', tksheet.__version__)"
python -c "import msoffcrypto; print('msoffcrypto-tool: OK')" 2>nul || echo msoffcrypto-tool: Khong co (khong bat buoc)

echo.
echo Hoan tat! Nhan phim bat ky de dong...
pause
