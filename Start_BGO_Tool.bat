@echo off
echo Dang chuan bi chay tool BGO...
echo.

REM Check python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python chua duoc cai dat!
    echo Vui long cai dat Python tu: https://www.python.org/downloads/
    echo Luu y: Khi cai dat, nho tick vao "Add Python to PATH"
    pause
    exit /b
)

REM Check pip
pip --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Pip chua duoc cai dat! Dang cai dat pip...
    python -m ensurepip --default-pip
    if %errorlevel% neq 0 (
        echo Loi: Khong tcai duoc pip. thu lai....
        pause
        exit /b
    )
)

REM Check selenium
python -c "import selenium" >nul 2>&1
if %errorlevel% neq 0 (
    echo Selenium chua duoc cai dat.
    echo Dang cai dat cac thu vien can thiet...
    pip install -r requirements.txt
    if %errorlevel% neq 0 (
        echo Loi:  thu lai.
        pause
        exit /b
    )
) else (
    echo Thu vien da duoc cai dat. Dang chay tool...
)

echo.
echo Dang chay tool...
python bgo_auto_tool.py

pause
