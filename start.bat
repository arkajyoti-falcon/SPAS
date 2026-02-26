@echo off
setlocal enabledelayedexpansion

echo ========================================
echo       SPAS Application Starter
echo ========================================
echo.

:: Check if venv exists
if exist "venv\Scripts\activate.bat" (
    echo [OK] Virtual environment found.
    call venv\Scripts\activate.bat
) else (
    echo [INFO] Virtual environment not found. Creating venv...
    py -m venv venv
    if errorlevel 1 (
        echo [ERROR] Failed to create virtual environment. Make sure Python is installed.
        pause
        exit /b 1
    )
    echo [OK] Virtual environment created.
    call venv\Scripts\activate.bat
)

:: Check and install required pip modules
echo.
echo [INFO] Installing pip modules...

:: Count total packages
set /a total=0
for /f "usebackq eol=#" %%i in ("requirements.txt") do set /a total+=1

:: Install each package with progress
set /a count=0
for /f "usebackq eol=#" %%i in ("requirements.txt") do (
    set /a count+=1
    echo.
    echo [!count!/%total%] Installing %%i...
    pip install %%i
    if errorlevel 1 (
        echo [ERROR] Failed to install %%i
        pause
        exit /b 1
    )
)
echo.
echo [OK] All %total% pip modules installed successfully.

:: Check if .env file exists
echo.
if exist ".env" (
    echo [OK] .env file found.
) else (
    echo [WARNING] .env file not found. Please create a .env file with required environment variables.
    pause
    exit /b 1
)

:: Run the application
echo.
echo [INFO] Starting Streamlit application...
echo ========================================
streamlit run app.py --server.maxUploadSize 5000 --server.address 0.0.0.0

pause
