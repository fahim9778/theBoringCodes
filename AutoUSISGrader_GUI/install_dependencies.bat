@echo off

:: Check for Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python is not installed. Please install Python first.
    pause
    exit /b
)

:: Check for pip
python -m pip --version >nul 2>&1
if %errorlevel% neq 0 (
    echo pip not found. Installing pip...
    python -c "import urllib.request; urllib.request.urlretrieve('https://bootstrap.pypa.io/get-pip.py', 'get-pip.py')"
    python get-pip.py
)

:: Install required packages
echo Installing required Python packages...
pip install pandas selenium webdriver_manager tk openpyxl
echo Installation complete.
pause
