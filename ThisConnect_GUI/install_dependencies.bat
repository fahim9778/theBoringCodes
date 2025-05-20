@echo off
echo *************************************************
echo *          ThisConnect Dependency Installer         *
echo *************************************************
echo.
echo This script will install the necessary Python packages for ThisConnect.
echo Required packages: pandas, selenium, webdriver-manager
echo.
echo Please make sure Python is installed and added to your PATH before proceeding.
echo.
pause

echo Checking Python installation...
python --version > nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo Python is not installed or not in your PATH.
    echo Please install Python from https://www.python.org/downloads/
    echo and make sure to check "Add Python to PATH" during installation.
    echo.
    pause
    exit /b 1
)

echo Python is installed. Installing required packages...
echo.

echo Installing pandas...
pip install pandas
if %ERRORLEVEL% NEQ 0 (
    echo Failed to install pandas. Please try manually: pip install pandas
)

echo Installing selenium...
pip install selenium
if %ERRORLEVEL% NEQ 0 (
    echo Failed to install selenium. Please try manually: pip install selenium
)

echo Installing webdriver-manager...
pip install webdriver-manager
if %ERRORLEVEL% NEQ 0 (
    echo Failed to install webdriver-manager. Please try manually: pip install webdriver-manager
)

echo.
echo *************************************************
echo *          Installation Complete!                *
echo *************************************************
echo.
echo If there were any errors, please try installing the packages manually:
echo pip install pandas selenium webdriver-manager
echo.
echo You can now run ThisConnect.py successfully!
echo.
pause
