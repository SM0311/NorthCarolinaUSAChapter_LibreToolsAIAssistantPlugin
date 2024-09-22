@echo off
:: Check if the script is running with administrator privileges
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo Requesting administrator privileges...
    :: Relaunch the script with administrator rights
    powershell -Command "Start-Process cmd -ArgumentList '/c, %~f0 %*' -Verb RunAs"
    exit /b
)

:: Hard-coded source and destination paths
set "source_path=%~dp0scriptlocatepython.py"
set "destination_path=C:\Program Files\LibreOffice\share\Scripts\python"

:: Check if source file exists
if not exist "%source_path%" (
    echo Error: Source file "%source_path%" does not exist.
    pause
    exit /b
)

:: Check if destination path exists
if not exist "%destination_path%" (
    echo Error: Destination path "%destination_path%" does not exist.
    pause
    exit /b
)

:: Copy the file
copy "%source_path%" "%destination_path%"
if %errorLevel% equ 0 (
    echo File "%source_path%" successfully copied to "%destination_path%".
) else (
    echo Error: Failed to copy the file.
)

pause


# Pause to keep terminal open
read -p "Press any key to exit..."
