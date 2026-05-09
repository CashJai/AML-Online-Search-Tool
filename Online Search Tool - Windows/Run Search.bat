@echo off
title Online-Search-Tool V2

echo Checking prerequisites...
node -v >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo ========================================================
    echo  ERROR: Node.js is not installed or not in your PATH!
    echo  This tool requires Node.js to run.
    echo  Please download and install it from: https://nodejs.org/
    echo ========================================================
    echo.
    pause
    exit /b
)

echo Setting up Online-Search-Tool V2...
if not exist node_modules (
    echo Installing dependencies...
    npm install
)
echo Starting CLI...
node collect_input.js
pause
