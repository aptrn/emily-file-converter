@echo off
setlocal enabledelayedexpansion

echo Emily File Converter - Batch Processor
echo =======================================

if "%~1"=="" (
    echo.
    echo Usage: 
    echo   For single file: Drag and drop a .emily file onto this batch file
    echo   For folder:      Drag and drop a folder containing .emily files onto this batch file
    echo   Command line:    %~nx0 "path\to\input" [output_directory]
    echo.
    echo Output files will be named: originalname_MOVEJ.txt
    echo.
    pause
    exit /b 1
)

set "inputPath=%~1"
set "outputPath=%~2"

if not exist "%inputPath%" (
    echo Error: Input path "%inputPath%" not found!
    pause
    exit /b 1
)

echo Input path: %inputPath%

if "%outputPath%"=="" (
    if exist "%inputPath%\" (
        rem Input is a directory, output to same directory
        set "outputPath=%inputPath%"
    ) else (
        rem Input is a file, let PowerShell script handle the output path
        set "outputPath="
    )
)

echo Output path: !outputPath!
echo.

powershell.exe -ExecutionPolicy Bypass -File "%~dp0convert_emily.ps1" -InputPath "%inputPath%" -OutputPath "!outputPath!"

if %errorlevel% equ 0 (
    echo.
    echo =======================================
    echo Conversion completed successfully!
    echo Check the output path for your converted files.
) else (
    echo.
    echo =======================================
    echo Conversion failed with error code %errorlevel%
)

echo.
pause 