@echo off
REM Build script for DemoApp
setlocal

set CONFIG=%1
set PLATFORM=%2

if "%CONFIG%"=="" set CONFIG=Debug
if "%PLATFORM%"=="" set PLATFORM=x64

cd /d "%~dp0"

echo Building DemoApp.sln (%CONFIG%^|%PLATFORM%)...

REM Find Visual Studio and MSBuild
for /f "usebackq tokens=*" %%i in (`"C:\Program Files (x86)\Microsoft Visual Studio\Installer\vswhere.exe" -latest -property installationPath`) do (
    set VS_PATH=%%i
)

if not defined VS_PATH (
    echo ERROR: Visual Studio not found
    exit /b 1
)

set MSBUILD=%VS_PATH%\MSBuild\Current\Bin\MSBuild.exe
if not exist "%MSBUILD%" set MSBUILD=%VS_PATH%\MSBuild\15.0\Bin\MSBuild.exe

if not exist "%MSBUILD%" (
    echo ERROR: MSBuild.exe not found
    exit /b 1
)

"%MSBUILD%" DemoApp.sln /p:Configuration=%CONFIG% /p:Platform=%PLATFORM% /v:minimal /nologo

if errorlevel 1 (
    echo Build failed!
    exit /b 1
)

echo Build completed successfully!
