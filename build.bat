@echo off
REM SuperPUWEtty2 Build Script for Windows
REM Requires MSBuild (Visual Studio Build Tools or .NET SDK)

echo ================================================
echo SuperPUWEtty2 Build Script
echo ================================================
echo.

REM Try to find MSBuild
set MSBUILD_PATH=

REM Check for Visual Studio 2022
if exist "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" (
    set MSBUILD_PATH=C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe
    goto :found
)

REM Check for Visual Studio 2022 Professional
if exist "C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe" (
    set MSBUILD_PATH=C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe
    goto :found
)

REM Check for Visual Studio 2022 Enterprise
if exist "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe" (
    set MSBUILD_PATH=C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe
    goto :found
)

REM Check for Visual Studio 2019
if exist "C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe" (
    set MSBUILD_PATH=C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe
    goto :found
)

REM Check for Build Tools
if exist "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe" (
    set MSBUILD_PATH=C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe
    goto :found
)

REM MSBuild not found
echo ERROR: MSBuild not found!
echo.
echo Please install one of the following:
echo   1. Visual Studio 2019 or 2022 (any edition)
echo   2. Visual Studio Build Tools
echo.
echo Download from: https://visualstudio.microsoft.com/downloads/
echo.
pause
exit /b 1

:found
echo Found MSBuild: %MSBUILD_PATH%
echo.

REM Clean old build
echo Cleaning old build artifacts...
if exist "SuperPUWEtty2\bin" rmdir /s /q "SuperPUWEtty2\bin"
if exist "SuperPUWEtty2\obj" rmdir /s /q "SuperPUWEtty2\obj"
if exist "SuperPUWEtty2UnitTests\bin" rmdir /s /q "SuperPUWEtty2UnitTests\bin"
if exist "SuperPUWEtty2UnitTests\obj" rmdir /s /q "SuperPUWEtty2UnitTests\obj"
echo.

REM Restore NuGet packages
echo Restoring NuGet packages...
nuget restore SuperPUWEtty2.sln
if errorlevel 1 (
    echo WARNING: NuGet restore failed. Trying to continue...
    echo.
)

REM Build the solution
echo.
echo ================================================
echo Building SuperPUWEtty2 (Release)...
echo ================================================
echo.

"%MSBUILD_PATH%" SuperPUWEtty2.sln /t:Rebuild /p:Configuration=Release /p:Platform="Any CPU" /v:minimal

if errorlevel 1 (
    echo.
    echo ================================================
    echo BUILD FAILED!
    echo ================================================
    pause
    exit /b 1
)

echo.
echo ================================================
echo BUILD SUCCESSFUL!
echo ================================================
echo.
echo Output: SuperPUWEtty2\bin\Release\SuperPUWEtty2.exe
echo.

REM Check if executable exists
if exist "SuperPUWEtty2\bin\Release\SuperPUWEtty2.exe" (
    echo Starting SuperPUWEtty2...
    echo.
    start "" "SuperPUWEtty2\bin\Release\SuperPUWEtty2.exe"
) else (
    echo Warning: Executable not found at expected location.
    echo Check the build output above for errors.
    pause
)

exit /b 0
