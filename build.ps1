# SuperPUWEtty2 Build Script for Windows (PowerShell)
# Requires MSBuild (Visual Studio Build Tools or .NET SDK)

Write-Host "================================================" -ForegroundColor Cyan
Write-Host "SuperPUWEtty2 Build Script" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

# Function to find MSBuild
function Find-MSBuild {
    $possiblePaths = @(
        "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files (x86)\Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files (x86)\Microsoft Visual Studio\2019\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files (x86)\Microsoft Visual Studio\2019\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
    )

    foreach ($path in $possiblePaths) {
        if (Test-Path $path) {
            return $path
        }
    }

    # Try to find via vswhere
    $vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
    if (Test-Path $vswhere) {
        $vsPath = & $vswhere -latest -products * -requires Microsoft.Component.MSBuild -property installationPath
        if ($vsPath) {
            $msbuildPath = Join-Path $vsPath "MSBuild\Current\Bin\MSBuild.exe"
            if (Test-Path $msbuildPath) {
                return $msbuildPath
            }
        }
    }

    return $null
}

# Find MSBuild
$msbuildPath = Find-MSBuild

if (-not $msbuildPath) {
    Write-Host "ERROR: MSBuild not found!" -ForegroundColor Red
    Write-Host ""
    Write-Host "Please install one of the following:"
    Write-Host "  1. Visual Studio 2019 or 2022 (any edition)"
    Write-Host "  2. Visual Studio Build Tools"
    Write-Host ""
    Write-Host "Download from: https://visualstudio.microsoft.com/downloads/" -ForegroundColor Yellow
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host "Found MSBuild: $msbuildPath" -ForegroundColor Green
Write-Host ""

# Clean old build artifacts
Write-Host "Cleaning old build artifacts..." -ForegroundColor Yellow
$cleanPaths = @(
    "SuperPUWEtty2\bin",
    "SuperPUWEtty2\obj",
    "SuperPUWEtty2UnitTests\bin",
    "SuperPUWEtty2UnitTests\obj",
    "SuperPUWEtty2Installer\bin",
    "SuperPUWEtty2Installer\obj"
)

foreach ($path in $cleanPaths) {
    if (Test-Path $path) {
        Remove-Item -Path $path -Recurse -Force
        Write-Host "  Removed: $path"
    }
}
Write-Host ""

# Check for NuGet
$nugetPath = "nuget.exe"
if (-not (Get-Command nuget -ErrorAction SilentlyContinue)) {
    Write-Host "NuGet not found in PATH. Trying to download..." -ForegroundColor Yellow
    $nugetUrl = "https://dist.nuget.org/win-x86-commandline/latest/nuget.exe"
    try {
        Invoke-WebRequest -Uri $nugetUrl -OutFile "nuget.exe"
        $nugetPath = ".\nuget.exe"
        Write-Host "NuGet downloaded successfully." -ForegroundColor Green
    } catch {
        Write-Host "WARNING: Could not download NuGet. Package restore may fail." -ForegroundColor Yellow
    }
}

# Restore NuGet packages
Write-Host "Restoring NuGet packages..." -ForegroundColor Yellow
& $nugetPath restore SuperPUWEtty2.sln
Write-Host ""

# Build configuration
$configuration = "Release"
$platform = "Any CPU"

Write-Host "================================================" -ForegroundColor Cyan
Write-Host "Building SuperPUWEtty2 ($configuration)..." -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

# Build the solution
$buildArgs = @(
    "SuperPUWEtty2.sln",
    "/t:Rebuild",
    "/p:Configuration=$configuration",
    "/p:Platform=`"$platform`"",
    "/v:minimal",
    "/m"
)

& $msbuildPath $buildArgs

if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "================================================" -ForegroundColor Red
    Write-Host "BUILD FAILED!" -ForegroundColor Red
    Write-Host "================================================" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""
Write-Host "================================================" -ForegroundColor Green
Write-Host "BUILD SUCCESSFUL!" -ForegroundColor Green
Write-Host "================================================" -ForegroundColor Green
Write-Host ""

# Find the executable
$exePath = "SuperPUWEtty2\bin\$configuration\SuperPUWEtty2.exe"
if (Test-Path $exePath) {
    Write-Host "Output: $exePath" -ForegroundColor Green
    Write-Host ""

    $response = Read-Host "Start SuperPUWEtty2 now? (Y/n)"
    if ($response -eq "" -or $response -eq "Y" -or $response -eq "y") {
        Write-Host "Starting SuperPUWEtty2..." -ForegroundColor Cyan
        Start-Process $exePath
    }
} else {
    Write-Host "Warning: Executable not found at expected location: $exePath" -ForegroundColor Yellow
    Write-Host "Check the build output above for errors." -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
}

exit 0
