# SuperPUWEtty2 Build Script (ohne Installer)
# Baut nur die Hauptanwendung, überspringt das Installer-Projekt

Write-Host "================================================" -ForegroundColor Cyan
Write-Host "SuperPUWEtty2 Build Script (ohne Installer)" -ForegroundColor Cyan
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
    Write-Host "Installiere Visual Studio Build Tools:" -ForegroundColor Yellow
    Write-Host "Invoke-WebRequest -Uri 'https://aka.ms/vs/17/release/vs_buildtools.exe' -OutFile 'vs_buildtools.exe'" -ForegroundColor Cyan
    Write-Host ".\vs_buildtools.exe --add Microsoft.VisualStudio.Workload.MSBuildTools --add Microsoft.Net.Component.4.5.TargetingPack --quiet --wait" -ForegroundColor Cyan
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
    "SuperPUWEtty2UnitTests\obj"
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
    if (-not (Test-Path "nuget.exe")) {
        Write-Host "Downloading NuGet..." -ForegroundColor Yellow
        $nugetUrl = "https://dist.nuget.org/win-x86-commandline/latest/nuget.exe"
        try {
            Invoke-WebRequest -Uri $nugetUrl -OutFile "nuget.exe"
            $nugetPath = ".\nuget.exe"
            Write-Host "NuGet downloaded successfully." -ForegroundColor Green
        } catch {
            Write-Host "WARNING: Could not download NuGet." -ForegroundColor Yellow
        }
    } else {
        $nugetPath = ".\nuget.exe"
    }
}

# Restore NuGet packages (nur für die Hauptprojekte)
Write-Host "Restoring NuGet packages..." -ForegroundColor Yellow
& $nugetPath restore SuperPUWEtty2.sln
Write-Host ""

# Build configuration
$configuration = "Release"

Write-Host "================================================" -ForegroundColor Cyan
Write-Host "Building SuperPUWEtty2 ($configuration)..." -ForegroundColor Cyan
Write-Host "Building only: SuperPUWEtty2 + SuperPUWEtty2UnitTests" -ForegroundColor Yellow
Write-Host "Skipping: SuperPUWEtty2Installer (requires WiX)" -ForegroundColor Yellow
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

# Build only the main project (skip installer)
$buildArgs = @(
    "SuperPUWEtty2\SuperPUWEtty2.csproj",
    "/t:Rebuild",
    "/p:Configuration=$configuration",
    "/v:minimal",
    "/m"
)

& $msbuildPath $buildArgs

if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "================================================" -ForegroundColor Red
    Write-Host "BUILD FAILED!" -ForegroundColor Red
    Write-Host "================================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "Mögliche Ursache: .NET Framework 4.5 Developer Pack fehlt" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Lösung 1: Developer Pack installieren:" -ForegroundColor Cyan
    Write-Host "  https://dotnet.microsoft.com/download/dotnet-framework/net45" -ForegroundColor White
    Write-Host ""
    Write-Host "Lösung 2: Build Tools mit .NET 4.5 neu installieren:" -ForegroundColor Cyan
    Write-Host "  Invoke-WebRequest -Uri 'https://aka.ms/vs/17/release/vs_buildtools.exe' -OutFile 'vs_buildtools.exe'" -ForegroundColor White
    Write-Host "  .\vs_buildtools.exe --add Microsoft.Net.Component.4.5.TargetingPack --quiet --wait" -ForegroundColor White
    Write-Host ""
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
    Write-Host "Warning: Executable not found at: $exePath" -ForegroundColor Yellow
    Write-Host "Check the build output above." -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
}

exit 0
