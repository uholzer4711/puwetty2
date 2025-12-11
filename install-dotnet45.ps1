# Install .NET Framework 4.5 Developer Pack
# Required for building SuperPUWEtty2

Write-Host "================================================" -ForegroundColor Cyan
Write-Host ".NET Framework 4.5 Developer Pack Installer" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "This will install .NET Framework 4.5 Developer Pack" -ForegroundColor Yellow
Write-Host "Required for building SuperPUWEtty2" -ForegroundColor Yellow
Write-Host ""

$response = Read-Host "Continue? (Y/n)"
if ($response -ne "" -and $response -ne "Y" -and $response -ne "y") {
    Write-Host "Aborted." -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Checking if .NET Framework 4.5 is already installed..." -ForegroundColor Yellow

# Check if .NET 4.5 is installed
$dotnet45 = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -ErrorAction SilentlyContinue
if ($dotnet45 -and $dotnet45.Release -ge 378389) {
    Write-Host ".NET Framework 4.5+ is already installed (Release: $($dotnet45.Release))" -ForegroundColor Green
    Write-Host ""
    Write-Host "But the Developer Pack (Targeting Pack) might still be missing..." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Installing via Visual Studio Build Tools..." -ForegroundColor Cyan
Write-Host "(This will take a few minutes)" -ForegroundColor Yellow
Write-Host ""

# Download VS Build Tools if not exists
$installerPath = "vs_buildtools.exe"
if (-not (Test-Path $installerPath)) {
    Write-Host "Downloading Visual Studio Build Tools installer..." -ForegroundColor Yellow
    $url = "https://aka.ms/vs/17/release/vs_buildtools.exe"
    try {
        Invoke-WebRequest -Uri $url -OutFile $installerPath
        Write-Host "Download complete." -ForegroundColor Green
    } catch {
        Write-Host "ERROR: Could not download installer!" -ForegroundColor Red
        Write-Host "Please download manually from: https://visualstudio.microsoft.com/downloads/" -ForegroundColor Yellow
        Read-Host "Press Enter to exit"
        exit 1
    }
}

Write-Host ""
Write-Host "Installing .NET Framework 4.5 Targeting Pack..." -ForegroundColor Cyan
Write-Host "This may take 5-10 minutes..." -ForegroundColor Yellow
Write-Host ""

# Install .NET 4.5 Targeting Pack
$arguments = @(
    "--add", "Microsoft.Net.Component.4.5.TargetingPack",
    "--add", "Microsoft.Net.Component.4.5.2.TargetingPack",
    "--add", "Microsoft.VisualStudio.Workload.MSBuildTools",
    "--quiet",
    "--wait",
    "--norestart"
)

$process = Start-Process -FilePath $installerPath -ArgumentList $arguments -Wait -PassThru

if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010) {
    Write-Host ""
    Write-Host "================================================" -ForegroundColor Green
    Write-Host "Installation successful!" -ForegroundColor Green
    Write-Host "================================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "You can now build SuperPUWEtty2:" -ForegroundColor Cyan
    Write-Host "  .\build-without-installer.ps1" -ForegroundColor White
    Write-Host ""

    if ($process.ExitCode -eq 3010) {
        Write-Host "NOTE: A system restart is recommended." -ForegroundColor Yellow
    }
} else {
    Write-Host ""
    Write-Host "Installation failed with exit code: $($process.ExitCode)" -ForegroundColor Red
    Write-Host ""
    Write-Host "Try manual installation:" -ForegroundColor Yellow
    Write-Host "1. Download from: https://dotnet.microsoft.com/download/dotnet-framework/net45" -ForegroundColor White
    Write-Host "2. Or use Visual Studio Installer to add .NET 4.5 workload" -ForegroundColor White
}

Read-Host "`nPress Enter to exit"
exit 0
