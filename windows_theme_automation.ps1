# ============================================
# Compatibility launcher for Windows Theme Automation V2
# ============================================

param(
    [switch]$AutoRun,
    [switch]$Install,
    [switch]$Uninstall,
    [switch]$Status,
    [switch]$Diagnose
)

function Find-ThemeAutoCli {
    $candidates = @(
        (Join-Path $env:LOCALAPPDATA "WindowsThemeAuto\themeauto.exe"),
        (Join-Path $PSScriptRoot "src\ThemeAutomation.Cli\bin\Release\net8.0-windows\themeauto.exe"),
        (Join-Path $PSScriptRoot "src\ThemeAutomation.Cli\bin\Debug\net8.0-windows\themeauto.exe")
    )

    foreach ($candidate in $candidates) {
        if (Test-Path $candidate) {
            return $candidate
        }
    }

    return $null
}

function Invoke-ThemeAutoCli {
    param([string]$Command)

    $cli = Find-ThemeAutoCli
    if ([string]::IsNullOrWhiteSpace($cli)) {
        Write-Host "themeauto.exe was not found." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Build or publish the CLI first:" -ForegroundColor Cyan
        Write-Host "  dotnet publish .\src\ThemeAutomation.Cli\ThemeAutomation.Cli.csproj -c Release -o `"%LOCALAPPDATA%\WindowsThemeAuto`""
        Write-Host ""
        Write-Host "After publishing, run this launcher again." -ForegroundColor Cyan
        exit 1
    }

    Write-Host "Running: $cli $Command" -ForegroundColor Gray
    & $cli $Command
    exit $LASTEXITCODE
}

if ($AutoRun) {
    Invoke-ThemeAutoCli -Command "apply"
}

if ($Install) {
    Invoke-ThemeAutoCli -Command "install"
}

if ($Uninstall) {
    Invoke-ThemeAutoCli -Command "uninstall"
}

if ($Status) {
    Invoke-ThemeAutoCli -Command "status"
}

if ($Diagnose) {
    Invoke-ThemeAutoCli -Command "diagnose"
}

Clear-Host
Write-Host ""
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "   WINDOWS THEME AUTOMATION V2" -ForegroundColor Cyan
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "1. Install automation" -ForegroundColor Green
Write-Host "2. Apply theme now" -ForegroundColor White
Write-Host "3. Show status" -ForegroundColor White
Write-Host "4. Diagnose Night Light" -ForegroundColor White
Write-Host "5. Uninstall automation" -ForegroundColor Red
Write-Host "6. Exit" -ForegroundColor Gray
Write-Host ""
Write-Host "Select an option (1-6): " -ForegroundColor Yellow -NoNewline
$choice = Read-Host

switch ($choice) {
    "1" { Invoke-ThemeAutoCli -Command "install" }
    "2" { Invoke-ThemeAutoCli -Command "apply" }
    "3" { Invoke-ThemeAutoCli -Command "status" }
    "4" { Invoke-ThemeAutoCli -Command "diagnose" }
    "5" { Invoke-ThemeAutoCli -Command "uninstall" }
    "6" {
        Write-Host "Exiting..." -ForegroundColor Gray
        exit 0
    }
    default {
        Write-Host "Invalid option" -ForegroundColor Red
        exit 1
    }
}
