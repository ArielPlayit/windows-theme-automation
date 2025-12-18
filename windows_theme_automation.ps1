# ============================================
# Auto-Installation Script for Windows Theme Automation
# Day/Night Mode and Night Light Automation
# ============================================

# Check if running as Administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "This script needs Administrator privileges" -ForegroundColor Yellow
    Write-Host "Relaunching with Administrator rights..." -ForegroundColor Cyan
    Start-Sleep -Seconds 2
    
    $scriptPath = $MyInvocation.MyCommand.Path
    Start-Process powershell.exe -Verb RunAs -ArgumentList "-ExecutionPolicy Bypass -File `"$scriptPath`""
    exit
}

# ============================================
# MAIN FUNCTIONS
# ============================================

function Set-WindowsTheme {
    param([string]$Mode)
    
    $ThemePath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"
    
    if ($Mode -eq "Light") {
        Set-ItemProperty -Path $ThemePath -Name "AppsUseLightTheme" -Value 1
        Set-ItemProperty -Path $ThemePath -Name "SystemUsesLightTheme" -Value 1
        Write-Host "Light mode activated" -ForegroundColor Green
    }
    elseif ($Mode -eq "Dark") {
        Set-ItemProperty -Path $ThemePath -Name "AppsUseLightTheme" -Value 0
        Set-ItemProperty -Path $ThemePath -Name "SystemUsesLightTheme" -Value 0
        Write-Host "Dark mode activated" -ForegroundColor Cyan
    }
    
    # Restart Windows Explorer to apply theme changes immediately
    Write-Host "Restarting Explorer to apply changes..." -ForegroundColor Yellow
    Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 500
    Start-Process explorer.exe
}

function Set-NightLight {
    param(
        [int]$Percentage,  # 0-100 (0 = neutral/off, 100 = warmest)
        [bool]$Enable
    )
    
    try {
        # If not enabled, reset to neutral
        if (-not $Enable) {
            $Percentage = 0
        }
        
        # Clamp percentage
        if ($Percentage -lt 0) { $Percentage = 0 }
        if ($Percentage -gt 100) { $Percentage = 100 }
        
        # Add required Win32 API for gamma control
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public class GammaController {
    [DllImport("user32.dll")]
    public static extern IntPtr GetDC(IntPtr hWnd);
    
    [DllImport("user32.dll")]
    public static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);
    
    [DllImport("gdi32.dll")]
    public static extern bool SetDeviceGammaRamp(IntPtr hDC, ref RAMP lpRamp);
    
    [DllImport("gdi32.dll")]
    public static extern bool GetDeviceGammaRamp(IntPtr hDC, ref RAMP lpRamp);
    
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
    public struct RAMP {
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 256)]
        public ushort[] Red;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 256)]
        public ushort[] Green;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 256)]
        public ushort[] Blue;
    }
    
    public static void SetGamma(double redMult, double greenMult, double blueMult) {
        RAMP ramp = new RAMP();
        ramp.Red = new ushort[256];
        ramp.Green = new ushort[256];
        ramp.Blue = new ushort[256];
        
        for (int i = 0; i < 256; i++) {
            int redVal = (int)(i * 255 * redMult);
            int greenVal = (int)(i * 255 * greenMult);
            int blueVal = (int)(i * 255 * blueMult);
            
            ramp.Red[i] = (ushort)Math.Min(65535, Math.Max(0, redVal));
            ramp.Green[i] = (ushort)Math.Min(65535, Math.Max(0, greenVal));
            ramp.Blue[i] = (ushort)Math.Min(65535, Math.Max(0, blueVal));
        }
        
        IntPtr hDC = GetDC(IntPtr.Zero);
        SetDeviceGammaRamp(hDC, ref ramp);
        ReleaseDC(IntPtr.Zero, hDC);
    }
}
'@ -ErrorAction SilentlyContinue
        
        # Convert percentage to color temperature (Kelvin)
        # 0% = 6500K (neutral daylight), 100% = 2700K (warm candlelight)
        $minTemp = 2700   # Warmest (100%)
        $maxTemp = 6500   # Neutral (0%)
        $temperature = $maxTemp - (($Percentage / 100) * ($maxTemp - $minTemp))
        
        # Calculate RGB multipliers based on color temperature
        # Algorithm based on Tanner Helland's work
        $temp = $temperature / 100
        
        # Calculate Red
        if ($temperature -le 6600) {
            $red = 1.0
        } else {
            $red = [Math]::Pow(($temp - 60), -0.1332047592) * 329.698727446 / 255
            $red = [Math]::Max(0, [Math]::Min(1, $red))
        }
        
        # Calculate Green
        if ($temperature -le 6600) {
            $green = [Math]::Log($temp) * 99.4708025861 - 161.1195681661
            $green = $green / 255
        } else {
            $green = [Math]::Pow(($temp - 60), -0.0755148492) * 288.1221695283 / 255
        }
        $green = [Math]::Max(0, [Math]::Min(1, $green))
        
        # Calculate Blue
        if ($temperature -ge 6600) {
            $blue = 1.0
        } elseif ($temperature -le 1900) {
            $blue = 0.0
        } else {
            $blue = [Math]::Log($temp - 10) * 138.5177312231 - 305.0447927307
            $blue = $blue / 255
            $blue = [Math]::Max(0, [Math]::Min(1, $blue))
        }
        
        # Apply gamma
        [GammaController]::SetGamma($red, $green, $blue)
        
        $statusText = if ($Enable -and $Percentage -gt 0) { "ENABLED at ${Percentage}% warmth (~${temperature}K)" } else { "DISABLED (neutral colors)" }
        Write-Host "Night Light: $statusText" -ForegroundColor Yellow
        
    } catch {
        Write-Host "Error setting Night Light: $_" -ForegroundColor Red
    }
}

function Apply-ThemeSettings {
    $CurrentHour = (Get-Date).Hour
    
    Write-Host "`n=== Applying theme settings ===" -ForegroundColor Magenta
    Write-Host "Current time: $(Get-Date -Format 'HH:mm')`n" -ForegroundColor White
    
    if ($CurrentHour -ge 7 -and $CurrentHour -lt 19) {
        Write-Host "DAY configuration (7am-7pm)" -ForegroundColor Green
        Set-WindowsTheme -Mode "Light"
        Set-NightLight -Percentage 20 -Enable $true
    }
    else {
        Write-Host "NIGHT configuration (7pm-7am)" -ForegroundColor Cyan
        Set-WindowsTheme -Mode "Dark"
        Set-NightLight -Percentage 50 -Enable $true
    }
    
    Write-Host "`nSuccessfully applied!`n" -ForegroundColor Magenta
}

function Install-AutoScheduler {
    param([string]$OriginalScriptPath)
    
    if ([string]::IsNullOrEmpty($OriginalScriptPath)) {
        $OriginalScriptPath = $PSCommandPath
    }
    
    if ([string]::IsNullOrEmpty($OriginalScriptPath)) {
        Write-Host "Error: Cannot detect script path" -ForegroundColor Red
        Write-Host "Please run the script by right-clicking and selecting 'Run with PowerShell'" -ForegroundColor Yellow
        Write-Host "`nPress any key to exit..."
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit
    }
    
    $scriptPath = $OriginalScriptPath
    $scriptName = Split-Path -Leaf $scriptPath
    
    $installPath = "$env:LOCALAPPDATA\WindowsThemeAuto"
    $finalScriptPath = "$installPath\$scriptName"
    
    if (-not (Test-Path $installPath)) {
        New-Item -ItemType Directory -Path $installPath -Force | Out-Null
    }
    
    if ($scriptPath -ne $finalScriptPath) {
        Copy-Item -Path $scriptPath -Destination $finalScriptPath -Force
        Write-Host "Script copied to: $finalScriptPath" -ForegroundColor Green
    }
    
    $taskNameHourly = "ThemeAutoSwitch_Hourly"
    $taskNameStartup = "ThemeAutoSwitch_Startup"
    
    Unregister-ScheduledTask -TaskName $taskNameHourly -Confirm:$false -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName $taskNameStartup -Confirm:$false -ErrorAction SilentlyContinue
    
    $actionHourly = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$finalScriptPath`" -AutoRun"
    $triggerHourly = New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Hours 1)
    $principalHourly = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Highest
    $settingsHourly = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
    
    Register-ScheduledTask -TaskName $taskNameHourly -Action $actionHourly -Trigger $triggerHourly -Principal $principalHourly -Settings $settingsHourly -Description "Automatically switches Windows theme every hour based on day/night schedule" | Out-Null
    
    # Set the task to repeat indefinitely
    $task = Get-ScheduledTask -TaskName $taskNameHourly
    $task.Triggers[0].Repetition.Duration = ""
    $task | Set-ScheduledTask | Out-Null
    
    $actionStartup = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$finalScriptPath`" -AutoRun"
    $triggerStartup = New-ScheduledTaskTrigger -AtLogOn -User "$env:USERNAME"
    $principalStartup = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Highest
    $settingsStartup = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
    
    Register-ScheduledTask -TaskName $taskNameStartup -Action $actionStartup -Trigger $triggerStartup -Principal $principalStartup -Settings $settingsStartup -Description "Applies correct theme when Windows starts" | Out-Null
    
    Write-Host "`nScheduled tasks created:" -ForegroundColor Green
    Write-Host "  - Runs automatically every hour" -ForegroundColor White
    Write-Host "  - Runs at Windows login" -ForegroundColor White
    Write-Host "`nINSTALLATION COMPLETED!" -ForegroundColor Magenta
    Write-Host "`nAutomatic switching is now active." -ForegroundColor Cyan
    Write-Host "You don't need to run this script again.`n" -ForegroundColor Yellow
}

function Uninstall-AutoScheduler {
    Write-Host "`n=== UNINSTALLATION ===" -ForegroundColor Red
    
    Unregister-ScheduledTask -TaskName "ThemeAutoSwitch_Hourly" -Confirm:$false -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName "ThemeAutoSwitch_Startup" -Confirm:$false -ErrorAction SilentlyContinue
    
    $installPath = "$env:LOCALAPPDATA\WindowsThemeAuto"
    if (Test-Path $installPath) {
        Remove-Item -Path $installPath -Recurse -Force -ErrorAction SilentlyContinue
    }
    
    Write-Host "Automation completely uninstalled" -ForegroundColor Green
    Write-Host "`nPress any key to exit..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}

# ============================================
# MAIN MENU
# ============================================

param([switch]$AutoRun)

if ($AutoRun) {
    Apply-ThemeSettings
    exit
}

Clear-Host
Write-Host "`n========================================================" -ForegroundColor Cyan
Write-Host "   WINDOWS DAY/NIGHT THEME AUTOMATION" -ForegroundColor Cyan
Write-Host "========================================================" -ForegroundColor Cyan

Write-Host "`nWhat would you like to do?`n" -ForegroundColor Yellow
Write-Host "1. Install automation (recommended)" -ForegroundColor Green
Write-Host "2. Apply theme now (without installing)" -ForegroundColor White
Write-Host "3. Uninstall automation" -ForegroundColor Red
Write-Host "4. Exit" -ForegroundColor Gray

Write-Host "`nSelect an option (1-4): " -ForegroundColor Yellow -NoNewline
$choice = Read-Host

switch ($choice) {
    "1" {
        Write-Host "`n=== INSTALLING AUTOMATION ===" -ForegroundColor Cyan
        Install-AutoScheduler -OriginalScriptPath $PSCommandPath
        Apply-ThemeSettings
        Write-Host "`nPress any key to exit..."
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
    "2" {
        Apply-ThemeSettings
        Write-Host "Press any key to exit..."
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
    "3" {
        Uninstall-AutoScheduler
    }
    "4" {
        Write-Host "`nExiting..." -ForegroundColor Gray
        exit
    }
    default {
        Write-Host "`nInvalid option" -ForegroundColor Red
        Start-Sleep -Seconds 2
    }
}