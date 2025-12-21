# ============================================
# Auto-Installation Script for Windows Theme Automation
# Day/Night Mode and Blue Light Filter
# ============================================

# ============================================
# PARAMETERS - MUST BE FIRST
# ============================================
param([switch]$AutoRun)

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

function Set-BlueLight {
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
        
        # Define gamma control Win32 API
        $gammaCode = @'
using System;
using System.Runtime.InteropServices;

public class ScreenGamma {
    [DllImport("user32.dll")]
    private static extern IntPtr GetDC(IntPtr hWnd);
    
    [DllImport("user32.dll")]
    private static extern bool ReleaseDC(IntPtr hWnd, IntPtr hDC);
    
    [DllImport("gdi32.dll")]
    private static extern bool SetDeviceGammaRamp(IntPtr hDC, ref GammaRampStruct lpRamp);
    
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
    public struct GammaRampStruct {
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 256)]
        public ushort[] Red;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 256)]
        public ushort[] Green;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 256)]
        public ushort[] Blue;
    }
    
    public static bool SetGamma(int brightness, int temperature) {
        IntPtr hDC = GetDC(IntPtr.Zero);
        if (hDC == IntPtr.Zero) return false;
        
        GammaRampStruct ramp = new GammaRampStruct();
        ramp.Red = new ushort[256];
        ramp.Green = new ushort[256];
        ramp.Blue = new ushort[256];
        
        // Blue light reduction
        double blueMult = 1.0 - (temperature / 100.0 * 0.7);  // Reduce up to 70%
        double greenMult = 1.0 - (temperature / 100.0 * 0.3); // Reduce up to 30%
        double redMult = 1.0;  // Keep red at 100%
        
        double brightFactor = brightness / 100.0;
        
        for (int i = 0; i < 256; i++) {
            int value = (i << 8) | i;
            
            int red = (int)(value * redMult * brightFactor);
            int green = (int)(value * greenMult * brightFactor);
            int blue = (int)(value * blueMult * brightFactor);
            
            ramp.Red[i] = (ushort)Math.Min(65535, Math.Max(0, red));
            ramp.Green[i] = (ushort)Math.Min(65535, Math.Max(0, green));
            ramp.Blue[i] = (ushort)Math.Min(65535, Math.Max(0, blue));
        }
        
        bool result = SetDeviceGammaRamp(hDC, ref ramp);
        ReleaseDC(IntPtr.Zero, hDC);
        return result;
    }
    
    public static bool ResetGamma() {
        IntPtr hDC = GetDC(IntPtr.Zero);
        if (hDC == IntPtr.Zero) return false;
        
        GammaRampStruct ramp = new GammaRampStruct();
        ramp.Red = new ushort[256];
        ramp.Green = new ushort[256];
        ramp.Blue = new ushort[256];
        
        for (int i = 0; i < 256; i++) {
            int value = (i << 8) | i;
            ramp.Red[i] = (ushort)value;
            ramp.Green[i] = (ushort)value;
            ramp.Blue[i] = (ushort)value;
        }
        
        bool result = SetDeviceGammaRamp(hDC, ref ramp);
        ReleaseDC(IntPtr.Zero, hDC);
        return result;
    }
}
'@
        
        # Load the type if not already loaded
        if (-not ([System.Management.Automation.PSTypeName]'ScreenGamma').Type) {
            Add-Type -TypeDefinition $gammaCode -ErrorAction Stop
        }
        
        # Apply gamma settings
        if ($Enable -and $Percentage -gt 0) {
            $brightness = 100
            $success = [ScreenGamma]::SetGamma($brightness, $Percentage)
            
            if ($success) {
                $minTemp = 2700
                $maxTemp = 6500
                $temperature = [int]($maxTemp - (($Percentage / 100.0) * ($maxTemp - $minTemp)))
                
                Write-Host "Blue Light Filter: ENABLED at ${Percentage}% (~${temperature}K)" -ForegroundColor Yellow
            } else {
                Write-Host "Warning: Could not apply blue light filter (driver limitation)" -ForegroundColor Yellow
            }
        } else {
            $success = [ScreenGamma]::ResetGamma()
            if ($success) {
                Write-Host "Blue Light Filter: DISABLED (neutral colors)" -ForegroundColor Yellow
            }
        }
        
    } catch {
        Write-Host "Note: Blue light filter could not be applied" -ForegroundColor Yellow
    }
}

function Apply-ThemeSettings {
    $CurrentHour = (Get-Date).Hour
    
    Write-Host "`n=== Applying theme settings ===" -ForegroundColor Magenta
    Write-Host "Current time: $(Get-Date -Format 'HH:mm')`n" -ForegroundColor White
    
    if ($CurrentHour -ge 7 -and $CurrentHour -lt 19) {
        Write-Host "DAY configuration (7am-7pm)" -ForegroundColor Green
        Set-WindowsTheme -Mode "Light"
        Set-BlueLight -Percentage 20 -Enable $true
    }
    else {
        Write-Host "NIGHT configuration (7pm-7am)" -ForegroundColor Cyan
        Set-WindowsTheme -Mode "Dark"
        Set-BlueLight -Percentage 50 -Enable $true
    }
    
    Write-Host "`nSuccessfully applied!`n" -ForegroundColor Magenta
}

function Install-AutoScheduler {
    param([string]$OriginalScriptPath)
    
    # Check if running as Administrator
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    
    if (-not $isAdmin) {
        Write-Host "Installation requires Administrator privileges" -ForegroundColor Yellow
        Write-Host "Relaunching with Administrator rights..." -ForegroundColor Cyan
        Start-Sleep -Seconds 2
        
        $scriptPath = $MyInvocation.MyCommand.Path
        if ([string]::IsNullOrEmpty($scriptPath)) {
            $scriptPath = $OriginalScriptPath
        }
        
        $args = "-ExecutionPolicy Bypass -File `"$scriptPath`""
        Start-Process powershell.exe -Verb RunAs -ArgumentList $args
        exit
    }
    
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
    
    # Remove existing tasks
    $taskNameDaily = "ThemeAutoSwitch_Daily"
    $taskName7AM = "ThemeAutoSwitch_7AM"
    $taskName7PM = "ThemeAutoSwitch_7PM"
    $taskNameStartup = "ThemeAutoSwitch_Startup"
    
    # Remove old hourly task if exists
    Unregister-ScheduledTask -TaskName "ThemeAutoSwitch_Hourly" -Confirm:$false -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName $taskNameDaily -Confirm:$false -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName $taskName7AM -Confirm:$false -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName $taskName7PM -Confirm:$false -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName $taskNameStartup -Confirm:$false -ErrorAction SilentlyContinue
    
    Write-Host "`nCreating scheduled tasks..." -ForegroundColor Cyan
    
    # Task 1: Run at 7 AM daily (switch to day mode)
    $action7AM = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$finalScriptPath`" -AutoRun"
    $trigger7AM = New-ScheduledTaskTrigger -Daily -At "07:00"
    $principal7AM = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive
    $settings7AM = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 5)
    
    Register-ScheduledTask -TaskName $taskName7AM -Action $action7AM -Trigger $trigger7AM -Principal $principal7AM -Settings $settings7AM -Description "Switches to Day Mode at 7 AM" -Force | Out-Null
    Write-Host "  Created: 7 AM daily task" -ForegroundColor Green
    
    # Task 2: Run at 7 PM daily (switch to night mode)
    $action7PM = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$finalScriptPath`" -AutoRun"
    $trigger7PM = New-ScheduledTaskTrigger -Daily -At "19:00"
    $principal7PM = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive
    $settings7PM = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 5)
    
    Register-ScheduledTask -TaskName $taskName7PM -Action $action7PM -Trigger $trigger7PM -Principal $principal7PM -Settings $settings7PM -Description "Switches to Night Mode at 7 PM" -Force | Out-Null
    Write-Host "  Created: 7 PM daily task" -ForegroundColor Green
    
    # Task 3: Run at startup/login
    $actionStartup = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$finalScriptPath`" -AutoRun"
    $triggerStartup = New-ScheduledTaskTrigger -AtLogOn -User "$env:USERNAME"
    $principalStartup = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive
    $settingsStartup = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 5)
    
    Register-ScheduledTask -TaskName $taskNameStartup -Action $actionStartup -Trigger $triggerStartup -Principal $principalStartup -Settings $settingsStartup -Description "Applies correct theme at Windows login" -Force | Out-Null
    Write-Host "  Created: Login task" -ForegroundColor Green
    
    Write-Host "`nScheduled tasks created:" -ForegroundColor Green
    Write-Host "  - Runs at 7:00 AM daily (Day Mode)" -ForegroundColor White
    Write-Host "  - Runs at 7:00 PM daily (Night Mode)" -ForegroundColor White
    Write-Host "  - Runs at Windows login" -ForegroundColor White
    Write-Host "`nINSTALLATION COMPLETED!" -ForegroundColor Magenta
    Write-Host "`nAutomatic switching is now active." -ForegroundColor Cyan
    Write-Host "You don't need to run this script again.`n" -ForegroundColor Yellow
}

function Uninstall-AutoScheduler {
    Write-Host "`n=== UNINSTALLATION ===" -ForegroundColor Red
    
    # Check if running as Administrator
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    
    if (-not $isAdmin) {
        Write-Host "Uninstallation requires Administrator privileges" -ForegroundColor Yellow
        Write-Host "Relaunching with Administrator rights..." -ForegroundColor Cyan
        Start-Sleep -Seconds 2
        
        $scriptPath = $MyInvocation.MyCommand.Path
        Start-Process powershell.exe -Verb RunAs -ArgumentList "-ExecutionPolicy Bypass -Command `"& '$scriptPath'`""
        exit
    }
    
    # Remove all possible task variants
    Unregister-ScheduledTask -TaskName "ThemeAutoSwitch_Hourly" -Confirm:$false -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName "ThemeAutoSwitch_Daily" -Confirm:$false -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName "ThemeAutoSwitch_7AM" -Confirm:$false -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName "ThemeAutoSwitch_7PM" -Confirm:$false -ErrorAction SilentlyContinue
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