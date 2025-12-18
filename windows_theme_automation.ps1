# ============================================
# Auto-Installation Script for Windows Theme Automation
# Day/Night Mode and Night Light Automation
# ============================================

# NOTE: elevation is performed only for install/uninstall operations.
# Allow `-AutoRun` (startup/triggered runs) to execute without forcing Administrator,
# so theme changes applied to the interactive user's HKCU take effect correctly.

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
    
    [DllImport("gdi32.dll")]
    private static extern bool GetDeviceGammaRamp(IntPtr hDC, ref GammaRampStruct lpRamp);
    
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
        
        // Calculate color temperature multipliers
        // temperature: 0 = neutral (6500K), 100 = warmest (2700K)
        double tempFactor = 1.0 - (temperature / 100.0);
        
        // Blue light reduction (most important for filtering blue light)
        double blueMult = 1.0 - (temperature / 100.0 * 0.7);  // Reduce up to 70%
        double greenMult = 1.0 - (temperature / 100.0 * 0.3); // Reduce up to 30%
        double redMult = 1.0;  // Keep red channel at 100%
        
        // Brightness adjustment
        double brightFactor = brightness / 100.0;
        
        for (int i = 0; i < 256; i++) {
            // Linear gamma ramp from 0 to 65535
            int value = (i << 8) | i;  // Equivalent to i * 257
            
            // Apply color temperature and brightness
            int red = (int)(value * redMult * brightFactor);
            int green = (int)(value * greenMult * brightFactor);
            int blue = (int)(value * blueMult * brightFactor);
            
            // Clamp values to valid range
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
            $brightness = 100  # Keep brightness at 100%
            $success = [ScreenGamma]::SetGamma($brightness, $Percentage)
            
            if ($success) {
                # Calculate approximate color temperature for display
                $minTemp = 2700   # Warmest (100%)
                $maxTemp = 6500   # Neutral (0%)
                $temperature = [int]($maxTemp - (($Percentage / 100.0) * ($maxTemp - $minTemp)))
                
                Write-Host "Blue Light Filter: ENABLED at ${Percentage}% (~${temperature}K)" -ForegroundColor Yellow
                Write-Host "  Red: 100% | Green: $([int](100 - $Percentage * 0.3))% | Blue: $([int](100 - $Percentage * 0.7))%" -ForegroundColor Cyan
            } else {
                Write-Host "Failed to apply blue light filter. Your display driver may not support gamma adjustments." -ForegroundColor Red
                Write-Host "Try updating your graphics drivers or using Windows Night Light instead." -ForegroundColor Yellow
            }
        } else {
            # Reset to neutral
            $success = [ScreenGamma]::ResetGamma()
            
            if ($success) {
                Write-Host "Blue Light Filter: DISABLED (neutral colors restored)" -ForegroundColor Yellow
            } else {
                Write-Host "Failed to reset gamma. Try restarting your computer to reset display settings." -ForegroundColor Red
            }
        }
        
    } catch {
        Write-Host "Error configuring blue light filter: $_" -ForegroundColor Red
        Write-Host "Details: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "`nTroubleshooting:" -ForegroundColor Yellow
        Write-Host "  1. Make sure you're running this on the main display" -ForegroundColor White
        Write-Host "  2. Update your graphics drivers" -ForegroundColor White
        Write-Host "  3. Some displays/drivers don't support gamma adjustments" -ForegroundColor White
        Write-Host "  4. Try using Windows built-in Night Light instead" -ForegroundColor White
    }
}

function Repair-NightLightRegistry {
    Write-Host "`n=== REPAIRING NIGHT LIGHT REGISTRY ===" -ForegroundColor Cyan
    Write-Host "This will completely rebuild the Night Light registry structure" -ForegroundColor Yellow
    Write-Host "Including the Cache entries that Windows requires" -ForegroundColor Yellow
    
    try {
        # Base paths
        $baseCurrent = "HKCU:\Software\Microsoft\Windows\CurrentVersion\CloudStore\Store\DefaultAccount\Current"
        $baseCache = "HKCU:\Software\Microsoft\Windows\CurrentVersion\CloudStore\Store\Cache\DefaultAccount"
        
        $stateKey = "default`$windows.data.bluelightreduction.bluelightreductionstate"
        $settingsKey = "default`$windows.data.bluelightreduction.settings"
        
        Write-Host "`nStep 1: Removing ALL Night Light registry entries..." -ForegroundColor Yellow
        
        # Remove from Current
        $currentStatePath = "$baseCurrent\$stateKey"
        $currentSettingsPath = "$baseCurrent\$settingsKey"
        
        if (Test-Path $currentStatePath) {
            Remove-Item -Path $currentStatePath -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "  Removed Current\State entries" -ForegroundColor Green
        }
        
        if (Test-Path $currentSettingsPath) {
            Remove-Item -Path $currentSettingsPath -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "  Removed Current\Settings entries" -ForegroundColor Green
        }
        
        # Remove from Cache
        $cacheStatePath = "$baseCache\$stateKey"
        $cacheSettingsPath = "$baseCache\$settingsKey"
        
        if (Test-Path $cacheStatePath) {
            Remove-Item -Path $cacheStatePath -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "  Removed Cache\State entries" -ForegroundColor Green
        }
        
        if (Test-Path $cacheSettingsPath) {
            Remove-Item -Path $cacheSettingsPath -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "  Removed Cache\Settings entries" -ForegroundColor Green
        }
        
        Start-Sleep -Milliseconds 500
        
        Write-Host "`nStep 2: Recreating registry structure..." -ForegroundColor Yellow
        
        # Create Current paths
        $currentStateFullPath = "$currentStatePath\windows.data.bluelightreduction.bluelightreductionstate"
        $currentSettingsFullPath = "$currentSettingsPath\windows.data.bluelightreduction.settings"
        
        New-Item -Path $currentStateFullPath -Force | Out-Null
        Write-Host "  Created Current\State key" -ForegroundColor Green
        
        New-Item -Path $currentSettingsFullPath -Force | Out-Null
        Write-Host "  Created Current\Settings key" -ForegroundColor Green
        
        # Create Cache paths (CRITICAL - this was missing!)
        $cacheStateFullPath = "$cacheStatePath\windows.data.bluelightreduction.bluelightreductionstate"
        $cacheSettingsFullPath = "$cacheSettingsPath\windows.data.bluelightreduction.settings"
        
        New-Item -Path $cacheStateFullPath -Force | Out-Null
        Write-Host "  Created Cache\State key" -ForegroundColor Green
        
        New-Item -Path $cacheSettingsFullPath -Force | Out-Null
        Write-Host "  Created Cache\Settings key" -ForegroundColor Green
        
        Write-Host "`nStep 3: Setting default Night Light values..." -ForegroundColor Yellow
        
        # Night Light OFF state data (valid format for Windows 11)
        $stateDataOff = [byte[]](
            0x02, 0x00, 0x00, 0x00,  # Version
            0x54, 0xBC, 0xCA, 0xE4, 0xC7, 0x3D, 0xDB, 0x01,  # Timestamp
            0x00, 0x00, 0x00, 0x00,  # Reserved
            0x43, 0x42, 0x01, 0x00,  # Header
            0xCA, 0x14, 0x0E, 0x10,  # State flags
            0xCA, 0x14, 0x00, 0x00,  # More flags
            0x00                     # Enabled = 0 (OFF)
        )
        
        # Night Light settings data (temperature 4000K)
        $settingsData = [byte[]](
            0x02, 0x00, 0x00, 0x00,  # Version
            0x54, 0xBC, 0xCA, 0xE4, 0xC7, 0x3D, 0xDB, 0x01,  # Timestamp
            0x00, 0x00, 0x00, 0x00,  # Reserved
            0x43, 0x42, 0x01, 0x00,  # Header
            0xCA, 0x1E,              # Settings type
            0x15, 0x00, 0x00, 0x00,  # More settings
            0xCF, 0x28,              # Schedule info
            0xA0, 0x0F, 0x00, 0x00,  # Temperature: 4000 (0x0FA0 in little-endian)
            0xCA, 0x32, 0x00, 0x00,  # More config
            0xCA, 0x3C, 0x00, 0x00,  # Sunset/sunrise
            0x00                     # End
        )
        
        # Apply to Current store
        Set-ItemProperty -Path $currentStateFullPath -Name "Data" -Value $stateDataOff -Type Binary -Force
        Set-ItemProperty -Path $currentSettingsFullPath -Name "Data" -Value $settingsData -Type Binary -Force
        Write-Host "  Set Current store values" -ForegroundColor Green
        
        # Apply to Cache store (CRITICAL!)
        Set-ItemProperty -Path $cacheStateFullPath -Name "Data" -Value $stateDataOff -Type Binary -Force
        Set-ItemProperty -Path $cacheSettingsFullPath -Name "Data" -Value $settingsData -Type Binary -Force
        Write-Host "  Set Cache store values" -ForegroundColor Green
        
        Write-Host "`nStep 4: Restarting Windows services..." -ForegroundColor Yellow
        
        # Stop Windows Explorer
        Write-Host "  Stopping Explorer..." -ForegroundColor Gray
        Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 1
        
        # Start Explorer
        Write-Host "  Starting Explorer..." -ForegroundColor Gray
        Start-Process explorer.exe
        Start-Sleep -Seconds 2
        
        Write-Host "`n=== REPAIR COMPLETE ===" -ForegroundColor Green
        Write-Host "`nNight Light registry has been fully rebuilt!" -ForegroundColor Green
        Write-Host "Both Current and Cache stores have been restored." -ForegroundColor Cyan
        Write-Host "`nIMPORTANT: Please do one of the following:" -ForegroundColor Yellow
        Write-Host "  Option A: Sign out and sign back in" -ForegroundColor White
        Write-Host "  Option B: Restart your computer" -ForegroundColor White
        Write-Host "`nThen check: Settings > System > Display > Night light" -ForegroundColor Yellow
        
    } catch {
        Write-Host "`nError: $_" -ForegroundColor Red
        Write-Host "Details: $($_.Exception.Message)" -ForegroundColor Red
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
    # Require Administrator for installation (registering scheduled tasks and copying files)
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Write-Host "Installer requires Administrator privileges" -ForegroundColor Yellow
        Write-Host "Relaunching installer with Administrator rights..." -ForegroundColor Cyan
        Start-Sleep -Seconds 1
        $scriptPath = $MyInvocation.MyCommand.Path
        Start-Process powershell.exe -Verb RunAs -ArgumentList "-ExecutionPolicy Bypass -File `"$scriptPath`" -OriginalScriptPath `"$OriginalScriptPath`""
        return
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
    # Require Administrator to remove scheduled tasks and files
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Write-Host "Uninstall requires Administrator privileges" -ForegroundColor Yellow
        Write-Host "Relaunching uninstaller with Administrator rights..." -ForegroundColor Cyan
        Start-Sleep -Seconds 1
        $scriptPath = $MyInvocation.MyCommand.Path
        Start-Process powershell.exe -Verb RunAs -ArgumentList "-ExecutionPolicy Bypass -File `"$scriptPath`" -Confirm:$false" -Wait
        return
    }
    
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
Write-Host "3. Repair Night Light registry (fix corrupted settings)" -ForegroundColor Magenta
Write-Host "4. Uninstall automation" -ForegroundColor Red
Write-Host "5. Exit" -ForegroundColor Gray

Write-Host "`nSelect an option (1-5): " -ForegroundColor Yellow -NoNewline
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
        Repair-NightLightRegistry
        Write-Host "`nPress any key to exit..."
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
    "4" {
        Uninstall-AutoScheduler
    }
    "5" {
        Write-Host "`nExiting..." -ForegroundColor Gray
        exit
    }
    default {
        Write-Host "`nInvalid option" -ForegroundColor Red
        Start-Sleep -Seconds 2
    }
}