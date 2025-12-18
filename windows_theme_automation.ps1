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
        [int]$Intensity,  # 0-100 (will be converted to color temperature)
        [bool]$Enable
    )
    
    try {
        # Registry paths for Night Light
        $BasePath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\CloudStore\Store\DefaultAccount\Current"
        $StatePath = "$BasePath\default`$windows.data.bluelightreduction.bluelightreductionstate\windows.data.bluelightreduction.bluelightreductionstate"
        $SettingsPath = "$BasePath\default`$windows.data.bluelightreduction.settings\windows.data.bluelightreduction.settings"
        
        # Convert intensity (0-100) to color temperature value
        # Range: 13 (0% = 6500K cool) to 68 (100% = 1200K warm)
        $tempValue = [Math]::Round(13 + (($Intensity / 100) * (68 - 13)))
        
        # ========== HANDLE NIGHT LIGHT STATE (ON/OFF) ==========
        $stateData = $null
        
        # Try to read existing state data
        if (Test-Path $StatePath) {
            try {
                $existingState = Get-ItemProperty -Path $StatePath -Name "Data" -ErrorAction SilentlyContinue
                if ($existingState -and $existingState.Data) {
                    $stateData = [byte[]]$existingState.Data
                }
            } catch { }
        }
        
        if ($stateData -and $stateData.Length -gt 18) {
            # Modify existing data - find and change the enable/disable byte
            # The state flag is typically after the "CB" marker (0x43, 0x42)
            for ($i = 0; $i -lt $stateData.Length - 2; $i++) {
                if ($stateData[$i] -eq 0x43 -and $stateData[$i+1] -eq 0x42) {
                    # Found CB marker, state byte is usually at offset +4 or +5
                    if ($i + 5 -lt $stateData.Length) {
                        if ($Enable) {
                            $stateData[$i + 4] = 0x02  # Enable flag
                            if ($i + 6 -lt $stateData.Length) {
                                $stateData[$i + 5] = 0x01
                            }
                        } else {
                            $stateData[$i + 4] = 0x00  # Disable flag
                            if ($i + 6 -lt $stateData.Length) {
                                $stateData[$i + 5] = 0x00
                            }
                        }
                    }
                    break
                }
            }
        } else {
            # No existing data - create minimal valid structure
            # This is a known-working structure for Windows 10/11
            if ($Enable) {
                $stateData = [byte[]](
                    0x02, 0x00, 0x00, 0x00,
                    0x56, 0x3A, 0xCC, 0x9A, 0xDE, 0xB8, 0xDA, 0x01,  # Timestamp
                    0x00, 0x00, 0x00, 0x00,
                    0x43, 0x42, 0x01, 0x00,
                    0x02, 0x01,  # Enabled state
                    0xCA, 0x14, 0x0E,
                    0x15,  # Active
                    0x00, 0x00, 0x00, 0x00
                )
            } else {
                $stateData = [byte[]](
                    0x02, 0x00, 0x00, 0x00,
                    0x56, 0x3A, 0xCC, 0x9A, 0xDE, 0xB8, 0xDA, 0x01,
                    0x00, 0x00, 0x00, 0x00,
                    0x43, 0x42, 0x01, 0x00,
                    0x00, 0x00,  # Disabled state
                    0xCA, 0x14, 0x0E,
                    0x00,
                    0x00, 0x00, 0x00, 0x00
                )
            }
            
            # Create registry path if needed
            $StateParent = Split-Path $StatePath -Parent
            if (-not (Test-Path $StateParent)) {
                New-Item -Path $StateParent -Force | Out-Null
            }
            if (-not (Test-Path $StatePath)) {
                New-Item -Path $StatePath -Force | Out-Null
            }
        }
        
        Set-ItemProperty -Path $StatePath -Name "Data" -Value $stateData -Type Binary -Force
        
        # ========== HANDLE INTENSITY SETTINGS ==========
        $settingsData = $null
        
        if (Test-Path $SettingsPath) {
            try {
                $existingSettings = Get-ItemProperty -Path $SettingsPath -Name "Data" -ErrorAction SilentlyContinue
                if ($existingSettings -and $existingSettings.Data) {
                    $settingsData = [byte[]]$existingSettings.Data
                }
            } catch { }
        }
        
        if ($settingsData -and $settingsData.Length -gt 15) {
            # Modify existing settings - find the temperature value after CA marker
            for ($i = 0; $i -lt $settingsData.Length - 3; $i++) {
                if ($settingsData[$i] -eq 0xCA -and ($settingsData[$i+1] -eq 0x0E -or $settingsData[$i+1] -eq 0x14)) {
                    # Found settings marker, temperature is next byte
                    if ($i + 2 -lt $settingsData.Length) {
                        $settingsData[$i + 2] = [byte]$tempValue
                    }
                    break
                }
            }
        } else {
            # Create new settings structure
            $settingsData = [byte[]](
                0x02, 0x00, 0x00, 0x00,
                0x56, 0x3A, 0xCC, 0x9A, 0xDE, 0xB8, 0xDA, 0x01,
                0x00, 0x00, 0x00, 0x00,
                0x43, 0x42, 0x01, 0x00,
                0xCA, 0x0E, [byte]$tempValue,
                0xCF, 0x28,
                0xCA, 0x32, 0x00, 0x00,
                0xCA, 0x2A, 0x00, 0x00
            )
            
            $SettingsParent = Split-Path $SettingsPath -Parent
            if (-not (Test-Path $SettingsParent)) {
                New-Item -Path $SettingsParent -Force | Out-Null
            }
            if (-not (Test-Path $SettingsPath)) {
                New-Item -Path $SettingsPath -Force | Out-Null
            }
        }
        
        Set-ItemProperty -Path $SettingsPath -Name "Data" -Value $settingsData -Type Binary -Force
        
        $statusText = if ($Enable) { "ENABLED" } else { "DISABLED" }
        Write-Host "Night Light: $statusText at $Intensity% intensity" -ForegroundColor Yellow
        
        # Broadcast settings change to refresh display
        Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class NightLightHelper {
    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint Msg, IntPtr wParam, string lParam, uint fuFlags, uint uTimeout, out IntPtr lpdwResult);
    
    public static void NotifySettingsChange() {
        IntPtr result;
        SendMessageTimeout((IntPtr)0xFFFF, 0x001A, IntPtr.Zero, "ImmersiveColorSet", 0x0002, 1000, out result);
    }
}
"@ -ErrorAction SilentlyContinue
        
        try {
            [NightLightHelper]::NotifySettingsChange()
        } catch { }
        
    } catch {
        Write-Host "Error configuring Night Light: $_" -ForegroundColor Red
        Write-Host "Try enabling Night Light manually once in Settings > Display > Night Light" -ForegroundColor Yellow
    }
}

function Apply-ThemeSettings {
    $CurrentHour = (Get-Date).Hour
    
    Write-Host "`n=== Applying theme settings ===" -ForegroundColor Magenta
    Write-Host "Current time: $(Get-Date -Format 'HH:mm')`n" -ForegroundColor White
    
    if ($CurrentHour -ge 7 -and $CurrentHour -lt 19) {
        Write-Host "DAY configuration (7am-7pm)" -ForegroundColor Green
        Set-WindowsTheme -Mode "Light"
        Set-NightLight -Intensity 20 -Enable $true
    }
    else {
        Write-Host "NIGHT configuration (7pm-7am)" -ForegroundColor Cyan
        Set-WindowsTheme -Mode "Dark"
        Set-NightLight -Intensity 50 -Enable $true
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