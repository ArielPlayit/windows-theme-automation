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
        [int]$Intensity,
        [bool]$Enable
    )
    
    try {
        # Method 1: Registry settings for Night Light
        $NightLightPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\CloudStore\Store\Cache\DefaultAccount"
        
        # Find the Night Light key
        $NightLightKeys = Get-ChildItem -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\CloudStore\Store" -Recurse -ErrorAction SilentlyContinue | 
            Where-Object { $_.Name -like "*windows.data.bluelightreduction.settings*" }
        
        if ($NightLightKeys) {
            foreach ($key in $NightLightKeys) {
                try {
                    $FullPath = $key.PSPath
                    
                    # Calculate intensity value (0-100 maps to 0-64 in hex)
                    $IntensityValue = [Math]::Round($Intensity * 0.64)
                    
                    # Enable Night Light with specified intensity
                    if ($Enable) {
                        # Binary data structure for Night Light
                        $header = [byte[]](0x43, 0x42, 0x01, 0x00, 0x0a, 0xd7, 0x23, 0xf1, 0x31, 0xd6, 0x01, 0x00)
                        $data1 = [byte[]](0x43, 0x42, 0x01, 0x00, 0xca, 0x32, 0x01, 0x5d, 0x31, 0xd6, 0x01, 0x00)
                        $data2 = [byte[]](0x15, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x2a, 0x07, 0x01, 0x00)
                        $data3 = [byte[]](0x1e, 0x00, 0x00, 0x00, 0xcb, 0x00, 0x00, 0x00, 0x14, 0x05, 0x00, 0x00)
                        $intensity = [byte[]]([byte]$IntensityValue, 0x00, 0x00, 0x00)
                        
                        $fullData = $header + $data1 + $data2 + $data3 + $intensity
                        
                        Set-ItemProperty -Path $FullPath -Name "Data" -Value $fullData -Type Binary -Force
                    }
                } catch {
                    # Continue to next key if this one fails
                }
            }
        }
        
        # Method 2: Additional registry settings
        $BlueLightPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\CloudStore\Store\DefaultAccount\Current\default`$windows.data.bluelightreduction.bluelightreductionstate\windows.data.bluelightreduction.bluelightreductionstate"
        
        if (Test-Path $BlueLightPath) {
            if ($Enable) {
                # Enable with 1, disable with 0
                Set-ItemProperty -Path $BlueLightPath -Name "Data" -Value ([byte[]](0x02, 0x00, 0x00, 0x00, 0x2f, 0x1f, 0x5a, 0x81, 0xf8, 0xd5, 0x01, 0x00, 0x43, 0x42, 0x01, 0x00, 0xca, 0x32, 0x5a, 0x81, 0xf8, 0xd5, 0x01, 0x00, 0x15, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x2b, 0x05, 0x01, 0x00, 0x14, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00)) -Type Binary -Force
            }
        }
        
        Write-Host "Night light set to $Intensity%" -ForegroundColor Yellow
        
        # Force a settings refresh
        $null = Start-Process -FilePath "reg.exe" -ArgumentList "add `"HKCU\Software\Microsoft\Windows\CurrentVersion\CloudStore\Store`" /f" -WindowStyle Hidden -Wait
        
    } catch {
        Write-Host "Note: Night light settings may require manual activation in Windows Settings" -ForegroundColor Yellow
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