# ============================================
# Auto-Installation Script for Windows Theme Automation
# Day/Night Mode and Night Light Automation
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

function Set-NightLight {
    param(
        [int]$Percentage,  # 0-100
        [bool]$Enable
    )
    
    try {
        # Registry path (using Current store which is the active one)
        $settingsPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\CloudStore\Store\DefaultAccount\Current\default`$windows.data.bluelightreduction.settings\windows.data.bluelightreduction.settings"
        
        # Check if Night Light is initialized
        if (-not (Test-Path $settingsPath)) {
            Write-Host "Night Light: Not initialized. Please enable it manually once in Settings." -ForegroundColor Yellow
            return
        }
        
        # Read existing settings data
        try {
            $prop = Get-ItemProperty -Path $settingsPath -Name "Data" -ErrorAction Stop
            $settingsData = [byte[]]$prop.Data
        } catch {
            Write-Host "Night Light: Could not read settings data" -ForegroundColor Yellow
            return
        }
        
        if ($settingsData.Length -lt 40) {
            Write-Host "Night Light: Invalid data structure" -ForegroundColor Yellow
            return
        }
        
        # Calculate temperature value based on percentage
        # Real values from your system:
        # 20% = 0x5580 (21888 decimal)
        # 50% = 0x3BAA (15274 decimal)
        
        # Extended formula for 0-100%:
        # At 0%: ~25600 (neutral/coolest)
        # At 20%: 21888
        # At 50%: 15274
        # At 100%: ~2560 (warmest)
        
        $tempValue = 0
        if ($Enable -and $Percentage -gt 0) {
            # Linear interpolation between known points
            if ($Percentage -le 20) {
                # 0% to 20% range
                $tempValue = [int](25600 - (($Percentage / 20.0) * (25600 - 21888)))
            } elseif ($Percentage -le 50) {
                # 20% to 50% range
                $tempValue = [int](21888 - ((($Percentage - 20) / 30.0) * (21888 - 15274)))
            } else {
                # 50% to 100% range
                $tempValue = [int](15274 - ((($Percentage - 50) / 50.0) * (15274 - 2560)))
            }
        } else {
            # Disabled or 0% - set to neutral
            $tempValue = 25600
        }
        
        # Clamp to valid range
        if ($tempValue -lt 2560) { $tempValue = 2560 }
        if ($tempValue -gt 25600) { $tempValue = 25600 }
        
        # Convert to little-endian 16-bit bytes
        $lowByte = [byte]($tempValue -band 0xFF)
        $highByte = [byte](($tempValue -shr 8) -band 0xFF)
        
        # Find and update the CF 28 marker (temperature setting)
        $found = $false
        for ($i = 0; $i -lt ($settingsData.Length - 3); $i++) {
            if ($settingsData[$i] -eq 0xCF -and $settingsData[$i+1] -eq 0x28) {
                # Found temperature marker at position $i
                # Next 2 bytes are the temperature value (little-endian)
                $settingsData[$i+2] = $lowByte
                $settingsData[$i+3] = $highByte
                $found = $true
                break
            }
        }
        
        if (-not $found) {
            Write-Host "Night Light: Could not find temperature marker in data" -ForegroundColor Yellow
            return
        }
        
        # ONLY modify the registry - NO gamma manipulation
        Set-ItemProperty -Path $settingsPath -Name "Data" -Value $settingsData -Type Binary -Force
        
        # Calculate approximate color temperature for display
        $colorTemp = [int](6500 - (($Percentage / 100.0) * (6500 - 2700)))
        
        $statusText = if ($Enable -and $Percentage -gt 0) { "ENABLED" } else { "DISABLED" }
        Write-Host "Night Light: $statusText at $Percentage% (~${colorTemp}K)" -ForegroundColor Yellow
        
    } catch {
        Write-Host "Night Light: Error - $_" -ForegroundColor Red
        Write-Host "Make sure Night Light is enabled in Settings > Display" -ForegroundColor Yellow
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
    $taskName7AM = "ThemeAutoSwitch_7AM"
    $taskName7PM = "ThemeAutoSwitch_7PM"
    $taskNameStartup = "ThemeAutoSwitch_Startup"
    
    # Remove old variants
    Unregister-ScheduledTask -TaskName "ThemeAutoSwitch_Hourly" -Confirm:$false -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName "ThemeAutoSwitch_Daily" -Confirm:$false -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName $taskName7AM -Confirm:$false -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName $taskName7PM -Confirm:$false -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName $taskNameStartup -Confirm:$false -ErrorAction SilentlyContinue
    
    Write-Host "`nCreating scheduled tasks..." -ForegroundColor Cyan
    
    # Task 1: Run at 7 AM daily
    $action7AM = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$finalScriptPath`" -AutoRun"
    $trigger7AM = New-ScheduledTaskTrigger -Daily -At "07:00"
    $principal7AM = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive
    $settings7AM = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 5)
    
    Register-ScheduledTask -TaskName $taskName7AM -Action $action7AM -Trigger $trigger7AM -Principal $principal7AM -Settings $settings7AM -Description "Switches to Day Mode at 7 AM" -Force | Out-Null
    Write-Host "  Created: 7 AM daily task" -ForegroundColor Green
    
    # Task 2: Run at 7 PM daily
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