# Windows Theme Automation

Windows Theme Automation switches Windows between day and night profiles and adjusts Night Light warmth automatically.

This repository is moving from a single PowerShell script to a .NET architecture:

- `ThemeAutomation.Core`: shared automation logic and Windows integrations.
- `ThemeAutomation.Cli`: `themeauto` command-line runner for scheduled tasks.
- `ThemeAutomation.App`: WPF configuration app shell.
- `windows_theme_automation.ps1`: compatibility launcher for existing users.

## Features

- Day profile from `07:00` to `18:59`
  - Light Windows theme
  - Night Light at 20 percent warmth
- Night profile from `19:00` to `06:59`
  - Dark Windows theme
  - Night Light at 50 percent warmth
- Native Night Light first, gamma fallback second
- Scheduled tasks at day start, night start, and user logon
- JSON configuration at `%LOCALAPPDATA%\WindowsThemeAuto\config.json`
- Diagnostics for scheduled tasks and Night Light registry availability

## Requirements

- Windows 10 or Windows 11
- .NET 8 SDK to build
- .NET 8 Desktop Runtime to run the WPF app

## Download

For normal use, download the latest GitHub release ZIP:

1. Go to **Releases**.
2. Download `WindowsThemeAutomation-v0.1.0-win-x64.zip`.
3. Extract the ZIP.
4. Open `ThemeAutomation.App.exe` for the visual app.

`themeauto.exe` is the command-line tool used by scheduled tasks. It is normal for it to close quickly when double-clicked.

## Build

```powershell
dotnet build .\ThemeAutomation.sln
```

Publish the CLI to the install directory:

```powershell
dotnet publish .\src\ThemeAutomation.Cli\ThemeAutomation.Cli.csproj -c Release -o "$env:LOCALAPPDATA\WindowsThemeAuto"
```

Publish a local release build:

```powershell
dotnet publish .\src\ThemeAutomation.App\ThemeAutomation.App.csproj -c Release -r win-x64 --self-contained false -o .\artifacts\WindowsThemeAutomation-v0.1.0-win-x64
dotnet publish .\src\ThemeAutomation.Cli\ThemeAutomation.Cli.csproj -c Release -r win-x64 --self-contained false -o .\artifacts\WindowsThemeAutomation-v0.1.0-win-x64
```

## CLI Usage

```powershell
themeauto apply
themeauto install
themeauto uninstall
themeauto status
themeauto diagnose
```

Commands:

- `apply`: applies the correct profile for the current local time.
- `install`: creates scheduled tasks for day start, night start, and logon.
- `uninstall`: removes managed scheduled tasks.
- `status`: shows active profile, config path, Night Light diagnostics, and task status.
- `diagnose`: prints Night Light registry paths found under CloudStore.

## Compatibility Launcher

Existing usage of `windows_theme_automation.ps1` still works as a launcher once `themeauto.exe` has been published.

Examples:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\windows_theme_automation.ps1 -AutoRun
powershell.exe -ExecutionPolicy Bypass -File .\windows_theme_automation.ps1 -Install
powershell.exe -ExecutionPolicy Bypass -File .\windows_theme_automation.ps1 -Status
```

If the launcher cannot find `themeauto.exe`, it prints the publish command.

## Night Light Strategy

Night Light settings are stored by Windows in CloudStore registry blobs. Older versions of this project updated only the default settings blob, which can leave Windows out of sync on systems that also use per-device blobs.

V2 updates every detected `bluelightreduction.settings` blob, including per-device settings, verifies the temperature bytes after writing, broadcasts display setting changes, and then uses a gamma fallback only if the native update does not verify.

If Night Light has never been enabled in Windows Settings, run:

1. Settings
2. System
3. Display
4. Night light
5. Turn it on once

Then run:

```powershell
themeauto diagnose
```

## Scheduled Tasks

Managed task names:

- `ThemeAutoSwitch_7AM`
- `ThemeAutoSwitch_7PM`
- `ThemeAutoSwitch_Startup`

The CLI creates tasks with explicit local wall-clock times from `config.json`, avoiding the previous observed `08:00` and `20:00` trigger drift.

## Development

Run the lightweight test project:

```powershell
dotnet run --project .\tests\ThemeAutomation.Tests\ThemeAutomation.Tests.csproj
```

The tests cover schedule selection, Night Light blob editing, and fallback decisions.

## Safety

The app modifies only local Windows theme registry values, Night Light CloudStore values, display gamma ramp when fallback is required, local config/log files, and Windows scheduled tasks.
