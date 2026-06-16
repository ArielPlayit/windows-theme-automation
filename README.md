# Windows Theme Automation

🌓 **Windows Theme Automation** is a small Windows desktop app that switches between day and night profiles, updates Night Light warmth, and keeps everything running on schedule.

It is built for people who want Windows to feel right automatically: bright and light during the day, darker and warmer at night.

## ✨ What It Does

- ☀️ Applies a **day profile** at `07:00`
- 🌙 Applies a **night profile** at `19:00`
- 🎨 Switches Windows between light and dark mode
- 🌡️ Adjusts Windows Night Light warmth
- 🗓️ Installs scheduled tasks for automatic switching
- 🩺 Shows diagnostics for Night Light and scheduler state
- ⚙️ Lets you personalize active days, logon behavior, delay, and fallback mode

## 📦 Download

Download the latest release:

👉 [Windows Theme Automation v0.1.0](https://github.com/ArielPlayit/windows-theme-automation/releases/tag/v0.1.0)

Then:

1. Download `WindowsThemeAutomation-v0.1.0-win-x64.zip`.
2. Extract the ZIP.
3. Open `ThemeAutomation.App.exe`.

> [!IMPORTANT]
> Open `ThemeAutomation.App.exe` for the visual app.  
> `themeauto.exe` is the command-line tool used by scheduled tasks, so it may close quickly when double-clicked.

## ✅ Requirements

- Windows 10 or Windows 11
- [.NET 8 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/8.0)

## 🖥️ App Sections

- **Profiles**: configure day/night theme mode, schedule times, Night Light warmth, and fallback mode.
- **Diagnostics**: check Night Light registry availability, config paths, logs, and service health.
- **Scheduler**: install/uninstall automation tasks and customize when automation is allowed to run.

## 🌡️ Night Light Handling

Windows Night Light does not have a stable public API, so this app uses a careful hybrid approach:

- updates Windows Night Light registry blobs;
- refreshes CloudStore metadata so Windows is more likely to apply the change visually;
- broadcasts Windows display setting changes;
- uses gamma fallback only when the native Night Light refresh looks unreliable.

If Night Light has never been enabled before, open Windows Settings once and turn Night Light on manually. After that, run diagnostics from the app.

## 🛡️ Safety

The app only modifies local Windows theme settings, Night Light CloudStore values, optional display gamma fallback, local config/log files, and Windows scheduled tasks created by the app.

Your config is stored locally at:

```text
%LOCALAPPDATA%\WindowsThemeAuto\config.json
```

<details>
<summary>🧑‍💻 Developer Notes</summary>

## Project Structure

- `ThemeAutomation.Core`: shared automation logic and Windows integrations.
- `ThemeAutomation.Cli`: `themeauto` command-line runner for scheduled tasks.
- `ThemeAutomation.App`: WPF desktop configuration app.
- `windows_theme_automation.ps1`: compatibility launcher for older PowerShell usage.

## Build From Source

```powershell
dotnet build .\ThemeAutomation.sln
```

Run the app from source:

```powershell
dotnet run --project .\src\ThemeAutomation.App\ThemeAutomation.App.csproj
```

Run tests:

```powershell
dotnet run --project .\tests\ThemeAutomation.Tests\ThemeAutomation.Tests.csproj
```

Publish a local release build:

```powershell
dotnet publish .\src\ThemeAutomation.App\ThemeAutomation.App.csproj -c Release -r win-x64 --self-contained false -o .\artifacts\WindowsThemeAutomation-v0.1.0-win-x64
dotnet publish .\src\ThemeAutomation.Cli\ThemeAutomation.Cli.csproj -c Release -r win-x64 --self-contained false -o .\artifacts\WindowsThemeAutomation-v0.1.0-win-x64
```

## CLI Commands

```powershell
themeauto apply
themeauto install
themeauto uninstall
themeauto status
themeauto diagnose
```

## Scheduled Task Names

- `ThemeAutoSwitch_7AM`
- `ThemeAutoSwitch_7PM`
- `ThemeAutoSwitch_Startup`

</details>
