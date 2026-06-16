# Windows Theme Automation

🌓 **Windows Theme Automation** switches Windows between day and night profiles, updates Night Light warmth, and keeps the whole routine on schedule.

It started as a PowerShell automation script and is now moving into a small .NET 8 desktop app with a reusable core, a CLI runner, and a modern WPF configuration UI.

## ✨ Highlights

- ☀️ **Day profile** from `07:00` to `18:59`
  - Light Windows theme
  - Night Light at 20 percent warmth
- 🌙 **Night profile** from `19:00` to `06:59`
  - Dark Windows theme
  - Night Light at 50 percent warmth
- 🎛️ **Modern WPF app** for profiles, diagnostics, and scheduler settings
- 🧰 **CLI runner** for scheduled tasks and scripting
- 🌡️ **Hybrid Night Light strategy**: native registry update first, gamma fallback second
- 🗓️ **Windows scheduled tasks** for day switch, night switch, and logon
- 🧾 **JSON config** stored at `%LOCALAPPDATA%\WindowsThemeAuto\config.json`
- 🩺 **Diagnostics** for Night Light registry blobs, Windows services, and task state

## 🧱 Project Structure

- `ThemeAutomation.Core`: shared automation logic, scheduling, configuration, diagnostics, and Windows integrations.
- `ThemeAutomation.Cli`: `themeauto` command-line runner for scheduled tasks.
- `ThemeAutomation.App`: WPF desktop configuration app.
- `windows_theme_automation.ps1`: compatibility launcher for existing PowerShell users.

## ✅ Requirements

- Windows 10 or Windows 11
- .NET 8 SDK to build from source
- .NET 8 Desktop Runtime to run the WPF app from the framework-dependent release ZIP

## 📦 Download

For normal use, download the latest GitHub release ZIP:

1. Open the **Releases** page.
2. Download `WindowsThemeAutomation-v0.1.0-win-x64.zip`.
3. Extract the ZIP.
4. Open `ThemeAutomation.App.exe` for the visual app.

> [!NOTE]
> `themeauto.exe` is the command-line tool used by scheduled tasks. If you double-click it, it may open and close quickly. That is expected. Use `ThemeAutomation.App.exe` for the desktop UI.

## 🚀 Build

Build the solution:

```powershell
dotnet build .\ThemeAutomation.sln
```

Run the WPF app from source:

```powershell
dotnet run --project .\src\ThemeAutomation.App\ThemeAutomation.App.csproj
```

Publish a local release build:

```powershell
dotnet publish .\src\ThemeAutomation.App\ThemeAutomation.App.csproj -c Release -r win-x64 --self-contained false -o .\artifacts\WindowsThemeAutomation-v0.1.0-win-x64
dotnet publish .\src\ThemeAutomation.Cli\ThemeAutomation.Cli.csproj -c Release -r win-x64 --self-contained false -o .\artifacts\WindowsThemeAutomation-v0.1.0-win-x64
```

Publish the CLI to the local install directory:

```powershell
dotnet publish .\src\ThemeAutomation.Cli\ThemeAutomation.Cli.csproj -c Release -o "$env:LOCALAPPDATA\WindowsThemeAuto"
```

## 🧑‍💻 CLI Usage

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
- `status`: shows active profile, config path, Night Light diagnostics, and task state.
- `diagnose`: prints Night Light registry and service diagnostics.

## 🔁 Compatibility Launcher

Existing usage of `windows_theme_automation.ps1` still works as a launcher once `themeauto.exe` has been published.

Examples:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\windows_theme_automation.ps1 -AutoRun
powershell.exe -ExecutionPolicy Bypass -File .\windows_theme_automation.ps1 -Install
powershell.exe -ExecutionPolicy Bypass -File .\windows_theme_automation.ps1 -Status
```

If the launcher cannot find `themeauto.exe`, it prints the publish command you need to run.

## 🌡️ Night Light Strategy

Night Light settings are stored by Windows in CloudStore registry blobs. Older versions of this project updated only the default settings blob, which could leave Windows out of sync on systems that also use per-device blobs.

V2 updates every detected `bluelightreduction.settings` blob, including per-device settings. It also bumps the CloudStore timestamp, broadcasts Windows display setting changes, and uses a small native "nudge" refresh so Windows is more likely to apply the requested warmth visually.

If Windows still does not visually apply the native Night Light change, the app can apply a gamma fallback filter when fallback mode is enabled.

If Night Light has never been enabled in Windows Settings, initialize it once:

1. Open **Settings**.
2. Go to **System**.
3. Open **Display**.
4. Open **Night light**.
5. Turn it on once.

Then run:

```powershell
themeauto diagnose
```

## 🗓️ Scheduled Tasks

Managed task names:

- `ThemeAutoSwitch_7AM`
- `ThemeAutoSwitch_7PM`
- `ThemeAutoSwitch_Startup`

The CLI creates tasks with explicit local wall-clock times from `config.json`, avoiding the previous observed `08:00` and `20:00` trigger drift.

## 🧪 Development

Run the lightweight test project:

```powershell
dotnet run --project .\tests\ThemeAutomation.Tests\ThemeAutomation.Tests.csproj
```

The tests cover schedule selection, config compatibility, scheduler behavior, Night Light blob editing, fallback decisions, and WPF smoke checks.

## 🛡️ Safety

The app modifies only local Windows theme registry values, Night Light CloudStore values, the display gamma ramp when fallback is required, local config/log files, and Windows scheduled tasks.

Because Night Light internals are undocumented by Microsoft, the app keeps diagnostics visible and falls back conservatively when the native refresh looks unreliable.
