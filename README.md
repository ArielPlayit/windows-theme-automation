# 🌓 Windows Theme Automation

Automatically switch between light and dark themes in Windows based on time of day, with configurable Night Light intensity.

## ✨ Features

- 🌞 **Automatic Day Mode** (7 AM - 7 PM)
  - Light theme
  - Night Light at 20% intensity

- 🌙 **Automatic Night Mode** (7 PM - 7 AM)
  - Dark theme
  - Night Light at 50% intensity

- ⚡ **Runs in Background** - No manual intervention needed
- 🔄 **Auto-start on Login** - Applies theme when Windows starts
- ⏰ **Hourly Checks** - Ensures theme is always correct
- 🎯 **One-Click Installation** - No need to configure Task Scheduler manually

## 📋 Requirements

- Windows 10 or Windows 11
- PowerShell 5.1 or later (pre-installed on Windows)
- Administrator privileges (script will request them automatically)

## 🚀 Installation

### Step 1: Download the Script

1. Download `AutoTheme.ps1` from this repository
2. Save it anywhere on your computer (Desktop, Downloads, etc.)

### Step 2: Run the Script

1. **Right-click** on `windows_theme_automation.ps1`
2. Select **"Run with PowerShell"**
3. If prompted, allow Administrator privileges
4. Select option **1** (Install automation)
5. Done! The script is now installed and running

## 🎮 Usage

### Menu Options

When you run the script, you'll see a menu:

# 🌓 Windows Theme Automation

Automatically switch between Light and Dark themes in Windows based on time of day, and adjust Night Light intensity via the Windows registry.

## ✨ Features

- 🌞 Automatic Day Mode (07:00 — 18:59)
  - Applies Light theme
  - Sets Night Light to ~20% warmth

- 🌙 Automatic Night Mode (19:00 — 06:59)
  - Applies Dark theme
  - Sets Night Light to ~50% warmth

- 🔁 Runs via Scheduled Tasks (runs at 07:00, 19:00 and at user logon)
  - Tasks created: `ThemeAutoSwitch_7AM`, `ThemeAutoSwitch_7PM`, `ThemeAutoSwitch_Startup`

- ⚠️ Night Light registry adjustment
  - The script writes binary Night Light settings directly into the user CloudStore registry blob. If Night Light hasn't been enabled once via Settings, the script will notify you.

## 📋 Requirements

- Windows 10 or Windows 11
- PowerShell 5.1 or later
- Administrator privileges for installation/uninstallation (the script will request elevation)

## 🚀 Installation

1. Place `windows_theme_automation.ps1` somewhere on your PC.
2. Right-click the file and choose **Run with PowerShell** (or run from an elevated PowerShell prompt).
3. Choose **1. Install automation (recommended)** from the menu.

What the installer does:
- Copies the script to `%LOCALAPPDATA%\\WindowsThemeAuto\\`
- Registers three scheduled tasks: `ThemeAutoSwitch_7AM`, `ThemeAutoSwitch_7PM`, and `ThemeAutoSwitch_Startup`
- Runs the script once to apply current settings

The scheduled tasks call the script with the `-AutoRun` switch so it executes without the interactive menu.

## 🎮 Usage

Run the script to see the interactive menu:

- `1` — Install automation
- `2` — Apply theme now (run `Apply-ThemeSettings` once)
- `3` — Uninstall automation (removes tasks and installed copy)
- `4` — Exit

To run the script once (for example from Task Scheduler) use:

```powershell
powershell.exe -ExecutionPolicy Bypass -File "C:\\Path\\to\\windows_theme_automation.ps1" -AutoRun
```

## ⚙️ Customization

Edit `Apply-ThemeSettings` inside `windows_theme_automation.ps1` to change schedule or intensity values. The function uses `Set-NightLight -Percentage` (0..100) and `Set-WindowsTheme -Mode "Light"|"Dark"`.

Example snippet from the script:

```powershell
if ($CurrentHour -ge 7 -and $CurrentHour -lt 19) {
    Set-WindowsTheme -Mode "Light"
    Set-NightLight -Percentage 20 -Enable $true
} else {
    Set-WindowsTheme -Mode "Dark"
    Set-NightLight -Percentage 50 -Enable $true
}
```

After editing, uninstall and reinstall so the scheduled tasks use the updated copy.

## 🔧 Troubleshooting

- "Installation requires Administrator privileges": allow the elevation prompt.
- "Night Light: Not initialized": open **Settings > System > Display > Night light** and enable it once; then rerun the script.
- If Night Light changes don't take effect, the registry blob format might differ on some systems — ensure Night Light is enabled and try different percentage values.
- If theme changes seem incomplete, signing out or restarting Windows will ensure settings apply.

## 🗑️ Uninstallation

Run the script and select **3. Uninstall automation**. This removes the scheduled tasks and deletes `%LOCALAPPDATA%\\WindowsThemeAuto`.

## 📁 What Gets Installed

Files are copied to:

```
%LOCALAPPDATA%\\WindowsThemeAuto\\
```

Scheduled tasks created:
- `ThemeAutoSwitch_7AM`
- `ThemeAutoSwitch_7PM`
- `ThemeAutoSwitch_Startup`

## ⚒️ Technical details

- Theme switching: updates `HKCU:\\\\SOFTWARE\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Themes\\\\Personalize` keys and restarts Explorer.
- Night Light: edits the binary `Data` value under the user CloudStore path used by Windows Night Light (searches for a marker and updates two bytes representing temperature).
- Scheduling: uses `Register-ScheduledTask` with an interactive principal so tasks run for the current user.

## 🛡️ Safety & Privacy

- The script only modifies local registry values and scheduled tasks.
- No network access or external connections.

## 📄 License

MIT License — see the `LICENSE` file.

---

If you'd like, I can also add a short checklist for verifying installation or a troubleshooting flow. Would you like that?

## ⭐ Support

If this script helped you, please consider:
- Giving it a ⭐ star on GitHub
- Sharing it with others
- Reporting any issues you find
