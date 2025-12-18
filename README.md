# 🌓 Windows Theme Automation


Automatically switch between light and dark themes in Windows based on time of day, with configurable Night Light intensity.

## âœ¨ Features

- ðŸŒž **Automatic Day Mode** (7 AM - 7 PM)
  - Light theme
  - Night Light at 20% intensity

- 🌙 **Automatic Night Mode** (7 PM - 7 AM)

  - Dark theme
  - Night Light at 50% intensity

- âš¡ **Runs in Background** - No manual intervention needed
- ðŸ”„ **Auto-start on Login** - Applies theme when Windows starts
- â° **Hourly Checks** - Ensures theme is always correct
- ðŸŽ¯ **One-Click Installation** - No need to configure Task Scheduler manually

## ðŸ“‹ Requirements

- Windows 10 or Windows 11
- PowerShell 5.1 or later (pre-installed on Windows)
- Administrator privileges (script will request them automatically)

## ðŸš€ Installation

### Step 1: Download the Script

1. Download `AutoTheme.ps1` from this repository
2. Save it anywhere on your computer (Desktop, Downloads, etc.)

### Step 2: Run the Script

1. **Right-click** on `windows_theme_automation.ps1`
2. Select **"Run with PowerShell"**
3. If prompted, allow Administrator privileges
4. Select option **1** (Install automation)
5. Done! The script is now installed and running

## ðŸŽ® Usage

### Menu Options

When you run the script, you'll see a menu:

```
1. Install automation (recommended)    - Sets up automatic theme switching
2. Apply theme now (without installing) - Applies current theme once
3. Uninstall automation                - Removes all scheduled tasks
4. Exit                                - Close the script
```

### After Installation

The script runs automatically:
- âœ… Every hour to check and apply the correct theme
- âœ… When you log in to Windows
- âœ… Works silently in the background

**You don't need to run the script again!**

## âš™ï¸ Customization

To change the schedule or intensity levels, edit the `Apply-ThemeSettings` function in the script:

```powershell
if ($CurrentHour -ge 7 -and $CurrentHour -lt 19) {
    # Day Mode: 7 AM - 7 PM
    Set-WindowsTheme -Mode "Light"
    Set-NightLight -Intensity 20 -Enable $true  # Change intensity here (0-100)
}
else {
    # Night Mode: 7 PM - 7 AM
    Set-WindowsTheme -Mode "Dark"
    Set-NightLight -Intensity 50 -Enable $true  # Change intensity here (0-100)
}
```

**Example customizations:**

- Change time: Modify `7` (7 AM) and `19` (7 PM) to your preferred hours
- Adjust intensity: Change `20` and `50` to any value between 0-100
- Disable Night Light during day: Set `Enable $false` for day mode

After making changes, reinstall the script (option 3 to uninstall, then option 1 to install).

## ðŸ”§ Troubleshooting

### "Cannot load file" error

Run this command in PowerShell (as Administrator):
```powershell
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Theme doesn't apply completely

The script automatically restarts Windows Explorer. If issues persist:
1. Log out and log back in
2. Restart your computer

### Night Light doesn't change or colors look wrong

The script uses direct gamma adjustment instead of Windows Night Light. If you experience issues:
1. Make sure the script is running with Administrator privileges
2. Try adjusting the intensity values in the script
3. Restart your computer to reset display gamma

### Script doesn't run automatically

Check that the scheduled tasks were created:
1. Open **Task Scheduler**
2. Look for tasks named:
   - `ThemeAutoSwitch_Hourly`
   - `ThemeAutoSwitch_Startup`

If missing, run the script again and select option 1 to reinstall.

## ðŸ—‘ï¸ Uninstallation

1. Run the script
2. Select option **3** (Uninstall automation)
3. The script will remove all scheduled tasks and files

Alternatively, manually delete:
- Scheduled tasks: `ThemeAutoSwitch_Hourly` and `ThemeAutoSwitch_Startup`
- Script folder: `%LOCALAPPDATA%\WindowsThemeAuto`

## ðŸ“ What Gets Installed

The script installs to:
```
C:\Users\YourUsername\AppData\Local\WindowsThemeAuto\
```

And creates two scheduled tasks in Windows Task Scheduler.

## ðŸ”’ Security

- The script only modifies Windows theme and display gamma settings
- No network access or external connections
- Open source - you can review all code
- Requires admin privileges for Task Scheduler and Explorer restart

## ðŸ› ï¸ Technical Details

The script uses:
- **Registry modification** for Windows theme (Light/Dark mode)
- **Direct gamma control via Win32 API** for warm color adjustment (Night Light effect)
- **Color temperature algorithm** to convert warmth percentage to RGB values (6500K to 2700K range)

## ðŸ¤ Contributing

Contributions are welcome! Feel free to:
- Report bugs
- Suggest new features
- Submit pull requests

## ðŸ“„ License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## â­ Support

If this script helped you, please consider:
- Giving it a â­ star on GitHub
- Sharing it with others
- Reporting any issues you find

## ðŸ“ Changelog
### Version 1.2.0
- Sincronización: se actualizaron y mejoraron funciones del script según cambios locales.
- Commit: Update windows_theme_automation.ps1  sync local changes



### Version 1.2.0
- **Script improvements** - Updated and enhanced script functions
- Code synchronization with latest local changes
- General stability and reliability improvements

### Version 1.1.0
- **Improved Night Light control** - Now uses direct gamma adjustment via Win32 API
- More reliable warm color application
- No longer requires manual Night Light activation
- Better PowerShell 5.1 compatibility

### Version 1.0.0
- Initial release
- Automatic theme switching based on time
- Night Light intensity control
- One-click installation
- Background execution

---

**Made with â¤ï¸ for Windows users who love automation**
