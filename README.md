# Windows Theme Automation

PowerShell script to automate Windows theme configuration.

## Usage

Open PowerShell and run:

`powershell
powershell -ExecutionPolicy Bypass -File .\\windows_theme_automation.ps1
`
"@ -Encoding UTF8; Set-Content -Path (Join-Path C:\Users\ArielPlayit\Documents\windows-theme-automation ".gitignore") -Value @"
# General
*.log
*.tmp
Thumbs.db
Desktop.ini

# Editors
.vscode/
.idea/

## License

This project is licensed under the MIT License - see the LICENSE file for details.
