# 🚀 Rocmine's Program Installer

Quick setup guide for automatically installing essential Windows programs using PowerShell and winget.

## ⚡ Quick Start

### Method 1: Direct Download & Run
```powershell
# Download and run in one command
iwr -useb https://fresh.rocmine.net/install.ps1 | iex
```

### Method 2: Download First, Then Run
```powershell
# Download the script
Invoke-WebRequest -Uri "https://fresh.rocmine.net/install.ps1" -OutFile "install.ps1"

# Run the installer
.\install.ps1
```

## 🔧 Advanced Usage

### Skip Confirmation Prompt
```powershell
.\install.ps1 -SkipConfirmation
```

### Continue Installing Even If Some Packages Fail
```powershell
.\install.ps1 -ContinueOnError
```

### Exclude Specific Packages
```powershell
.\install.ps1 -ExcludePackages @('Discord.Discord', 'Valve.Steam')
```

### Combine Multiple Options
```powershell
.\install.ps1 -SkipConfirmation -ContinueOnError -ExcludePackages @('Docker.DockerDesktop')
```

## 📋 Prerequisites

- **Windows 10/11** with PowerShell 5.1+
- **Windows Package Manager (winget)** - Usually pre-installed on Windows 11
- **Administrator privileges** recommended for some packages

## 📦 What Gets Installed

The script installs 36+ essential programs across categories:
- **Development**: VS Code, Git, Node.js, Python, Docker
- **System Utilities**: PowerShell, Windows Terminal, PowerToys
- **Media**: OBS Studio, Audacity, Steam, Discord
- **Productivity**: 7-Zip, LibreOffice, Foxit Reader
- **And much more!**

## ⚠️ Important Notes

- First run may take 15-30 minutes depending on your internet speed
- Some packages may require a system restart
- If a package fails, the script will continue with the next one
- Already installed packages will be skipped automatically

## 🆘 Troubleshooting

**If you get execution policy errors:**
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

**If winget is missing:**
- Install from Microsoft Store: "App Installer"
- Or download from [GitHub](https://github.com/microsoft/winget-cli)

---

*Made with ❤️ by Rocmine | [Report Issues](https://github.com/rocmine/installer)*
