#Requires -Version 5.1
<#
.SYNOPSIS
    Rocmine's Automated Program Installer
.DESCRIPTION
    Installs a curated list of essential programs using Windows Package Manager (winget)
.AUTHOR
    Rocmine
#>

[CmdletBinding()]
param(
    [switch]$SkipConfirmation,
    [switch]$ContinueOnError
)

# Script configuration
$ErrorActionPreference = 'Continue'
$ProgressPreference = 'SilentlyContinue'

# Package definitions with categories for better organization
$PackageCategories = @{
    'Development' = @(
        @{Name='Microsoft.VisualStudioCode'; Display='Visual Studio Code'; Selected=$true},
        @{Name='Git.Git'; Display='Git'; Selected=$true},
        @{Name='OpenJS.NodeJS'; Display='Node.js'; Selected=$true},
        @{Name='Yarn.Yarn'; Display='Yarn'; Selected=$false},
        @{Name='Python.Python.3.12'; Display='Python 3.12'; Selected=$true},
        @{Name='Docker.DockerDesktop'; Display='Docker Desktop'; Selected=$false},
        @{Name='Postman.Postman'; Display='Postman'; Selected=$false},
        @{Name='CoreyButler.NVMforWindows'; Display='NVM for Windows'; Selected=$false}
    )
    'System Utilities' = @(
        @{Name='Microsoft.PowerShell'; Display='PowerShell'; Selected=$true},
        @{Name='Microsoft.WindowsTerminal'; Display='Windows Terminal'; Selected=$true},
        @{Name='Microsoft.PowerToys'; Display='PowerToys'; Selected=$true},
        @{Name='JanDeDobbeleer.OhMyPosh'; Display='Oh My Posh'; Selected=$true},
        @{Name='Fastfetch-cli.Fastfetch'; Display='Fastfetch'; Selected=$true},
        @{Name='REALiX.HWiNFO'; Display='HWiNFO'; Selected=$false},
        @{Name='Rufus.Rufus'; Display='Rufus'; Selected=$false},
        @{Name='WiresharkFoundation.Wireshark'; Display='Wireshark'; Selected=$false}
    )
    'Media & Entertainment' = @(
        @{Name='OBSProject.OBSStudio'; Display='OBS Studio'; Selected=$false},
        @{Name='Audacity.Audacity'; Display='Audacity'; Selected=$false},
        @{Name='Valve.Steam'; Display='Steam'; Selected=$true},
        @{Name='EpicGames.EpicGamesLauncher'; Display='Epic Games Launcher'; Selected=$false},
        @{Name='Discord.Discord'; Display='Discord'; Selected=$true},
        @{Name='Stremio.Stremio'; Display='Stremio'; Selected=$false},
        @{Name='qBittorrent.qBittorrent'; Display='qBittorrent'; Selected=$false},
        @{Name='yt-dlp.yt-dlp'; Display='yt-dlp'; Selected=$false}
    )
    'Productivity' = @(
        @{Name='7zip.7zip'; Display='7-Zip'; Selected=$true},
        @{Name='TheDocumentFoundation.LibreOffice'; Display='LibreOffice'; Selected=$false},
        @{Name='Foxit.FoxitReader'; Display='Foxit PDF Reader'; Selected=$false},
        @{Name='PuTTY.PuTTY'; Display='PuTTY'; Selected=$false},
        @{Name='Parsec.Parsec'; Display='Parsec'; Selected=$false}
    )
    'System Components' = @(
        @{Name='Microsoft.VCRedist.2015+.x86'; Display='Visual C++ Redistributable (x86)'; Selected=$true},
        @{Name='Microsoft.VCRedist.2015+.x64'; Display='Visual C++ Redistributable (x64)'; Selected=$true},
        @{Name='Microsoft.DirectX'; Display='DirectX Runtime'; Selected=$false}
    )
    'Network & Security' = @(
        @{Name='angryziber.AngryIPScanner'; Display='Angry IP Scanner'; Selected=$false},
        @{Name='ChatterinoTeam.Chatterino'; Display='Chatterino'; Selected=$false}
    )
    'Graphics & Drivers' = @(
        @{Name='Wagnardsoft.DisplayDriverUninstaller'; Display='Display Driver Uninstaller'; Selected=$false},
        @{Name='TechPowerUp.NVCleanstall'; Display='NVCleanstall'; Selected=$false}
    )
}

# Global variables for selected packages
$SelectedPackages = @()

# Statistics tracking
$Stats = @{
    Total = 0
    Successful = 0
    Failed = 0
    Skipped = 0
    StartTime = Get-Date
}

function Write-ColoredBanner {
    param([string]$Text, [string]$Color = 'Cyan')
    
    $border = '=' * 60
    Clear-Host
    Write-Host $border -ForegroundColor $Color
    Write-Host ("  {0,-54}  " -f $Text) -ForegroundColor White -BackgroundColor DarkBlue
    Write-Host $border -ForegroundColor $Color
}

function Write-ProgressInfo {
    param(
        [string]$PackageName,
        [int]$Current,
        [int]$Total,
        [string]$Status = 'Installing'
    )
    
    $percentComplete = if ($Total -gt 0) { [math]::Round(($Current / $Total) * 100, 1) } else { 0 }
    $progressBar = ('█' * [math]::Floor($percentComplete / 4)) + ('░' * (25 - [math]::Floor($percentComplete / 4)))
    
    Write-Host "`n┌─ Package Progress ─────────────────────────────────────┐" -ForegroundColor DarkGray
    Write-Host ("│ {0,-54} │" -f "$Status`: $PackageName") -ForegroundColor Yellow
    Write-Host ("│ Progress: [{0}] {1}% ({2}/{3})" -f $progressBar, $percentComplete, $Current, $Total) -ForegroundColor Cyan
    Write-Host "└────────────────────────────────────────────────────────┘" -ForegroundColor DarkGray
}

function Test-WingetAvailability {
    try {
        $wingetCommand = Get-Command winget -ErrorAction Stop
        $wingetOutput = & winget --version 2>$null
        if ($wingetOutput) {
            $wingetVersion = $wingetOutput.ToString().Trim('v')
            Write-Host "✓ Winget found (Version: $wingetVersion)" -ForegroundColor Green
            return $true
        } else {
            throw "Winget command failed"
        }
    }
    catch {
        Write-Host "✗ Winget not found or not accessible" -ForegroundColor Red
        Write-Host "Please install Windows Package Manager from Microsoft Store or GitHub" -ForegroundColor Yellow
        Write-Host "GitHub: https://github.com/microsoft/winget-cli/releases" -ForegroundColor Blue
        return $false
    }
}

function Install-Package {
    param([string]$PackageName)
    
    $installArgs = @(
        'install', $PackageName,
        '--accept-source-agreements',
        '--accept-package-agreements',
        '--silent',
        '--disable-interactivity'
    )
    
    try {
        $process = Start-Process -FilePath 'winget' -ArgumentList $installArgs -NoNewWindow -Wait -PassThru -RedirectStandardOutput 'NUL' -RedirectStandardError 'NUL'
        return $process.ExitCode
    }
    catch {
        Write-Host "✗ Exception occurred: $($_.Exception.Message)" -ForegroundColor Red
        return -1
    }
}

function Install-CaskaydiaCoveFont {
    Write-Host "`n🔤 Installing CaskaydiaCove Nerd Font..." -ForegroundColor Cyan
    
    try {
        # Check if already installed
        $installedFonts = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts" -ErrorAction SilentlyContinue
        if ($installedFonts -and ($installedFonts.PSObject.Properties | Where-Object { $_.Name -like "*CaskaydiaCove*" -or $_.Name -like "*CascadiaCode*" })) {
            Write-Host "✓ CaskaydiaCove Nerd Font already installed" -ForegroundColor Green
            return $true
        }
        
        # Create temp directory
        $tempDir = Join-Path $env:TEMP "CaskaydiaCove_$(Get-Random)"
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
        
        # Download CaskaydiaCove Nerd Font
        $fontUrl = "https://github.com/ryanoasis/nerd-fonts/releases/download/v3.1.1/CascadiaCode.zip"
        $tempPath = Join-Path $tempDir "CascadiaCode.zip"
        $extractPath = Join-Path $tempDir "extracted"
        
        Write-Host "   Downloading font..." -ForegroundColor Yellow
        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Invoke-WebRequest -Uri $fontUrl -OutFile $tempPath -UseBasicParsing -UserAgent "PowerShell/5.1"
        }
        catch {
            Write-Host "   Failed to download font: $($_.Exception.Message)" -ForegroundColor Red
            return $false
        }
        
        Write-Host "   Extracting font files..." -ForegroundColor Yellow
        try {
            Expand-Archive -Path $tempPath -DestinationPath $extractPath -Force
        }
        catch {
            Write-Host "   Failed to extract font: $($_.Exception.Message)" -ForegroundColor Red
            return $false
        }
        
        # Install fonts using Add-Type
        Add-Type -AssemblyName System.Drawing
        $fonts = New-Object System.Drawing.Text.PrivateFontCollection
        
        # Get CaskaydiaCove font files specifically
        $fontFiles = Get-ChildItem -Path $extractPath -Filter "*.ttf" -Recurse | Where-Object { 
            $_.Name -match "CaskaydiaCove.*Nerd.*Font.*(Regular|Bold)" -and $_.Name -notmatch "(Italic|Light|Thin)" 
        }
        
        if ($fontFiles.Count -eq 0) {
            Write-Host "   No suitable font files found" -ForegroundColor Red
            return $false
        }
        
        # Install fonts by copying to Fonts directory
        $fontsPath = [Environment]::GetFolderPath("Fonts")
        $installedCount = 0
        
        foreach ($fontFile in $fontFiles) {
            try {
                $destPath = Join-Path $fontsPath $fontFile.Name
                if (-not (Test-Path $destPath)) {
                    Copy-Item $fontFile.FullName $destPath -Force
                    $installedCount++
                    Write-Host "   Installed: $($fontFile.Name)" -ForegroundColor Yellow
                }
            }
            catch {
                Write-Host "   Failed to install: $($fontFile.Name)" -ForegroundColor Red
            }
        }
        
        # Register fonts in registry (for system recognition)
        try {
            foreach ($fontFile in $fontFiles) {
                $fontName = [System.Drawing.FontFamily]::new($fontFile.FullName).Name
                if ($fontName) {
                    $regName = "$fontName (TrueType)"
                    $regValue = $fontFile.Name
                    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts" -Name $regName -Value $regValue -ErrorAction SilentlyContinue
                }
            }
        }
        catch {
            # Font registration failed, but files are copied
        }
        
function Configure-WindowsTerminal {
    Write-Host "`n⚙️  Configuring Windows Terminal..." -ForegroundColor Cyan
    
    $terminalSettingsPath = "$env:LOCALAPPDATA\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState\settings.json"
    
    if (-not (Test-Path $terminalSettingsPath)) {
        Write-Host "   Windows Terminal settings not found, skipping configuration" -ForegroundColor Yellow
        return $false
    }
    
    try {
        Write-Host "   Updating Windows Terminal settings..." -ForegroundColor Yellow
        $settingsContent = Get-Content $terminalSettingsPath -Raw -Encoding UTF8
        $settings = $settingsContent | ConvertFrom-Json
        
        # Ensure structure exists
        if (-not $settings.profiles) {
            $settings | Add-Member -NotePropertyName "profiles" -NotePropertyValue @{} -Force
        }
        if (-not $settings.profiles.defaults) {
            $settings.profiles | Add-Member -NotePropertyName "defaults" -NotePropertyValue @{} -Force
        }
        
        # Configure font for all profiles
        $fontConfig = @{
            face = "CaskaydiaCove Nerd Font"
            size = 11
            weight = "normal"
        }
        
        $settings.profiles.defaults | Add-Member -NotePropertyName "font" -NotePropertyValue $fontConfig -Force
        
        # Also update individual profiles if they exist
        if ($settings.profiles.list) {
            foreach ($profile in $settings.profiles.list) {
                if (-not $profile.font) {
                    $profile | Add-Member -NotePropertyName "font" -NotePropertyValue $fontConfig -Force
                } else {
                    $profile.font.face = "CaskaydiaCove Nerd Font"
                }
            }
        }
        
        # Save settings
        $updatedSettings = $settings | ConvertTo-Json -Depth 10 -Compress:$false
        Set-Content -Path $terminalSettingsPath -Value $updatedSettings -Encoding UTF8
        
        Write-Host "   ✓ Windows Terminal configured successfully" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "   ✗ Failed to configure Windows Terminal: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Configure-PowerShell7 {
    Write-Host "`n⚙️  Configuring PowerShell 7..." -ForegroundColor Cyan
    
    try {
        # Find PowerShell 7 installation
        $ps7Paths = @(
            "$env:ProgramFiles\PowerShell\7\pwsh.exe",
            "$env:ProgramFiles(x86)\PowerShell\7\pwsh.exe",
            (Get-Command pwsh -ErrorAction SilentlyContinue).Source
        ) | Where-Object { $_ -and (Test-Path $_) }
        
        if (-not $ps7Paths) {
            Write-Host "   PowerShell 7 not found, skipping configuration" -ForegroundColor Yellow
            return $false
        }
        
        $ps7Path = $ps7Paths[0]
        Write-Host "   Found PowerShell 7 at: $ps7Path" -ForegroundColor Yellow
        
        # Create PowerShell profile directory
        $profilePath = Split-Path $PROFILE.AllUsersAllHosts -Parent
        if (-not (Test-Path $profilePath)) {
            New-Item -ItemType Directory -Path $profilePath -Force | Out-Null
        }
        
        # Configure PowerShell profile for font and telemetry
        $profileContent = @"
# PowerShell 7 Configuration - Generated by Rocmine Installer
# Disable telemetry
`$env:POWERSHELL_TELEMETRY_OPTOUT = 1

# Set console font (for console applications)
if (`$Host.UI.RawUI.WindowTitle -notlike "*ISE*") {
    try {
        # Try to set console font
        Add-Type -TypeDefinition @"
            using System;
            using System.Runtime.InteropServices;
            public class ConsoleFont {
                [DllImport("kernel32.dll", SetLastError = true)]
                public static extern bool SetCurrentConsoleFontEx(IntPtr hConsoleOutput, bool bMaximumWindow, ref CONSOLE_FONT_INFOEX lpConsoleCurrentFontEx);
                
                [DllImport("kernel32.dll", SetLastError = true)]
                public static extern IntPtr GetStdHandle(int nStdHandle);
                
                [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
                public struct CONSOLE_FONT_INFOEX {
                    public uint cbSize;
                    public uint nFont;
                    public short dwFontSizeX;
                    public short dwFontSizeY;
                    public uint FontFamily;
                    public uint FontWeight;
                    [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
                    public string FaceName;
                }
            }
"@ -ErrorAction SilentlyContinue
    } catch {
        # Font setting failed, continue anyway
    }
}

# Oh My Posh configuration (if installed)
if (Get-Command oh-my-posh) {
        try {
        oh-my-posh init pwsh | Invoke-Expression
    } catch {
        Write-Warning "Failed to initialize oh-my-posh: $_"
    }
}

Write-Host "PowerShell 7 configured with CaskaydiaCove Nerd Font and telemetry disabled" -ForegroundColor Green
"@
        
        $profileFile = "$profilePath\Microsoft.PowerShell_profile.ps1"
        Set-Content -Path $profileFile -Value $profileContent -Encoding UTF8
        
        Write-Host "   ✓ PowerShell 7 profile configured" -ForegroundColor Green
        Write-Host "   ✓ Telemetry disabled" -ForegroundColor Green
        Write-Host "   Profile location: $profileFile" -ForegroundColor Gray
        
        return $true
    }
    catch {
        Write-Host "   ✗ Failed to configure PowerShell 7: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Disable-PowerShellTelemetry {
    Write-Host "`n🚫 Disabling PowerShell telemetry..." -ForegroundColor Cyan
    
    try {
        # Set environment variable for current session
        $env:POWERSHELL_TELEMETRY_OPTOUT = "1"
        
        # Set system-wide environment variable
        [Environment]::SetEnvironmentVariable("POWERSHELL_TELEMETRY_OPTOUT", "1", "Machine")
        
        # Set user environment variable as backup
        [Environment]::SetEnvironmentVariable("POWERSHELL_TELEMETRY_OPTOUT", "1", "User")
        
        Write-Host "   ✓ PowerShell telemetry disabled system-wide" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "   ✗ Failed to disable telemetry: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}
        
        # Cleanup
        Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        
        if ($installedCount -gt 0) {
            Write-Host "✓ CaskaydiaCove Nerd Font installed successfully! ($installedCount files)" -ForegroundColor Green
            return $true
        } else {
            Write-Host "⚠️  Font files may already be installed" -ForegroundColor Yellow
            return $true
        }
    }
    catch {
        Write-Host "✗ Failed to install CaskaydiaCove Nerd Font: $($_.Exception.Message)" -ForegroundColor Red
        # Cleanup on error
        if ($tempDir -and (Test-Path $tempDir)) {
            Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        }
        return $false
    }
}

function Show-PackageSelector {
    Write-ColoredBanner "Package Selection"
    
    Write-Host "`n📦 Select packages to install:" -ForegroundColor Cyan
    Write-Host "   Use SPACE to toggle, ENTER to continue, A to select all, N to select none" -ForegroundColor Gray
    Write-Host "   Use UP/DOWN arrows to navigate, Q to quit" -ForegroundColor Gray
    Write-Host ""
    
    $allPackages = @()
    foreach ($category in $PackageCategories.Keys | Sort-Object) {
        foreach ($package in $PackageCategories[$category]) {
            $allPackages += [PSCustomObject]@{
                Category = $category
                Name = $package.Name
                Display = $package.Display
                Selected = $package.Selected
            }
        }
    }
    
    if ($allPackages.Count -eq 0) {
        Write-Host "No packages available for selection." -ForegroundColor Red
        return @()
    }
    
    $currentSelection = 0
    $maxDisplay = [Math]::Min(15, $allPackages.Count)
    $scrollOffset = 0
    
    do {
        # Calculate scroll bounds
        if ($currentSelection -lt $scrollOffset) { 
            $scrollOffset = $currentSelection 
        }
        if ($currentSelection -ge ($scrollOffset + $maxDisplay)) { 
            $scrollOffset = $currentSelection - $maxDisplay + 1 
        }
        
        # Clear and redraw
        $startY = $Host.UI.RawUI.CursorPosition.Y
        
        # Display packages
        for ($i = 0; $i -lt $maxDisplay; $i++) {
            $packageIndex = $i + $scrollOffset
            if ($packageIndex -ge $allPackages.Count) { break }
            
            $package = $allPackages[$packageIndex]
            $prefix = if ($package.Selected) { "[✓]" } else { "[ ]" }
            $highlight = if ($packageIndex -eq $currentSelection) { "► " } else { "  " }
            
            $color = switch ($true) {
                ($packageIndex -eq $currentSelection) { 'Yellow' }
                ($package.Selected) { 'Green' }
                default { 'White' }
            }
            
            $line = "{0}{1} {2} ({3})" -f $highlight, $prefix, $package.Display, $package.Category
            Write-Host ("{0,-78}" -f $line) -ForegroundColor $color
        }
        
        # Show scroll indicator if needed
        if ($allPackages.Count -gt $maxDisplay) {
            $scrollPercent = [math]::Round(($scrollOffset / ($allPackages.Count - $maxDisplay)) * 100)
            Write-Host "[$scrollPercent% - Scroll: $($scrollOffset+1)-$([math]::Min($scrollOffset + $maxDisplay, $allPackages.Count)) of $($allPackages.Count)]" -ForegroundColor DarkGray
        }
        
        # Navigation info
        Write-Host "`n[$($currentSelection + 1)/$($allPackages.Count)] " -NoNewline -ForegroundColor DarkGray
        Write-Host "SPACE=Toggle | A=All | N=None | ENTER=Continue | Q=Quit" -ForegroundColor Gray
        Write-Host -NoNewline # Position cursor for input
        
        $key = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        
        # Move cursor back to start of display area
        $Host.UI.RawUI.CursorPosition = @{X=0; Y=$startY}
        
        switch ($key.VirtualKeyCode) {
            38 { # Up arrow
                $currentSelection = [Math]::Max(0, $currentSelection - 1)
            }
            40 { # Down arrow
                $currentSelection = [Math]::Min($allPackages.Count - 1, $currentSelection + 1)
            }
            32 { # Space - Toggle selection
                $allPackages[$currentSelection].Selected = !$allPackages[$currentSelection].Selected
            }
            65 { # A - Select all
                foreach ($package in $allPackages) { $package.Selected = $true }
            }
            78 { # N - Select none
                foreach ($package in $allPackages) { $package.Selected = $false }
            }
            81 { # Q - Quit
                Write-Host "`nInstallation cancelled by user." -ForegroundColor Yellow
                exit 0
            }
            13 { # Enter - Continue
                break
            }
        }
        
    } while ($true)
    
    # Return selected packages
    return ($allPackages | Where-Object { $_.Selected })
}

function Show-InstallationSummary {
    $duration = (Get-Date) - $Stats.StartTime
    $durationFormatted = "{0:mm}m {0:ss}s" -f $duration
    
    Write-ColoredBanner "Installation Summary"
    
    Write-Host "`n📊 Results:" -ForegroundColor Cyan
    Write-Host ("   ✓ Successful: {0}" -f $Stats.Successful) -ForegroundColor Green
    Write-Host ("   ✗ Failed:     {0}" -f $Stats.Failed) -ForegroundColor Red
    Write-Host ("   ⊘ Skipped:    {0}" -f $Stats.Skipped) -ForegroundColor Yellow
    Write-Host ("   📦 Total:     {0}" -f $Stats.Total) -ForegroundColor White
    Write-Host ("   ⏱️  Duration:  {0}" -f $durationFormatted) -ForegroundColor Blue
    
    $successRate = if ($Stats.Total -gt 0) { [math]::Round(($Stats.Successful / $Stats.Total) * 100, 1) } else { 0 }
    Write-Host ("`n🎯 Success Rate: {0}%" -f $successRate) -ForegroundColor $(if ($successRate -ge 80) { 'Green' } elseif ($successRate -ge 60) { 'Yellow' } else { 'Red' })
}

# Main execution
try {
    Write-ColoredBanner "Rocmine Program Installer"
    
    Write-Host "`n🔍 Checking system requirements..." -ForegroundColor Cyan
    
    # Check if running as administrator
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
    if (-not $isAdmin) {
        Write-Host "⚠️  Warning: Not running as Administrator. Some installations may fail." -ForegroundColor Yellow
        Write-Host "   Consider running as Administrator for better compatibility." -ForegroundColor Gray
    } else {
        Write-Host "✓ Running as Administrator" -ForegroundColor Green
    }
    
    if (-not (Test-WingetAvailability)) {
        Write-Host "`nPress any key to exit..." -ForegroundColor Gray
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        exit 1
    }
    
    # Show package selector
    $SelectedPackages = Show-PackageSelector
    $Stats.Total = $SelectedPackages.Count
    
    if ($Stats.Total -eq 0) {
        Write-Host "`n⚠️  No packages selected. Exiting..." -ForegroundColor Yellow
        Write-Host "`nPress any key to exit..." -ForegroundColor Gray
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        exit 0
    }
    
    Clear-Host
    Write-ColoredBanner "Installation Configuration"
    Write-Host "`n📋 Configuration:" -ForegroundColor Cyan
    Write-Host ("   • Selected packages: {0}" -f $Stats.Total) -ForegroundColor White
    Write-Host ("   • Continue on error: {0}" -f $ContinueOnError) -ForegroundColor White
    Write-Host ("   • Skip confirmation: {0}" -f $SkipConfirmation) -ForegroundColor White
    
    Write-Host "`n📦 Selected packages:" -ForegroundColor Cyan
    $SelectedPackages | Group-Object Category | ForEach-Object {
        Write-Host "   $($_.Name):" -ForegroundColor Yellow
        $_.Group | ForEach-Object { Write-Host "     • $($_.Display)" -ForegroundColor Gray }
    }
    
    if (-not $SkipConfirmation) {
        Write-Host "`n❓ " -NoNewline -ForegroundColor Yellow
        $confirmation = Read-Host "Proceed with installation? (Y/n)"
        if ($confirmation -match '^n(o)?$') {
            Write-Host "Installation cancelled by user." -ForegroundColor Yellow
            exit 0
        }
    }
    
    Write-Host "`n🚀 Starting installation process..." -ForegroundColor Green
    Start-Sleep -Seconds 2
    
    # Install CaskaydiaCove Nerd Font first
    $fontInstalled = Install-CaskaydiaCoveFont
    
    # Disable PowerShell telemetry
    $telemetryDisabled = Disable-PowerShellTelemetry
    
    $current = 0
    foreach ($package in $SelectedPackages) {
        $current++
        
        Write-ColoredBanner "Installing Applications ($current/$($Stats.Total))"
        Write-ProgressInfo -PackageName $package.Display -Current $current -Total $Stats.Total
        
        $exitCode = Install-Package -PackageName $package.Name
        
        switch ($exitCode) {
            0 {
                $Stats.Successful++
                Write-Host "`n✓ Successfully installed: $($package.Display)" -ForegroundColor Green
            }
            -1978335189 {
                $Stats.Skipped++
                Write-Host "`n⊘ Already installed: $($package.Display)" -ForegroundColor Yellow
            }
            default {
                $Stats.Failed++
                Write-Host "`n✗ Failed to install: $($package.Display) (Exit code: $exitCode)" -ForegroundColor Red
                
                if (-not $ContinueOnError) {
                    Write-Host "Stopping installation due to error. Use -ContinueOnError to skip failures." -ForegroundColor Yellow
                    break
                }
            }
        }
        
        Start-Sleep -Milliseconds 500
    }
    
    Show-InstallationSummary
    
    # Configure applications
    Write-Host "`n🔧 Configuring applications..." -ForegroundColor Cyan
    $terminalConfigured = Configure-WindowsTerminal
    $ps7Configured = Configure-PowerShell7
    
    Write-Host "`n💡 Post-installation summary:" -ForegroundColor Blue
    if ($fontInstalled) {
        Write-Host "   ✓ CaskaydiaCove Nerd Font installed" -ForegroundColor Green
    }
    if ($telemetryDisabled) {
        Write-Host "   ✓ PowerShell telemetry disabled" -ForegroundColor Green
    }
    if ($terminalConfigured) {
        Write-Host "   ✓ Windows Terminal configured with CaskaydiaCove font" -ForegroundColor Green
    }
    if ($ps7Configured) {
        Write-Host "   ✓ PowerShell 7 profile configured" -ForegroundColor Green
    }
    
    Write-Host "`n📝 Important notes:" -ForegroundColor Blue
    Write-Host "   • Restart Windows Terminal and PowerShell to see font changes" -ForegroundColor Gray
    Write-Host "   • PowerShell telemetry is now disabled system-wide" -ForegroundColor Gray
    Write-Host "   • Some applications may require a system restart" -ForegroundColor Gray
    if ($ps7Configured) {
        Write-Host "   • PowerShell 7 profile includes Oh My Posh integration" -ForegroundColor Gray
    }
    
    Write-Host "`n🎉 Installation process completed!" -ForegroundColor Green
    Write-Host "Thank you for using Rocmine Program Installer!" -ForegroundColor Magenta
}
catch {
    Write-Host "`n💥 Critical error occurred: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack trace: $($_.ScriptStackTrace)" -ForegroundColor DarkRed
    exit 1
}
finally {
    Write-Host "`nPress any key to exit..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
}
