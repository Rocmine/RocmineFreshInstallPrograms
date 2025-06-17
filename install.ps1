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
        @{Name='equalizerapo'; Display='Equalizer APO'; Selected=$false}
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
    
    $border = '═' * 60
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
    
    $percentComplete = [math]::Round(($Current / $Total) * 100, 1)
    $progressBar = ('█' * [math]::Floor($percentComplete / 4)) + ('░' * (25 - [math]::Floor($percentComplete / 4)))
    
    Write-Host "`n┌─ Package Progress ─────────────────────────────────────┐" -ForegroundColor DarkGray
    Write-Host ("│ {0,-54} │" -f "$Status`: $PackageName") -ForegroundColor Yellow
    Write-Host ("│ Progress: [{0}] {1}% ({2}/{3})" -f $progressBar, $percentComplete, $Current, $Total) -ForegroundColor Cyan
    Write-Host "└────────────────────────────────────────────────────────┘" -ForegroundColor DarkGray
}

function Test-WingetAvailability {
    try {
        $null = Get-Command winget -ErrorAction Stop
        $wingetVersion = (winget --version).Trim('v')
        Write-Host "✓ Winget found (Version: $wingetVersion)" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "✗ Winget not found or not accessible" -ForegroundColor Red
        Write-Host "Please install Windows Package Manager from Microsoft Store" -ForegroundColor Yellow
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
        $process = Start-Process -FilePath 'winget' -ArgumentList $installArgs -NoNewWindow -Wait -PassThru -RedirectStandardError 'NUL'
        return $process.ExitCode
    }
    catch {
        Write-Host "✗ Exception occurred: $($_.Exception.Message)" -ForegroundColor Red
        return -1
    }
}

function Install-CascadiaCodeFont {
    Write-Host "`n🔤 Installing CascadiaCode Nerd Font..." -ForegroundColor Cyan
    
    try {
        # Download CascadiaCode Nerd Font
        $fontUrl = "https://github.com/ryanoasis/nerd-fonts/releases/download/v3.1.1/CascadiaCode.zip"
        $tempPath = "$env:TEMP\CascadiaCode.zip"
        $extractPath = "$env:TEMP\CascadiaCode"
        
        Write-Host "   Downloading font..." -ForegroundColor Yellow
        Invoke-WebRequest -Uri $fontUrl -OutFile $tempPath -UseBasicParsing
        
        Write-Host "   Extracting font files..." -ForegroundColor Yellow
        Expand-Archive -Path $tempPath -DestinationPath $extractPath -Force
        
        # Install fonts
        $shell = New-Object -ComObject Shell.Application
        $fontsFolder = $shell.Namespace(0x14)  # Special folder for Fonts
        
        $fontFiles = Get-ChildItem -Path $extractPath -Filter "*.ttf" | Where-Object { $_.Name -like "*Regular*" -or $_.Name -like "*Bold*" }
        
        foreach ($fontFile in $fontFiles) {
            Write-Host "   Installing: $($fontFile.Name)" -ForegroundColor Yellow
            $fontsFolder.CopyHere($fontFile.FullName, 0x10)
        }
        
        # Configure Windows Terminal to use the font
        $terminalSettingsPath = "$env:LOCALAPPDATA\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState\settings.json"
        
        if (Test-Path $terminalSettingsPath) {
            Write-Host "   Configuring Windows Terminal..." -ForegroundColor Yellow
            $settings = Get-Content $terminalSettingsPath -Raw | ConvertFrom-Json
            
            # Set the font for all profiles
            if ($settings.profiles -and $settings.profiles.defaults) {
                $settings.profiles.defaults | Add-Member -NotePropertyName "font" -NotePropertyValue @{
                    face = "CaskaydiaCove Nerd Font"
                    size = 11
                } -Force
            }
            
            $settings | ConvertTo-Json -Depth 10 | Set-Content $terminalSettingsPath
        }
        
        # Cleanup
        Remove-Item $tempPath -Force -ErrorAction SilentlyContinue
        Remove-Item $extractPath -Recurse -Force -ErrorAction SilentlyContinue
        
        Write-Host "✓ CascadiaCode Nerd Font installed successfully!" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "✗ Failed to install CascadiaCode Nerd Font: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Show-PackageSelector {
    Write-ColoredBanner "Package Selection"
    
    Write-Host "`n📦 Select packages to install:" -ForegroundColor Cyan
    Write-Host "   Use SPACE to toggle, ENTER to continue, A to select all, N to select none" -ForegroundColor Gray
    Write-Host ""
    
    $allPackages = @()
    foreach ($category in $PackageCategories.Keys) {
        foreach ($package in $PackageCategories[$category]) {
            $allPackages += [PSCustomObject]@{
                Category = $category
                Name = $package.Name
                Display = $package.Display
                Selected = $package.Selected
            }
        }
    }
    
    $currentSelection = 0
    $maxDisplay = [Math]::Min(20, $allPackages.Count)
    $scrollOffset = 0
    
    do {
        # Clear the selection area
        $currentPos = $Host.UI.RawUI.CursorPosition
        $Host.UI.RawUI.CursorPosition = @{X=0; Y=$currentPos.Y}
        
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
            
            $line = "{0}{1} {2} - {3}" -f $highlight, $prefix, $package.Display, $package.Category
            Write-Host ("{0,-80}" -f $line) -ForegroundColor $color
        }
        
        # Show navigation info
        Write-Host "`n[$($currentSelection + 1)/$($allPackages.Count)] " -NoNewline -ForegroundColor DarkGray
        Write-Host "SPACE=Toggle, A=All, N=None, ENTER=Continue, Q=Quit" -ForegroundColor Gray
        
        $key = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        
        switch ($key.VirtualKeyCode) {
            38 { # Up arrow
                $currentSelection = [Math]::Max(0, $currentSelection - 1)
                if ($currentSelection -lt $scrollOffset) { $scrollOffset = [Math]::Max(0, $scrollOffset - 1) }
            }
            40 { # Down arrow
                $currentSelection = [Math]::Min($allPackages.Count - 1, $currentSelection + 1)
                if ($currentSelection -ge ($scrollOffset + $maxDisplay)) { $scrollOffset = $currentSelection - $maxDisplay + 1 }
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
        
        # Move cursor back up to redraw
        $Host.UI.RawUI.CursorPosition = @{X=0; Y=$currentPos.Y}
        
    } while ($true)
    
    # Return selected packages
    return $allPackages | Where-Object { $_.Selected }
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
    Write-ColoredBanner "Rocmine's Automated Program Installer"
    
    Write-Host "`n🔍 Checking system requirements..." -ForegroundColor Cyan
    if (-not (Test-WingetAvailability)) {
        exit 1
    }
    
    # Show package selector
    $SelectedPackages = Show-PackageSelector
    $Stats.Total = $SelectedPackages.Count
    
    if ($Stats.Total -eq 0) {
        Write-Host "`n⚠️  No packages selected. Exiting..." -ForegroundColor Yellow
        exit 0
    }
    
    Write-Host "`n📋 Configuration:" -ForegroundColor Cyan
    Write-Host ("   • Selected packages: {0}" -f $Stats.Total) -ForegroundColor White
    Write-Host ("   • Continue on error: {0}" -f $ContinueOnError) -ForegroundColor White
    
    if (-not $SkipConfirmation) {
        Write-Host "`n❓ " -NoNewline -ForegroundColor Yellow
        $confirmation = Read-Host "Proceed with installation? (Y/n)"
        if ($confirmation -match '^n(o)?
        
        if ($exitCode -eq 0) {
            $Stats.Successful++
            Write-Host "`n✓ Successfully installed: $($package.Display)" -ForegroundColor Green
        }
        elseif ($exitCode -eq -1978335189) {
            $Stats.Skipped++
            Write-Host "`n⊘ Already installed: $($package.Display)" -ForegroundColor Yellow
        }
        else {
            $Stats.Failed++
            Write-Host "`n✗ Failed to install: $($package.Display) (Exit code: $exitCode)" -ForegroundColor Red
            
            if (-not $ContinueOnError) {
                Write-Host "Stopping installation due to error. Use -ContinueOnError to skip failures." -ForegroundColor Yellow
                break
            }
        }
        
        Start-Sleep -Milliseconds 500
    }
    
    Show-InstallationSummary
    
    if ($fontInstalled) {
        Write-Host "`n💡 Tip: Restart Windows Terminal to see the new CascadiaCode Nerd Font!" -ForegroundColor Blue
    }
    
    Write-Host "`n🎉 Installation process completed!" -ForegroundColor Green
    Write-Host "Thank you for using Rocmine's Program Installer!" -ForegroundColor Magenta
}
catch {
    Write-Host "`n💥 Critical error occurred: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
finally {
    Write-Host "`nPress any key to exit..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
}) {
            Write-Host "Installation cancelled by user." -ForegroundColor Yellow
            exit 0
        }
    }
    
    Write-Host "`n🚀 Starting installation process..." -ForegroundColor Green
    Start-Sleep -Seconds 2
    
    # Install CascadiaCode Nerd Font first
    $fontInstalled = Install-CascadiaCodeFont
    
    $current = 0
    foreach ($package in $SelectedPackages) {
        $current++
        
        Write-ColoredBanner "Installing Applications ($current/$($Stats.Total))"
        Write-ProgressInfo -PackageName $package.Display -Current $current -Total $Stats.Total
        
        $exitCode = Install-Package -PackageName $package.Name
        
        if ($exitCode -eq 0) {
            $Stats.Successful++
            Write-Host "`n✓ Successfully installed: $package" -ForegroundColor Green
        }
        elseif ($exitCode -eq -1978335189) {
            $Stats.Skipped++
            Write-Host "`n⊘ Already installed: $package" -ForegroundColor Yellow
        }
        else {
            $Stats.Failed++
            Write-Host "`n✗ Failed to install: $package (Exit code: $exitCode)" -ForegroundColor Red
            
            if (-not $ContinueOnError) {
                Write-Host "Stopping installation due to error. Use -ContinueOnError to skip failures." -ForegroundColor Yellow
                break
            }
        }
        
        Start-Sleep -Milliseconds 500
    }
    
    Show-InstallationSummary
    
    Write-Host "`n🎉 Installation process completed!" -ForegroundColor Green
    Write-Host "Thank you for using Rocmine's Program Installer!" -ForegroundColor Magenta
}
catch {
    Write-Host "`n💥 Critical error occurred: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
finally {
    Write-Host "`nPress any key to exit..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
}
