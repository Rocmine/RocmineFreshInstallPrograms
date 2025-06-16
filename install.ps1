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
    [switch]$ContinueOnError,
    [string[]]$ExcludePackages = @()
)

# Script configuration
$ErrorActionPreference = 'Continue'
$ProgressPreference = 'SilentlyContinue'

# Package definitions with categories for better organization
$PackageCategories = @{
    'Development' = @(
        'Microsoft.VisualStudioCode',
        'Git.Git',
        'OpenJS.NodeJS',
        'Yarn.Yarn',
        'Python.Python.3.12',
        'Docker.DockerDesktop',
        'Postman.Postman',
        'CoreyButler.NVMforWindows'
    )
    'System Utilities' = @(
        'Microsoft.PowerShell',
        'Microsoft.WindowsTerminal',
        'Microsoft.PowerToys',
        'JanDeDobbeleer.OhMyPosh',
        'Fastfetch-cli.Fastfetch',
        'REALiX.HWiNFO',
        'Rufus.Rufus',
        'WiresharkFoundation.Wireshark'
    )
    'Media & Entertainment' = @(
        'OBSProject.OBSStudio',
        'Audacity.Audacity',
        'Valve.Steam',
        'EpicGames.EpicGamesLauncher',
        'Discord.Discord',
        'Stremio.Stremio',
        'qBittorrent.qBittorrent',
        'yt-dlp.yt-dlp'
    )
    'Productivity' = @(
        '7zip.7zip',
        'TheDocumentFoundation.LibreOffice',
        'Foxit.FoxitReader',
        'PuTTY.PuTTY',
        'Parsec.Parsec'
    )
    'System Components' = @(
        'Microsoft.VCRedist.2015+.x86',
        'Microsoft.VCRedist.2015+.x64',
        'equalizerapo'
    )
    'Network & Security' = @(
        'angryziber.AngryIPScanner',
        'ChatterinoTeam.Chatterino'
    )
    'Graphics & Drivers' = @(
        'Wagnardsoft.DisplayDriverUninstaller',
        'TechPowerUp.NVCleanstall'
    )
}

# Flatten package list for processing
$AllPackages = $PackageCategories.Values | ForEach-Object { $_ } | Where-Object { $_ -notin $ExcludePackages }

# Statistics tracking
$Stats = @{
    Total = $AllPackages.Count
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
    
    Write-Host "`n📋 Configuration:" -ForegroundColor Cyan
    Write-Host ("   • Total packages: {0}" -f $Stats.Total) -ForegroundColor White
    Write-Host ("   • Excluded packages: {0}" -f ($ExcludePackages.Count)) -ForegroundColor White
    Write-Host ("   • Continue on error: {0}" -f $ContinueOnError) -ForegroundColor White
    
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
    
    $current = 0
    foreach ($package in $AllPackages) {
        $current++
        
        Write-ColoredBanner "Installing Applications ($current/$($Stats.Total))"
        Write-ProgressInfo -PackageName $package -Current $current -Total $Stats.Total
        
        $exitCode = Install-Package -PackageName $package
        
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
