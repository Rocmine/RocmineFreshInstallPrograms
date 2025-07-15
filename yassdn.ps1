# Rocmine's Program Installer - PowerShell Script
Clear-Host
$packages = @(
    "Valve.Steam",
    "Discord.Discord",
    "Mozilla.Firefox.Nightly",
    "RARLab.WinRAR",
    "Logitech.GHUB",
    "OBSProject.OBSStudio",
    "Microsoft.VisualStudioCode",
    "JanDeDobbeleer.OhMyPosh",
    "equalizerapo",
    "REALiX.HWiNFO"
)
$total = $packages.Count
$current = 0
function Write-Header {
    Clear-Host
    Write-Host ('=' * 50) -ForegroundColor Cyan
    Write-Host "        Rocmine's Program Installer" -ForegroundColor Magenta
    Write-Host ('=' * 50) -ForegroundColor Cyan
    Write-Host
}
Write-Header
foreach ($package in $packages) {
    $current++
    Write-Header
    Write-Host ("Installing package $package  [$current/$total]") -ForegroundColor Yellow
    Write-Host ""
    # Run winget install with accepted agreements and silent
    $installArgs = @(
        "install", $package,
        "--accept-source-agreements",
        "--accept-package-agreements",
        "--silent"
    )
    # Start winget process and wait
    $process = Start-Process -FilePath "winget" -ArgumentList $installArgs -NoNewWindow -Wait -PassThru
    if ($process.ExitCode -ne 0) {
        Write-Host ""
        Write-Host "Error installing $package. Aborting package, skipping..." -ForegroundColor Red
        Start-Sleep -Seconds 5
    } else {
        Write-Host ""
        Write-Host "Successfully installed $package." -ForegroundColor Green
        Start-Sleep -Seconds 2
    }
}
Write-Header
Write-Host "All installations completed successfully!" -ForegroundColor Green
Write-Host
Start-Sleep -Seconds 3
exit 0
