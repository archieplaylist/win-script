<#
.SYNOPSIS
    A native PowerShell script that acts like a driver update utility.
.DESCRIPTION
    This script utilizes the Windows Update COM object to scan the Microsoft Update 
    Catalog for missing or outdated hardware drivers, downloads them, and installs them.
    It does not require any third-party software or modules.
#>

# 1. Check for Administrator Privileges
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "Requesting Administrator privileges..." -ForegroundColor Yellow
    # Relaunch the script with Administrator privileges
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    Exit
}

Clear-Host
Write-Host "===================================================" -ForegroundColor Cyan
Write-Host "       PowerShell Automated Driver Updater         " -ForegroundColor Cyan
Write-Host "===================================================" -ForegroundColor Cyan
Write-Host ""

# 2. Register Microsoft Update as the source (to get full OEM drivers, not just basic Windows ones)
Write-Host "[*] Connecting to Microsoft Update Catalog..." -ForegroundColor Yellow
$UpdateSvc = New-Object -ComObject Microsoft.Update.ServiceManager
$UpdateSvc.AddService2("7971f918-a847-4430-9279-4a52d1efe18d", 7, "") | Out-Null

$Session = New-Object -ComObject Microsoft.Update.Session
$Searcher = $Session.CreateUpdateSearcher()
$Searcher.ServiceID = '7971f918-a847-4430-9279-4a52d1efe18d'
$Searcher.SearchScope = 1
$Searcher.ServerSelection = 3 

# 3. Scan for missing or outdated drivers
Write-Host "[*] Scanning hardware for missing or outdated drivers. This may take a few minutes..." -ForegroundColor Yellow
$Criteria = "IsInstalled=0 and Type='Driver' and IsHidden=0"
$SearchResult = $Searcher.Search($Criteria)
$Updates = $SearchResult.Updates

if ($Updates.Count -eq 0) {
    Write-Host "[+] Your system is fully up to date! No missing drivers found." -ForegroundColor Green
    Write-Host ""
    Pause
    Exit
}

# 4. List the drivers found
Write-Host ""
Write-Host "Found $($Updates.Count) driver update(s):" -ForegroundColor Green
foreach ($Update in $Updates) {
    Write-Host "  -> $($Update.Title)" -ForegroundColor White
}
Write-Host ""

# 5. Prepare to Download
$UpdatesToDownload = New-Object -ComObject Microsoft.Update.UpdateColl
foreach ($Update in $Updates) {
    $UpdatesToDownload.Add($Update) | Out-Null
}

Write-Host "[*] Downloading Drivers..." -ForegroundColor Yellow
$Downloader = $Session.CreateUpdateDownloader()
$Downloader.Updates = $UpdatesToDownload
$Downloader.Download() | Out-Null
Write-Host "[+] Download Complete." -ForegroundColor Green

# 6. Filter for successfully downloaded drivers to install
$UpdatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl
foreach ($Update in $Updates) {
    if ($Update.IsDownloaded) {
        $UpdatesToInstall.Add($Update) | Out-Null
    }
}

# 7. Install the Drivers
if ($UpdatesToInstall.Count -gt 0) {
    Write-Host "[*] Installing Drivers..." -ForegroundColor Yellow
    $Installer = $Session.CreateUpdateInstaller()
    $Installer.Updates = $UpdatesToInstall
    
    # Trigger the installation
    $InstallationResult = $Installer.Install()
    
    Write-Host "[+] Installation Process Finished." -ForegroundColor Green
    
    # 8. Check if a reboot is required
    if ($InstallationResult.RebootRequired) {
        Write-Host "===================================================" -ForegroundColor Red
        Write-Host "[!] REBOOT REQUIRED: Please restart your computer to apply the new drivers." -ForegroundColor Red
        Write-Host "===================================================" -ForegroundColor Red
    } else {
        Write-Host "[+] All drivers installed successfully. No reboot required." -ForegroundColor Green
    }
} else {
    Write-Host "[-] Could not verify downloaded drivers. Installation aborted." -ForegroundColor Red
}

Write-Host ""
Write-Host "Press any key to exit..." -ForegroundColor Cyan
$host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null