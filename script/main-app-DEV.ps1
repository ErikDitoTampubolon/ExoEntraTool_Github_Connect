# =========================================================================
# AUTHOR   : Erik Dito Tampubolon - TelkomSigma
# VERSION  : 2.0 (dengan Auto-Update Terintegrasi)
# DESKRIPSI: ExoEntraTool dengan Auto-Update dari GitHub
# =========================================================================

# ==========================================================================
# AUTO-UPDATE SYSTEM - SELALU MENDAPATKAN VERSI TERBARU
# ==========================================================================

# 1. Tentukan Path Dasar
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }

# Konfigurasi URL GitHub
$GitHubMainScript = "https://raw.githubusercontent.com/ErikDitoTampubolon/ExoEntraTool_Github_Connect/dev/script/main-app-DEV.ps1"
$LocalMainScript = Join-Path -Path $scriptDir -ChildPath "main-app-DEV-latest.ps1"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  ExoEntraTool - Auto Update Check" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Fungsi untuk mengecek dan update script
function Check-AndUpdate {
    Write-Host "[*] Memeriksa pembaruan dari GitHub..." -ForegroundColor Yellow
    
    try {
        # Tentukan path script saat ini (work untuk .ps1 dan .exe)
        $currentScript = if ($PSCommandPath) { 
            $PSCommandPath 
        } elseif ($MyInvocation.MyCommand.Path) { 
            $MyInvocation.MyCommand.Path 
        } else {
            Join-Path -Path $scriptDir -ChildPath "main-app-DEV.ps1"
        }
        
        # Skip update jika dijalankan sebagai EXE
        if ($currentScript -match '\.exe$') {
            Write-Host "[*] Berjalan sebagai EXE. Auto-update dinonaktifkan." -ForegroundColor Gray
            Write-Host "[*] Untuk update, jalankan ulang ExoEntraTool_Updater.ps1" -ForegroundColor Cyan
            Start-Sleep -Seconds 2
            return
        }
        
        # Download versi terbaru
        $ProgressPreference = 'SilentlyContinue'
        Invoke-WebRequest -Uri $GitHubMainScript -OutFile $LocalMainScript -UseBasicParsing -ErrorAction Stop
        $ProgressPreference = 'Continue'
        
        if (Test-Path $LocalMainScript) {
            $latestContent = Get-Content $LocalMainScript -Raw -ErrorAction SilentlyContinue
            $currentContent = Get-Content $currentScript -Raw -ErrorAction SilentlyContinue
            
            if ($latestContent -and $currentContent -and ($latestContent -ne $currentContent)) {
                Write-Host "[!] Versi baru tersedia! Memperbarui..." -ForegroundColor Green
                
                # Backup versi lama
                $backupPath = "$currentScript.backup"
                Copy-Item $currentScript -Destination $backupPath -Force -ErrorAction SilentlyContinue
                
                # Update ke versi terbaru
                Copy-Item $LocalMainScript -Destination $currentScript -Force
                Remove-Item $LocalMainScript -Force
                
                Write-Host "[OK] Script berhasil diperbarui!" -ForegroundColor Green
                Write-Host "[*] Memulai ulang dengan versi terbaru..." -ForegroundColor Cyan
                Start-Sleep -Seconds 2
                
                # Jalankan ulang script yang sudah diupdate
                & $currentScript
                exit
            } else {
                Write-Host "[OK] Anda sudah menggunakan versi terbaru." -ForegroundColor Green
                Remove-Item $LocalMainScript -Force -ErrorAction SilentlyContinue
            }
        }
    } catch {
        Write-Host "[!] Tidak dapat memeriksa update: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "[*] Melanjutkan dengan versi lokal..." -ForegroundColor Gray
    }
    
    Write-Host ""
}

# Jalankan auto-update check
Check-AndUpdate

# ==========================================================================
# INFRASTRUCTURE & REPOSITORY SYNC SYSTEM
# ==========================================================================

# 2. Daftar Folder yang Akan Dibuat
$folders = @(
    "exported_data",
    "script",
    "script/exchange_online",
    "script/entra"
)

Write-Host "--- 1. Menyiapkan Struktur Direktori ---" -ForegroundColor Blue
foreach ($folder in $folders) {
    $targetPath = Join-Path -Path $scriptDir -ChildPath $folder
    if (-not (Test-Path -Path $targetPath)) {
        New-Item -Path $targetPath -ItemType Directory | Out-Null
        Write-Host "[OK] Folder dibuat: $folder" -ForegroundColor Green
    } else {
        Write-Host "[Selesai] Folder sudah ada: $folder" -ForegroundColor Gray
    }
}

# 3. Fungsi untuk Download File dari GitHub
function Sync-GitHubRepo {
    param (
        [string]$RepoPath,
        [string]$GitHubUrl
    )

    Write-Host "`n[*] Sinkronisasi file ke folder: $RepoPath" -ForegroundColor Cyan
    
    $apiUrl = $GitHubUrl -replace "github.com", "api.github.com/repos" -replace "tree/main", "contents"

    try {
        $files = Invoke-RestMethod -Uri $apiUrl -Method Get -ErrorAction Stop
        foreach ($file in $files) {
            if ($file.type -eq "file") {
                $destination = Join-Path -Path (Join-Path $scriptDir $RepoPath) -ChildPath $file.name
                
                Write-Host " -> Mendownload: $($file.name) . . ." -ForegroundColor White -NoNewline
                Invoke-WebRequest -Uri $file.download_url -OutFile $destination
                Write-Host " [BERHASIL]" -ForegroundColor Green
            }
        }
    } catch {
        Write-Host " [GAGAL] Tidak dapat mengakses repositori: $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host "`n--- Infrastruktur Siap. Melanjutkan ke Logika Utama ---" -ForegroundColor Blue

# --- Memeriksa Lingkungan PowerShell ---
Set-ExecutionPolicy RemoteSigned -Scope Process -Force -ErrorAction SilentlyContinue

function Check-Module {
    param($ModuleName)
    Write-Host "Memeriksa Modul '$ModuleName'" -ForegroundColor Cyan
    if (Get-Module -Name $ModuleName -ListAvailable) {
        Write-Host "Terinstal." -ForegroundColor Green
    } else {
        Write-Host "Belum ada. Menginstal..." -ForegroundColor Yellow
        Install-Module $ModuleName -Force -AllowClobber -Scope CurrentUser -ErrorAction SilentlyContinue
    }
}

Write-Host "`n--- Prasyarat Modul ---" -ForegroundColor Blue
Check-Module -ModuleName "PowerShellGet"
Check-Module -ModuleName "ExchangeOnlineManagement"
Check-Module -ModuleName "Microsoft.Graph"
Check-Module -ModuleName "Microsoft.Entra"
Check-Module -ModuleName "Microsoft.Entra.Beta"

# --- Membangun Koneksi Multi-Service ---
$requiredScopes = "User.ReadWrite.All", "Organization.Read.All"
Write-Host "`n--- Membangun Koneksi ke Microsoft Graph ---" -ForegroundColor Blue

if (Get-MgContext -ErrorAction SilentlyContinue) {
    Write-Host "Sesi Graph yang ada akan diputus untuk koneksi ulang dengan scopes baru." -ForegroundColor DarkYellow
    Disconnect-MgGraph
}

Write-Host "Anda akan diminta untuk login menggunakan akun Global Administrator." -ForegroundColor Cyan
Write-Host "Menghubungkan ke Microsoft Graph" -ForegroundColor Yellow

try {
    Connect-MgGraph -Scopes $requiredScopes -ErrorAction Stop | Out-Null
    Write-Host "Koneksi ke Microsoft Graph berhasil" -ForegroundColor Green
} catch {
    Write-Error "Gagal terhubung ke Microsoft Graph."
    exit 1
}

# 2.2 KONEKSI ENTRA
Write-Host "Menghubungkan ke Microsoft Entra" -ForegroundColor Yellow
try {
    Connect-Entra -scope 'User.Read.All', 'UserAuthenticationMethod.Read.All' -ErrorAction Stop
    Write-Host "Koneksi Microsoft Entra Berhasil." -ForegroundColor Green
} catch {
    Write-Host "Peringatan: Gagal terkoneksi ke Entra." -ForegroundColor Yellow
}

# 2.3 KONEKSI EXCHANGE ONLINE
Write-Host "Menghubungkan ke Exchange Online" -ForegroundColor Yellow

$existingSession = Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" }

if (-not $existingSession) {
    try {
        Connect-ExchangeOnline -ShowProgress $true -ErrorAction Stop
        Write-Host "Koneksi ke Exchange Online berhasil" -ForegroundColor Green
    }
    catch {
        Write-Host "`nGagal terhubung ke Exchange Online!" -ForegroundColor Red
        Write-Host "Detail Error: $($_.Exception.Message)" -ForegroundColor Yellow
        exit 1
    }
}

Write-Host "`nSemua layanan terhubung. Memuat antarmuka..." -ForegroundColor Green
Start-Sleep -Seconds 1

# --- FUNGSI DOWNLOAD ON DEMAND ---
function Get-AndExecute {
    param (
        [string]$SubFolder,
        [string]$FileName
    )

    $localPath = Join-Path $scriptDir "script\$SubFolder\$FileName"
    $githubUrl = "https://raw.githubusercontent.com/ErikDitoTampubolon/ExchangeOnlineTools-ErikDito/main/script/$SubFolder/$FileName"

    if (-not (Test-Path $localPath)) {
        Write-Host "`n[!] Download Script." -ForegroundColor Cyan
        try {
            $destFolder = Split-Path $localPath
            if (-not (Test-Path $destFolder)) { New-Item -Path $destFolder -ItemType Directory | Out-Null }
            
            Invoke-WebRequest -Uri $githubUrl -OutFile $localPath -ErrorAction Stop
            Write-Host "[OK] Download Berhasil." -ForegroundColor Green
        } catch {
            Write-Host "[GAGAL] Tidak dapat mendownload file: $($_.Exception.Message)" -ForegroundColor Red
            Pause
            return
        }
    }

    & $localPath
    Pause
}

## -----------------------------------------------------------------------
## FUNGSI HEADER DENGAN ASCII ART
## -----------------------------------------------------------------------

function Show-Header {
    Clear-Host
    $ascii = @'
  _____              ______       _               _______           _ 
 |  ___|            |  ____|     | |             |__   __|         | | 
 | |__  __  _____   | |__   _ __ | |_ _ __ __ _     | | ___   ___  | |
 |  __| \ \/ / _ \  |  __| | '_ \| __| '__/ _` |    | |/ _ \ / _ \ | |
 | |___  >  < (_) | | |____| | | | |_| | | (_| |    | | (_) | (_) || |
 \____/ /_/\_\___/  |______|_| |_|\__|_|  \__,_|    |_|\___/ \___/ |_|
'@

    Write-Host $ascii -ForegroundColor Cyan
    Write-Host "======================================================================" -ForegroundColor DarkCyan
    Write-Host " Author   : Erik Dito Tampubolon - TelkomSigma" -ForegroundColor White
    Write-Host " Version  : 2.0 (ExoEntraTool Suite - Auto-Update Enabled)" -ForegroundColor White
    Write-Host "----------------------------------------------------------------------" -ForegroundColor DarkCyan
    Write-Host " Location : ${scriptDir}" -ForegroundColor Gray
    Write-Host " Time     : $(Get-Date -Format 'dd-MM-yyyy HH:mm:ss')" -ForegroundColor Gray  
    Write-Host "======================================================================" -ForegroundColor DarkCyan
    Write-Host ""
}

## -----------------------------------------------------------------------
## LOGIKA EKSEKUSI LOOP
## -----------------------------------------------------------------------

$mainRunning = $true
while ($mainRunning) {
    Show-Header
    Write-Host "Menu Utama:" -ForegroundColor Yellow
    Write-Host "  1. Microsoft Exchange Online"
    Write-Host "  2. Microsoft Entra ID"
    Write-Host ""
    Write-Host "  10. Keluar & Putus Koneksi" -ForegroundColor Red
    Write-Host "======================================================================" -ForegroundColor DarkCyan
    
    $mainChoice = Read-Host "Pilih nomor menu"

    switch ($mainChoice) {
        "1" { 
            $subRunning = $true
            while ($subRunning) {
                Show-Header
                Write-Host "Sub-Menu: Microsoft Exchange Online" -ForegroundColor Yellow
                Write-Host "  1. Assign or Remove License User Testing 2"
                Write-Host "  2. Export List License Availability"
                Write-Host "  3. Export List All Mailbox"
                Write-Host "  4. Export List All Active User"
                Write-Host "  5. Export List Active User (DisplayName,UPN,Contact)"
                Write-Host "  6. Export List Active User Last Password Changes"
                Write-Host "  7. Export List Mailbox Storage Usage"
                Write-Host "  8. Export List Mailbox Last Logon"
                Write-Host "  9. Export List Transport Rules"
                Write-Host "  10. Export List OneDrive Usage"
                Write-Host "  11. Export List Spam Email (10 days)"
                Write-Host "  12. Export List Full Access"
                Write-Host "  13. Export List Group Mailbox"
                Write-Host "  14. Export List OneDrive Shared Items"
                Write-Host "  15. Export List OneDrive Folder and File"
                Write-Host ""
                Write-Host "  B. Kembali ke Menu Utama" -ForegroundColor Yellow
                Write-Host "---------------------------------------------" -ForegroundColor DarkCyan
                
                $subChoice = Read-Host "Pilih nomor menu"
                if ($subChoice.ToUpper() -eq "B") { $subRunning = $false }
                elseif ($subChoice -eq "1") { Get-AndExecute -SubFolder "exchange_online" -FileName "assign-or-remove-license-user-final.ps1" }
                elseif ($subChoice -eq "2") { Get-AndExecute -SubFolder "exchange_online" -FileName "check-license-name-and-quota-final.ps1" }
                elseif ($subChoice -eq "3") { Get-AndExecute -SubFolder "exchange_online" -FileName "export-mailbox-final.ps1" }
                elseif ($subChoice -eq "4") { Get-AndExecute -SubFolder "exchange_online" -FileName "export-active-users-final.ps1" }
                elseif ($subChoice -eq "5") { Get-AndExecute -SubFolder "exchange_online" -FileName "export-alluser-userprincipalname-contact-final.ps1" }
                elseif ($subChoice -eq "6") { Get-AndExecute -SubFolder "exchange_online" -FileName "check-lastpasswordchange-user-final.ps1" }
                elseif ($subChoice -eq "7") { Get-AndExecute -SubFolder "exchange_online" -FileName "check-mailbox-storage-user-final.ps1" }
                elseif ($subChoice -eq "8") { Get-AndExecute -SubFolder "exchange_online" -FileName "check-lastlogon-user-final.ps1" }
                elseif ($subChoice -eq "9") { Get-AndExecute -SubFolder "exchange_online" -FileName "check-transport-rules-final.ps1" }
                elseif ($subChoice -eq "10") { Get-AndExecute -SubFolder "exchange_online" -FileName "export-onedrive-user-final.ps1" }
                elseif ($subChoice -eq "11") { Get-AndExecute -SubFolder "exchange_online" -FileName "export-spam-email-final.ps1" }
                elseif ($subChoice -eq "12") { Get-AndExecute -SubFolder "exchange_online" -FileName "export-full-access-final.ps1" }
                elseif ($subChoice -eq "13") { Get-AndExecute -SubFolder "exchange_online" -FileName "export-group-mailbox-final.ps1" }
                elseif ($subChoice -eq "14") { Get-AndExecute -SubFolder "exchange_online" -FileName "export-onedrive-shared-items-final.ps1" }
                elseif ($subChoice -eq "15") { Get-AndExecute -SubFolder "exchange_online" -FileName "export-onedrive-folder-file-user.ps1" }
            }
        }
        "2" { 
            $subRunning = $true
            while ($subRunning) {
                Show-Header
                Write-Host "Sub-Menu: Microsoft Entra ID" -ForegroundColor Yellow
                Write-Host "  1. Enable or Disable MFA User by .csv"
                Write-Host "  2. Force Change Password User by .csv"
                Write-Host "  3. Export List All User MFA Status"
                Write-Host "  4. Export List All Device"
                Write-Host "  5. Export List All User Owned Device"
                Write-Host "  6. Export List All Application"
                Write-Host "  7. Export List All Deleted User"
                Write-Host "  8. Export List All Inactive User (30 days)"
                Write-Host "  9. Export List Duplicate Device"
                Write-Host "  10. Export List Conditional Access Policy"
                Write-Host "  11. Export List User Auth Method"
                Write-Host "  12. Export List Permission Grant Policy"
                Write-Host "  13. Export List Entra Auth Policy"
                Write-Host "  14. Export List Entra Identity Provider"
                Write-Host "  15. Export List All Group"
                Write-Host "  16. Export List All Domain"
                Write-Host "  17. Export List Mail Traffic ATP"
                Write-Host ""
                Write-Host "  B. Kembali ke Menu Utama" -ForegroundColor Yellow
                Write-Host "---------------------------------------------" -ForegroundColor DarkCyan
                
                $subChoice = Read-Host "Pilih nomor menu"
                if ($subChoice.ToUpper() -eq "B") { $subRunning = $false }
                elseif ($subChoice -eq "1") { Get-AndExecute -SubFolder "entra" -FileName "enable-or-disable-mfa-user-by-csv-final.ps1" }
                elseif ($subChoice -eq "2") { Get-AndExecute -SubFolder "entra" -FileName "force-password-change-alluser-by-csv-final.ps1" }
                elseif ($subChoice -eq "3") { Get-AndExecute -SubFolder "entra" -FileName "export-alluser-mfa-final.ps1" }
                elseif ($subChoice -eq "4") { Get-AndExecute -SubFolder "entra" -FileName "export-alldevice-final.ps1" }
                elseif ($subChoice -eq "5") { Get-AndExecute -SubFolder "entra" -FileName "export-alluser-owned-device-final.ps1" }
                elseif ($subChoice -eq "6") { Get-AndExecute -SubFolder "entra" -FileName "export-allapplication-final.ps1" }
                elseif ($subChoice -eq "7") { Get-AndExecute -SubFolder "entra" -FileName "export-alldeleted-user-final.ps1" }
                elseif ($subChoice -eq "8") { Get-AndExecute -SubFolder "entra" -FileName "export-alluser-inactive-30days-final.ps1" }
                elseif ($subChoice -eq "9") { Get-AndExecute -SubFolder "entra" -FileName "export-list-alldevice-final.ps1" }
                elseif ($subChoice -eq "10") { Get-AndExecute -SubFolder "entra" -FileName "check-conditional-access-policy-final.ps1" }
                elseif ($subChoice -eq "11") { Get-AndExecute -SubFolder "entra" -FileName "check-user-auth-method-final.ps1" }
                elseif ($subChoice -eq "12") { Get-AndExecute -SubFolder "entra" -FileName "check-permission-grant-policy-final.ps1" }
                elseif ($subChoice -eq "13") { Get-AndExecute -SubFolder "entra" -FileName "check-entra-auth-policy-final.ps1" }
                elseif ($subChoice -eq "14") { Get-AndExecute -SubFolder "entra" -FileName "check-entra-identity-provider-final.ps1" }
                elseif ($subChoice -eq "15") { Get-AndExecute -SubFolder "entra" -FileName "export-allgroup-final.ps1" }
                elseif ($subChoice -eq "16") { Get-AndExecute -SubFolder "entra" -FileName "export-domain-final.ps1" }
                elseif ($subChoice -eq "17") { Get-AndExecute -SubFolder "entra" -FileName "export-mailtraffic-atp-report-final.ps1" }
            }
        }
        "10" {
            Write-Host "`nClosing sessions..." -ForegroundColor Cyan
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            Disconnect-Entra -ErrorAction SilentlyContinue
            Disconnect-MgGraph -ErrorAction SilentlyContinue
            $mainRunning = $false
        }
        default { 
            Write-Host "Pilihan tidak valid!" -ForegroundColor Red
            Start-Sleep -Seconds 1 
        }
    }
}