# =========================================================================  
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.1)  
# Deskripsi: Inventarisasi Seluruh Nama Folder & File OneDrive (Rekursif)
# =========================================================================  

# Variabel Global dan Output  
$scriptName = "OneDrive_Complete_Inventory" 
$scriptOutput = @() 

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }  
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"  
$outputFileName = "Output_$($scriptName)_$($timestamp).csv"  

# Definisi parentDir (2 tingkat di atas)  
$parentDir = (Get-Item $scriptDir).Parent.Parent.FullName  


## ==========================================================================  
#                           INFORMASI SCRIPT                  
## ==========================================================================  

Write-Host "`n================================================" -ForegroundColor Yellow  
Write-Host "                INFORMASI SCRIPT                " -ForegroundColor Yellow  
Write-Host "================================================" -ForegroundColor Yellow  
Write-Host " Nama Skrip        : $scriptName.ps1" -ForegroundColor Yellow  
Write-Host " Field Kolom       : Owner, ParentPath, ItemName, Type, Size_MB" -ForegroundColor Yellow  
Write-Host " Deskripsi Singkat : Mengekspor daftar lengkap seluruh file dan" -ForegroundColor Cyan  
Write-Host "                     folder di OneDrive setiap user secara mendalam." -ForegroundColor Cyan  
Write-Host "==========================================================" -ForegroundColor Yellow  


## ==========================================================================  
#                           KONFIRMASI EKSEKUSI  
## ==========================================================================  

$confirmation = Read-Host "Apakah Anda ingin menjalankan skrip ini untuk SELURUH USER? (Y/N)"  
if ($confirmation -ne "Y") {  
    Write-Host "`nEksekusi skrip dibatalkan." -ForegroundColor Red  
    return  
}  


## ==========================================================================  
#                     PRASYARAT DAN INSTALASI MODUL  
## ==========================================================================  

Write-Host "`n--- 1. Memeriksa Lingkungan PowerShell ---" -ForegroundColor Blue  
Set-ExecutionPolicy RemoteSigned -Scope Process -Force -ErrorAction SilentlyContinue  

if (-not (Get-Module -Name Microsoft.Graph.Files -ListAvailable)) {  
    Write-Host "Menginstal modul Microsoft.Graph.Files..." -ForegroundColor Yellow
    Install-Module -Name Microsoft.Graph.Files -Force -AllowClobber -Scope CurrentUser  
}  


## ==========================================================================  
##                    KONEKSI KE LAYANAN MICROSOFT
## ==========================================================================  

Write-Host "`n--- 2. Membangun Koneksi ke Microsoft Graph ---" -ForegroundColor Blue  
try {  
    # Memastikan scope Files.Read.All tersedia untuk mengakses OneDrive user lain
    Connect-MgGraph -Scopes "User.Read.All", "Files.Read.All", "Sites.Read.All" -NoWelcome -ContextScope Process
} catch {  
    Write-Error "Gagal Login: $($_.Exception.Message)" ; exit 1  
}  


## ==========================================================================
##                          LOGIKA UTAMA SCRIPT
## ==========================================================================

Write-Host "`n--- 3. Memulai Inventarisasi OneDrive (Mode Rekursif + Sharing Detail) ---" -ForegroundColor Magenta

# --- FUNGSI REKURSIF INVENTARISASI ---
function Get-DeepOneDriveItems {
    param (
        [string]$UserId,
        [string]$UserUPN,
        [string]$DriveId,
        [string]$ParentId = "root",
        [string]$FolderPath = "/",
        [int]$Depth = 0
    )

    if ($Depth -gt 50) { return }

    try {
        # Ambil konten folder
        $items = Get-MgDriveItemChild -DriveId $DriveId -DriveItemId $ParentId -All -ErrorAction Stop
        
        foreach ($item in $items) {
            $isFolder = $null -ne $item.Folder
            $type = if ($isFolder) { "Folder" } else { "File" }
            $sizeMB = if ($item.Size) { [Math]::Round($item.Size / 1MB, 2) } else { 0 }
            $extension = if (-not $isFolder) { [System.IO.Path]::GetExtension($item.Name) } else { "" }

            # --- LOGIKA SHARING (Mendapatkan Email Penerima) ---
            $sharedWithEmails = @()
            if ($null -ne $item.Shared) {
                try {
                    $permissions = Get-MgDriveItemPermission -DriveId $DriveId -DriveItemId $item.Id -ErrorAction SilentlyContinue
                    foreach ($perm in $permissions) {
                        if ($perm.Roles -notcontains "owner") {
                            $email = "N/A"
                            if ($perm.Link) { $email = "Link ($($perm.Link.Scope))" }
                            elseif ($perm.GrantedToV2.User.AdditionalProperties.email) { $email = $perm.GrantedToV2.User.AdditionalProperties.email }
                            elseif ($perm.GrantedToV2.User.UserPrincipalName) { $email = $perm.GrantedToV2.User.UserPrincipalName }
                            elseif ($perm.GrantedToV2.Group.AdditionalProperties.email) { $email = "[Group] " + $perm.GrantedToV2.Group.AdditionalProperties.email }
                            
                            if ($email -ne "N/A") { $sharedWithEmails += $email }
                        }
                    }
                } catch {}
            }

            # Masukkan ke Output Global
            $global:scriptOutput += [PSCustomObject]@{
                Owner          = $UserUPN
                ParentPath     = $FolderPath
                ItemName       = $item.Name
                Extension      = $extension
                Type           = $type
                Size_MB        = $sizeMB
                SharedWith     = ($sharedWithEmails -join ", ")
                LastModified   = $item.LastModifiedDateTime
                WebUrl         = $item.WebUrl
            }

            # Rekursi jika folder (kecuali folder sistem Apps)
            if ($isFolder -and $item.Name -ne "Apps") {
                Get-DeepOneDriveItems -UserId $UserId -UserUPN $UserUPN -DriveId $DriveId -ParentId $item.Id -FolderPath "$FolderPath$($item.Name)/" -Depth ($Depth + 1)
            }
        }
    } catch {
        Write-Host "   [!] Skip path: $FolderPath" -ForegroundColor Yellow
    }
}

# --- PROSES IDENTIFIKASI USER ---
$targetUsers = @()
if ($useAllUsers) {
    $targetUsers = Get-MgUser -All -Property "Id","UserPrincipalName"
} else {
    foreach ($email in $targetUserEmails) {
        try {
            # Menggunakan -UserId langsung untuk menghindari error UnsupportedQuery
            $u = Get-MgUser -UserId $email -Property "Id","UserPrincipalName" -ErrorAction Stop
            if ($u) { $targetUsers += $u }
        } catch {
            Write-Host "   [!] User tidak ditemukan: $email" -ForegroundColor Red
        }
    }
}

# --- EKSEKUSI PEMINDAIAN ---
foreach ($user in $targetUsers) {
    Write-Host "`n>> Memeriksa: $($user.UserPrincipalName)" -ForegroundColor Cyan
    try {
        # Mengambil drive tanpa filter ketat 'business' jika data tidak muncul
        $drives = Get-MgUserDrive -UserId $user.Id -ErrorAction Stop
        $drive = $drives | Where-Object { $_.DriveType -eq "business" } | Select-Object -First 1
        
        # Fallback jika 'business' tidak ditemukan (beberapa tenant terbaca berbeda)
        if (-not $drive) { $drive = $drives | Select-Object -First 1 }

        if ($drive) {
            Write-Host "   [OK] Drive ditemukan (ID: $($drive.Id)). Menelusuri..." -ForegroundColor Green
            Get-DeepOneDriveItems -UserId $user.Id -UserUPN $user.UserPrincipalName -DriveId $drive.Id
        } else {
            Write-Host "   [-] OneDrive tidak ditemukan/aktif." -ForegroundColor Gray
        }
    } catch {
        Write-Host "   [!] Gagal akses Drive: $($_.Exception.Message)" -ForegroundColor Red
    }
}


## ==========================================================================  
##                               EKSPOR HASIL  
## ==========================================================================  

if ($scriptOutput.Count -gt 0) {  
    $exportFolderPath = Join-Path -Path $parentDir -ChildPath "exported_data"  
    if (-not (Test-Path $exportFolderPath)) { New-Item $exportFolderPath -ItemType Directory | Out-Null }  
    
    $resultsFilePath = Join-Path -Path $exportFolderPath -ChildPath $outputFileName  
    $scriptOutput | Export-Csv -Path $resultsFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8  
    Write-Host "`nLaporan selesai! Tersimpan di: ${resultsFilePath}" -ForegroundColor Green  
} else {  
    Write-Host "`nTidak ada data file/folder ditemukan." -ForegroundColor Yellow  
}  