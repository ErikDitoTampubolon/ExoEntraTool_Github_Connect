# =========================================================================  
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.1)  
# Deskripsi: Mengekspor Struktur Folder dan File dari OneDrive User
# =========================================================================  

# Variabel Global dan Output  
$scriptName = "OneDrive_File_Folder_Inventory" 
$scriptOutput = @() # Array tempat semua data hasil skrip dikumpulkan  

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
Write-Host " Field Kolom       : Owner, ParentPath, ItemName, Type, Size_MB, LastModified" -ForegroundColor Yellow  
Write-Host " Deskripsi Singkat : Melakukan inventarisasi seluruh folder dan file" -ForegroundColor Cyan  
Write-Host "                     di OneDrive pengguna secara rekursif." -ForegroundColor Cyan  
Write-Host "==========================================================" -ForegroundColor Yellow  




## ==========================================================================  
#                           KONFIRMASI EKSEKUSI  
## ==========================================================================  

$confirmation = Read-Host "Apakah Anda ingin menjalankan skrip ini? (Y/N)"  

if ($confirmation -ne "Y") {  
    Write-Host "`nEksekusi skrip dibatalkan oleh pengguna." -ForegroundColor Red  
    return  
}  




# ==========================================================================  
# PILIHAN METODE INPUT (CSV vs ALL USERS)  
# ==========================================================================  
$useAllUsers = $false  
$validInput = $false  
# Loop hingga input valid (1 atau 2)  
while (-not $validInput) {  
    Write-Host "`n--- Metode Input Data ---" -ForegroundColor Blue  
    Write-Host "1. Gunakan Daftar Email dari File CSV" -ForegroundColor Cyan  
    Write-Host "2. Proses Seluruh Pengguna (All Users) di Tenant" -ForegroundColor Cyan  
    $inputMethod = Read-Host "`nPilih metode (1/2)"  
    if ($inputMethod -eq "1") {  
        $useAllUsers = $false  
        $validInput = $true  
    }  
    elseif ($inputMethod -eq "2") {  
        $useAllUsers = $true  
        $validInput = $true  
        Write-Host "[OK] Mode: Seluruh Pengguna OneDrive terpilih." -ForegroundColor Green  
    }  
    else {  
        Write-Host "[ERROR] Pilihan tidak valid! Masukkan angka 1 atau 2." -ForegroundColor Red  
    }  
}  
# Logika Tambahan Jika Menggunakan CSV  
if (-not $useAllUsers) {  
    # Logika Deteksi CSV Existing  
    $csvFiles = Get-ChildItem -Path $parentDir -Filter "*.csv"  
    if ($csvFiles.Count -eq 0) {  
        $newFileName = "daftar_email.csv"  
        $newFilePath = Join-Path -Path $parentDir -ChildPath $newFileName  
        Write-Host "Membuat file CSV baru: $newFileName" -ForegroundColor Cyan  
        "UserPrincipalName" | Out-File -FilePath $newFilePath -Encoding utf8  
          
        Write-Host "`n==========================================================" -ForegroundColor Yellow  
        $checkFill = Read-Host "Apakah Anda sudah mengisi daftar email di file $newFileName? (Y/N)"  
        if ($checkFill -ne "Y") {  
            Write-Host "`nSilakan isi file CSV terlebih dahulu." -ForegroundColor Red  
            Start-Process notepad.exe $newFilePath  
            return  
        }  
        $inputFilePath = $newFilePath  
        $inputFileName = $newFileName  
    } else {  
        $validFileChoice = $false  
        while (-not $validFileChoice) {  
            Write-Host "`nFile CSV yang ditemukan:" -ForegroundColor Yellow  
            for ($i = 0; $i -lt $csvFiles.Count; $i++) {  
                Write-Host "$($i + 1). $($csvFiles[$i].Name)" -ForegroundColor Cyan  
            }  
              
            $fileChoice = Read-Host "`nPilih nomor file CSV yang ingin digunakan"  
              
            if ($fileChoice -as [int] -and [int]$fileChoice -ge 1 -and [int]$fileChoice -le $csvFiles.Count) {  
                $selectedFile = $csvFiles[[int]$fileChoice - 1]  
                $inputFilePath = $selectedFile.FullName  
                $inputFileName = $selectedFile.Name  
                $validFileChoice = $true  
            } else {  
                Write-Host "[ERROR] Pilihan tidak valid! Masukkan angka antara 1 sampai $($csvFiles.Count)." -ForegroundColor Red  
            }  
        }  
    }  
    # Hitung Total Email untuk CSV  
    try {  
        $tempData = Import-Csv -Path $inputFilePath -Header "UserPrincipalName" -ErrorAction SilentlyContinue  
        $totalEmail = if ($tempData) { $tempData.Count } else { 0 }  
        Write-Host "`nFile Terpilih: $inputFileName" -ForegroundColor Green  
        Write-Host "Total email yang terdeteksi: $totalEmail email"  
    } catch {  
        Write-Host "Gagal membaca file CSV." -ForegroundColor Red  
        return  
    }  
}




## ==========================================================================  
#                     PRASYARAT DAN INSTALASI MODUL  
## ==========================================================================  

Write-Host "`n--- 1. Memeriksa dan Menyiapkan Lingkungan PowerShell ---" -ForegroundColor Blue  

# 1.1. Mengatur Execution Policy  
Write-Host "1.1. Mengatur Execution Policy ke RemoteSigned..." -ForegroundColor Cyan  
try {  
    Set-ExecutionPolicy RemoteSigned -Scope Process -Force -ErrorAction Stop  
    Write-Host " Execution Policy berhasil diatur." -ForegroundColor Green  
} catch {  
    Write-Error "Gagal mengatur Execution Policy: $($_.Exception.Message)"  
    exit 1  
}  

# 1.2. Fungsi Pembantu untuk Cek dan Instal Modul  
function CheckAndInstallModule {  
    param(  
        [Parameter(Mandatory=$true)]  
        [string]$ModuleName  
    )  

    Write-Host "1.$(++$global:moduleStep). Memeriksa Modul '$ModuleName'..." -ForegroundColor Cyan  

    if (Get-Module -Name $ModuleName -ListAvailable) {  
        Write-Host " Modul '$ModuleName' sudah terinstal. Melewati instalasi." -ForegroundColor Green  
    } else {  
        Write-Host " Modul '$ModuleName' belum ditemukan. Memulai instalasi..." -ForegroundColor Yellow  
        try {  
            Install-Module -Name $ModuleName -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop  
            Write-Host " Modul '$ModuleName' berhasil diinstal." -ForegroundColor Green  
        } catch {  
            Write-Error "Gagal menginstal modul '$ModuleName'. Pastikan PowerShellGet sudah terinstal dan koneksi internet tersedia."  
            exit 1  
        }  
    }  
}  

$global:moduleStep = 1  
CheckAndInstallModule -ModuleName "PowerShellGet"  
CheckAndInstallModule -ModuleName "Microsoft.Graph"  
CheckAndInstallModule -ModuleName "ExchangeOnlineManagement"



## ==========================================================================  
##                    KONEKSI KE SCOPES YANG DIBUTUHKAN
## ==========================================================================  

Write-Host "`n--- 2. Membangun Koneksi ke Layanan Microsoft ---" -ForegroundColor Blue  

# 2.1 Microsoft Graph  
try {  
    $requiredScopes = "User.Read.All", "Files.Read.All", "Sites.Read.All"
    Connect-MgGraph -NoWelcome -Scopes $requiredScopes -ErrorAction Stop  
    Write-Host "Koneksi ke Microsoft Graph berhasil." -ForegroundColor Green  
} catch {  
    Write-Error "Gagal terhubung ke Microsoft Graph: $($_.Exception.Message)"  
    exit 1  
}  



## ==========================================================================
##                          LOGIKA UTAMA SCRIPT
## ==========================================================================

Write-Host "`n--- 3. Memulai Logika Utama Skrip: $($scriptName) ---" -ForegroundColor Magenta

# Fungsi Rekursif untuk menelusuri folder
function Get-OneDriveItems {
    param (
        [string]$UserId,
        [string]$UserUPN,
        [string]$DriveId,
        [string]$ParentId = "root",
        [string]$FolderPath = "/"
    )

    try {
        $items = Get-MgDriveItemChild -DriveId $DriveId -DriveItemId $ParentId -ErrorAction Stop
        foreach ($item in $items) {
            $type = if ($item.Folder) { "Folder" } else { "File" }
            $sizeMB = if ($item.Size) { [Math]::Round($item.Size / 1MB, 2) } else { 0 }
            
            # Tambahkan ke Output
            $global:scriptOutput += [PSCustomObject]@{
                Owner        = $UserUPN
                ParentPath   = $FolderPath
                ItemName     = $item.Name
                Type         = $type
                Size_MB      = $sizeMB
                LastModified = $item.LastModifiedDateTime
            }

            # Jika Folder, telusuri isinya (Rekursif)
            if ($item.Folder) {
                Get-OneDriveItems -UserId $UserId -UserUPN $UserUPN -DriveId $DriveId -ParentId $item.Id -FolderPath "$FolderPath$($item.Name)/"
            }
        }
    } catch {
        Write-Host "   [!] Gagal mengakses path $FolderPath untuk $UserUPN" -ForegroundColor Red
    }
}

# Ambil Daftar User
if ($useAllUsers) {
    $targetUsers = Get-MgUser -All -Property "Id", "UserPrincipalName"
} else {
    $csvData = Import-Csv -Path $inputFilePath
    $targetUsers = foreach ($row in $csvData) {
        Get-MgUser -UserId $row.UserPrincipalName -Property "Id", "UserPrincipalName" -ErrorAction SilentlyContinue
    }
}

$totalUsers = $targetUsers.Count
$uCount = 0

foreach ($user in $targetUsers) {
    $uCount++
    if ($null -eq $user) { continue }
    
    Write-Host "[$uCount/$totalUsers] Memproses OneDrive: $($user.UserPrincipalName)..." -ForegroundColor Cyan
    
    try {
        # Ambil Drive ID OneDrive User
        $drive = Get-MgUserDrive -UserId $user.Id -ErrorAction Stop | Where-Object { $_.DriveType -eq "business" } | Select-Object -First 1
        
        if ($drive) {
            Get-OneDriveItems -UserId $user.Id -UserUPN $user.UserPrincipalName -DriveId $drive.Id
        } else {
            Write-Host "   [-] OneDrive tidak ditemukan untuk user ini." -ForegroundColor Yellow
        }
    } catch {
        Write-Host "   [!] Error saat mengakses Drive user: $($user.UserPrincipalName)" -ForegroundColor Red
    }
}




## ==========================================================================  
##                               EKSPOR HASIL  
## ==========================================================================  

if ($scriptOutput.Count -gt 0) {  
    # Memastikan folder penampung ada di 2 tingkat di atas skrip  
    $exportFolderName = "exported_data"  
    $exportFolderPath = Join-Path -Path $parentDir -ChildPath $exportFolderName  

    if (-not (Test-Path -Path $exportFolderPath)) {   
        New-Item -Path $exportFolderPath -ItemType Directory | Out-Null   
    }  

    $resultsFilePath = Join-Path -Path $exportFolderPath -ChildPath $outputFileName  

    $scriptOutput | Export-Csv -Path $resultsFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8  
    Write-Host "`nLaporan tersimpan di: ${resultsFilePath}" -ForegroundColor Cyan  
} else {
    Write-Host "`n[INFO] Tidak ada data yang ditemukan untuk diekspor." -ForegroundColor Yellow
}