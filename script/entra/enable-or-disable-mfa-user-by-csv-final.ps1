# =========================================================================  
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.1)  
# Nama Skrip: Bulk-MFA-Manager
# Deskripsi: Mengelola MFA (Enable/Disable) via CSV atau All Users.
# =========================================================================  

# Variabel Global dan Output  
$scriptName = "BulkMFAManager" 
$scriptOutput = [System.Collections.ArrayList]::new() 

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
Write-Host " Field Kolom       : [Timestamp], [UserAccount], [Operation], [Status], [Message]" -ForegroundColor Yellow  
Write-Host " Deskripsi Singkat : Mengelola status MFA (Enable/Disable) secara massal " -ForegroundColor Cyan  
Write-Host "                     berdasarkan CSV atau seluruh user di tenant." -ForegroundColor Cyan
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
while (-not $validInput) {  
    Write-Host "`n--- Metode Input Data ---" -ForegroundColor Blue  
    Write-Host "1. Gunakan Daftar Email dari File CSV (Tanpa Header)" -ForegroundColor Cyan  
    Write-Host "2. Proses Seluruh Pengguna (All Users) di Tenant" -ForegroundColor Cyan  
    $inputMethod = Read-Host "`nPilih metode (1/2)"  
    if ($inputMethod -eq "1") {  
        $useAllUsers = $false  
        $validInput = $true  
    }  
    elseif ($inputMethod -eq "2") {  
        $useAllUsers = $true  
        $validInput = $true  
        Write-Host "[OK] Mode: Seluruh User terpilih." -ForegroundColor Green  
    }  
    else {  
        Write-Host "[ERROR] Pilihan tidak valid! Masukkan angka 1 atau 2." -ForegroundColor Red  
    }  
}  

if (-not $useAllUsers) {  
    $csvFiles = Get-ChildItem -Path $parentDir -Filter "*.csv"  
    if ($csvFiles.Count -eq 0) {  
        $newFileName = "daftar_email.csv"  
        $newFilePath = Join-Path -Path $parentDir -ChildPath $newFileName  
        Write-Host "Membuat file CSV baru: $newFileName" -ForegroundColor Cyan  
        $null | Out-File -FilePath $newFilePath -Encoding utf8  
          
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
                Write-Host "[ERROR] Pilihan tidak valid!" -ForegroundColor Red  
            }  
        }  
    }  
}

## ==========================================================================  
#                     PRASYARAT DAN INSTALASI MODUL  
## ==========================================================================  

Write-Host "`n--- 1. Memeriksa dan Menyiapkan Lingkungan PowerShell ---" -ForegroundColor Blue  

Set-ExecutionPolicy RemoteSigned -Scope Process -Force -ErrorAction SilentlyContinue

function CheckAndInstallModule {  
    param([Parameter(Mandatory=$true)][string]$ModuleName)  
    Write-Host "1.$(++$global:moduleStep). Memeriksa Modul '$ModuleName'..." -ForegroundColor Cyan  
    if (-not (Get-Module -Name $ModuleName -ListAvailable)) {  
        Install-Module -Name $ModuleName -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop  
    }  
    Write-Host " Modul '$ModuleName' OK." -ForegroundColor Green  
}  

$global:moduleStep = 0  
CheckAndInstallModule -ModuleName "Microsoft.Graph"  
CheckAndInstallModule -ModuleName "Microsoft.Entra.Beta"

## ==========================================================================  
##                    KONEKSI DAN PEMILIHAN OPERASI
## ==========================================================================  

Write-Host "`n--- 2. Membangun Koneksi & Konfigurasi Operasi ---" -ForegroundColor Blue  

# Pilih Operasi Terlebih Dahulu
$operationChoice = Read-Host "Pilih operasi: (1) Enable MFA | (2) Disable MFA"
$targetMfaState = if ($operationChoice -eq "1") { "enabled" } else { "disabled" }
$operationType = if ($operationChoice -eq "1") { "ENABLE-MFA" } else { "DISABLE-MFA" }

try {  
    Connect-Entra -Scopes 'Policy.ReadWrite.AuthenticationMethod', 'User.ReadWrite.All' -ErrorAction Stop  
    Write-Host "Koneksi ke Microsoft Entra berhasil." -ForegroundColor Green  
} catch {  
    Write-Error "Gagal terhubung: $($_.Exception.Message)"  
    exit 1  
}  

## ==========================================================================
##                          LOGIKA UTAMA SCRIPT
## ==========================================================================

Write-Host "`n--- 3. Memulai Logika Utama Skrip: $($operationType) ---" -ForegroundColor Magenta

# Mengambil Data User
if ($useAllUsers) {
    Write-Host "Mengambil seluruh data user dari tenant..." -ForegroundColor Yellow
    $users = Get-EntraUser -All | Select-Object @{Name="UserPrincipalName"; Expression={$_.UserPrincipalName}}
} else {
    $users = Import-Csv $inputFilePath -Header "UserPrincipalName" | Where-Object { $_.UserPrincipalName -ne $null -and $_.UserPrincipalName.Trim() -ne "" }
}

$totalUsers = $users.Count
$counter = 0

foreach ($user in $users) {
    $counter++
    $targetUser = $user.UserPrincipalName.Trim()
    Write-Host "`r-> [$counter/$totalUsers] Memproses: $targetUser" -ForegroundColor Green -NoNewline

    try {
        # Update MFA menggunakan Microsoft Entra Beta
        Update-EntraBetaUserAuthenticationRequirement -UserId $targetUser -PerUserMfaState $targetMfaState -ErrorAction Stop
        
        $res = [PSCustomObject]@{
            Timestamp   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            UserAccount = $targetUser
            Operation   = $operationType
            Status      = "SUCCESS"
            Message     = "Status set to $targetMfaState"
        }
    } catch {
        # Menangani error dengan ekstraksi pesan yang lebih aman
        $errorMessage = $_.Exception.Message
        if ($_.Exception.InnerException) {
            $errorMessage += " | " + $_.Exception.InnerException.Message
        }

        $res = [PSCustomObject]@{
            Timestamp   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            UserAccount = $targetUser
            Operation   = $operationType
            Status      = "FAILED"
            Message     = $errorMessage.Replace("`n", " ").Replace("`r", "") # Membersihkan baris baru agar CSV tetap rapi
        }
    }
    [void]$scriptOutput.Add($res)
}

Write-Host "`n`nPemrosesan Selesai." -ForegroundColor Green

## ==========================================================================  
##                               EKSPOR HASIL  
## ==========================================================================  

if ($scriptOutput.Count -gt 0) {  
    $exportFolderName = "exported_data"  
    $exportFolderPath = Join-Path -Path $parentDir -ChildPath $exportFolderName  

    if (-not (Test-Path -Path $exportFolderPath)) {   
        New-Item -Path $exportFolderPath -ItemType Directory | Out-Null   
    }  

    $resultsFilePath = Join-Path -Path $exportFolderPath -ChildPath $outputFileName  
    $scriptOutput | Export-Csv -Path $resultsFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8  
    
    Write-Host "Laporan tersimpan di: ${resultsFilePath}" -ForegroundColor Cyan  
}