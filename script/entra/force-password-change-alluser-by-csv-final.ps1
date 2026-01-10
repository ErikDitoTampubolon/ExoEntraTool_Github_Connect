# =========================================================================  
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.1)  
# Nama Skrip: Bulk-PasswordReset-AutoGenerate
# Deskripsi: Reset password massal dengan password acak dan ekspor hasil ke CSV.
# =========================================================================  

# Variabel Global dan Output  
$scriptName = "BulkPasswordReset" 
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
Write-Host " Field Kolom       : [Timestamp], [UserPrincipalName], [TemporaryPassword], [Status], [Message]" -ForegroundColor Yellow  
Write-Host " Deskripsi Singkat : Melakukan reset password massal (CSV/All Users) dengan password acak 12 karakter." -ForegroundColor Cyan  
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
        Write-Host "[OK] Mode: Seluruh Pengguna terpilih." -ForegroundColor Green  
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

Write-Host "`n--- Memeriksa dan Menyiapkan Lingkungan PowerShell ---" -ForegroundColor Blue  

Set-ExecutionPolicy RemoteSigned -Scope Process -Force -ErrorAction SilentlyContinue

function CheckAndInstallModule {  
    param([string]$ModuleName)  
    Write-Host "Memeriksa Modul '$ModuleName'..." -ForegroundColor Cyan  
    if (-not (Get-Module -Name $ModuleName -ListAvailable)) {  
        Write-Host "Menginstal '$ModuleName'..." -ForegroundColor Yellow  
        Install-Module -Name $ModuleName -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop  
    }  
}  

CheckAndInstallModule -ModuleName "Microsoft.Graph"  

## ==========================================================================  
##                    KONEKSI KE SCOPES YANG DIBUTUHKAN
## ==========================================================================  

Write-Host "`n--- Membangun Koneksi ke Layanan Microsoft ---" -ForegroundColor Blue  

try {  
    $requiredScopes = @("User.ReadWrite.All", "Directory.ReadWrite.All")  
    Connect-MgGraph -Scopes $requiredScopes -ErrorAction Stop  
    Write-Host "Koneksi ke Microsoft Graph berhasil." -ForegroundColor Green  
} catch {  
    Write-Error "Gagal terhubung: $($_.Exception.Message)"  
    exit 1  
}  

## ==========================================================================
##                          FUNGSI PEMBANTU
## ==========================================================================

function Generate-RandomPassword {
    # Meningkatkan panjang ke 16 karakter dan memastikan variasi karakter yang lebih kuat
    $length = 16
    $upper   = "ABCDEFGHJKLMNPQRSTUVWXYZ"
    $lower   = "abcdefghijkmnopqrstuvwxyz"
    $numbers = "23456789"
    $symbols = "!@#$%^&*"
    
    # Memastikan setidaknya ada satu karakter dari setiap kategori
    $pass = @(
        $upper[(Get-Random -Maximum $upper.Length)],
        $lower[(Get-Random -Maximum $lower.Length)],
        $numbers[(Get-Random -Maximum $numbers.Length)],
        $symbols[(Get-Random -Maximum $symbols.Length)]
    )

    # Melengkapi sisa karakter secara acak
    $allChars = $upper + $lower + $numbers + $symbols
    for ($i = 1; $i -le ($length - 4); $i++) {
        $pass += $allChars[(Get-Random -Maximum $allChars.Length)]
    }

    # Mengacak urutan karakter agar tidak terprediksi
    return (-join ($pass | Sort-Object { Get-Random }))
}

## ==========================================================================
##                          LOGIKA UTAMA SCRIPT
## ==========================================================================

Write-Host "`n--- Memulai Logika Utama Skrip: $($scriptName) ---" -ForegroundColor Magenta

# Ambil Data Pengguna
if ($useAllUsers) {
    Write-Host "Mengambil data seluruh pengguna..." -ForegroundColor Yellow
    $targetUsers = Get-MgUser -All -Select "UserPrincipalName,Id"
} else {
    $csvData = Import-Csv -Path $inputFilePath -Header "UserPrincipalName" | Where-Object { $_.UserPrincipalName -ne $null -and $_.UserPrincipalName.Trim() -ne "" }
    $targetUsers = $csvData
}

$totalUsers = $targetUsers.Count
$counter = 0

foreach ($user in $targetUsers) {
    $counter++
    $upn = $user.UserPrincipalName.Trim()
    $newPassword = Generate-RandomPassword
    
    Write-Host "[$counter/$totalUsers] Memproses: $upn" -ForegroundColor Green

    try {
        # Parameter Reset Password untuk Microsoft Graph
        $params = @{
            passwordProfile = @{
                forceChangePasswordNextSignIn = $true
                password = $newPassword
            }
        }
        Update-MgUser -UserId $upn -BodyParameter $params -ErrorAction Stop

        $res = [PSCustomObject]@{
            Timestamp         = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            UserPrincipalName = $upn
            TemporaryPassword = $newPassword
            Status            = "SUCCESS"
            Message           = "Password berhasil direset"
        }
    } catch {
        $res = [PSCustomObject]@{
            Timestamp         = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            UserPrincipalName = $upn
            TemporaryPassword = "-"
            Status            = "FAILED"
            Message           = $_.Exception.Message
        }
    }
    [void]$scriptOutput.Add($res)
}

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
    
    Write-Host "`nSemua proses selesai!" -ForegroundColor Green
    Write-Host "Laporan tersimpan di: ${resultsFilePath}" -ForegroundColor Cyan  
}