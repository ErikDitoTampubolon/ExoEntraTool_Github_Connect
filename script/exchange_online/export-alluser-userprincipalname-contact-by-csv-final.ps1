# =========================================================================
# AUTHOR: Erik Dito Tampubolon - TelkomSigma
# VERSION: 2.9 (UI Enhanced Output)
# Deskripsi: Fix ParserError & Support No Header CSV dengan Output Progres Hijau.
# =========================================================================

# Variabel Global dan Output
$scriptName = "AllActiveUsersDnUpnContactByCSVReport" 
$scriptOutput = [System.Collections.ArrayList]::new() 
$global:moduleStep = 1

# Konfigurasi File Input]

# Tentukan jalur dan nama file output dinamis
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$parentDir = (Get-Item $scriptDir).Parent.Parent.FullName

$outputFileName = "Output_$($scriptName)_$($timestamp).csv"
$outputFilePath = Join-Path -Path $scriptDir -ChildPath $outputFileName

# ==========================================================================
#                    DETEKSI DAN PEMILIHAN FILE CSV
# ==========================================================================

Write-Host "`n--- Mencari file CSV di: $parentDir ---" -ForegroundColor Blue

# Mencari semua file .csv di folder 2 tingkat di atas script
$csvFiles = Get-ChildItem -Path $parentDir -Filter "*.csv"

if ($csvFiles.Count -eq 0) {
    Write-Host "Tidak ditemukan file CSV di direktori: $parentDir" -ForegroundColor Yellow
    
    $newFileName = "daftar.email.csv"
    $newFilePath = Join-Path -Path $parentDir -ChildPath $newFileName
    
    Write-Host "Membuat file CSV baru: $newFileName" -ForegroundColor Cyan
    $null | Out-File -FilePath $newFilePath -Encoding utf8
    
    Write-Host "File berhasil dibuat." -ForegroundColor Green
    
    # ALERT MESSAGE: Memastikan pengguna mengisi file sebelum lanjut
    Write-Host "`n==========================================================" -ForegroundColor Yellow
    $checkFill = Read-Host "Apakah Anda sudah mengisi daftar email di file $newFileName? (Y/N)"
    Write-Host "==========================================================" -ForegroundColor Yellow

    if ($checkFill -ne "Y") {
        Write-Host "`nSilakan isi file CSV terlebih dahulu, lalu jalankan ulang skrip." -ForegroundColor Red
        # Membuka file secara otomatis agar pengguna bisa langsung mengisi
        Start-Process notepad.exe $newFilePath
        return
    }

    $inputFilePath = $newFilePath
    $inputFileName = $newFileName
}
else {
    Write-Host "File CSV yang ditemukan:" -ForegroundColor Yellow
    for ($i = 0; $i -lt $csvFiles.Count; $i++) {
        Write-Host "$($i + 1). $($csvFiles[$i].Name)" -ForegroundColor Cyan
    }

    $fileChoice = Read-Host "`nPilih nomor file CSV yang ingin digunakan"

    if (-not ($fileChoice -as [int]) -or [int]$fileChoice -lt 1 -or [int]$fileChoice -gt $csvFiles.Count) {
        Write-Host "Pilihan tidak valid. Skrip dibatalkan." -ForegroundColor Red
        return
    }

    $selectedFile = $csvFiles[[int]$fileChoice - 1]
    $inputFilePath = $selectedFile.FullName
    $inputFileName = $selectedFile.Name
}

# --- LOGIKA HITUNG TOTAL EMAIL ---
try {
    # Ambil data untuk verifikasi isi
    $tempData = Import-Csv -Path $inputFilePath -Header "TempColumn" -ErrorAction SilentlyContinue
    $totalEmail = if ($tempData) { $tempData.Count } else { 0 }
    
    Write-Host "`nFile Terpilih: $inputFileName" -ForegroundColor Green
    Write-Host "Total email yang terdeteksi: $totalEmail email" -ForegroundColor Cyan

    # Proteksi Tambahan: Jika file terdeteksi masih 0 baris setelah konfirmasi Y
    if ($totalEmail -eq 0) {
        Write-Host "`nPERINGATAN: File $inputFileName terdeteksi masih KOSONG." -ForegroundColor Red
        $reconfirm = Read-Host "Apakah Anda yakin ingin tetap melanjutkan? (Y/N)"
        if ($reconfirm -ne "Y") { 
            Write-Host "Eksekusi dibatalkan. Silakan isi data terlebih dahulu." -ForegroundColor Yellow
            return 
        }
    }
    Write-Host "----------------------------------------------------------"
} catch {
    Write-Host "Gagal membaca file CSV: $($_.Exception.Message)" -ForegroundColor Red
    return
}

# ==========================================================================
#                            INFORMASI SCRIPT                
# ==========================================================================

Write-Host "`n================================================" -ForegroundColor Yellow
Write-Host "                INFORMASI SCRIPT                " -ForegroundColor Yellow
Write-Host "================================================" -ForegroundColor Yellow
Write-Host " Nama Skrip        : UserContactReport_Final_Fixed" -ForegroundColor Yellow
Write-Host " Field Kolom       : [InputUser]
                     [DisplayName]
                     [UPN]
                     [Contact]" -ForegroundColor Yellow
Write-Host " Deskripsi Singkat : Script ini berfungsi untuk membuat laporan kontak pengguna Microsoft Entra ID berdasarkan daftar UPN dari file CSV tanpa header. Script menampilkan progres eksekusi di konsol, mengambil informasi DisplayName, UPN, serta nomor telepon (BusinessPhones dan MobilePhone), lalu mengekspor hasil ke file CSV." -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Yellow

## ==========================================================================
##                          KONFIRMASI EKSEKUSI
## ==========================================================================

$confirmation = Read-Host "Apakah Anda ingin menjalankan skrip ini? (Y/N)"

if ($confirmation -ne "Y") {
    Write-Host "`nEksekusi skrip dibatalkan oleh pengguna." -ForegroundColor Red
    return
}

## ==========================================================================
##                          PRASYARAT DAN KONEKSI
## ==========================================================================

Write-Host "--- 1. Menyiapkan Lingkungan ---" -ForegroundColor Blue 
if (-not (Get-MgContext -ErrorAction SilentlyContinue)) { 
    Connect-MgGraph -Scopes "User.Read.All" -ErrorAction Stop | Out-Null
}

## ==========================================================================
##                          LOGIKA UTAMA SCRIPT
## ==========================================================================

Write-Host "`n--- 2. Memulai Logika Utama Skrip ---" -ForegroundColor Magenta 

if (-not (Test-Path $inputFilePath)) {
    Write-Host " ERROR: File '$inputFileName' tidak ditemukan!" -ForegroundColor Red
    exit 1
}

# Membaca CSV dengan Header manual karena file asli tidak memiliki judul kolom
$csvData = Import-Csv -Path $inputFilePath -Header "Email" -ErrorAction SilentlyContinue

if ($null -eq $csvData -or $csvData.Count -eq 0) {
    Write-Host " ERROR: File CSV kosong." -ForegroundColor Red
    exit 1
}

$totalData = $csvData.Count
$i = 0

foreach ($row in $csvData) {
    $i++
    
    # Ambil nilai email
    $targetUser = if ($row.Email) { $row.Email.Trim() } else { $null }
    
    if ([string]::IsNullOrWhiteSpace($targetUser)) { continue }

    # FORMAT OUTPUT SESUAI PERMINTAAN: -> [i/total] Memproses: email@domain.com
    Write-Host "-> [$($i)/$($totalData)] Memproses: $($targetUser)" -ForegroundColor White

    try {
        $userObj = Get-MgUser -UserId $targetUser -Property "UserPrincipalName","DisplayName","BusinessPhones","MobilePhone" -ErrorAction Stop
        
        $phones = @()
        if ($userObj.BusinessPhones) { $phones += ($userObj.BusinessPhones -join ", ") }
        if ($userObj.MobilePhone) { $phones += $userObj.MobilePhone }
        
        $contactInfo = if ($phones.Count -gt 0) { $phones -join " | " } else { "-" }

        [void]$scriptOutput.Add([PSCustomObject]@{
            InputUser   = $targetUser
            DisplayName = $userObj.DisplayName
            UPN         = $userObj.UserPrincipalName
            Contact     = $contactInfo
        })
    }
    catch {
        [void]$scriptOutput.Add([PSCustomObject]@{
            InputUser   = $targetUser
            DisplayName = "NOT FOUND"
            UPN         = "-"
            Contact     = "-"
        })
    }
}

## ==========================================================================
##                              EKSPOR HASIL
## ==========================================================================

if ($scriptOutput.Count -gt 0) {
    # 1. Tentukan nama folder
    $exportFolderName = "exported_data"
    
    # 2. Ambil jalur dua tingkat di atas direktori skrip
    # Contoh: Jika skrip di C:\Users\Erik\Project\Scripts, maka ini ke C:\Users\Erik\
    $parentDir = (Get-Item $scriptDir).Parent.Parent.FullName
    
    # 3. Gabungkan menjadi jalur folder ekspor
    $exportFolderPath = Join-Path -Path $parentDir -ChildPath $exportFolderName

    # 4. Cek apakah folder 'exported_data' sudah ada di lokasi tersebut, jika belum buat baru
    if (-not (Test-Path -Path $exportFolderPath)) {
        New-Item -Path $exportFolderPath -ItemType Directory | Out-Null
        Write-Host "`nFolder '$exportFolderName' berhasil dibuat di: $parentDir" -ForegroundColor Yellow
    }

    # 5. Tentukan nama file dan jalur lengkap
    $outputFileName = "Output_$($scriptName)_$($timestamp).csv"
    $resultsFilePath = Join-Path -Path $exportFolderPath -ChildPath $outputFileName
    
    # 6. Ekspor data
    $scriptOutput | Export-Csv -Path $resultsFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8
    
    Write-Host "`nSemua proses selesai!" -ForegroundColor Green
    Write-Host "Laporan tersimpan di: ${resultsFilePath}" -ForegroundColor Cyan
}