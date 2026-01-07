# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.6)
# Nama Skrip: Bulk-EntraPasswordReset-AutoGenerate
# Deskripsi: Reset password massal dengan password acak yang dibuat sistem.
# =========================================================================

# 1. Konfigurasi File Input & Path
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$parentDir = (Get-Item $scriptDir).Parent.Parent.FullName

# Variabel Global dan Output
$scriptName = "AutoPasswordResetReport" 
$scriptOutput = [System.Collections.ArrayList]::new() 
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
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

## ==========================================================================
#                           INFORMASI SCRIPT                
## ==========================================================================

Write-Host "`n================================================" -ForegroundColor Yellow
Write-Host "                INFORMASI SCRIPT                " -ForegroundColor Yellow
Write-Host "================================================" -ForegroundColor Yellow
Write-Host " Nama Skrip        : Bulk-EntraPasswordReset-AutoGenerate" -ForegroundColor Yellow
Write-Host " Field Kolom       : [Timestamp]
                     [UserPrincipalName]
                     [TemporaryPassword]
                     [Status]
                     [Message]" -ForegroundColor Yellow
Write-Host " Deskripsi Singkat : Script ini berfungsi untuk melakukan reset password massal pada pengguna Microsoft Entra ID dengan password acak yang digenerate otomatis oleh sistem. Password baru dicatat dalam laporan CSV bersama status eksekusi dan pesan hasil." -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Yellow

## ==========================================================================
#                           KONFIRMASI EKSEKUSI
## ==========================================================================

$confirmation = Read-Host "Apakah Anda ingin menjalankan skrip ini? (Y/N)"

if ($confirmation -ne "Y") {
    Write-Host "`nEksekusi skrip dibatalkan oleh pengguna." -ForegroundColor Red
    return
}

## ==========================================================================
##                  FUNGSI GENERATOR PASSWORD & KONEKSI
## ==========================================================================

# Fungsi untuk membuat password acak 12 karakter yang aman
function Generate-RandomPassword {
    $length = 12
    $chars = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789!@#$%^&*"
    $randomPassword = ""
    for ($i = 0; $i -lt $length; $i++) {
        $randomPassword += $chars[(Get-Random -Minimum 0 -Maximum $chars.Length)]
    }
    return $randomPassword
}

Write-Host "`n--- 2. Membangun Koneksi ke Microsoft Entra ---" -ForegroundColor Blue

try {
    Write-Host "Menghubungkan ke Microsoft Entra..." -ForegroundColor Yellow
    Connect-Entra -Scopes 'User.ReadWrite.All' -ErrorAction Stop
    Write-Host "Koneksi Berhasil." -ForegroundColor Green
} catch {
    Write-Error "Gagal terhubung: $($_.Exception.Message)"
    exit 1
}

## ==========================================================================
##                          LOGIKA UTAMA SCRIPT
## ==========================================================================

Write-Host "`n--- 3. Memulai Proses Reset Password Otomatis ---" -ForegroundColor Magenta

if (Test-Path $inputFilePath) {
    # Import CSV (Asumsi: File hanya berisi daftar email tanpa header)
    $users = Import-Csv $inputFilePath -Header "TempColumn" | Where-Object { $_.TempColumn -ne $null -and $_.TempColumn.Trim() -ne "" }
    $totalUsers = $users.Count
    $counter = 0

    foreach ($user in $users) {
        $counter++
        $targetUser = $user.TempColumn.Trim()
        
        # Buat Password Baru secara acak untuk user ini
        $autoGeneratedPassword = Generate-RandomPassword
        
        # Tampilan progres baris tunggal
        $statusText = "-> [$counter/$totalUsers] Memproses: $targetUser . . ."
        Write-Host "`r$statusText" -ForegroundColor Green -NoNewline

        try {
            # Eksekusi Reset Password
            Set-EntraUser -UserId $targetUser -PasswordProfile @{
                Password = $autoGeneratedPassword
                ForceChangePasswordNextSignIn = $true
            } -ErrorAction Stop

            $res = [PSCustomObject]@{
                Timestamp         = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                UserPrincipalName = $targetUser
                TemporaryPassword = $autoGeneratedPassword # Simpan password di log CSV
                Status            = "SUCCESS"
                Message           = "Password auto-generated and reset"
            }
        } catch {
            $res = [PSCustomObject]@{
                Timestamp         = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                UserPrincipalName = $targetUser
                TemporaryPassword = "-"
                Status            = "FAILED"
                Message           = $_.Exception.Message
            }
        }
        [void]$scriptOutput.Add($res)
    }
    Write-Host "`n`nPemrosesan Selesai." -ForegroundColor Green
} else {
    Write-Host "ERROR: File '$inputFileName' tidak ditemukan!" -ForegroundColor Red
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