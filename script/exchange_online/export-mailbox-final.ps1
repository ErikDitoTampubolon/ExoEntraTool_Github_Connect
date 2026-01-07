# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V3.3)
# Menyimpan output skrip ke file CSV dinamis di folder skrip.
# =========================================================================

# Variabel Global dan Output
$scriptName = "AllMailboxesToCSV"
$scriptOutput = @()

# Tentukan jalur dan nama file output dinamis
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
# Menggunakan $PSScriptRoot memastikan file disimpan di folder yang sama dengan skrip
$outputFileName = "Output_$($scriptName)_$($timestamp).csv"

# Penanganan kasus $PSScriptRoot tidak ada saat dijalankan dari konsol
$scriptDir = if ($PSScriptRoot) {$PSScriptRoot} else {(Get-Location).Path}
$outputFilePath = Join-Path -Path $scriptDir -ChildPath $outputFileName

## ==========================================================================
#                           INFORMASI SCRIPT                
## ==========================================================================

Write-Host "`n================================================" -ForegroundColor Yellow
Write-Host "                INFORMASI SCRIPT                " -ForegroundColor Yellow
Write-Host "================================================" -ForegroundColor Yellow
Write-Host " Nama Skrip        : ExportAllMailboxesToCSV" -ForegroundColor Yellow
Write-Host " Field Kolom       : [DisplayName]
                     [SamAccountName]
                     [PrimarySmtpAddress]" -ForegroundColor Yellow
Write-Host " Deskripsi Singkat : Script ini berfungsi untuk mengambil semua mailbox pengguna dari Exchange Online, termasuk informasi DisplayName, SamAccountName, dan Primary SMTP Address. Hasil laporan ditampilkan di konsol dan diekspor otomatis ke file CSV." -ForegroundColor Cyan
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
##                          LOGIKA UTAMA SCRIPT
## ==========================================================================

Write-Host "`n--- 3. Memulai Logika Utama Skrip: $($scriptName) ---" -ForegroundColor Magenta

try {
    Write-Host "3.1. Mengambil semua Mailbox Pengguna..." -ForegroundColor Cyan
    
    # Logika inti: Mendapatkan mailbox pengguna dan memilih properti yang relevan.
    $mailboxData = Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited -ErrorAction Stop | 
                   Select-Object DisplayName, SamAccountName, PrimarySmtpAddress | 
                   Sort-Object PrimarySmtpAddress

    $scriptOutput = $mailboxData
    $totalMailboxes = $scriptOutput.Count
    Write-Host "  Ditemukan $($totalMailboxes) Mailbox untuk diekspor." -ForegroundColor Green
    
}
catch {
    $reason = "FATAL ERROR: Gagal mengambil Mailbox. Error: $($_.Exception.Message)"
    Write-Error $reason
    
    if ($scriptOutput.Count -eq 0) {
        $scriptOutput += [PSCustomObject]@{
            DisplayName = "ERROR"; SamAccountName = "N/A"; PrimarySmtpAddress = $reason
        }
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