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

Write-Host "`n--- Memulai Logika Utama Skrip: $($scriptName) ---" -ForegroundColor Magenta

try {
    Write-Host "Mengambil semua Mailbox Pengguna..." -ForegroundColor Cyan
    
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
    $targetDir = (Get-Item $scriptDir).Parent.Parent.FullName

    $exportFolderName = "exported_data"
    $exportFolderPath = Join-Path -Path $targetDir -ChildPath $exportFolderName
    
    if (-not (Test-Path -Path $exportFolderPath)) { 
        New-Item -Path $exportFolderPath -ItemType Directory | Out-Null 
    }

    $resultsFilePath = Join-Path -Path $exportFolderPath -ChildPath $outputFileName
    $scriptOutput | Export-Csv -Path $resultsFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8
    Write-Host "`nLaporan tersimpan di: ${resultsFilePath}" -ForegroundColor Cyan
}