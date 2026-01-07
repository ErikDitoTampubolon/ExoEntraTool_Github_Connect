# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.0)
# Nama Skrip: Get-EntraGroupsInfo
# Deskripsi: Menarik daftar semua grup dari Microsoft Entra ID.
# =========================================================================

# Variabel Global dan Output
$scriptName = "AllGroupsReport" 
$scriptOutput = New-Object System.Collections.Generic.List[PSCustomObject]

# Tentukan jalur dan nama file output dinamis
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$outputFileName = "Output_$($scriptName)_$($timestamp).csv"
$outputFilePath = Join-Path -Path $scriptDir -ChildPath $outputFileName

## ==========================================================================
#                           INFORMASI SCRIPT                
## ==========================================================================

Write-Host "`n================================================" -ForegroundColor Yellow
Write-Host "                INFORMASI SCRIPT                " -ForegroundColor Yellow
Write-Host "================================================" -ForegroundColor Yellow
Write-Host " Nama Skrip        : Get-EntraGroupsInfo" -ForegroundColor Yellow
Write-Host " Field Kolom       : [GroupId]
                     [DisplayName]
                     [Description]
                     [MailEnabled]
                     [SecurityEnabled]
                     [GroupTypes]" -ForegroundColor Yellow
Write-Host " Deskripsi Singkat : Script ini berfungsi untuk menarik daftar semua grup dari Microsoft Entra ID, termasuk informasi ID grup, nama tampilan, deskripsi, status mail-enabled, status security-enabled, serta tipe grup, kemudian mengekspor hasilnya ke file CSV." -ForegroundColor Cyan
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
    Write-Host "Sedang mengambil data Grup..." -ForegroundColor Cyan
    
    # Mengambil semua grup
    $groups = Get-EntraGroup -All -ErrorAction Stop
    
    if ($groups) {
        $total = $groups.Count
        Write-Host "Ditemukan $total grup." -ForegroundColor Green
        $counter = 0

        foreach ($group in $groups) {
            $counter++
            # Progres baris tunggal
            Write-Host "`r-> [$counter/$total] Memproses: $($group.DisplayName) . . ." -ForegroundColor Green -NoNewline
            
            # Membuat objek data kustom
            $obj = [PSCustomObject]@{
                GroupId          = $group.Id
                DisplayName      = $group.DisplayName
                Description      = $group.Description
                MailEnabled      = $group.MailEnabled
                SecurityEnabled  = $group.SecurityEnabled
                GroupTypes       = ($group.GroupTypes -join ", ")
            }
            $scriptOutput.Add($obj)
        }
        Write-Host "`n`nData grup berhasil dikumpulkan." -ForegroundColor Green
    } else {
        Write-Host "`nTidak ada grup yang ditemukan." -ForegroundColor Yellow
    }
} catch {
    Write-Error "Terjadi kesalahan saat mengambil data grup: $($_.Exception.Message)"
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