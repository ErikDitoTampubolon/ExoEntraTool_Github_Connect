# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.0)
# Nama Skrip: Export-DuplicateEntraDevices
# Deskripsi: Mengidentifikasi dan mengekspor perangkat dengan nama duplikat.
# =========================================================================

# Variabel Global dan Output
$scriptName = "DuplicateDevicesReport"
$scriptOutput = @()

# Tentukan jalur dan nama file output dinamis
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$outputFileName = "Output_$($scriptName)_$($timestamp).csv"
$outputFilePath = Join-Path -Path $PSScriptRoot -ChildPath $outputFileName

## ==========================================================================
#                           INFORMASI SCRIPT                
## ==========================================================================

Write-Host "`n================================================" -ForegroundColor Yellow
Write-Host "                INFORMASI SCRIPT                " -ForegroundColor Yellow
Write-Host "================================================" -ForegroundColor Yellow
Write-Host " Nama Skrip        : Export-DuplicateEntraDevices" -ForegroundColor Yellow
Write-Host " Field Kolom       : [DeviceName]
                     [OperatingSystem]
                     [Count]" -ForegroundColor Yellow
Write-Host " Deskripsi Singkat : Script ini berfungsi untuk mengidentifikasi perangkat di Microsoft Entra ID yang memiliki nama duplikat. Hasil laporan menampilkan nama perangkat, sistem operasi, serta jumlah duplikat, kemudian diekspor otomatis ke file CSV." -ForegroundColor Cyan
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
    Write-Host "Sedang mengambil dan memproses data perangkat, harap tunggu..." -ForegroundColor Cyan
    
    # Menjalankan logika pencarian duplikat
    $results = Get-EntraDevice -All -Select DisplayName, OperatingSystem |
        Group-Object DisplayName |
        Where-Object { $_.Count -gt 1 } |
        Select-Object @{Name = "DeviceName"; Expression = { $_.Name }}, 
                      @{Name = "OperatingSystem"; Expression = { ($_.Group | Select-Object -First 1).OperatingSystem } }, 
                      Count | 
        Sort-Object Count -Descending

    # Memasukkan hasil ke variabel framework untuk ekspor
    if ($null -ne $results) {
        $scriptOutput += $results
        
        # Tampilkan tabel di layar sesuai permintaan
        $results | Format-Table -AutoSize
    } else {
        Write-Host "Tidak ditemukan perangkat duplikat." -ForegroundColor Green
    }
} catch {
    Write-Error "Terjadi kesalahan saat memproses data: $($_.Exception.Message)"
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