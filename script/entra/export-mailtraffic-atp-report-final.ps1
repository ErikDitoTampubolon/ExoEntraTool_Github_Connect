# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.0)
# Menyimpan output skrip ke file CSV dinamis di folder skrip.
# =========================================================================

# Variabel Global dan Output
$scriptName = "ATPMailTrafficReport" 
$scriptOutput = @() 

# Tentukan jalur dan nama file output dinamis
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$outputFileName = "Output_$($scriptName)_$($timestamp).csv"
$outputFilePath = Join-Path -Path $PSScriptRoot -ChildPath $outputFileName

## ==========================================================================
##                          LOGIKA UTAMA SCRIPT
## ==========================================================================

Write-Host "`n--- 3. Memulai Logika Utama Skrip: $($scriptName) ---" -ForegroundColor Magenta

try {
    Write-Host "Mengambil data ATP Report (Lalu lintas terbaru)..." -ForegroundColor Cyan
    
    # Menarik data dari Exchange Online
    $atpData = Get-MailTrafficATPReport -ErrorAction Stop
    
    if ($null -ne $atpData) {
        $atpList = @($atpData)
        $totalItems = $atpList.Count
        $indexCount = 0

        # Kolom yang ingin dihilangkan (sesuai permintaan Bapak)
        $excludedFields = "SummarizeBy", "PivotBy", "StartDate", "EndDate", "AggregateBy", "Index"

        foreach ($report in $atpList) {
            $indexCount++
            Write-Host "-> [$indexCount/$totalItems] Memproses Laporan: $($report.Date) . . . Event: $($report.Event)" -ForegroundColor White
            
            # Memfilter objek untuk membuang kolom yang tidak diinginkan
            $filteredReport = $report | Select-Object * -ExcludeProperty $excludedFields
            
            # Memasukkan ke array output
            $scriptOutput += $filteredReport
        }
    } else {
        Write-Host "Tidak ada data ATP ditemukan untuk periode ini." -ForegroundColor Yellow
    }
} catch {
    Write-Host "   Gagal mengambil laporan ATP: $($_.Exception.Message)" -ForegroundColor Red
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
