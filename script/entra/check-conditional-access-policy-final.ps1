# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.0)
# Nama Skrip: Get-EntraConditionalAccessPolicies
# Deskripsi: Menarik semua kebijakan Conditional Access ke file CSV.
# =========================================================================

# Variabel Global dan Output
$scriptName = "EntraCAPoliciesReport" 
$scriptOutput = New-Object System.Collections.Generic.List[PSCustomObject]

# Tentukan jalur dan nama file output dinamis
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$outputFileName = "Output_$($scriptName)_$($timestamp).csv"
$outputFilePath = Join-Path -Path $scriptDir -ChildPath $outputFileName

## ==========================================================================
#                               INFORMASI SCRIPT                
## ==========================================================================

Write-Host "`n================================================" -ForegroundColor Yellow
Write-Host "                INFORMASI SCRIPT                " -ForegroundColor Yellow
Write-Host "================================================" -ForegroundColor Yellow
Write-Host " Nama Skrip        : Get-EntraConditionalAccessPolicies" -ForegroundColor Yellow
Write-Host " Field Kolom       : [Id]
                     [DisplayName]
                     [State]
                     [CreatedTime]
                     [ModifiedTime]" -ForegroundColor Yellow
Write-Host " Deskripsi Singkat : Script ini berfungsi untuk menarik semua kebijakan Conditional Access dari Microsoft Entra ID, termasuk informasi nama kebijakan, status aktif/nonaktif, serta waktu pembuatan dan modifikasi, kemudian mengekspor hasilnya ke file CSV." -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Yellow

## ==========================================================================
#                               KONFIRMASI EKSEKUSI
## ==========================================================================

$confirmation = Read-Host "Apakah Anda ingin menjalankan skrip ini? (Y/N)"

if ($confirmation -ne "Y") {
    Write-Host "`nEksekusi skrip dibatalkan oleh pengguna." -ForegroundColor Red
    return
}

## ==========================================================================
##                              LOGIKA UTAMA SCRIPT
## ==========================================================================

Write-Host "`n--- 3. Memulai Logika Utama Skrip: $($scriptName) ---" -ForegroundColor Magenta

try {
    Write-Host "Mengambil data Conditional Access Policies..." -ForegroundColor Cyan
    $policies = Get-EntraConditionalAccessPolicy -ErrorAction Stop
    $total = $policies.Count
    $counter = 0

    foreach ($policy in $policies) {
        $counter++
        # Tampilan progres baris tunggal
        $statusText = "-> [$counter/$total] Memproses: $($policy.DisplayName) . . ."
        Write-Host "`r$statusText" -ForegroundColor Green -NoNewline

        # Mengumpulkan data ke dalam objek
        $obj = [PSCustomObject]@{
            Id           = $policy.Id
            DisplayName  = $policy.DisplayName
            State        = $policy.State
            CreatedTime  = $policy.CreatedDateTime
            ModifiedTime = $policy.ModifiedDateTime
        }
        $scriptOutput.Add($obj)
    }
    Write-Host "`n`nData berhasil dikumpulkan." -ForegroundColor Green
} catch {
    Write-Error "Gagal mengambil kebijakan: $($_.Exception.Message)"
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