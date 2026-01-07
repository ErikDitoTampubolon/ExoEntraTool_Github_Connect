# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.0)
# Menyimpan output skrip ke file CSV dinamis di folder skrip.
# =========================================================================

# Variabel Global dan Output
$scriptName = "TransportRulesReport"
$scriptOutput = @()

# Tentukan jalur dan nama file output dinamis
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputFileName = "Output_$($scriptName)_$($timestamp).csv"

# Menentukan lokasi file output
$scriptDir = if ($PSScriptRoot) {$PSScriptRoot} else {(Get-Location).Path}
$outputFilePath = Join-Path -Path $scriptDir -ChildPath $outputFileName

# ==========================================================================
#                           INFORMASI SCRIPT                
# ==========================================================================
Write-Host "`n================================================" -ForegroundColor Yellow
Write-Host "                INFORMASI SCRIPT                " -ForegroundColor Yellow
Write-Host "================================================" -ForegroundColor Yellow
Write-Host " Nama Skrip        : ExportTransportRulesToCSV" -ForegroundColor Yellow
Write-Host " Field Kolom       : [RuleName]
                     [State]
                     [Priority]
                     [WhenCreated]
                     [WhenChanged]
                     [SenderRestrictions]
                     [Description]
                     [AllConditions]
                     [AllActions]" -ForegroundColor Yellow
Write-Host " Deskripsi Singkat : Script ini berfungsi untuk mengambil semua aturan Mail Flow (Transport Rules) dari Exchange Online, termasuk detail kondisi, aksi, status, prioritas, serta metadata pembuatan dan perubahan, lalu mengekspor hasilnya ke file CSV." -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Yellow

# ==========================================================================
#                           KONFIRMASI EKSEKUSI
# ==========================================================================
$confirmation = Read-Host "Apakah Anda ingin menjalankan skrip ini? (Y/N)"

if ($confirmation -ne "Y") {
    Write-Host "`nEksekusi skrip dibatalkan oleh pengguna." -ForegroundColor Red
    return
}

## ==========================================================================
##                          LOGIKA UTAMA SCRIPT
## ==========================================================================

Write-Host "`n--- 3. Memulai Logika Utama Skrip: $($scriptName) ---" -ForegroundColor Magenta
Write-Host "PENTING: Pastikan Anda telah menjalankan Connect-ExchangeOnline secara manual sebelum melanjutkan." -ForegroundColor Red

try {
    Write-Host "3.1. Mengambil semua Mail Flow Rules..." -ForegroundColor Cyan
    
    # Ambil semua Mail Flow Rules. Pilih properti yang relevan.
    $rules = Get-TransportRule | Select-Object Name, State, Priority, *Conditions, *Actions, Description, WhenCreated, WhenChanged, SenderRestrictions

    $totalRules = $rules.Count
    Write-Host "Ditemukan $($totalRules) Aturan Transport." -ForegroundColor Green
    
    $i = 0
    foreach ($rule in $rules) {
        $i++
        
        Write-Progress -Activity "Collecting Transport Rule Data" `
                       -Status "Processing ${i} of ${totalRules}: $($rule.Name)" `
                       -PercentComplete ([int](($i / $totalRules) * 100))
        
        # Inisialisasi variabel untuk properti yang kompleks
        $conditions = @()
        $actions = @()

        # Iterasi melalui semua properti untuk menemukan Conditions dan Actions
        $rule.PSObject.Properties | ForEach-Object {
            $propName = $_.Name
            $propValue = $_.Value
            
            if ($propValue -is [System.Collections.ICollection] -and $propValue.Count -gt 0) {
                # Properti adalah koleksi (misalnya: SentToMemberOf)
                $stringValue = $propValue -join "; "
            } elseif ($propValue -ne $null) {
                # Properti adalah nilai tunggal yang valid
                $stringValue = $propValue.ToString()
            } else {
                # Properti null atau kosong
                $stringValue = ""
            }

            if ($propName -like "*Conditions*") {
                # Hanya simpan Kondisi jika ada nilai
                if (-not [string]::IsNullOrEmpty($stringValue)) {
                    $conditions += "$propName : $stringValue"
                }
            } elseif ($propName -like "*Actions*") {
                 # Hanya simpan Actions jika ada nilai
                if (-not [string]::IsNullOrEmpty($stringValue)) {
                    $actions += "$propName : $stringValue"
                }
            }
        }
        
        # Gabungkan semua Conditions dan Actions menjadi satu string
        $conditionsString = $conditions -join "`r`n"
        $actionsString = $actions -join "`r`n"
        
        # Bangun objek kustom untuk diekspor
        $scriptOutput += [PSCustomObject]@{
            RuleName = $rule.Name
            State = $rule.State
            Priority = $rule.Priority
            WhenCreated = $rule.WhenCreated
            WhenChanged = $rule.WhenChanged
            SenderRestrictions = $rule.SenderRestrictions
            Description = $rule.Description
            
            # Kolom Conditions dan Actions
            AllConditions = $conditionsString
            AllActions = $actionsString
        }
    }

    Write-Progress -Activity "Collecting Transport Rule Data Complete" -Status "Exporting Results" -Completed

}
catch {
    $reason = "Gagal fatal saat mengambil Mail Flow Rules: $($_.Exception.Message)"
    Write-Error $reason
    $scriptOutput += [PSCustomObject]@{
        RuleName = "FATAL ERROR"; State = "N/A"; Priority = "N/A";
        AllConditions = $reason; AllActions = "N/A"
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