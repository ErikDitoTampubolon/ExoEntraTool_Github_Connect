# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.0)
# Menyimpan output skrip ke file CSV dinamis di folder skrip.
# =========================================================================

# Variabel Global dan Output
$scriptName = "OneDriveUsageReport" 
$scriptOutput = @() 

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$parentDir = (Get-Item $scriptDir).Parent.Parent.FullName

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

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

# ==========================================================
#                INFORMASI SCRIPT                
# ==========================================================
Write-Host "`n================================================" -ForegroundColor Yellow
Write-Host "                INFORMASI SCRIPT                " -ForegroundColor Yellow
Write-Host "================================================" -ForegroundColor Yellow
Write-Host " Nama Skrip        : Get-OneDriveSizeReport.ps1" -ForegroundColor Yellow
Write-Host " Field Kolom       : [Owner]
                     [UserPrincipalName]
                     [SiteId]
                     [IsDeleted]
                     [LastActivityDate]
                     [FileCount]
                     [ActiveFileCount]
                     [QuotaGB]
                     [UsedGB]
                     [PercentUsed]
                     [City]
                     [Country]
                     [Department]
                     [JobTitle]" -ForegroundColor Yellow
Write-Host " Deskripsi Singkat : Script ini berfungsi untuk mengekspor laporan penggunaan storage OneDrive for Business di Microsoft 365. Laporan mencakup informasi pemilik, UPN, detail site, status penghapusan, aktivitas terakhir, jumlah file, kuota, pemakaian storage, persentase penggunaan, serta atribut tambahan pengguna (lokasi, departemen, jabatan). Hasil laporan diekspor ke file CSV dan ditampilkan dalam GridView." -ForegroundColor Cyan
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
##                       KONEKSI KE MICROSOFT GRAPH
## ==========================================================================

Write-Host "`n--- 2. Membangun Koneksi ke Layanan Microsoft ---" -ForegroundColor Blue

# 2.1 Microsoft Graph
try {
    $requiredScopes = "User.Read.All", "Reports.Read.All", "ReportSettings.ReadWrite.All"
    Connect-MgGraph -NoWelcome -Scopes $requiredScopes -ErrorAction Stop
    Write-Host "Koneksi ke Microsoft Graph berhasil." -ForegroundColor Green
} catch {
    Write-Error "Gagal terhubung ke Microsoft Graph: $($_.Exception.Message)"
    exit 1
}

# 2.2 Exchange Online (Wajib Framework)
if (-not (Get-PSSession | Where-Object {$_.ConfigurationName -eq "Microsoft.Exchange"})) {
    try {
        Connect-ExchangeOnline -ShowProgress $false -ErrorAction Stop | Out-Null
        Write-Host "Koneksi ke Exchange Online berhasil." -ForegroundColor Green
    } catch {
        Write-Error "Gagal terhubung ke Exchange Online: $($_.Exception.Message)"
        exit 1
    }
}

## ==========================================================================
##                           LOGIKA UTAMA SCRIPT
## ==========================================================================

Write-Host "`n--- 3. Memulai Logika Utama Skrip: $($scriptName) ---" -ForegroundColor Magenta

try {
    # 2.1 Bypass Concealment (Penting agar UPN tidak berupa ID acak)
    Update-MgAdminReportSetting -BodyParameter @{ displayConcealedNames = $false } -ErrorAction SilentlyContinue

    # 2.2 Ambil Data User untuk Mapping City/Dept
    Write-Host "Mengambil data detail akun user..." -ForegroundColor Cyan
    $users = Get-MgUser -All -Property Id,displayName,userPrincipalName,city,department,jobTitle -ErrorAction SilentlyContinue
    $UserHash = @{}
    foreach ($u in $users) { $UserHash[$u.userPrincipalName] = $u }

    # 2.3 Ambil Report OneDrive
    $TempExportFile = Join-Path $scriptDir "temp_usage.csv"
    Get-MgReportOneDriveUsageAccountDetail -Period D30 -Outfile $TempExportFile -ErrorAction Stop

    if (Test-Path $TempExportFile) {
        # URUTAN HEADER YANG BENAR UNTUK GRAPH REPORT:
        # 1. Report Refresh Date, 2. Site URL, 3. Owner Display Name, 4. Is Deleted, 
        # 5. Last Activity Date, 6. File Count, 7. Active File Count, 8. Storage Allocated (Byte), 
        # 9. Storage Used (Byte), 10. Owner Principal Name
        
        $rawCSV = Import-CSV $TempExportFile
        $totalItems = $rawCSV.Count
        $currentIndex = 0

        foreach ($row in $rawCSV) {
            $currentIndex++
            
            # Mengambil properti langsung dari objek CSV (Graph SDK biasanya menyertakan header otomatis)
            $targetUPN = $row.'Owner Principal name'
            $usedByte  = if ($row.'Storage Used (Byte)') { [double]$row.'Storage Used (Byte)' } else { 0 }
            $quotaByte = if ($row.'Storage Allocated (Byte)') { [double]$row.'Storage Allocated (Byte)' } else { 0 }
            
            $usedGB  = [Math]::Round($usedByte / 1GB, 2)
            $quotaGB = [Math]::Round($quotaByte / 1GB, 2)

            Write-Host ("-> [{0}/{1}] Memproses: {2} . . . Usage: {3} GB" -f $currentIndex, $totalItems, $targetUPN, $usedGB) -ForegroundColor White

            $userData = $UserHash[$targetUPN]

            $scriptOutput += [PSCustomObject]@{
                ReportDate        = $row.'Report Refresh Date'
                Owner             = $row.'Owner display name'
                UserPrincipalName = $targetUPN
                IsDeleted         = $row.'Is Deleted'
                LastActivityDate  = $row.'Last Activity Date'
                FileCount         = $row.'File Count'
                QuotaGB           = $quotaGB
                UsedGB            = $usedGB
                City              = if ($userData.city) { $userData.city } else { "N/A" }
                Department        = if ($userData.department) { $userData.department } else { "N/A" }
                JobTitle          = if ($userData.jobTitle) { $userData.jobTitle } else { "N/A" }
            }
        }
    }
} catch {
    Write-Host "Terjadi kesalahan: $($_.Exception.Message)" -ForegroundColor Red
} finally {
    if (Test-Path $TempExportFile) { Remove-Item $TempExportFile -Force }
}

## ==========================================================================
##                               EKSPOR HASIL
## ==========================================================================

if ($scriptOutput.Count -gt 0) {
    $exportFolderName = "exported_data"
    $parentDir = (Get-Item $scriptDir).Parent.Parent.FullName
    $exportFolderPath = Join-Path -Path $parentDir -ChildPath $exportFolderName

    if (-not (Test-Path -Path $exportFolderPath)) {
        New-Item -Path $exportFolderPath -ItemType Directory | Out-Null
        Write-Host "`nFolder '$exportFolderName' berhasil dibuat." -ForegroundColor Yellow
    }

    $outputFileName = "Output_$($scriptName)_$($timestamp).csv"
    $resultsFilePath = Join-Path -Path $exportFolderPath -ChildPath $outputFileName

    # Ekspor menggunakan pemisah titik koma sesuai framework
    $scriptOutput | Export-Csv -Path $resultsFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8
    
    Write-Host "`nSemua proses selesai!" -ForegroundColor Green
    Write-Host "Laporan tersimpan di: ${resultsFilePath}" -ForegroundColor Cyan
} else {
    Write-Host "Tidak ada data yang diproses untuk diekspor." -ForegroundColor Yellow
}