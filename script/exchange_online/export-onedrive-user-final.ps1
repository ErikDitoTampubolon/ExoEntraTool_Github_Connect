# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.0)
# Menyimpan output skrip ke file CSV dinamis di folder skrip.
# =========================================================================

# Variabel Global dan Output
$scriptName = "OneDriveUsageReport" 
$scriptOutput = @() 

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

# --- TAMBAHKAN BARIS INI UNTUK MEMPERBAIKI ERROR ---
$outputFileName = "Output_$($scriptName)_$($timestamp).csv"
# ----------------------------------------------------

# Definisi parentDir (2 tingkat di atas)
$parentDir = (Get-Item $scriptDir).Parent.Parent.FullName

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

Write-Host "`n--- Membangun Koneksi ke Layanan Microsoft ---" -ForegroundColor Blue

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

Write-Host "`n--- Memulai Logika Utama Skrip: $($scriptName) ---" -ForegroundColor Magenta

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
    # Memastikan folder penampung ada di 2 tingkat di atas skrip
    $exportFolderName = "exported_data"
    $exportFolderPath = Join-Path -Path $parentDir -ChildPath $exportFolderName
    
    if (-not (Test-Path -Path $exportFolderPath)) { 
        New-Item -Path $exportFolderPath -ItemType Directory | Out-Null 
    }

    # Menggunakan $outputFileName yang sudah didefinisikan di atas
    $resultsFilePath = Join-Path -Path $exportFolderPath -ChildPath $outputFileName
    
    $scriptOutput | Export-Csv -Path $resultsFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8
    Write-Host "`nLaporan tersimpan di: ${resultsFilePath}" -ForegroundColor Cyan
}