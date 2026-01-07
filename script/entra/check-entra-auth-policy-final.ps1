# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.0)
# Nama Skrip: Get-EntraAuthorizationPolicies
# Deskripsi: Mengambil detail kebijakan otorisasi Microsoft Entra ID.
# =========================================================================

# Variabel Global dan Output
$scriptName = "EntraAuthPolicyReport" 
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
Write-Host " Nama Skrip        : Get-EntraAuthorizationPolicies" -ForegroundColor Yellow
Write-Host " Field Kolom       : [Id]
                     [DisplayName]
                     [Description]
                     [AllowUserConsentForApp]
                     [BlockMsolIdp]
                     [DefaultUserRolePermissions]" -ForegroundColor Yellow
Write-Host " Deskripsi Singkat : Script ini berfungsi untuk mengambil detail kebijakan otorisasi dari Microsoft Entra ID, termasuk informasi nama kebijakan, deskripsi, pengaturan consent aplikasi, blok MSOL IdP, serta default user role permissions, kemudian mengekspor hasilnya ke file CSV." -ForegroundColor Cyan
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
    Write-Host "Mengambil data Authorization Policy..." -ForegroundColor Cyan
    
    # Menjalankan perintah utama
    $policies = Get-EntraAuthorizationPolicy -ErrorAction Stop
    
    if ($policies) {
        $total = $policies.Count
        $counter = 0

        foreach ($policy in $policies) {
            $counter++
            # Visualisasi progres baris tunggal
            Write-Host "`r-> [$counter/$total] Memproses: $($policy.DisplayName) . . ." -ForegroundColor Green -NoNewline
            
            # Memasukkan data ke array output
            $scriptOutput += [PSCustomObject]@{
                Id                           = $policy.Id
                DisplayName                  = $policy.DisplayName
                Description                  = $policy.Description
                AllowUserConsentForApp       = $policy.AllowedToUseSSPR
                BlockMsolIdp                 = $policy.BlockMsolIdp
                DefaultUserRolePermissions   = $policy.DefaultUserRolePermissions | ConvertTo-Json -Compress
            }
        }
        Write-Host "`nData berhasil dikumpulkan." -ForegroundColor Green
    }
} catch {
    Write-Error "Terjadi kesalahan saat mengambil kebijakan: $($_.Exception.Message)"
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