# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.0)
# Nama Skrip: Export-AllAppsOwnersReport
# Deskripsi: Mengambil daftar pemilik (owners) untuk SEMUA aplikasi di Entra.
# =========================================================================

# Variabel Global dan Output
$scriptName = "AllAppsOwnersReport" 
$scriptOutput = [System.Collections.ArrayList]::new() 

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
Write-Host " Nama Skrip        : Export-AllAppsOwnersReport" -ForegroundColor Yellow
Write-Host " Field Kolom       : [ApplicationName]
                     [ApplicationId]
                     [OwnerObjectId]
                     [OwnerDisplayName]
                     [UserPrincipalName]
                     [CreatedDateTime]
                     [UserType]
                     [AccountEnabled]" -ForegroundColor Yellow
Write-Host " Deskripsi Singkat : Script ini berfungsi untuk mengambil daftar semua aplikasi di Microsoft Entra ID beserta informasi pemiliknya (owners). Jika aplikasi tidak memiliki owner, data tetap dicatat dengan keterangan kosong. Hasil laporan diekspor otomatis ke file CSV." -ForegroundColor Cyan
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
##                  KONEKSI WAJIB (MICROSOFT ENTRA)
## ==========================================================================

Write-Host "`n--- 2. Membangun Koneksi ke Microsoft Entra ---" -ForegroundColor Blue

try {
    Write-Host "Menghubungkan ke Microsoft Entra. Selesaikan login pada pop-up..." -ForegroundColor Yellow
    
    # Menangani potensi konflik DLL dengan mencoba Disconnect terlebih dahulu
    Disconnect-Entra -ErrorAction SilentlyContinue
    
    # Koneksi utama
    Connect-Entra -Scopes 'Application.Read.All' -ErrorAction Stop
    Write-Host "Koneksi ke Microsoft Entra berhasil dibuat." -ForegroundColor Green
} catch {
    Write-Error "Gagal terhubung: $($_.Exception.Message)"
    Write-Host "`nTIP: Jika error library berlanjut, tutup SEMUA jendela PowerShell lalu buka kembali." -ForegroundColor Yellow
    exit 1
}


## ==========================================================================
##                          LOGIKA UTAMA SCRIPT
## ==========================================================================

Write-Host "`n--- 3. Memulai Logika Utama Skrip: $($scriptName) ---" -ForegroundColor Magenta

try {
    Write-Host "Mengambil daftar semua aplikasi..." -ForegroundColor Cyan
    $allApps = Get-EntraApplication -All
    $totalApps = $allApps.Count
    $counter = 0

    Write-Host "Ditemukan $totalApps aplikasi. Memulai pengambilan data pemilik..." -ForegroundColor White

    foreach ($app in $allApps) {
        $counter++
        # Output progres hijau satu baris sesuai style yang Anda sukai
        Write-Host "-> [$counter/$totalApps] Memproses: $($app.DisplayName)" -ForegroundColor Green
        
        try {
            # Mengambil owner untuk aplikasi saat ini
            $owners = Get-EntraApplicationOwner -ApplicationId $app.Id -All -ErrorAction SilentlyContinue

            if ($owners) {
                foreach ($owner in $owners) {
                    $obj = [PSCustomObject]@{
                        ApplicationName   = $app.DisplayName
                        ApplicationId     = $app.AppId
                        OwnerObjectId     = $owner.Id
                        OwnerDisplayName  = $owner.DisplayName
                        UserPrincipalName = $owner.UserPrincipalName
                        CreatedDateTime   = $owner.CreatedDateTime
                        UserType          = $owner.UserType
                        AccountEnabled    = $owner.AccountEnabled
                    }
                    [void]$scriptOutput.Add($obj)
                }
            } else {
                # Jika aplikasi tidak memiliki owner, tetap catat dengan keterangan kosong
                $obj = [PSCustomObject]@{
                    ApplicationName   = $app.DisplayName
                    ApplicationId     = $app.AppId
                    OwnerObjectId     = "NO OWNER"
                    OwnerDisplayName  = "-"
                    UserPrincipalName = "-"
                    CreatedDateTime   = "-"
                    UserType          = "-"
                    AccountEnabled    = "-"
                }
                [void]$scriptOutput.Add($obj)
            }
        } catch {
            Write-Host "   Gagal mengambil owner untuk aplikasi: $($app.DisplayName)" -ForegroundColor Red
        }
    }
} catch {
    Write-Error "Terjadi kesalahan sistem: $($_.Exception.Message)"
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