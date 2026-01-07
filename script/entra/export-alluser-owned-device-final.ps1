# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.1)
# Nama Skrip: Export-AllUserOwnedDevice
# Deskripsi: Mengekspor daftar semua perangkat milik semua pengguna Entra ID.
# =========================================================================

# Variabel Global dan Output
$scriptName = "AllUserOwnedDeviceReport" 
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
Write-Host " Nama Skrip        : Export-AllUserOwnedDevice" -ForegroundColor Yellow
Write-Host " Field Kolom       : [UserPrincipalName]
                     [UserDisplayName]
                     [DeviceDisplayName]
                     [DeviceId]
                     [OperatingSystem]
                     [OSVersion]
                     [AccountEnabled]
                     [TrustType]" -ForegroundColor Yellow
Write-Host " Deskripsi Singkat : Script ini berfungsi untuk mengekspor daftar semua perangkat yang dimiliki oleh setiap pengguna di Microsoft Entra ID. Laporan mencakup informasi pengguna (UPN, DisplayName) serta detail perangkat (nama, ID, OS, versi, status akun, dan jenis trust). Hasil laporan ditampilkan di konsol dengan progres bar dan diekspor otomatis ke file CSV." -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Yellow

## ==========================================================================
#                       KONFIRMASI EKSEKUSI
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
    Connect-Entra -Scopes 'User.Read.All' -ErrorAction Stop
    Write-Host "Koneksi ke Microsoft Entra berhasil dibuat." -ForegroundColor Green
} catch {
    Write-Error "Gagal terhubung: $($_.Exception.Message)"
    Write-Host "`nTIP: Jika error library berlanjut, tutup SEMUA jendela PowerShell lalu buka kembali." -ForegroundColor Yellow
    exit 1
}

## ==========================================================================
##                       LOGIKA UTAMA SCRIPT
## ==========================================================================

Write-Host "`n--- 3. Memulai Logika Utama Skrip: $($scriptName) ---" -ForegroundColor Magenta

try {
    Write-Host "Mengambil daftar semua pengguna..." -ForegroundColor Cyan
    $allUsers = Get-EntraUser -All
    $totalUsers = $allUsers.Count
    $counter = 0

    Write-Host "Ditemukan $totalUsers pengguna. Memulai pemindaian perangkat..." -ForegroundColor White

    foreach ($user in $allUsers) {
        $counter++
        
        # Tampilan progres di satu baris (UI Refresh)
        $statusText = "-> [$counter/$totalUsers] Memproses: $($user.UserPrincipalName) . . ."
        Write-Host "`r$statusText" -ForegroundColor Green -NoNewline

        try {
            # Mengambil perangkat milik user tersebut
            $ownedDevices = Get-EntraUserOwnedDevice -UserId $user.Id -All -ErrorAction SilentlyContinue

            if ($ownedDevices) {
                foreach ($device in $ownedDevices) {
                    $obj = [PSCustomObject]@{
                        UserPrincipalName = $user.UserPrincipalName
                        UserDisplayName   = $user.DisplayName
                        DeviceDisplayName = $device.DisplayName
                        DeviceId          = $device.DeviceId
                        OperatingSystem   = $device.OperatingSystem
                        OSVersion         = $device.OperatingSystemVersion
                        AccountEnabled    = $device.AccountEnabled
                        TrustType         = $device.TrustType
                    }
                    [void]$scriptOutput.Add($obj)
                }
            }
        } catch {
            # Mengabaikan error jika satu user gagal diakses (misal: permission)
            continue
        }
    }
    Write-Host "`n`nPemrosesan selesai." -ForegroundColor Green
} catch {
    Write-Error "Terjadi kesalahan sistem: $($_.Exception.Message)"
}

## ==========================================================================
##                               EKSPOR HASIL
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