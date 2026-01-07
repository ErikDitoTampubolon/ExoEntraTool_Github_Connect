# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.1)
# Deskripsi: Mendapatkan status Per-User MFA dengan tampilan Progres Konsol
# =========================================================================

# Variabel Global dan Output
$scriptName = "AllUserMFAStatusReport" 
$scriptOutput = [System.Collections.ArrayList]::new() 

# Tentukan jalur dan nama file output dinamis
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputFileName = "Output_$($scriptName)_$($timestamp).csv"

# Menentukan lokasi file output
$scriptDir = if ($PSScriptRoot) {$PSScriptRoot} else {(Get-Location).Path}
$outputFilePath = Join-Path -Path $scriptDir -ChildPath $outputFileName

## ==========================================================================
#                       INFORMASI SCRIPT                
## ==========================================================================

Write-Host "`n================================================" -ForegroundColor Yellow
Write-Host "                INFORMASI SCRIPT                " -ForegroundColor Yellow
Write-Host "================================================" -ForegroundColor Yellow
Write-Host " Nama Skrip        : Export-AllUser-MFA" -ForegroundColor Yellow
Write-Host " Field Kolom       : [Id]
                     [DisplayName]
                     [UserPrincipalName]
                     [PerUserMFAState]" -ForegroundColor Yellow
Write-Host " Deskripsi Singkat : Script ini berfungsi untuk mendapatkan status MFA per-user dari Microsoft Entra ID, menampilkan progres eksekusi di konsol, serta mengekspor hasil ke file CSV." -ForegroundColor Cyan
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
##                   KONEKSI WAJIB (MICROSOFT ENTRA)
## ==========================================================================

Write-Host "`n--- 2. Membangun Koneksi ke Microsoft Entra ---" -ForegroundColor Blue

try {
    Write-Host "Menghubungkan ke Microsoft Entra. Selesaikan login pada pop-up..." -ForegroundColor Yellow
    
    # Menangani potensi konflik DLL dengan mencoba Disconnect terlebih dahulu
    Disconnect-Entra -ErrorAction SilentlyContinue
    
    # Koneksi utama
    Connect-Entra -Scopes 'User.Read.All', 'UserAuthenticationMethod.Read.All' -ErrorAction Stop
    Write-Host "Koneksi ke Microsoft Entra berhasil dibuat." -ForegroundColor Green
} catch {
    Write-Error "Gagal terhubung: $($_.Exception.Message)"
    Write-Host "`nTIP: Jika error library berlanjut, tutup SEMUA jendela PowerShell lalu buka kembali." -ForegroundColor Yellow
    exit 1
}

## ==========================================================================
##                           LOGIKA UTAMA SCRIPT
## ==========================================================================

Write-Host "`n--- 3. Memulai Logika Utama Skrip ---" -ForegroundColor Magenta

try {
    $users = Get-EntraUser -All -Select Id, UserPrincipalName, DisplayName
    $totalUsers = $users.Count
    $counter = 1

    foreach ($u in $users) {
        # FORMAT TAMPILAN SESUAI PERMINTAAN: -> [1/24] Memproses: email@domain.com
        Write-Host "-> [$($counter)/$($totalUsers)] Memproses: $($u.UserPrincipalName)" -ForegroundColor Green
        
        $mfaState = "Unknown"
        try {
            $mfaReq = Get-EntraBetaUserAuthenticationRequirement -UserId $u.Id -ErrorAction Stop
            $mfaState = if ($null -ne $mfaReq.PerUserMFAState) { $mfaReq.PerUserMFAState } else { "Disabled" }
        } catch {
            $mfaState = "None/Disabled"
        }

        $userProperties = [PSCustomObject]@{
            Id                = $u.Id
            DisplayName       = $u.DisplayName
            UserPrincipalName = $u.UserPrincipalName
            PerUserMFAState   = $mfaState
        }

        [void]$scriptOutput.Add($userProperties)
        $counter++
    }
} catch {
    Write-Error "Terjadi kesalahan: $($_.Exception.Message)"
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