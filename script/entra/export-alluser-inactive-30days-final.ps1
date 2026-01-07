# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.1 - FIXED)
# Nama Skrip: Export-EntraInactiveGuestUsers
# Deskripsi: Mengambil daftar Guest User yang tidak aktif > 30 hari.
# =========================================================================

# Variabel Global dan Output
$scriptName = "InactiveGuestUsers30DaysReport" 
$scriptOutput = [System.Collections.Generic.List[PSCustomObject]]::new() 

# Tentukan jalur dan nama file output dinamis
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }

## ==========================================================================
#                           INFORMASI SCRIPT                
## ==========================================================================

Write-Host "`n================================================" -ForegroundColor Yellow
Write-Host "                INFORMASI SCRIPT                " -ForegroundColor Yellow
Write-Host "================================================" -ForegroundColor Yellow
Write-Host " Nama Skrip        : Export-EntraInactiveGuestUsers" -ForegroundColor Yellow
Write-Host " Field Kolom       : [DisplayName]
                     [UserPrincipalName]
                     [UserType]
                     [AccountEnabled]
                     [LastSignInDateTime]
                     [LastNonInteractiveSignIn]
                     [Id]" -ForegroundColor Yellow
Write-Host " Deskripsi Singkat : Script ini berfungsi untuk mengambil daftar Guest User di Microsoft Entra ID yang tidak aktif lebih dari 30 hari. Laporan mencakup informasi DisplayName, UPN, tipe user, status akun, serta detail aktivitas sign-in terakhir. Hasil laporan ditampilkan di konsol dengan progres bar dan diekspor otomatis ke file CSV." -ForegroundColor Cyan
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
##                   KONEKSI WAJIB (MICROSOFT ENTRA)
## ==========================================================================

Write-Host "`n--- 2. Membangun Koneksi ke Microsoft Entra ---" -ForegroundColor Blue

try {
    Write-Host "Menghubungkan ke Microsoft Entra..." -ForegroundColor Yellow
    Disconnect-Entra -ErrorAction SilentlyContinue
    Connect-Entra -Scopes 'AuditLog.Read.All','User.Read.All' -ErrorAction Stop
    Write-Host "Koneksi berhasil." -ForegroundColor Green
} catch {
    Write-Error "Gagal terhubung: $($_.Exception.Message)"
    exit 1
}

## ==========================================================================
##                           LOGIKA UTAMA SCRIPT
## ==========================================================================

Write-Host "`n--- 3. Memulai Logika Utama Skrip ---" -ForegroundColor Magenta

try {
    Write-Host "Menganalisis pengguna tidak aktif (LastSignIn < 30 hari lalu)..." -ForegroundColor Cyan
    
    # PERBAIKAN: Menghapus parameter -All karena tidak didukung oleh cmdlet ini
    $inactiveUsers = Get-EntraInactiveSignInUser -LastSignInBeforeDaysAgo 30 -ErrorAction Stop
    
    $totalData = $inactiveUsers.Count
    
    if ($totalData -gt 0) {
        $i = 0
        foreach ($user in $inactiveUsers) {
            $i++
            
            # Update Progres UI
            Write-Host "`r-> [$i/$totalData] Memproses: $($user.UserPrincipalName)" -ForegroundColor Green -NoNewline
            
            # Mapping data
            $obj = [PSCustomObject]@{
                DisplayName              = $user.DisplayName
                UserPrincipalName        = $user.UserPrincipalName
                UserType                 = $user.UserType
                AccountEnabled           = $user.AccountEnabled
                LastSignInDateTime       = $user.SignInActivity.LastSignInDateTime
                LastNonInteractiveSignIn = $user.SignInActivity.LastNonInteractiveSignInDateTime
                Id                       = $user.Id
            }
            $scriptOutput.Add($obj)
        }
        Write-Host "`n`nBerhasil memproses $totalData pengguna." -ForegroundColor Green
    } else {
        Write-Host "Tidak ditemukan pengguna tidak aktif." -ForegroundColor Yellow
    }
} catch {
    Write-Error "Terjadi kesalahan saat mengambil data: $($_.Exception.Message)"
}

## ==========================================================================
##                              EKSPOR HASIL
## ==========================================================================

if ($scriptOutput.Count -gt 0) {
    Write-Host "`n--- 4. Mengekspor Hasil ---" -ForegroundColor Blue
    
    try {
        $exportFolderName = "exported_data"
        # Jalur folder: 2 tingkat di atas folder skrip
        $parentDir = (Get-Item $scriptDir).Parent.Parent.FullName
        $exportFolderPath = Join-Path -Path $parentDir -ChildPath $exportFolderName

        if (-not (Test-Path -Path $exportFolderPath)) {
            New-Item -Path $exportFolderPath -ItemType Directory | Out-Null
        }

        $resultsFilePath = Join-Path -Path $exportFolderPath -ChildPath "Output_$($scriptName)_$($timestamp).csv"
        
        $scriptOutput | Export-Csv -Path $resultsFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8
        
        Write-Host "Laporan tersimpan di: ${resultsFilePath}" -ForegroundColor Cyan
    } catch {
        Write-Error "Gagal mengekspor CSV: $($_.Exception.Message)"
    }
}

Disconnect-Entra -ErrorAction SilentlyContinue
Write-Host "`nProses Selesai." -ForegroundColor Yellow