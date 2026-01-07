# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.1)
# Nama Skrip: Export-EntraDeletedUsers
# Deskripsi: Mengambil daftar pengguna yang dihapus dengan UI Progress.
# =========================================================================

# Variabel Global dan Output
$scriptName = "DeletedUsersReport" 
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
Write-Host " Nama Skrip        : Export-EntraDeletedUsers" -ForegroundColor Yellow
Write-Host " Field Kolom       : [Id]
                     [UserPrincipalName]
                     [DisplayName]
                     [AccountEnabled]
                     [DeletedDateTime]
                     [DeletionAgeInDays]
                     [UserType]" -ForegroundColor Yellow
Write-Host " Deskripsi Singkat : Script ini berfungsi untuk mengambil daftar pengguna yang telah dihapus dari Microsoft Entra ID, menampilkan progres eksekusi di konsol, serta mengekspor hasil detail (termasuk informasi UPN, nama tampilan, status akun, tanggal penghapusan, usia penghapusan, dan tipe user) ke file CSV." -ForegroundColor Cyan
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
##                          LOGIKA UTAMA SCRIPT
## ==========================================================================

Write-Host "`n--- 3. Memulai Logika Utama Skrip: $($scriptName) ---" -ForegroundColor Magenta

try {
    Write-Host "Mengambil data pengguna yang dihapus..." -ForegroundColor Cyan
    
    $rawDeleted = Get-EntraDeletedUser -All -ErrorAction Stop
    $totalData = $rawDeleted.Count
    
    if ($totalData -gt 0) {
        $i = 0
        foreach ($user in $rawDeleted) {
            $i++
            
            # OUTPUT PROGRES BARIS TUNGGAL SESUAI PERMINTAAN
            # Menggunakan -NoNewline dan `r untuk menjaga di satu baris
            $statusText = "-> [$i/$totalData] Memproses: $($user.UserPrincipalName) . . ."
            Write-Host "`r$statusText" -ForegroundColor White -NoNewline
            
            # Mapping data ke objek hasil
            $obj = [PSCustomObject]@{
                Id                 = $user.Id
                UserPrincipalName  = $user.UserPrincipalName
                DisplayName        = $user.DisplayName
                AccountEnabled     = $user.AccountEnabled
                DeletedDateTime    = $user.DeletedDateTime
                DeletionAgeInDays  = $user.DeletionAgeInDays
                UserType           = $user.UserType
            }
            [void]$scriptOutput.Add($obj)
        }
        Write-Host "`n`nBerhasil memproses $totalData pengguna." -ForegroundColor Green
    } else {
        Write-Host "Tidak ditemukan pengguna yang dihapus." -ForegroundColor Yellow
    }
} catch {
    Write-Error "Terjadi kesalahan: $($_.Exception.Message)"
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