# ==========================================================================
# LISENSI MICROSOFT GRAPH ASSIGNMENT/REMOVAL SCRIPT V19.3
# AUTHOR: Erik Dito Tampubolon
# ==========================================================================

# Ambil lokasi script saat ini
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$parentDir = (Get-Item $scriptDir).Parent.Parent.FullName
$defaultUsageLocation = 'ID'
$operationType = ""

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

# ==========================================================================
#                               INFORMASI SCRIPT                
# ==========================================================================

Write-Host "`n================================================" -ForegroundColor Yellow
Write-Host "                INFORMASI SCRIPT                " -ForegroundColor Yellow
Write-Host "================================================" -ForegroundColor Yellow
Write-Host " Nama Skrip        : Assign or Remove License User" -ForegroundColor Yellow
Write-Host " Field Kolom       : [UserPrincipalName]
                     [DisplayName]
                     [Status]
                     [Reason]" -ForegroundColor Yellow
Write-Host " Deskripsi Singkat : Script ini berfungsi untuk melakukan otomatisasi proses pemberian (assign) atau penghapusan (remove) lisensi menggunakan daftar CSV." -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Yellow

# ==========================================================================
#                             KONFIRMASI EKSEKUSI
# ==========================================================================

$confirmation = Read-Host "Apakah Anda ingin menjalankan skrip ini? (Y/N)"

if ($confirmation -ne "Y") {
    Write-Host "`nEksekusi skrip dibatalkan oleh pengguna." -ForegroundColor Red
    return
}

## ==========================================================================
##                      KONEKSI KE MICROSOFT GRAPH
## ==========================================================================

$requiredScopes = "User.ReadWrite.All", "Organization.Read.All"
Write-Host "`n--- 2. Membangun Koneksi ke Microsoft Graph ---" -ForegroundColor Blue

if (Get-MgContext -ErrorAction SilentlyContinue) {
    Write-Host "Sesi Microsoft Graph aktif." -ForegroundColor Green
} else {
    Write-Host "Menghubungkan ke Microsoft Graph..." -ForegroundColor Cyan
    try {
        Connect-MgGraph -Scopes $requiredScopes -ErrorAction Stop | Out-Null
        Write-Host "Koneksi Berhasil." -ForegroundColor Green
    } catch {
        Write-Error "Gagal terhubung ke Microsoft Graph."
        return
    }
}

## ==========================================================================
##                      PEMILIHAN OPERASI DAN LISENSI
## ==========================================================================

Write-Host "`n--- 3. Pemilihan Operasi ---" -ForegroundColor Blue
$operationChoice = Read-Host "Pilih operasi: (1) Assign License | (2) Remove License"

switch ($operationChoice) {
    "1" { $operationType = "Assign" }
    "2" { $operationType = "Remove" }
    default { Write-Host "Pilihan tidak valid." -ForegroundColor Red; return }
}

try {
    $availableLicenses = Get-MgSubscribedSku | Select-Object SkuPartNumber, SkuId -ErrorAction Stop
    Write-Host "`nLisensi yang Tersedia:" -ForegroundColor Yellow
    [int]$index = 1
    $promptOptions = @{}
    foreach ($lic in $availableLicenses) {
        Write-Host "${index}. $($lic.SkuPartNumber)" -ForegroundColor Magenta
        $promptOptions.Add($index, $lic)
        $index++
    }
    
    $choiceInput = Read-Host "`nMasukkan nomor lisensi"
    if (-not $promptOptions.ContainsKey([int]$choiceInput)) { throw "Nomor tidak valid." }
    
    $selectedLicense = $promptOptions[[int]$choiceInput]
    $skuPartNumberTarget = $selectedLicense.SkuPartNumber
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    return
}

## ==========================================================================
##                            LOGIKA UTAMA SCRIPT
## ==========================================================================

$allResults = @()
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

if (-not (Test-Path -Path $inputFilePath)) {
    Write-Host "File ${inputFileName} tidak ditemukan di ${scriptDir}!" -ForegroundColor Red
    return
}

# Import CSV tanpa header
$users = Import-Csv -Path $inputFilePath -Header "UserPrincipalName" -ErrorAction SilentlyContinue
$totalUsers = $users.Count
$userCount = 0 

Write-Host "`n--- 4. Memproses ${totalUsers} Pengguna ---" -ForegroundColor Blue

foreach ($entry in $users) {
    $userCount++
    $userUpn = if ($entry.UserPrincipalName) { $entry.UserPrincipalName.Trim() } else { $null }
    if ([string]::IsNullOrWhiteSpace($userUpn)) { continue }

    # FIX: Menggunakan ${} untuk menghindari error 'Variable reference is not valid'
    Write-Progress -Activity "${operationType} License: ${skuPartNumberTarget}" `
                   -Status "User ${userCount} of ${totalUsers}: ${userUpn}" `
                   -PercentComplete ([int](($userCount / $totalUsers) * 100))
    
    Write-Host "-> [${userCount}/${totalUsers}] Memproses: ${userUpn}" -ForegroundColor White

    try {
        # Ambil User dan simpan ke variabel (agar tidak tumpah ke layar)
        $user = Get-MgUser -UserId $userUpn -Property 'Id', 'DisplayName', 'UsageLocation' -ErrorAction Stop
        
        # Penanganan UsageLocation
        if ($operationType -eq "ASSIGN" -and -not $user.UsageLocation) {
            $null = Update-MgUser -UserId $user.Id -UsageLocation $defaultUsageLocation -ErrorAction Stop
            $user.UsageLocation = $defaultUsageLocation
        }

        # Cek Lisensi
        $userLicense = Get-MgUserLicenseDetail -UserId $user.Id | Where-Object { $_.SkuId -eq $selectedLicense.SkuId }

        if ($operationType -eq "ASSIGN") {
            if ($userLicense) {
                $status = "ALREADY_ASSIGNED"; $reason = "Sudah memiliki lisensi."
            } else {
                $null = Set-MgUserLicense -UserId $user.Id -AddLicenses @(@{ SkuId = $selectedLicense.SkuId }) -RemoveLicenses @() -ErrorAction Stop
                $status = "SUCCESS"; $reason = "Lisensi berhasil diberikan."
            }
        } else {
            if (-not $userLicense) {
                $status = "ALREADY_REMOVED"; $reason = "User tidak memiliki lisensi ini."
            } else {
                $null = Set-MgUserLicense -UserId $user.Id -RemoveLicenses @($selectedLicense.SkuId) -AddLicenses @() -ErrorAction Stop
                $status = "SUCCESS_REMOVED"; $reason = "Lisensi berhasil dihapus."
            }
        }

        $allResults += [PSCustomObject]@{
            UserPrincipalName = $userUpn
            DisplayName       = $user.DisplayName
            Status            = $status
            Reason            = $reason
        }
    }
    catch {
        Write-Host "Gagal: $($_.Exception.Message)" -ForegroundColor Red
        $allResults += [PSCustomObject]@{
            UserPrincipalName = $userUpn
            DisplayName       = "Error/Not Found"
            Status            = "FAIL"
            Reason            = $_.Exception.Message
        }
    }
}
Write-Progress -Activity "Selesai" -Completed

## ==========================================================================
##                              EKSPOR HASIL
## ==========================================================================

if ($allResults.Count -gt 0) {
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
    $outputFileName = "${operationType}_License_Results_${timestamp}.csv"
    $resultsFilePath = Join-Path -Path $exportFolderPath -ChildPath $outputFileName
    
    # 6. Ekspor data
    $allResults | Export-Csv -Path $resultsFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8
    
    Write-Host "`nSemua proses selesai!" -ForegroundColor Green
    Write-Host "Laporan tersimpan di: ${resultsFilePath}" -ForegroundColor Cyan
}