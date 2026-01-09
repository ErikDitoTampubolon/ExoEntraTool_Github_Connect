# =========================================================================
# LISENSI MICROSOFT GRAPH ASSIGNMENT/REMOVAL SCRIPT V20.0
# AUTHOR: Erik Dito Tampubolon
# =========================================================================

# Variabel Global dan Output
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$parentDir = (Get-Item $scriptDir).Parent.Parent.FullName
$defaultUsageLocation = 'ID'
$allResults = @()
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

# ==========================================================================
#                   PILIHAN METODE INPUT (CSV vs ALL USERS)
# ==========================================================================

$useAllUsers = $false
$validInput = $false

while (-not $validInput) {
    Write-Host "`n--- Metode Input Data ---" -ForegroundColor Blue
    Write-Host "1. Gunakan Daftar Email dari File CSV" -ForegroundColor Cyan
    Write-Host "2. Proses Seluruh Pengguna (All Users) di Tenant" -ForegroundColor Cyan
    $inputMethod = Read-Host "`nPilih metode (1/2)"

    if ($inputMethod -eq "1") {
        $useAllUsers = $false
        $validInput = $true
    }
    elseif ($inputMethod -eq "2") {
        $useAllUsers = $true
        $validInput = $true
        Write-Host "[OK] Mode: Seluruh Pengguna terpilih." -ForegroundColor Green
    }
    else {
        Write-Host "[ERROR] Pilihan tidak valid! Masukkan angka 1 atau 2." -ForegroundColor Red
    }
}

# Logika Pemilihan CSV (Hanya jika mode 1)
if (-not $useAllUsers) {
    $csvFiles = Get-ChildItem -Path $parentDir -Filter "*.csv"
    if ($csvFiles.Count -eq 0) {
        $newFileName = "daftar_email.csv"
        $newFilePath = Join-Path -Path $parentDir -ChildPath $newFileName
        $null | Out-File -FilePath $newFilePath -Encoding utf8
        Write-Host "`n[!] File CSV kosong baru dibuat: $newFileName" -ForegroundColor Yellow
        $checkFill = Read-Host "Apakah Anda sudah mengisi daftar email? (Y/N)"
        if ($checkFill -ne "Y") { return }
        $inputFilePath = $newFilePath
        $inputFileName = $newFileName
    } else {
        $validFileChoice = $false
        while (-not $validFileChoice) {
            Write-Host "`nFile CSV yang ditemukan:" -ForegroundColor Blue
            for ($i = 0; $i -lt $csvFiles.Count; $i++) { Write-Host "$($i + 1). $($csvFiles[$i].Name)" -ForegroundColor Cyan }
            $fileChoice = Read-Host "`nPilih nomor file CSV"
            if ($fileChoice -as [int] -and [int]$fileChoice -ge 1 -and [int]$fileChoice -le $csvFiles.Count) {
                $selectedFile = $csvFiles[[int]$fileChoice - 1]
                $inputFilePath = $selectedFile.FullName
                $inputFileName = $selectedFile.Name
                $validFileChoice = $true 
            } else {
                Write-Host "[ERROR] Nomor file tidak valid!" -ForegroundColor Red
            }
        }
    }
}

# ==========================================================================
#                       KONFIGURASI OPERASI LISENSI
# ==========================================================================

$validOp = $false
while (-not $validOp) {
    Write-Host "`n--- Konfigurasi Operasi ---" -ForegroundColor Blue
    Write-Host "1. Assign License (Tambah)"
    Write-Host "2. Remove License (Hapus)"
    $opChoice = Read-Host "Pilih operasi (1/2)"

    if ($opChoice -eq "1") {
        $operationType = "Assign"
        $validOp = $true
    }
    elseif ($opChoice -eq "2") {
        $operationType = "Remove"
        $validOp = $true
    }
    else {
        Write-Host "[ERROR] Pilihan tidak valid! Masukkan angka 1 atau 2." -ForegroundColor Red
    }
}

Write-Host "[OK] Anda memilih untuk: $operationType License" -ForegroundColor Green

## ==========================================================================
##                      KONEKSI KE MICROSOFT GRAPH
## ==========================================================================

$requiredScopes = "User.ReadWrite.All", "Organization.Read.All"
Write-Host "`n--- Membangun Koneksi ke Microsoft Graph ---" -ForegroundColor Blue

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

# Ambil Daftar Lisensi yang Tersedia
$availableLicenses = Get-MgSubscribedSku | Select-Object SkuId, SkuPartNumber

$validSku = $false
while (-not $validSku) {
    Write-Host "`nDaftar Lisensi di Tenant Anda:" -ForegroundColor Yellow
    for ($i = 0; $i -lt $availableLicenses.Count; $i++) {
        Write-Host "$($i + 1). $($availableLicenses[$i].SkuPartNumber)" -ForegroundColor Cyan
    }

    $skuChoice = Read-Host "`nPilih nomor lisensi yang ingin di-$operationType"

    # Validasi: Apakah input adalah angka DAN berada dalam jangkauan daftar lisensi
    if ($skuChoice -as [int] -and [int]$skuChoice -ge 1 -and [int]$skuChoice -le $availableLicenses.Count) {
        $selectedSku = $availableLicenses[[int]$skuChoice - 1]
        $validSku = $true
        Write-Host "[OK] Lisensi terpilih: $($selectedSku.SkuPartNumber)" -ForegroundColor Green
    }
    else {
        Write-Host "[ERROR] Pilihan tidak valid! Masukkan nomor antara 1 sampai $($availableLicenses.Count)." -ForegroundColor Red
    }
}

# ==========================================================================
#                              LOGIKA UTAMA
# ==========================================================================

$usersToProcess = @()
if ($useAllUsers) {
    Write-Host "`nMengambil seluruh pengguna dari Tenant..." -ForegroundColor Yellow
    $usersToProcess = Get-MgUser -All -Property Id, UserPrincipalName, DisplayName
} else {
    $usersToProcess = Import-Csv -Path $inputFilePath -Header "UserPrincipalName"
}

# Validasi Data Kosong
if ($null -eq $usersToProcess -or $usersToProcess.Count -eq 0) {
    Write-Host "`n[ERROR] Tidak ada data pengguna untuk diproses!" -BackgroundColor Red
    return
}

$totalUsers = $usersToProcess.Count
$userCount = 0

Write-Host "`nMemulai proses $operationType untuk $totalUsers user..." -ForegroundColor Magenta

foreach ($entry in $usersToProcess) {
    $userCount++
    $upn = if ($useAllUsers) { $entry.UserPrincipalName } else { $entry.UserPrincipalName.Trim() }
    
    if ([string]::IsNullOrWhiteSpace($upn)) { continue }
    
    # Menggunakan logika counter sesuai permintaan Anda
    Write-Host "-> [$userCount/$totalUsers] Memproses: $upn" -ForegroundColor White
    
    try {
        if ($operationType -eq "Assign") {
            Set-MgUserLicense -UserId $upn -AddLicenses @{SkuId = $selectedSku.SkuId} -RemoveLicenses @() | Out-Null
            $status = "Success (Assigned)"
        } else {
            Set-MgUserLicense -UserId $upn -RemoveLicenses @($selectedSku.SkuId) -AddLicenses @() | Out-Null
            $status = "Success (Removed)"
        }
        Write-Host "   Hasil: $status" -ForegroundColor Green
    } catch {
        $status = "Failed: $($_.Exception.Message)"
        Write-Host "   Hasil: $status" -ForegroundColor Red
    }

    $allResults += [PSCustomObject]@{
        UserPrincipalName = $upn
        Operation         = $operationType
        License           = $selectedSku.SkuPartNumber
        Status            = $status
        Timestamp         = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    }
}

# ==========================================================================
#                              EKSPOR HASIL
# ==========================================================================

if ($scriptOutput.Count -gt 0) {
    $exportFolderName = "exported_data"
    $exportFolderPath = Join-Path -Path $parentDir -ChildPath $exportFolderName
    if (-not (Test-Path -Path $exportFolderPath)) { New-Item -Path $exportFolderPath -ItemType Directory | Out-Null }

    $outputFileName = "Output_$($scriptName)_$($timestamp).csv"
    $resultsFilePath = Join-Path -Path $exportFolderPath -ChildPath $outputFileName
    
    $scriptOutput | Export-Csv -Path $resultsFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8
    Write-Host "`nLaporan tersimpan di: ${resultsFilePath}" -ForegroundColor Cyan
}