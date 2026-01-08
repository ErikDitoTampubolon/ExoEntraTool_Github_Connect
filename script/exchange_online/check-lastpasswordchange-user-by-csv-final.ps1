# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.4)
# =========================================================================

# Variabel Global dan Output
$scriptName = "PasswordChangeReport" 
$scriptOutput = @() 

# Penanganan Jalur Aman
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$parentDir = (Get-Item $scriptDir).Parent.Parent.FullName

# Tentukan jalur output
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

# ==========================================================================
#                               INFORMASI SCRIPT                
# ==========================================================================

Write-Host "`n================================================" -ForegroundColor Yellow
Write-Host "                INFORMASI SCRIPT                " -ForegroundColor Yellow
Write-Host "================================================" -ForegroundColor Yellow
Write-Host " Nama Skrip        : $scriptName" -ForegroundColor Yellow
Write-Host " Mode Eksekusi     : $(if($useAllUsers){"All Users"}else{"CSV Input"})" -ForegroundColor Yellow
Write-Host " Field Kolom       : [UserPrincipalName], [DisplayName], [LastPasswordChangeWIB], [Status]" -ForegroundColor Yellow
Write-Host "==========================================================" -ForegroundColor Yellow

$confirmation = Read-Host "Apakah Anda ingin menjalankan skrip ini? (Y/N)"
if ($confirmation -ne "Y") { return }

# ==========================================================================
#                   PILIHAN METODE INPUT (CSV vs ALL USERS)
# ==========================================================================

$useAllUsers = $false
$validInput = $false

# Loop hingga input valid (1 atau 2)
while (-not $validInput) {
    Write-Host "`n--- Metode Input Data ---" -ForegroundColor Blue
    Write-Host "1. Gunakan Daftar Email dari File CSV"
    Write-Host "2. Proses Seluruh Pengguna (All Users) di Tenant"
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

# Logika Tambahan Jika Menggunakan CSV
if (-not $useAllUsers) {
    # Logika Deteksi CSV Existing
    $csvFiles = Get-ChildItem -Path $parentDir -Filter "*.csv"

    if ($csvFiles.Count -eq 0) {
        $newFileName = "daftar_email.csv"
        $newFilePath = Join-Path -Path $parentDir -ChildPath $newFileName
        Write-Host "Membuat file CSV baru: $newFileName" -ForegroundColor Cyan
        $null | Out-File -FilePath $newFilePath -Encoding utf8
        
        Write-Host "`n==========================================================" -ForegroundColor Yellow
        $checkFill = Read-Host "Apakah Anda sudah mengisi daftar email di file $newFileName? (Y/N)"
        if ($checkFill -ne "Y") {
            Write-Host "`nSilakan isi file CSV terlebih dahulu." -ForegroundColor Red
            Start-Process notepad.exe $newFilePath
            return
        }
        $inputFilePath = $newFilePath
        $inputFileName = $newFileName
    } 
    else {
        $validFileChoice = $false
        while (-not $validFileChoice) {
            Write-Host "`nFile CSV yang ditemukan:" -ForegroundColor Yellow
            for ($i = 0; $i -lt $csvFiles.Count; $i++) {
                Write-Host "$($i + 1). $($csvFiles[$i].Name)" -ForegroundColor Cyan
            }
            
            $fileChoice = Read-Host "`nPilih nomor file CSV yang ingin digunakan"
            
            if ($fileChoice -as [int] -and [int]$fileChoice -ge 1 -and [int]$fileChoice -le $csvFiles.Count) {
                $selectedFile = $csvFiles[[int]$fileChoice - 1]
                $inputFilePath = $selectedFile.FullName
                $inputFileName = $selectedFile.Name
                $validFileChoice = $true 
            } else {
                Write-Host "[ERROR] Pilihan tidak valid! Masukkan angka antara 1 sampai $($csvFiles.Count)." -ForegroundColor Red
            }
        }
    }

    # Hitung Total Email untuk CSV (Pastikan ini masih di DALAM blok "if (-not $useAllUsers)")
    try {
        $tempData = Import-Csv -Path $inputFilePath -Header "UserPrincipalName" -ErrorAction SilentlyContinue
        $totalEmail = if ($tempData) { $tempData.Count } else { 0 }
        Write-Host "`nFile Terpilih: $inputFileName" -ForegroundColor Green
        Write-Host "Total email yang terdeteksi: $totalEmail email"
    } catch {
        Write-Host "Gagal membaca file CSV." -ForegroundColor Red
        return
    }
}

## ==========================================================================
##                          KONEKSI KE MICROSOFT GRAPH
## ==========================================================================

if (-not (Get-MgContext -ErrorAction SilentlyContinue)) {
    Write-Host "`n--- Membangun Koneksi ke Microsoft Graph ---" -ForegroundColor Blue
    Connect-MgGraph -Scopes "User.Read.All" -ErrorAction Stop | Out-Null
}

## ==========================================================================
##                              LOGIKA UTAMA SCRIPT
## ==========================================================================

Write-Host "`n--- Memulai Logika Utama Skrip: ${scriptName} ---" -ForegroundColor Magenta

$usersToProcess = @()

if ($useAllUsers) {
    Write-Host "Mengambil seluruh data pengguna dari Tenant... Mohon tunggu." -ForegroundColor Yellow
    $usersToProcess = Get-MgUser -All -Property UserPrincipalName, DisplayName, LastPasswordChangeDateTime -ErrorAction Stop
} else {
    if (-not (Test-Path -Path $inputFilePath)) {
        Write-Host "[ERROR] File input CSV tidak ditemukan!" -ForegroundColor Red
        return
    }
    $usersToProcess = Import-Csv -Path $inputFilePath -Header "UserPrincipalName" -ErrorAction SilentlyContinue
}

# --- VALIDASI KETAT DATA KOSONG ---
if ($null -eq $usersToProcess -or $usersToProcess.Count -eq 0) {
    Write-Host " ERROR: DATA TIDAK DITEMUKAN " -ForegroundColor White -BackgroundColor Red
    
    if ($useAllUsers) {
        Write-Host "Penyebab: Tidak ada objek pengguna terdeteksi di Tenant ini." -ForegroundColor Yellow
    } else {
        Write-Host "Penyebab: File '$inputFileName' kosong atau tidak memiliki data email." -ForegroundColor Yellow
    }
    
    Write-Host "`nSkrip dihentikan secara otomatis untuk mencegah error sistem." -ForegroundColor Red
    Write-Host "Silakan periksa kembali data Anda sebelum menjalankan ulang." -ForegroundColor Cyan
    Write-Host "==========================================================" -ForegroundColor Red
    
    # Berhenti total
    return 
}

$totalUsers = $usersToProcess.Count
$userCount = 0

foreach ($entry in $usersToProcess) {
    $userCount++
    $upn = if ($useAllUsers) { $entry.UserPrincipalName } else { $entry.UserPrincipalName.Trim() }
    
    if ([string]::IsNullOrWhiteSpace($upn)) { continue }

    Write-Progress -Activity "Generating Report" -Status "User ${userCount}/${totalUsers}"
    Write-Host "-> [${userCount}/${totalUsers}] Memproses: ${upn}..." -ForegroundColor White
    
    try {
        # Jika mode All Users, data sudah ada di variabel $entry
        $userData = if ($useAllUsers) { $entry } else { Get-MgUser -UserId $upn -Property DisplayName, LastPasswordChangeDateTime -ErrorAction Stop }
        
        $lastChangeRaw = $userData.LastPasswordChangeDateTime
        $lastChangeWIB = "N/A"

        if ($lastChangeRaw) {
            $dateTimeUTC = [System.DateTime]::SpecifyKind($lastChangeRaw, [System.DateTimeKind]::Utc)
            $wibTimeZone = [System.TimeZoneInfo]::FindSystemTimeZoneById("SE Asia Standard Time")
            $dateTimeWIB = [System.TimeZoneInfo]::ConvertTimeFromUtc($dateTimeUTC, $wibTimeZone)
            $lastChangeWIB = $dateTimeWIB.ToString("yyyy-MM-dd HH:mm:ss")
        }
        
        Write-Host "Last Password Changes: ${lastChangeWIB}" -ForegroundColor Green
        
        $scriptOutput += [PSCustomObject]@{
            UserPrincipalName     = $upn
            DisplayName           = $userData.DisplayName
            LastPasswordChangeWIB = $lastChangeWIB
            Status                = "SUCCESS"
        }
    } 
    catch {
        Write-Host "   Gagal mengambil data." -ForegroundColor Red
        $scriptOutput += [PSCustomObject]@{
            UserPrincipalName = $upn; Status = "FAIL"; Reason = $_.Exception.Message
        }
    }
}

## ==========================================================================
##                              EKSPOR HASIL
## ==========================================================================

if ($scriptOutput.Count -gt 0) {
    $exportFolderName = "exported_data"
    $exportFolderPath = Join-Path -Path $parentDir -ChildPath $exportFolderName

    if (-not (Test-Path -Path $exportFolderPath)) {
        New-Item -Path $exportFolderPath -ItemType Directory | Out-Null
    }

    $outputFileName = "Output_$($scriptName)_$($timestamp).csv"
    $resultsFilePath = Join-Path -Path $exportFolderPath -ChildPath $outputFileName
    
    $scriptOutput | Export-Csv -Path $resultsFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8
    
    Write-Host "`nSemua proses selesai!" -ForegroundColor Green
    Write-Host "Laporan tersimpan di: ${resultsFilePath}" -ForegroundColor Cyan
}