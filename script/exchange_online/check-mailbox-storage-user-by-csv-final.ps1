# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V5.1 - No Header)
# =========================================================================

# Variabel Global dan Output
$scriptName = "MailboxStorageByCSVReport" 
$scriptOutput = @() 

# Penanganan Jalur Aman
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$parentDir = (Get-Item $scriptDir).Parent.Parent.FullName

# Tentukan jalur output
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputFileName = "Output_${scriptName}_${timestamp}.csv"
$outputFilePath = Join-Path -Path $scriptDir -ChildPath $outputFileName

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
#                           INFORMASI SCRIPT                
# ==========================================================================
Write-Host "`n================================================" -ForegroundColor Yellow
Write-Host "                INFORMASI SCRIPT                " -ForegroundColor Yellow
Write-Host "================================================" -ForegroundColor Yellow
Write-Host " Nama Skrip        : MailboxStorageReport" -ForegroundColor Yellow
Write-Host " Field Kolom       : [UserPrincipalName]
                     [DisplayName]
                     [TotalItemSizeGB]
                     [ItemCount]
                     [WarningQuota]
                     [SendQuota]
                     [LastLogonTime]
                     [Status]
                     [Reason]" -ForegroundColor Yellow
Write-Host " Deskripsi Singkat : Script ini berfungsi untuk membuat laporan penggunaan storage mailbox berdasarkan daftar UPN dari file CSV tanpa header. Script akan memvalidasi mailbox, mengambil statistik ukuran, jumlah item, kuota peringatan dan kuota kirim, serta waktu logon terakhir, kemudian mengekspor hasil ke file CSV." -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Yellow

# ==========================================================================
#                           KONFIRMASI EKSEKUSI
# ==========================================================================
$confirmation = Read-Host "Apakah Anda ingin menjalankan skrip ini? (Y/N)"

if ($confirmation -ne "Y") {
    Write-Host "`nEksekusi skrip dibatalkan oleh pengguna." -ForegroundColor Red
    return
}

## ==========================================================================
##                          LOGIKA UTAMA SCRIPT
## ==========================================================================

Write-Host "`n--- 3. Memulai Logika Utama Skrip: ${scriptName} ---" -ForegroundColor Magenta

if (-not (Test-Path -Path $inputFilePath)) {
    Write-Error "File input CSV tidak ditemukan di: $inputFilePath"
    $scriptOutput += [PSCustomObject]@{ UserPrincipalName = $inputFileName; Status = "FAIL"; Reason = "Input file not found." }
} else {
    Write-Host "Memuat data dari '${inputFileName}' (Mode: No Header)..." -ForegroundColor Cyan
    
    # MODIFIKASI: Menggunakan -Header "UserPrincipalName" karena CSV tidak memiliki judul kolom
    $users = Import-Csv -Path $inputFilePath -Header "UserPrincipalName" -ErrorAction SilentlyContinue
    
    $totalUsers = $users.Count
    $userCount = 0

    if ($totalUsers -eq 0) {
        Write-Host "File CSV kosong." -ForegroundColor Yellow
    }
    
    Write-Host "Total ${totalUsers} pengguna ditemukan." -ForegroundColor Yellow

    foreach ($entry in $users) {
        $userCount++
        
        # Trim email untuk membersihkan spasi
        $upn = if ($entry.UserPrincipalName) { $entry.UserPrincipalName.Trim() } else { $null }
        
        if ([string]::IsNullOrWhiteSpace($upn)) { continue }

        # Update Progress Bar dengan kurung kurawal ${}
        Write-Progress -Activity "Generating Storage Report" `
                       -Status "Processing User ${userCount} of ${totalUsers}: ${upn}" `
                       -PercentComplete ([int](($userCount / $totalUsers) * 100))
        
        Write-Host "-> [${userCount}/${totalUsers}] Memproses: ${upn}..." -ForegroundColor White
        
        try {
            # 3.2.1. Validasi Keberadaan Mailbox
            $recipient = Get-Recipient -Identity $upn -ErrorAction Stop | Select-Object RecipientType

            if ($recipient.RecipientType -like "*UserMailbox*") {
                # 3.2.2. Ambil Statistik dan Quota
                $stats = Get-MailboxStatistics -Identity $upn -ErrorAction Stop | Select-Object TotalItemSize, ItemCount, LastLogonTime
                $mailbox = Get-Mailbox -Identity $upn -ErrorAction Stop | Select-Object DisplayName, ProhibitSendQuota, ProhibitSendReceiveQuota, IssueWarningQuota
                
                # Konversi Size ke GB secara aman (Handle Deserialized Object)
                $TotalItemSizeGB = try {
                    [math]::Round(($stats.TotalItemSize.Value.ToBytes() / 1GB), 2)
                } catch {
                    $rawSize = $stats.TotalItemSize.ToString()
                    if ($rawSize -match '(\d+\.?\d*)\s*(\w+)') {
                        $val = [double]$Matches[1]
                        $unit = $Matches[2].ToUpper()
                        switch ($unit) {
                            "KB" { [math]::Round($val / 1MB, 4) }
                            "MB" { [math]::Round($val / 1024, 2) }
                            "GB" { $val }
                            default { $rawSize }
                        }
                    } else { $rawSize }
                }
                
                $scriptOutput += [PSCustomObject]@{
                    UserPrincipalName = $upn
                    DisplayName       = $mailbox.DisplayName
                    TotalItemSizeGB   = $TotalItemSizeGB
                    ItemCount         = $stats.ItemCount
                    WarningQuota      = $mailbox.IssueWarningQuota.ToString()
                    SendQuota         = $mailbox.ProhibitSendQuota.ToString()
                    LastLogonTime     = $stats.LastLogonTime
                    Status            = "SUCCESS"
                    Reason            = "Storage data collected."
                }
                Write-Host "Size: ${TotalItemSizeGB} GB" -ForegroundColor DarkGreen

            } else {
                $reason = "Recipient type is $($recipient.RecipientType) (Not a UserMailbox)."
                Write-Host "Gagal: ${reason}" -ForegroundColor Yellow
                $scriptOutput += [PSCustomObject]@{
                    UserPrincipalName = $upn; Status = "FAIL"; Reason = $reason
                }
            }
        } 
        catch {
            $errMsg = $_.Exception.Message
            Write-Host "ERROR: ${errMsg}" -ForegroundColor Red
            $scriptOutput += [PSCustomObject]@{
                UserPrincipalName = $upn; Status = "FAIL"; Reason = $errMsg
            }
        }
    }
    Write-Progress -Activity "Storage Report" -Completed
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