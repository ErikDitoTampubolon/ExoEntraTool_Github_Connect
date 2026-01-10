# =========================================================================  
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.1)  
# Deskripsi: Mengekspor daftar Full Access Permission (Seluruh User & Empty)
# =========================================================================  

# Variabel Global dan Output  
$scriptName = "Export_All_Mailboxes_FullAccess_Status" 
$scriptOutput = @() # Array tempat semua data hasil skrip dikumpulkan  

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }  
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"  

$outputFileName = "Output_$($scriptName)_$($timestamp).csv"  

# Definisi parentDir (2 tingkat di atas lokasi skrip dijalankan)
$parentDir = (Get-Item $scriptDir).Parent.Parent.FullName  


## ==========================================================================  
#                           INFORMASI SCRIPT                  
## ==========================================================================  

Write-Host "`n================================================" -ForegroundColor Yellow  
Write-Host "                INFORMASI SCRIPT                " -ForegroundColor Yellow  
Write-Host "================================================" -ForegroundColor Yellow  
Write-Host " Nama Skrip        : $scriptName.ps1" -ForegroundColor Yellow  
Write-Host " Field Kolom       : MailboxIdentity, DelegateUser, AccessRights, Status" -ForegroundColor Yellow  
Write-Host " Deskripsi Singkat : Mengekspor status Full Access untuk SELURUH Mailbox" -ForegroundColor Cyan  
Write-Host "                     (Termasuk yang tidak memiliki delegasi)." -ForegroundColor Cyan  
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

Write-Host "`n--- 3. Memulai Analisis Seluruh Mailbox ---" -ForegroundColor Magenta

# Mengambil semua mailbox kategori User dan Shared
$mailboxesToProcess = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox,SharedMailbox

$counter = 0
$total = $mailboxesToProcess.Count

Write-Host "Ditemukan total $total mailbox. Memulai audit..." -ForegroundColor White

foreach ($mbx in $mailboxesToProcess) {
    $counter++
    $id = $mbx.UserPrincipalName
    
    Write-Progress -Activity "Menganalisis Izin Mailbox" -Status "Memproses: $id ($counter/$total)" -PercentComplete (($counter/$total)*100)

    try {
        # Ambil izin dan filter user sistem/internal
        $permissions = Get-MailboxPermission -Identity $id -ErrorAction Stop | Where-Object { 
            ($_.AccessRights -match "FullAccess") -and 
            ($_.User -notmatch "NT AUTHORITY|MSOL|AdminAudit|ExchangeBackEnd") -and
            ($_.User -ne $id) # Abaikan jika user adalah dirinya sendiri (Self)
        }

        if ($permissions) {
            # Jika ditemukan user yang memiliki Full Access
            foreach ($perm in $permissions) {
                $scriptOutput += [PSCustomObject]@{
                    MailboxIdentity = $id
                    DelegateUser    = $perm.User
                    AccessRights    = $perm.AccessRights -join ", "
                    Status          = "Has FullAccess"
                }
            }
        } else {
            # Jika TIDAK ADA user luar yang memiliki Full Access
            $scriptOutput += [PSCustomObject]@{
                MailboxIdentity = $id
                DelegateUser    = "-"
                AccessRights    = "-"
                Status          = "No FullAccess Assigned"
            }
        }
    } catch {
        Write-Host "[SKIP] Gagal mengakses mailbox: $id" -ForegroundColor Red
    }
}


## ==========================================================================  
##                               EKSPOR HASIL  
## ==========================================================================  

if ($scriptOutput.Count -gt 0) {  
    $exportFolderName = "exported_data"  
    $exportFolderPath = Join-Path -Path $parentDir -ChildPath $exportFolderName  

    if (-not (Test-Path -Path $exportFolderPath)) {   
        New-Item -Path $exportFolderPath -ItemType Directory | Out-Null   
    }  

    $resultsFilePath = Join-Path -Path $exportFolderPath -ChildPath $outputFileName  

    $scriptOutput | Export-Csv -Path $resultsFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8  
    Write-Host "`n[SELESAI] Laporan berhasil dibuat." -ForegroundColor Green
    Write-Host "Lokasi file: ${resultsFilePath}" -ForegroundColor Cyan  
}