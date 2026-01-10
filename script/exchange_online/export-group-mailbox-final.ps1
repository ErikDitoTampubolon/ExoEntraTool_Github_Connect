# =========================================================================  
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.2)  
# Deskripsi: Menarik data M365 Groups, Distribution List, Dynamic DL, 
#            dan Mail-enabled Security Groups.
# =========================================================================  

# Variabel Global dan Output  
$scriptName = "Export_All_Group_Types" 
$scriptOutput = @() 

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }  
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"  
$outputFileName = "Output_$($scriptName)_$($timestamp).csv"  

# Definisi parentDir (2 tingkat di atas)  
$parentDir = (Get-Item $scriptDir).Parent.Parent.FullName  

## ==========================================================================  
#                           INFORMASI SCRIPT                  
## ==========================================================================  

Write-Host "`n================================================" -ForegroundColor Yellow  
Write-Host "                INFORMASI SCRIPT                " -ForegroundColor Yellow  
Write-Host "================================================" -ForegroundColor Yellow  
Write-Host " Nama Skrip        : $scriptName.ps1" -ForegroundColor Yellow  
Write-Host " Tipe Group        : M365, DL, Dynamic DL, Security" -ForegroundColor Cyan  
Write-Host " Deskripsi Singkat : Ekspor semua tipe grup yang memiliki mailbox/email." -ForegroundColor Cyan  
Write-Host "==========================================================" -ForegroundColor Yellow  

## ==========================================================================  
#                           KONFIRMASI EKSEKUSI  
## ==========================================================================  

$confirmation = Read-Host "Mulai tarik data dari seluruh tipe grup? (Y/N)"  
if ($confirmation -ne "Y") {  
    Write-Host "`nEksekusi dibatalkan." -ForegroundColor Red  
    return  
}  

# ## ==========================================================================  
# #                     PRASYARAT DAN INSTALASI MODUL  
# ## ==========================================================================  

# Write-Host "`n--- 1. Memeriksa Lingkungan PowerShell ---" -ForegroundColor Blue  
# Set-ExecutionPolicy RemoteSigned -Scope Process -Force -ErrorAction SilentlyContinue  

# if (Get-Module -Name "ExchangeOnlineManagement" -ListAvailable) {  
#     Write-Host " Modul ExchangeOnlineManagement tersedia." -ForegroundColor Green  
# } else {  
#     Write-Host " Menginstal modul ExchangeOnlineManagement..." -ForegroundColor Yellow  
#     Install-Module -Name "ExchangeOnlineManagement" -Force -AllowClobber -Scope CurrentUser  
# }  

# ## ==========================================================================  
# ##                    KONEKSI KE LAYANAN
# ## ==========================================================================  

# Write-Host "`n--- 2. Membangun Koneksi ke Exchange Online ---" -ForegroundColor Blue  
# try {  
#     Connect-ExchangeOnline -ShowProgress $true -ErrorAction Stop  
#     Write-Host "Koneksi Berhasil." -ForegroundColor Green  
# } catch {  
#     Write-Error "Gagal terhubung: $($_.Exception.Message)"  
#     exit 1  
# }  

## ==========================================================================
##                          LOGIKA UTAMA SCRIPT
## ==========================================================================

Write-Host "`n--- 3. Memulai Penarikan Data Multi-Tipe Grup ---" -ForegroundColor Magenta

# 3.1 Ambil Microsoft 365 Groups
$m365Groups = Get-UnifiedGroup -ResultSize Unlimited | Select-Object DisplayName, PrimarySmtpAddress, @{N="GroupType";E={"Microsoft 365"}}, ManagedBy, RecipientTypeDetails, WhenCreated, AccessType

# 3.2 Ambil Distribution Lists
$dlGroups = Get-DistributionGroup -ResultSize Unlimited | Where-Object {$_.RecipientTypeDetails -eq "MailUniversalDistributionGroup"} | Select-Object DisplayName, PrimarySmtpAddress, @{N="GroupType";E={"Distribution List"}}, ManagedBy, RecipientTypeDetails, WhenCreated, @{N="AccessType";E={"N/A"}}

# 3.3 Ambil Dynamic Distribution Lists
$dynamicDL = Get-DynamicDistributionGroup -ResultSize Unlimited | Select-Object DisplayName, PrimarySmtpAddress, @{N="GroupType";E={"Dynamic Distribution List"}}, ManagedBy, RecipientTypeDetails, WhenCreated, @{N="AccessType";E={"Dynamic"}}

# 3.4 Ambil Mail-enabled Security Groups
$securityGroups = Get-DistributionGroup -ResultSize Unlimited | Where-Object {$_.RecipientTypeDetails -eq "MailUniversalSecurityGroup"} | Select-Object DisplayName, PrimarySmtpAddress, @{N="GroupType";E={"Mail-enabled Security"}}, ManagedBy, RecipientTypeDetails, WhenCreated, @{N="AccessType";E={"Security"}}

# Menggabungkan semua hasil
$allGroups = $m365Groups + $dlGroups + $dynamicDL + $securityGroups
$total = ($allGroups | Measure-Object).Count

foreach ($g in $allGroups) {
    $object = [PSCustomObject]@{
        DisplayName          = $g.DisplayName
        EmailAddress         = $g.PrimarySmtpAddress
        GroupType            = $g.GroupType
        AccessType            = $g.AccessType
        RecipientTypeDetails = $g.RecipientTypeDetails
        CreatedDate          = $g.WhenCreated
    }
    $scriptOutput += $object
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
    
    Write-Host "`n==========================================================" -ForegroundColor Green
    Write-Host "PROSES SELESAI" -ForegroundColor Green
    Write-Host "Total Grup Ditemukan: $total"
    Write-Host "Lokasi File: ${resultsFilePath}" -ForegroundColor Cyan  
    Write-Host "==========================================================" -ForegroundColor Green
} else {
    Write-Host "`n[!] Tidak ada data grup yang ditemukan." -ForegroundColor Yellow
}