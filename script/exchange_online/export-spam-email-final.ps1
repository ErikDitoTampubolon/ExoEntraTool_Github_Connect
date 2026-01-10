# =========================================================================  
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.3)  
# Detail Per Email: Tanggal, Sender, Recipient, Subject, Status
# =========================================================================  

# Variabel Global dan Output  
$scriptName = "Detail_Spam_Tenant" 
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
Write-Host " Detail Kolom      : Date, Sender, Recipient, Subject, Status" -ForegroundColor Yellow  
Write-Host " Engine            : Get-MessageTraceV2" -ForegroundColor Yellow  
Write-Host "==========================================================" -ForegroundColor Yellow  

## ==========================================================================  
#                           KONFIRMASI EKSEKUSI  
## ==========================================================================  

$confirmation = Read-Host "Jalankan pemindaian detail spam untuk seluruh tenant? (Y/N)"  
if ($confirmation -ne "Y") { exit }  

## ==========================================================================  
#                     PRASYARAT DAN MODUL  
## ==========================================================================  

Write-Host "`n--- 1. Memeriksa Lingkungan ---" -ForegroundColor Blue  
if (-not (Get-Module -Name "ExchangeOnlineManagement" -ListAvailable)) {  
    Write-Host "Menginstal modul ExchangeOnlineManagement..." -ForegroundColor Yellow  
    Install-Module -Name "ExchangeOnlineManagement" -Force -AllowClobber -Scope CurrentUser  
}  

## ==========================================================================  
##                    KONEKSI KE LAYANAN
## ==========================================================================  

Write-Host "`n--- 2. Membangun Koneksi ---" -ForegroundColor Blue  
if (-not (Get-PSSession | Where-Object {$_.ConfigurationName -eq "Microsoft.Exchange"})) {  
    Connect-ExchangeOnline -ShowProgress $false | Out-Null  
}

## ==========================================================================
##                          LOGIKA UTAMA SCRIPT
## ==========================================================================

Write-Host "`n--- 3. Memulai Pemindaian Detail Spam ---" -ForegroundColor Magenta

# Rentang Waktu (10 Hari Terakhir)
$endDate = Get-Date
$startDate = $endDate.AddDays(-10)

Write-Host "Mengambil daftar mailbox pengguna..." -ForegroundColor Cyan
$targetUsersList = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox | Select-Object -ExpandProperty UserPrincipalName
$totalTargets = $targetUsersList.Count

$counter = 0

foreach ($userEmail in $targetUsersList) {
    $counter++
    $percentComplete = ($counter / $totalTargets) * 100
    Write-Progress -Activity "Extracting Detail Spam" -Status "Checking: $userEmail ($counter/$totalTargets)" -PercentComplete $percentComplete
    
    try {
        # Mengambil data mentah dari Message Trace V2
        $traces = Get-MessageTraceV2 -RecipientAddress $userEmail -StartDate $startDate -EndDate $endDate -ErrorAction Stop
        
        # Filter email yang ditandai Spam atau Masuk Karantina
        $spamItems = $traces | Where-Object { $_.Status -eq "FilteredAsSpam" -or $_.Status -eq "Quarantined" }

        if ($spamItems) {
            foreach ($item in $spamItems) {
                $scriptOutput += [PSCustomObject]@{
                    ReceivedDate = $item.Received
                    Sender       = $item.SenderAddress
                    Recipient    = $item.RecipientAddress
                    Subject      = $item.Subject
                    Status       = $item.Status
                    MessageId    = $item.MessageId
                }
            }
            Write-Host " [!] $userEmail : Ditemukan $(($spamItems | Measure-Object).Count) item spam" -ForegroundColor Red
        }
    } catch {
        Write-Host " [ERR] Gagal scan $userEmail : $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

Write-Progress -Activity "Scanning" -Completed

## ==========================================================================  
##                               EKSPOR HASIL  
## ==========================================================================  

if ($scriptOutput.Count -gt 0) {  
    $exportFolderPath = Join-Path -Path $parentDir -ChildPath "exported_data"  
    if (-not (Test-Path -Path $exportFolderPath)) { New-Item -Path $exportFolderPath -ItemType Directory | Out-Null }  

    $resultsFilePath = Join-Path -Path $exportFolderPath -ChildPath $outputFileName  
    
    # Ekspor dengan detail per baris email
    $scriptOutput | Export-Csv -Path $resultsFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8  
    
    Write-Host "`n================================================" -ForegroundColor Yellow
    Write-Host "EKSPOR SELESAI" -ForegroundColor Green
    Write-Host "Total Item Spam : $($scriptOutput.Count)" -ForegroundColor Cyan
    Write-Host "Lokasi File     : $resultsFilePath" -ForegroundColor Cyan  
    Write-Host "================================================" -ForegroundColor Yellow
} else {
    Write-Host "`nTidak ditemukan email spam dalam 10 hari terakhir." -ForegroundColor Green
}