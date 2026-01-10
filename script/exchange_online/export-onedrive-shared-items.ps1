# =========================================================================  
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.1)  
# Deskripsi: Mengekspor daftar File/Folder OneDrive yang dibagikan (Shared)
# =========================================================================  

# Variabel Global dan Output  
$scriptName = "OneDrive_Shared_Items_Report" 
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
Write-Host " Field Kolom       : Owner, ItemName, WebUrl, SharedWith, ShareType" -ForegroundColor Yellow  
Write-Host " Deskripsi Singkat : Mendeteksi file/folder OneDrive yang memiliki" -ForegroundColor Cyan  
Write-Host "                     izin berbagi (Shared) ke pihak lain." -ForegroundColor Cyan  
Write-Host "==========================================================" -ForegroundColor Yellow  


## ==========================================================================  
#                           KONFIRMASI EKSEKUSI  
## ==========================================================================  

$confirmation = Read-Host "Apakah Anda ingin menjalankan skrip ini? (Y/N)"  
if ($confirmation -ne "Y") {  
    Write-Host "`nEksekusi skrip dibatalkan oleh pengguna." -ForegroundColor Red  
    return  
}  


# ==========================================================================  
# PILIHAN METODE INPUT (CSV vs ALL USERS)  
# ==========================================================================  
$useAllUsers = $false  
$validInput = $false  
while (-not $validInput) {  
    Write-Host "`n--- Metode Input Data ---" -ForegroundColor Blue  
    Write-Host "1. Gunakan Daftar Email dari File CSV" -ForegroundColor Cyan  
    Write-Host "2. Proses Seluruh Pengguna (All Users) di Tenant" -ForegroundColor Cyan  
    $inputMethod = Read-Host "`nPilih metode (1/2)"  
    if ($inputMethod -eq "1") { $useAllUsers = $false; $validInput = $true }  
    elseif ($inputMethod -eq "2") { $useAllUsers = $true; $validInput = $true; Write-Host "[OK] Mode: Seluruh Pengguna." -ForegroundColor Green }  
    else { Write-Host "[ERROR] Pilihan tidak valid!" -ForegroundColor Red }  
}  

if (-not $useAllUsers) {  
    $csvFiles = Get-ChildItem -Path $parentDir -Filter "*.csv"  
    if ($csvFiles.Count -eq 0) {  
        $newFileName = "daftar_email.csv"  
        $newFilePath = Join-Path -Path $parentDir -ChildPath $newFileName  
        "UserPrincipalName" | Out-File -FilePath $newFilePath -Encoding utf8  
        Write-Host "Silakan isi file CSV: $newFileName" -ForegroundColor Yellow  
        Start-Process notepad.exe $newFilePath ; return  
    } else {  
        Write-Host "`nPilih file CSV:" -ForegroundColor Yellow  
        for ($i = 0; $i -lt $csvFiles.Count; $i++) { Write-Host "$($i + 1). $($csvFiles[$i].Name)" -ForegroundColor Cyan }  
        $fileChoice = Read-Host "Nomor file"  
        $selectedFile = $csvFiles[[int]$fileChoice - 1]  
        $inputFilePath = $selectedFile.FullName  
    }  
}


## ==========================================================================  
#                     PRASYARAT DAN INSTALASI MODUL  
## ==========================================================================  

Write-Host "`n--- 1. Memeriksa Lingkungan PowerShell ---" -ForegroundColor Blue  
Set-ExecutionPolicy RemoteSigned -Scope Process -Force -ErrorAction SilentlyContinue  

function CheckAndInstallModule {  
    param([string]$ModuleName)  
    if (-not (Get-Module -Name $ModuleName -ListAvailable)) {  
        Install-Module -Name $ModuleName -Force -AllowClobber -Scope CurrentUser  
    }  
}  

CheckAndInstallModule -ModuleName "Microsoft.Graph"  


## ==========================================================================  
##                    KONEKSI KE LAYANAN MICROSOFT
## ==========================================================================  

Write-Host "`n--- 2. Membangun Koneksi ke Microsoft Graph ---" -ForegroundColor Blue  
try {  
    $requiredScopes = "User.Read.All", "Files.Read.All", "Sites.Read.All"  
    Connect-MgGraph -Scopes $requiredScopes -NoWelcome  
} catch {  
    Write-Error "Gagal Login: $($_.Exception.Message)" ; exit 1  
}  


## ==========================================================================
##                          LOGIKA UTAMA SCRIPT
## ==========================================================================

Write-Host "`n--- 3. Memulai Scan Shared Items di OneDrive ---" -ForegroundColor Magenta

if ($useAllUsers) {  
    $users = Get-MgUser -All -Property "UserPrincipalName", "Id"  
} else {  
    $csvData = Import-Csv -Path $inputFilePath  
    $users = foreach ($row in $csvData) { Get-MgUser -UserId $row.UserPrincipalName -ErrorAction SilentlyContinue }  
}  

foreach ($user in $users) {  
    if ($null -eq $user) { continue }
    Write-Host "Memproses OneDrive: $($user.UserPrincipalName)..." -ForegroundColor Gray  
    
    try {  
        $userDrive = Get-MgUserDrive -UserId $user.Id -ErrorAction Stop | Where-Object { $_.DriveType -eq "business" } | Select-Object -First 1
        
        if (-not $userDrive) { continue }

        # Mengambil item root
        $allItems = Get-MgUserDriveItemChild -UserId $user.Id -DriveId $userDrive.Id -DriveItemId "root" -All -ErrorAction Stop
        $sharedItems = $allItems | Where-Object { $_.Shared -ne $null }

        foreach ($item in $sharedItems) {  
            $permissions = Get-MgUserDriveItemPermission -UserId $user.Id -DriveId $userDrive.Id -DriveItemId $item.Id -ErrorAction SilentlyContinue
            
            foreach ($perm in $permissions) {  
                if ($perm.Roles -notcontains "owner") {  
                    
                    # --- LOGIKA EKSTRAKSI EMAIL ---
                    $targetEmail = "N/A"
                    if ($perm.Link) {
                        $targetEmail = "Sharing Link ($($perm.Link.Scope))"
                    } elseif ($perm.GrantedToV2.User.AdditionalProperties.email) {
                        # Email untuk akun Guest/Eksternal
                        $targetEmail = $perm.GrantedToV2.User.AdditionalProperties.email
                    } elseif ($perm.GrantedToV2.User.UserPrincipalName) {
                        # Email untuk akun Internal
                        $targetEmail = $perm.GrantedToV2.User.UserPrincipalName
                    } elseif ($perm.GrantedToV2.Group.AdditionalProperties.email) {
                        # Email jika dibagikan ke Grup
                        $targetEmail = "[Group] " + $perm.GrantedToV2.Group.AdditionalProperties.email
                    }

                    $scriptOutput += [PSCustomObject]@{  
                        Owner            = $user.UserPrincipalName  
                        ItemName         = $item.Name  
                        Type             = if ($item.Folder) { "Folder" } else { "File" }  
                        SharedWith_Name  = $perm.GrantedToV2.User.DisplayName
                        SharedWith_Email = $targetEmail
                        Role             = $perm.Roles -join ", "  
                        WebUrl           = $item.WebUrl  
                        LastModified     = $item.LastModifiedDateTime
                    }  
                }  
            }  
        }  
    } catch {  
        Write-Host "  [!] Error pada $($user.UserPrincipalName): $($_.Exception.Message)" -ForegroundColor Red  
    }  
}


## ==========================================================================  
##                               EKSPOR HASIL  
## ==========================================================================  

if ($scriptOutput.Count -gt 0) {  
    $exportFolderPath = Join-Path -Path $parentDir -ChildPath "exported_data"  
    if (-not (Test-Path $exportFolderPath)) { New-Item $exportFolderPath -ItemType Directory }  
    $resultsFilePath = Join-Path -Path $exportFolderPath -ChildPath $outputFileName  
    $scriptOutput | Export-Csv -Path $resultsFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8  
    Write-Host "`nLaporan selesai! Tersimpan di: ${resultsFilePath}" -ForegroundColor Cyan  
} else {  
    Write-Host "`nTidak ditemukan item yang dibagikan." -ForegroundColor Yellow  
}  