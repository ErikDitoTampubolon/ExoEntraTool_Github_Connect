# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.0)
# Nama Skrip: Get-EntraUserAuthMethods
# Deskripsi: Menarik daftar metode autentikasi user dari CSV.
# =========================================================================

# Variabel Global dan Output
$scriptName = "AuthMethodsReport" 
$scriptOutput = New-Object System.Collections.Generic.List[PSCustomObject]

# Konfigurasi File Input (Pastikan file ini ada di folder yang sama dengan skrip)
$inputFileName = "UserPrincipalName.csv"
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$inputFilePath = Join-Path -Path $scriptDir -ChildPath $inputFileName

# Tentukan jalur dan nama file output dinamis
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputFileName = "Output_$($scriptName)_$($timestamp).csv"
$outputFilePath = Join-Path -Path $scriptDir -ChildPath $outputFileName

## ==========================================================================
#                           INFORMASI SCRIPT                
## ==========================================================================

Write-Host "`n================================================" -ForegroundColor Yellow
Write-Host "                INFORMASI SCRIPT                " -ForegroundColor Yellow
Write-Host "================================================" -ForegroundColor Yellow
Write-Host " Nama Skrip        : Get-EntraUserAuthMethods" -ForegroundColor Yellow
Write-Host " Field Kolom       : [UserPrincipalName]
                     [AuthenticationMethodId]
                     [DisplayName]
                     [AuthenticationMethodType]
                     [Status]
                     [ErrorMessage]" -ForegroundColor Yellow
Write-Host " Deskripsi Singkat : Script ini berfungsi untuk menarik daftar metode autentikasi yang dimiliki user berdasarkan daftar UPN dari file CSV tanpa header. Script menampilkan progres eksekusi di konsol, lalu mengekspor hasil ke folder 'exported_data' dua tingkat di atas direktori skrip." -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Yellow

## ==========================================================================
##                          LOGIKA UTAMA SCRIPT
## ==========================================================================

Write-Host "`n--- 3. Memulai Logika Utama Skrip: $($scriptName) ---" -ForegroundColor Magenta

if (Test-Path $inputFilePath) {
    # Mengimpor CSV tanpa header (menggunakan TempColumn) atau pastikan header bernama 'UserPrincipalName'
    $users = Import-Csv $inputFilePath -Header "TempColumn" | Where-Object { $_.TempColumn -ne $null -and $_.TempColumn.Trim() -ne "" }
    $totalUsers = $users.Count
    $counter = 0

    foreach ($row in $users) {
        $counter++
        $upn = $row.TempColumn.Trim()
        
        # UI Progres Baris Tunggal
        $statusText = "-> [$counter/$totalUsers] Memproses: $upn . . ."
        Write-Host "`r$statusText" -ForegroundColor Green -NoNewline

        try {
            # Mengambil metode autentikasi untuk user terkait
            $authMethods = Get-EntraUserAuthenticationMethod -UserId $upn -ErrorAction Stop
            
            foreach ($method in $authMethods) {
                $obj = [PSCustomObject]@{
                    UserPrincipalName        = $upn
                    AuthenticationMethodId   = $method.Id
                    DisplayName              = $method.DisplayName
                    AuthenticationMethodType = $method.AuthenticationMethodType
                    Status                   = "SUCCESS"
                }
                $scriptOutput.Add($obj)
            }
            
            # Jika user tidak memiliki metode yang terdaftar sama sekali
            if ($null -eq $authMethods) {
                $scriptOutput.Add([PSCustomObject]@{
                    UserPrincipalName = $upn
                    Status            = "NO_METHODS_FOUND"
                })
            }

        } catch {
            $scriptOutput.Add([PSCustomObject]@{
                UserPrincipalName = $upn
                Status            = "FAILED"
                ErrorMessage      = $_.Exception.Message
            })
        }
    }
    Write-Host "`n`nPemrosesan data selesai." -ForegroundColor Green
} else {
    Write-Host "ERROR: File '$inputFileName' tidak ditemukan di folder skrip!" -ForegroundColor Red
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