# ===================================================================
# Auto Updater untuk main-app-DEV (FIXED PATH & OUTPUT)
# ===================================================================

# 1. Penentuan Path yang Stabil (PSScriptRoot Guard)
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }

# Konfigurasi Path Absolut agar tidak lari ke C:\
$GitHubRawUrl = "https://raw.githubusercontent.com/ErikDitoTampubolon/ExoEntraTool_Github_Connect/dev/script/main-app-DEV.ps1"
$GitHubIconUrl = "https://raw.githubusercontent.com/ErikDitoTampubolon/ExoEntraTool_Github_Connect/dev/script/logo.ico"

$LocalScriptPath = Join-Path -Path $scriptDir -ChildPath "main-app-DEV.ps1"
$LocalIconPath   = Join-Path -Path $scriptDir -ChildPath "logo.ico"
$OutputExePath   = Join-Path -Path $scriptDir -ChildPath "ExoEntraTool.exe"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Auto Updater - main-app-DEV" -ForegroundColor Cyan
Write-Host "  Lokasi: $scriptDir" -ForegroundColor Gray
Write-Host "========================================" -ForegroundColor Cyan

# ===================================================================
# Step 1: Cek dan Install Module PS2EXE
# ===================================================================
Write-Host "`n[1/5] Memeriksa modul PS2EXE..." -ForegroundColor Yellow
if (-not (Get-Module -ListAvailable -Name PS2EXE)) {
    Write-Host "      Menginstal PS2EXE..." -ForegroundColor Yellow
    Install-Module -Name PS2EXE -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
}
Import-Module PS2EXE -Force
Write-Host "      Modul PS2EXE siap." -ForegroundColor Green

# ===================================================================
# Step 2: Download Logo
# ===================================================================
Write-Host "`n[2/5] Mendownload logo.ico..." -ForegroundColor Yellow
try {
    Invoke-WebRequest -Uri $GitHubIconUrl -OutFile $LocalIconPath -UseBasicParsing -ErrorAction SilentlyContinue
    if (Test-Path $LocalIconPath) { Write-Host "      Logo berhasil diunduh." -ForegroundColor Green }
} catch {
    Write-Host "      Peringatan: Gagal download logo, lanjut tanpa icon." -ForegroundColor Yellow
}

# ===================================================================
# Step 3: Download Script Utama
# ===================================================================
Write-Host "`n[3/5] Mendownload main-app-DEV.ps1..." -ForegroundColor Yellow
try {
    if (Test-Path $LocalScriptPath) { Remove-Item $LocalScriptPath -Force }
    Invoke-WebRequest -Uri $GitHubRawUrl -OutFile $LocalScriptPath -UseBasicParsing -ErrorAction Stop
    Write-Host "      Script terbaru berhasil diunduh." -ForegroundColor Green
} catch {
    Write-Host "      ERROR: Gagal download script utama!" -ForegroundColor Red
    pause; exit 1
}

# ===================================================================
# Step 4: Konversi ke EXE (Splatting Method)
# ===================================================================
Write-Host "`n[4/5] Mengkonversi ke .exe..." -ForegroundColor Yellow
try {
    # Hapus EXE lama jika masih ada (dan tidak sedang berjalan)
    if (Test-Path $OutputExePath) { 
        Remove-Item $OutputExePath -Force -ErrorAction SilentlyContinue 
    }

    $ps2exeParams = @{
        inputFile  = $LocalScriptPath
        outputFile = $OutputExePath
    }

    if (Test-Path $LocalIconPath) {
        $ps2exeParams.Add("iconFile", $LocalIconPath)
    }

    # Eksekusi konversi
    Invoke-ps2exe @ps2exeParams

    if (Test-Path $OutputExePath) {
        Write-Host "      BERHASIL: File EXE dibuat di $OutputExePath" -ForegroundColor Green
    } else {
        throw "File EXE tidak muncul setelah proses konversi."
    }
} catch {
    Write-Host "      ERROR Konversi: $($_.Exception.Message)" -ForegroundColor Red
    pause; exit 1
}

# ===================================================================
# Step 5: Cleanup
# ===================================================================
Write-Host "`n[5/5] Menghapus file sementara..." -ForegroundColor Yellow
Remove-Item $LocalScriptPath -Force -ErrorAction SilentlyContinue
# Icon tetap dihapus agar folder bersih sesuai permintaan Anda
Remove-Item $LocalIconPath -Force -ErrorAction SilentlyContinue 

Write-Host "`nUpdate Selesai! Silakan jalankan ExoEntraTool.exe" -ForegroundColor Green
pause