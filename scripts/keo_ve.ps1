# ==============================================================================
# HỆ THỐNG KÉO CODE TỪ GITHUB VỀ MÁY (KEO_VE)
# Tự động tải bản mới nhất từ https://github.com/bachhue128-create/SCT
# Chạy độc lập 100% trên mọi máy tính Windows (Không cần cài đặt Git)
# ==============================================================================

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Host.UI.RawUI.WindowTitle = "KÉO CODE VÀ DỮ LIỆU TỪ GITHUB VỀ MÁY"

Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host "     FPT MVNO - HE THONG KEO CODE & DU LIEU TU GITHUB VE MAY          " -ForegroundColor Yellow
Write-Host "======================================================================" -ForegroundColor Cyan

# 1. Xác định thư mục hiện tại
$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
if (-not $ScriptDir) { $ScriptDir = (Get-Location).Path }
$BaseDir = Split-Path -Parent $ScriptDir
if (-not (Test-Path (Join-Path $BaseDir "scripts"))) {
    $BaseDir = $ScriptDir
}

Write-Host "`n[1/3] Dang ket noi toi GitHub Repository..." -ForegroundColor Cyan
$repoUrl = "https://github.com/bachhue128-create/SCT/archive/refs/heads/main.zip"
$tempZip = Join-Path $env:TEMP "SCT_main_latest.zip"
$tempExtract = Join-Path $env:TEMP "SCT_extract"

if (Test-Path $tempZip) { Remove-Item $tempZip -Force -ErrorAction SilentlyContinue }
if (Test-Path $tempExtract) { Remove-Item $tempExtract -Recurse -Force -ErrorAction SilentlyContinue }

# 2. Tải file ZIP từ GitHub
Write-Host "[2/3] Dang tai ban cap nhat moi nhat tu GitHub ve..." -ForegroundColor Cyan
try {
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
    $webClient = New-Object System.Net.WebClient
    $webClient.Headers.Add("User-Agent", "PowerShell-Downloader")
    $webClient.DownloadFile($repoUrl, $tempZip)
    Write-Host "      Tai ve thanh cong ($([math]::Round((Get-Item $tempZip).Length / 1KB, 1)) KB)" -ForegroundColor Green
} catch {
    Write-Host "[LOI] Khong the tai file tu GitHub: $_" -ForegroundColor Red
    pause
    exit 1
}

# 3. Giải nén và cập nhật vào thư mục hiện tại
Write-Host "[3/3] Dang dong bo du lieu vao may..." -ForegroundColor Cyan
try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::ExtractToDirectory($tempZip, $tempExtract)
    
    $extractedFolder = Get-ChildItem -Path $tempExtract -Directory | Select-Object -First 1
    if ($extractedFolder) {
        $sourcePath = $extractedFolder.FullName
        
        # Sao chép tất cả các thư mục và file mới về thư mục dự án
        Get-ChildItem -Path $sourcePath -Recurse | ForEach-Object {
            $rel = $_.FullName.Substring($sourcePath.Length).TrimStart("\", "/")
            $dest = Join-Path $BaseDir $rel
            
            if ($_.PSIsContainer) {
                if (-not (Test-Path $dest)) { New-Item -ItemType Directory -Path $dest -Force | Out-Null }
            } else {
                $destParent = Split-Path -Parent $dest
                if (-not (Test-Path $destParent)) { New-Item -ItemType Directory -Path $destParent -Force | Out-Null }
                Copy-Item -Path $_.FullName -Destination $dest -Force
            }
        }
    }
    
    # Dọn dẹp temp
    Remove-Item $tempZip -Force -ErrorAction SilentlyContinue
    Remove-Item $tempExtract -Recurse -Force -ErrorAction SilentlyContinue

    Write-Host "`n======================================================================" -ForegroundColor Green
    Write-Host " [THANH CONG] DA KEO TOAN BO CODE VA DU LIEU MOI NHAT VE MAY!" -ForegroundColor Green
    Write-Host " Thu muc hien tai: $BaseDir" -ForegroundColor Yellow
    Write-Host "======================================================================" -ForegroundColor Green
} catch {
    Write-Host "[LOI] Khong the giai nen va cap nhat: $_" -ForegroundColor Red
    pause
    exit 1
}
