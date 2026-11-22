# ==============================================================================
# HỆ THỐNG KÉO CODE LÊN GITHUB (KEO_LEN / PUSH)
# Tự động đẩy toàn bộ code, file chạy, template và báo cáo lên:
# https://github.com/bachhue128-create/SCT
# Chạy độc lập 100% trên mọi máy tính Windows (Không cần cài đặt Git)
# ==============================================================================

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Host.UI.RawUI.WindowTitle = "KÉO CODE VÀ DỮ LIỆU LÊN GITHUB (SCT)"

Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host "     FPT MVNO - HE THONG DAY CODE & DU LIEU LEN GITHUB (SCT)          " -ForegroundColor Yellow
Write-Host "======================================================================" -ForegroundColor Cyan

# 1. Xác định thư mục dự án
$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
if (-not $ScriptDir) { $ScriptDir = (Get-Location).Path }
$BaseDir = Split-Path -Parent $ScriptDir
if (-not (Test-Path (Join-Path $BaseDir "scripts"))) {
    $BaseDir = $ScriptDir
}

$username = "bachhue128-create"
$repoName = "SCT"
$tokenFile = Join-Path $ScriptDir "github_token.txt"

# 2. Lấy GitHub Token
$token = ""
if (Test-Path $tokenFile) {
    $token = (Get-Content $tokenFile -Raw).Trim()
}
if (-not $token -and $env:GITHUB_TOKEN) {
    $token = $env:GITHUB_TOKEN.Trim()
}
if (-not $token) {
    Write-Host "`n[NHAP TOKEN] Chua tim thay ma GitHub Personal Access Token." -ForegroundColor Yellow
    Write-Host "Vui long paste ma Token (ghp_...) vao day va nhan Enter:" -ForegroundColor Cyan
    $token = Read-Host
    if ($token) {
        $token = $token.Trim()
        Set-Content -Path $tokenFile -Value $token -Encoding UTF8
        Write-Host "      Da luu Token cho cac lan su dung tiep theo vao: $tokenFile" -ForegroundColor Green
    } else {
        Write-Host "[LOI] Token trong. Khong the day len GitHub!" -ForegroundColor Red
        pause
        exit 1
    }
}

$headers = @{
    "Authorization" = "Bearer $token"
    "Accept"        = "application/vnd.github.v3+json"
    "User-Agent"    = "Antigravity-Agent"
}

# 3. Kiểm tra kết nối tới repo
Write-Host "`n[1/4] Kiem tra ket noi voi GitHub Repository: $username/$repoName..." -ForegroundColor Cyan
$parentCommitSha = $null
try {
    $refRes = Invoke-RestMethod -Uri "https://api.github.com/repos/$username/$repoName/git/refs/heads/main" -Method Get -Headers $headers -ErrorAction Stop
    $parentCommitSha = $refRes.object.sha
    Write-Host "      Ket noi thanh cong! Commit gan nhat: $($parentCommitSha.Substring(0,7))" -ForegroundColor Green
} catch {
    Write-Host "[LOI] Khong the ket noi toi repo! Kiem tra lai token hoac internet: $_" -ForegroundColor Red
    pause
    exit 1
}

# 4. Quét các file cần đẩy lên
$files = Get-ChildItem -Path $BaseDir -Recurse -File | Where-Object {
    $_.FullName -notmatch "\\\.git\\" -and
    $_.FullName -notmatch "\\\.vscode\\" -and
    $_.Name -ne "github_token.txt" -and
    $_.Name -notmatch "\.tmp$"
}

$treeEntries = @()
$counter = 0
$total = $files.Count
Write-Host "`n[2/4] Dang upload $total files len GitHub..." -ForegroundColor Cyan

foreach ($f in $files) {
    $counter++
    $relPath = $f.FullName.Substring($BaseDir.Length).TrimStart("\", "/").Replace("\", "/")
    $fileSizeKB = [math]::Round($f.Length / 1KB, 1)
    Write-Host "      [$counter/$total] $relPath ($fileSizeKB KB)..." -NoNewline

    $bytes = [System.IO.File]::ReadAllBytes($f.FullName)
    $b64 = [Convert]::ToBase64String($bytes)
    
    $blobBody = @{
        content  = $b64
        encoding = "base64"
    } | ConvertTo-Json -Compress

    try {
        $blobRes = Invoke-RestMethod -Uri "https://api.github.com/repos/$username/$repoName/git/blobs" -Method Post -Headers $headers -Body $blobBody -ContentType "application/json; charset=utf-8"
        
        $treeEntries += @{
            path = $relPath
            mode = "100644"
            type = "blob"
            sha  = $blobRes.sha
        }
        Write-Host " OK" -ForegroundColor Green
    } catch {
        Write-Host " LOI: $_" -ForegroundColor Red
    }
}

# 5. Tạo Git Tree
Write-Host "`n[3/4] Dang tao Git Tree moi..." -ForegroundColor Cyan
$treeBody = @{
    tree = $treeEntries
} | ConvertTo-Json -Depth 5 -Compress

$treeRes = Invoke-RestMethod -Uri "https://api.github.com/repos/$username/$repoName/git/trees" -Method Post -Headers $headers -Body $treeBody -ContentType "application/json; charset=utf-8"

# 6. Tạo Commit & Cập nhật nhánh main
Write-Host "`n[4/4] Dang tao Commit va cap nhat nhanh main..." -ForegroundColor Cyan
$nowStr = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
$commitBody = @{
    message = "Cap nhat du lieu tu dong ($nowStr)"
    tree    = $treeRes.sha
    parents = @($parentCommitSha)
} | ConvertTo-Json -Compress

$commitRes = Invoke-RestMethod -Uri "https://api.github.com/repos/$username/$repoName/git/commits" -Method Post -Headers $headers -Body $commitBody -ContentType "application/json; charset=utf-8"

$patchBody = @{
    sha   = $commitRes.sha
    force = $true
} | ConvertTo-Json -Compress

$finalRef = Invoke-RestMethod -Uri "https://api.github.com/repos/$username/$repoName/git/refs/heads/main" -Method Patch -Headers $headers -Body $patchBody -ContentType "application/json; charset=utf-8"

Write-Host "`n======================================================================" -ForegroundColor Green
Write-Host " [THANH CONG] DA DAY TOAN BO DU AN LEN GITHUB!" -ForegroundColor Green
Write-Host " Repository URL: https://github.com/$username/$repoName" -ForegroundColor Yellow
Write-Host " Commit SHA    : $($commitRes.sha)" -ForegroundColor Yellow
Write-Host " Thoi gian     : $nowStr" -ForegroundColor Yellow
Write-Host "======================================================================" -ForegroundColor Green
