# ==============================================================================
# HỆ THỐNG CẬP NHẬT TỰ ĐỘNG INFOGRAPHIC THI ĐUA SIM FPT
# Tự động đọc dữ liệu từ sheet "Thi đua Sim" (File .xlsx hoặc .csv)
# Xác định Top 2 giải Nhất/Nhì (Màu xanh lá) & Top 3-5 bám đuổi (Màu vàng)
# Tương thích 100% mọi máy tính Windows (Tự nhận diện Chrome hoặc Microsoft Edge)
# ==============================================================================

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Host.UI.RawUI.WindowTitle = "FPT MVNO - Cập nhật Bảng Vàng Thi Đua SIM FPT"

Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host "     FPT MVNO - HỆ THỐNG CẬP NHẬT BẢNG VÀNG THI ĐUA SIM FPT 3 TUỔI    " -ForegroundColor Yellow
Write-Host "======================================================================" -ForegroundColor Cyan

# 1. Xác định đường dẫn thư mục an toàn (Dùng PSScriptRoot để chạy được trên mọi máy)
$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
if (-not $ScriptDir) { $ScriptDir = (Get-Location).Path }
$BaseDir   = Split-Path -Parent $ScriptDir
$InputDir  = Join-Path $BaseDir "input_data"
$OutputDir = Join-Path $BaseDir "output"
$TplPath   = Join-Path $BaseDir "template\infographic_template.html"
$AssetsDir = Join-Path $BaseDir "template\assets"

if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }

# 2. Tìm file Excel (.xlsx hoặc .csv) mới nhất
$excelFiles = Get-ChildItem -Path $InputDir -Include *.xlsx, *.csv -Recurse -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending
if (-not $excelFiles -or $excelFiles.Count -eq 0) {
    # Thử tìm ở thư mục gốc
    $excelFiles = Get-ChildItem -Path $BaseDir -Include *.xlsx, *.csv -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending
}

if (-not $excelFiles -or $excelFiles.Count -eq 0) {
    Write-Host "`n[LỖI] Không tìm thấy file Excel hoặc CSV nào trong thư mục input_data!" -ForegroundColor Red
    Write-Host "Vui lòng copy file Excel vào thư mục: $InputDir" -ForegroundColor Yellow
    exit 1
}

$targetFile = $excelFiles[0].FullName
Write-Host "`n[1/4] Đang đọc file dữ liệu: $(Split-Path -Leaf $targetFile)" -ForegroundColor Green

# 3. Đọc dữ liệu từ file Excel .xlsx
Add-Type -AssemblyName System.IO.Compression.FileSystem

$shops = @()
$reportDate = ""

if ($targetFile.EndsWith(".xlsx", [System.StringComparison]::OrdinalIgnoreCase)) {
    $tempCopy = Join-Path $env:TEMP ("excel_read_" + [Guid]::NewGuid().ToString("N") + ".xlsx")
    $zip = $null
    try {
        # Dùng FileShare.ReadWrite để copy và đọc được file ngay cả khi người dùng đang mở xem trong Excel
        $fsSrc = New-Object System.IO.FileStream($targetFile, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        $fsDst = New-Object System.IO.FileStream($tempCopy, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
        $fsSrc.CopyTo($fsDst)
        $fsSrc.Close()
        $fsDst.Close()
        $zip = [System.IO.Compression.ZipFile]::OpenRead($tempCopy)
    } catch {
        Write-Host "`n======================================================================" -ForegroundColor Red
        Write-Host " [LỖI TRUY CẬP FILE EXCEL] Không thể đọc được file Excel!" -ForegroundColor Red
        Write-Host " File: $(Split-Path -Leaf $targetFile)" -ForegroundColor Yellow
        Write-Host " Nguyên nhân: File Excel đang bị khóa (có thể do đang mở trong Excel hoặc đang lưu)." -ForegroundColor Yellow
        Write-Host " -> GIẢI PHÁP: VUI LÒNG ĐÓNG FILE EXCEL LẠI RỒI CHẠY LẠI FILE BAT!" -ForegroundColor Green
        Write-Host "======================================================================" -ForegroundColor Red
        if (Test-Path $tempCopy) { Remove-Item $tempCopy -Force -ErrorAction SilentlyContinue }
        pause
        exit 1
    }
    
    # Đọc workbook.xml để tìm sheet "Thi đua Sim"
    $wbEntry = $zip.GetEntry("xl/workbook.xml")
    $wbReader = New-Object System.IO.StreamReader($wbEntry.Open())
    $wbXml = [xml]$wbReader.ReadToEnd()
    $wbReader.Close()
    
    $sheetTargetNode = $wbXml.workbook.sheets.sheet | Where-Object { 
        $_.name -like "*Thi đua Sim*" -or $_.name -like "*Thi dua Sim*" -or $_.name -like "*Sim*"
    } | Select-Object -First 1

    if (-not $sheetTargetNode) {
        Write-Host "[LỖI] Không tìm thấy sheet 'Thi đua Sim' trong file Excel!" -ForegroundColor Red
        if ($zip) { $zip.Dispose() }
        if (Test-Path $tempCopy) { Remove-Item $tempCopy -Force -ErrorAction SilentlyContinue }
        pause
        exit 1
    }
    
    $rId = $sheetTargetNode.id
    Write-Host "      Đã tìm thấy sheet: '$($sheetTargetNode.name)' (ID: $rId)" -ForegroundColor DarkCyan
    
    # Tìm file xml tương ứng từ _rels
    $relsEntry = $zip.GetEntry("xl/_rels/workbook.xml.rels")
    $relsReader = New-Object System.IO.StreamReader($relsEntry.Open())
    $relsXml = [xml]$relsReader.ReadToEnd()
    $relsReader.Close()
    
    $sheetRel = $relsXml.Relationships.Relationship | Where-Object { $_.Id -eq $rId }
    $sheetXmlPath = "xl/" + $sheetRel.Target.Replace("\", "/")
    $sheetEntry = $zip.GetEntry($sheetXmlPath)
    
    if (-not $sheetEntry) {
        $sheetXmlPath = "xl/worksheets/" + (Split-Path -Leaf $sheetRel.Target)
        $sheetEntry = $zip.GetEntry($sheetXmlPath)
    }

    # Đọc Shared Strings
    $sharedStrings = @()
    $sstEntry = $zip.GetEntry("xl/sharedStrings.xml")
    if ($sstEntry) {
        $sstReader = New-Object System.IO.StreamReader($sstEntry.Open())
        $sstXml = [xml]$sstReader.ReadToEnd()
        $sstReader.Close()
        $sharedStrings = @($sstXml.sst.si | ForEach-Object { 
            if ($_.t) { $_.t } else { ($_.r | ForEach-Object { $_.t }) -join '' }
        })
    }

    # Đọc Sheet 1 để tìm ngày báo cáo (nếu có ô ngày số OADate)
    try {
        $s1Entry = $zip.GetEntry("xl/worksheets/sheet1.xml")
        if ($s1Entry) {
            $s1Reader = New-Object System.IO.StreamReader($s1Entry.Open())
            $s1Xml = [xml]$s1Reader.ReadToEnd()
            $s1Reader.Close()
            $dateCell = $s1Xml.worksheet.sheetData.row | Where-Object { $_.r -eq "2" } | Select-Object -ExpandProperty c | Where-Object { $_.r -eq "B2" -or $_.r -eq "A2" -or $_.r -eq "P2" } | Select-Object -First 1
            if ($dateCell -and $dateCell.v) {
                $oaDate = 0
                if ([double]::TryParse($dateCell.v, [ref]$oaDate) -and $oaDate -gt 40000 -and $oaDate -lt 60000) {
                    $reportDate = [DateTime]::FromOADate($oaDate).ToString("dd/MM/yyyy")
                }
            }
        }
    } catch {}

    # Đọc Sheet XML chính
    $sheetReader = New-Object System.IO.StreamReader($sheetEntry.Open())
    $sheetXml = [xml]$sheetReader.ReadToEnd()
    $sheetReader.Close()
    if ($zip) { $zip.Dispose() }
    if (Test-Path $tempCopy) { Remove-Item $tempCopy -Force -ErrorAction SilentlyContinue }

    # Phân tích từng dòng dữ liệu (từ dòng 4 trở đi)
    foreach ($r in $sheetXml.worksheet.sheetData.row) {
        $rowIdx = [int]$r.r
        if ($rowIdx -lt 4) { continue }
        
        $cMap = @{}
        foreach ($c in $r.c) {
            $colName = $c.r -replace '\d+',''
            $val = $c.v
            if ($c.t -eq "s" -and $sharedStrings.Count -gt 0) {
                $val = $sharedStrings[[int]$val]
            }
            $cMap[$colName] = $val
        }
        
        $shop = $cMap['A']
        if ([string]::IsNullOrWhiteSpace($shop)) { continue }
        if ($shop -like "*Grand Total*" -or $shop -like "*Tổng*" -or $shop -like "*Total*") { continue }
        
        $slMay = 0; [double]::TryParse($cMap['C'], [ref]$slMay) | Out-Null
        $slSim = 0; [double]::TryParse($cMap['D'], [ref]$slSim) | Out-Null
        $tlSim = 0; [double]::TryParse($cMap['E'], [ref]$tlSim) | Out-Null
        $tangTruong = 0; [double]::TryParse($cMap['F'], [ref]$tangTruong) | Out-Null
        
        $shops += [PSCustomObject]@{
            Shop = $shop.Trim()
            SLMay = $slMay
            SLSim = $slSim
            TLSim = $tlSim
            TangTruong = $tangTruong
        }
    }
}
elseif ($targetFile.EndsWith(".csv", [System.StringComparison]::OrdinalIgnoreCase)) {
    # Xử lý file CSV
    $lines = Get-Content $targetFile -Encoding UTF8
    $inTable = $false
    foreach ($line in $lines) {
        if ($line -like "*MỪNG 3 TUỔI*" -or $line -like "*MUNG 3 TUOI*") {
            $inTable = $true
            continue
        }
        if (-not $inTable) { 
            if ($line -match '(\d{4}-\d{2}-\d{2})') {
                try {
                    $reportDate = [DateTime]::Parse($matches[1]).ToString("dd/MM/yyyy")
                } catch {}
            }
            continue 
        }
        if ($line -like "*TARGET THI ĐUA*" -or $line -like "*CHỈ SỐ BASE*" -or $line -like "*Grand Total*") { break }
        
        $parts = $line.Split(",")
        if ($parts.Count -ge 6) {
            $shopName = $parts[0].Trim()
            if ($shopName -eq "TÊN SHOP" -or $shopName -like "*Top 2*" -or [string]::IsNullOrWhiteSpace($shopName)) { continue }
            
            $slMay = 0; [double]::TryParse($parts[2].Trim(), [ref]$slMay) | Out-Null
            $slSim = 0; [double]::TryParse($parts[3].Trim(), [ref]$slSim) | Out-Null
            $tlSim = 0; [double]::TryParse($parts[4].Trim(), [ref]$tlSim) | Out-Null
            $tangTruong = 0; [double]::TryParse($parts[5].Trim(), [ref]$tangTruong) | Out-Null
            
            $shops += [PSCustomObject]@{
                Shop = $shopName
                SLMay = $slMay
                SLSim = $slSim
                TLSim = $tlSim
                TangTruong = $tangTruong
            }
        }
    }
}

if (-not $reportDate) {
    $reportDate = (Get-Date).ToString("dd/MM/yyyy")
}

Write-Host "      Ngày báo cáo: $reportDate" -ForegroundColor DarkCyan
Write-Host "      Tổng số shop đã tổng hợp: $($shops.Count) shop" -ForegroundColor DarkCyan

if ($shops.Count -eq 0) {
    Write-Host "[LỖI] Không trích xuất được dữ liệu shop từ sheet 'Thi đua Sim'!" -ForegroundColor Red
    exit 1
}

# 4. XẾP HẠNG 3 HẠNG MỤC THI ĐUA
Write-Host "[2/4] Đang xếp hạng các hạng mục thi đua..." -ForegroundColor Green

# Hạng mục 1: VUA SIM (Cột D - Số lượng SIM FPT)
$vuaSimRank = $shops | Sort-Object SLSim -Descending
$vuaSimTop = $vuaSimRank | Select-Object -First 5

# Hạng mục 2: CAO THỦ CHỐT SIM (Cột E - Tỷ lệ SIM/Máy)
$tyLeRank = $shops | Sort-Object TLSim -Descending
$tyLeTop = $tyLeRank | Select-Object -First 5

# Hạng mục 3: BỨT PHÁ SIM (Cột F - Tăng trưởng so với cùng kỳ 2025)
$tangTruongRank = $shops | Sort-Object TangTruong -Descending
$tangTruongTop = $tangTruongRank | Select-Object -First 5

# 4B. TỔNG HỢP CHỈ SỐ CỘT C & CỘT D TOÀN QUỐC
$totalSim = 0
$totalMay = 0
foreach ($s in $shops) {
    $totalSim += $s.SLSim
    $totalMay += $s.SLMay
}
$totalRate = if ($totalMay -gt 0) { $totalSim / $totalMay } else { 0 }

# Helper định dạng số
function Format-Number([double]$num) {
    return [string]::Format("{0:N0}", $num)
}
function Format-Percent([double]$num, [bool]$showSign = $false) {
    if ($showSign -and $num -gt 0) {
        return "+" + [string]::Format("{0:N1}%", $num * 100)
    }
    return [string]::Format("{0:N1}%", $num * 100)
}

# 5. IN KẾT QUẢ RA BẢNG CONSOLE
Write-Host "
=== 🌟 CHỈ SỐ TOÀN QUỐC HÀNG NGÀY (CỘT C & CỘT D) ===" -ForegroundColor Yellow
Write-Host "  * Tổng SIM FPT (Cột D)        : $(Format-Number $totalSim) SIM" -ForegroundColor Cyan
Write-Host "  * Tổng Máy tính SIM (Cột C)   : $(Format-Number $totalMay) Máy" -ForegroundColor Cyan
Write-Host "  * Tỷ lệ khai thác toàn quốc   : $(Format-Percent $totalRate)" -ForegroundColor Green

Write-Host "
=== 👑 HẠNG MỤC 1: VUA SIM (SỐ LƯỢNG) ===" -ForegroundColor Red
Write-Host "  [GIẢI NHẤT] 1. $($vuaSimTop[0].Shop) : $(Format-Number $vuaSimTop[0].SLSim) SIM (Tỷ lệ: $(Format-Percent $vuaSimTop[0].TLSim))" -ForegroundColor Green
Write-Host "  [GIẢI NHÌ]  2. $($vuaSimTop[1].Shop) : $(Format-Number $vuaSimTop[1].SLSim) SIM (Tỷ lệ: $(Format-Percent $vuaSimTop[1].TLSim))" -ForegroundColor Green
Write-Host "  [BÁM ĐUỔI]  3. $($vuaSimTop[2].Shop) : $(Format-Number $vuaSimTop[2].SLSim) SIM" -ForegroundColor Yellow
Write-Host "  [BÁM ĐUỔI]  4. $($vuaSimTop[3].Shop) : $(Format-Number $vuaSimTop[3].SLSim) SIM" -ForegroundColor Yellow
Write-Host "  [BÁM ĐUỔI]  5. $($vuaSimTop[4].Shop) : $(Format-Number $vuaSimTop[4].SLSim) SIM" -ForegroundColor Yellow

Write-Host "`n=== 🎯 HẠNG MỤC 2: CAO THỦ CHỐT SIM (TỶ LỆ) ===" -ForegroundColor Cyan
Write-Host "  [GIẢI NHẤT] 1. $($tyLeTop[0].Shop) : $(Format-Percent $tyLeTop[0].TLSim) ($($tyLeTop[0].SLSim) SIM / $($tyLeTop[0].SLMay) Máy)" -ForegroundColor Green
Write-Host "  [GIẢI NHÌ]  2. $($tyLeTop[1].Shop) : $(Format-Percent $tyLeTop[1].TLSim) ($($tyLeTop[1].SLSim) SIM / $($tyLeTop[1].SLMay) Máy)" -ForegroundColor Green
Write-Host "  [BÁM ĐUỔI]  3. $($tyLeTop[2].Shop) : $(Format-Percent $tyLeTop[2].TLSim) ($($tyLeTop[2].SLSim) SIM)" -ForegroundColor Yellow
Write-Host "  [BÁM ĐUỔI]  4. $($tyLeTop[3].Shop) : $(Format-Percent $tyLeTop[3].TLSim) ($($tyLeTop[3].SLSim) SIM)" -ForegroundColor Yellow
Write-Host "  [BÁM ĐUỔI]  5. $($tyLeTop[4].Shop) : $(Format-Percent $tyLeTop[4].TLSim) ($($tyLeTop[4].SLSim) SIM)" -ForegroundColor Yellow

Write-Host "`n=== 📈 HẠNG MỤC 3: BỨT PHÁ SIM (TĂNG TRƯỞNG) ===" -ForegroundColor Green
Write-Host "  [GIẢI NHẤT] 1. $($tangTruongTop[0].Shop) : $(Format-Percent $tangTruongTop[0].TangTruong $true) ($($tangTruongTop[0].SLSim) SIM)" -ForegroundColor Green
Write-Host "  [GIẢI NHÌ]  2. $($tangTruongTop[1].Shop) : $(Format-Percent $tangTruongTop[1].TangTruong $true) ($($tangTruongTop[1].SLSim) SIM)" -ForegroundColor Green
Write-Host "  [BÁM ĐUỔI]  3. $($tangTruongTop[2].Shop) : $(Format-Percent $tangTruongTop[2].TangTruong $true)" -ForegroundColor Yellow
Write-Host "  [BÁM ĐUỔI]  4. $($tangTruongTop[3].Shop) : $(Format-Percent $tangTruongTop[3].TangTruong $true)" -ForegroundColor Yellow
Write-Host "  [BÁM ĐUỔI]  5. $($tangTruongTop[4].Shop) : $(Format-Percent $tangTruongTop[4].TangTruong $true)" -ForegroundColor Yellow

# 6. CẬP NHẬT TEMPLATE HTML
Write-Host "`n[3/4] Đang tạo file Infographic HTML..." -ForegroundColor Green
$htmlTemplate = [System.IO.File]::ReadAllText($TplPath, [System.Text.Encoding]::UTF8)

# Chuyển ảnh banner & mascot sang Base64 để nhúng hoàn toàn
$headerBannerPath = Join-Path $AssetsDir "header_banner_ultra.png"
if (-not (Test-Path $headerBannerPath)) {
    $headerBannerPath = Join-Path $AssetsDir "header_banner_hd.png"
}
if (-not (Test-Path $headerBannerPath)) {
    $headerBannerPath = Join-Path $AssetsDir "header_banner.jpg"
}
$headerBannerB64 = ""
if (Test-Path $headerBannerPath) {
    $b = [System.IO.File]::ReadAllBytes($headerBannerPath)
    $ext = [System.IO.Path]::GetExtension($headerBannerPath).ToLower()
    $mime = if ($ext -eq ".png") { "image/png" } else { "image/jpeg" }
    $headerBannerB64 = "data:$mime;base64," + [Convert]::ToBase64String($b)
    Write-Host "      Đã nhúng Banner tiêu đề siêu nét: $(Split-Path -Leaf $headerBannerPath)" -ForegroundColor DarkCyan
}

$mascotPath = Join-Path $AssetsDir "mascot.png"
$mascotB64 = ""
if (Test-Path $mascotPath) {
    $b = [System.IO.File]::ReadAllBytes($mascotPath)
    $mascotB64 = "data:image/png;base64," + [Convert]::ToBase64String($b)
}

# Thay thế các placeholder
$replacements = @{
    "__HEADER_BANNER_BASE64__" = $headerBannerB64
    "__MASCOT_BASE64__"        = $mascotB64
    "__NGAY_CAP_NHAT__"        = $reportDate
    
    # Chỉ số Cột C & Cột D toàn quốc
    "__TONG_SL_SIM__"          = Format-Number $totalSim
    "__TONG_SL_MAY__"          = Format-Number $totalMay
    "__TONG_TY_LE__"           = Format-Percent $totalRate
    
    # Hạng mục 1: VUA SIM
    "__VUA_SIM_TOP1_NAME__"  = $vuaSimTop[0].Shop
    "__VUA_SIM_TOP1_VALUE__" = (Format-Number $vuaSimTop[0].SLSim) + " SIM"
    "__VUA_SIM_TOP1_MAY__"   = Format-Number $vuaSimTop[0].SLMay
    "__VUA_SIM_TOP1_RATE__"  = Format-Percent $vuaSimTop[0].TLSim
    
    "__VUA_SIM_TOP2_NAME__"  = $vuaSimTop[1].Shop
    "__VUA_SIM_TOP2_VALUE__" = (Format-Number $vuaSimTop[1].SLSim) + " SIM"
    "__VUA_SIM_TOP2_MAY__"   = Format-Number $vuaSimTop[1].SLMay
    "__VUA_SIM_TOP2_RATE__"  = Format-Percent $vuaSimTop[1].TLSim
    
    "__VUA_SIM_TOP3_NAME__"  = $vuaSimTop[2].Shop
    "__VUA_SIM_TOP3_VALUE__" = (Format-Number $vuaSimTop[2].SLSim) + " SIM"
    "__VUA_SIM_TOP4_NAME__"  = $vuaSimTop[3].Shop
    "__VUA_SIM_TOP4_VALUE__" = (Format-Number $vuaSimTop[3].SLSim) + " SIM"
    "__VUA_SIM_TOP5_NAME__"  = $vuaSimTop[4].Shop
    "__VUA_SIM_TOP5_VALUE__" = (Format-Number $vuaSimTop[4].SLSim) + " SIM"
    "__VUA_SIM_TOP3_MAY__"   = Format-Number $vuaSimTop[2].SLMay
    "__VUA_SIM_TOP4_MAY__"   = Format-Number $vuaSimTop[3].SLMay
    "__VUA_SIM_TOP5_MAY__"   = Format-Number $vuaSimTop[4].SLMay

    # Hạng mục 2: CAO THỦ CHỐT SIM
    "__TY_LE_TOP1_NAME__"  = $tyLeTop[0].Shop
    "__TY_LE_TOP1_VALUE__" = Format-Percent $tyLeTop[0].TLSim
    "__TY_LE_TOP1_SIM__"   = (Format-Number $tyLeTop[0].SLSim) + " SIM"
    "__TY_LE_TOP1_MAY__"   = Format-Number $tyLeTop[0].SLMay
    
    "__TY_LE_TOP2_NAME__"  = $tyLeTop[1].Shop
    "__TY_LE_TOP2_VALUE__" = Format-Percent $tyLeTop[1].TLSim
    "__TY_LE_TOP2_SIM__"   = (Format-Number $tyLeTop[1].SLSim) + " SIM"
    "__TY_LE_TOP2_MAY__"   = Format-Number $tyLeTop[1].SLMay
    
    "__TY_LE_TOP3_NAME__"  = $tyLeTop[2].Shop
    "__TY_LE_TOP3_VALUE__" = Format-Percent $tyLeTop[2].TLSim
    "__TY_LE_TOP4_NAME__"  = $tyLeTop[3].Shop
    "__TY_LE_TOP4_VALUE__" = Format-Percent $tyLeTop[3].TLSim
    "__TY_LE_TOP5_NAME__"  = $tyLeTop[4].Shop
    "__TY_LE_TOP5_VALUE__" = Format-Percent $tyLeTop[4].TLSim
    "__TY_LE_TOP3_SIM__"   = Format-Number $tyLeTop[2].SLSim
    "__TY_LE_TOP3_MAY__"   = Format-Number $tyLeTop[2].SLMay
    "__TY_LE_TOP4_SIM__"   = Format-Number $tyLeTop[3].SLSim
    "__TY_LE_TOP4_MAY__"   = Format-Number $tyLeTop[3].SLMay
    "__TY_LE_TOP5_SIM__"   = Format-Number $tyLeTop[4].SLSim
    "__TY_LE_TOP5_MAY__"   = Format-Number $tyLeTop[4].SLMay

    # Hạng mục 3: BỨT PHÁ SIM
    "__TANG_TRUONG_TOP1_NAME__"  = $tangTruongTop[0].Shop
    "__TANG_TRUONG_TOP1_VALUE__" = Format-Percent $tangTruongTop[0].TangTruong $true
    "__TANG_TRUONG_TOP1_SIM__"   = (Format-Number $tangTruongTop[0].SLSim) + " SIM"
    "__TANG_TRUONG_TOP1_MAY__"   = Format-Number $tangTruongTop[0].SLMay
    "__TANG_TRUONG_TOP1_RATE__"  = Format-Percent $tangTruongTop[0].TLSim
    
    "__TANG_TRUONG_TOP2_NAME__"  = $tangTruongTop[1].Shop
    "__TANG_TRUONG_TOP2_VALUE__" = Format-Percent $tangTruongTop[1].TangTruong $true
    "__TANG_TRUONG_TOP2_SIM__"   = (Format-Number $tangTruongTop[1].SLSim) + " SIM"
    "__TANG_TRUONG_TOP2_MAY__"   = Format-Number $tangTruongTop[1].SLMay
    "__TANG_TRUONG_TOP2_RATE__"  = Format-Percent $tangTruongTop[1].TLSim
    
    "__TANG_TRUONG_TOP3_NAME__"  = $tangTruongTop[2].Shop
    "__TANG_TRUONG_TOP3_VALUE__" = Format-Percent $tangTruongTop[2].TangTruong $true
    "__TANG_TRUONG_TOP3_SIM__"   = Format-Number $tangTruongTop[2].SLSim
    "__TANG_TRUONG_TOP3_MAY__"   = Format-Number $tangTruongTop[2].SLMay

    "__TANG_TRUONG_TOP4_NAME__"  = $tangTruongTop[3].Shop
    "__TANG_TRUONG_TOP4_VALUE__" = Format-Percent $tangTruongTop[3].TangTruong $true
    "__TANG_TRUONG_TOP4_SIM__"   = Format-Number $tangTruongTop[3].SLSim
    "__TANG_TRUONG_TOP4_MAY__"   = Format-Number $tangTruongTop[3].SLMay

    "__TANG_TRUONG_TOP5_NAME__"  = $tangTruongTop[4].Shop
    "__TANG_TRUONG_TOP5_VALUE__" = Format-Percent $tangTruongTop[4].TangTruong $true
    "__TANG_TRUONG_TOP5_SIM__"   = Format-Number $tangTruongTop[4].SLSim
    "__TANG_TRUONG_TOP5_MAY__"   = Format-Number $tangTruongTop[4].SLMay
}

$outputHtmlContent = $htmlTemplate
foreach ($key in $replacements.Keys) {
    $outputHtmlContent = $outputHtmlContent.Replace($key, $replacements[$key])
}

# Tạo file HTML tạm trong thư mục TEMP (để render ảnh, không lưu lại file HTML trong thư mục output theo yêu cầu)
$tempHtmlPath = Join-Path $env:TEMP ("KetQua_ThiDua_Sim_" + [Guid]::NewGuid().ToString("N") + ".html")
[System.IO.File]::WriteAllText($tempHtmlPath, $outputHtmlContent, [System.Text.Encoding]::UTF8)

# Xóa bất kỳ file HTML nào còn tồn tại trong thư mục output (chỉ giữ đúng 1 file PNG)
Get-ChildItem -Path $OutputDir -Filter "*.html" -File -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue

# 7. RENDER ẢNH INFOGRAPHIC (HỖ TRỢ ĐƯỜNG DẪN CÓ DẤU CÁCH & TỰ ĐỘNG CHUYỂN ĐỔI TRÌNH DUYỆT)
Write-Host "[4/4] Đang render ảnh Infographic chất lượng cao (PNG)..." -ForegroundColor Green
$outPngPath = Join-Path $OutputDir "KetQua_ThiDua_Sim.png"
$tempPngPath = Join-Path $env:TEMP ("render_" + [Guid]::NewGuid().ToString("N") + ".png")

# Chuẩn hóa đường dẫn HTML thành file:// URI có mã hóa %20 để không bị lỗi dấu cách
$htmlUri = [System.Uri]::new($tempHtmlPath).AbsoluteUri

function Get-AvailableBrowsers {
    $list = @()
    $candidates = @(
        "C:\Program Files\Google\Chrome\Application\chrome.exe",
        "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe",
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
        "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
        "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        "C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        "$env:LOCALAPPDATA\Microsoft\Edge\Application\msedge.exe",
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe"
    )
    foreach ($c in $candidates) {
        if ($c -and (Test-Path $c) -and -not ($list -contains $c)) { $list += $c }
    }
    $cmds = Get-Command chrome.exe, msedge.exe -ErrorAction SilentlyContinue
    if ($cmds) {
        foreach ($m in $cmds) {
            if ($m.Source -and -not ($list -contains $m.Source)) { $list += $m.Source }
        }
    }
    return $list
}

$browsers = Get-AvailableBrowsers
$renderSuccess = $false

foreach ($browserPath in $browsers) {
    $bName = Split-Path -Leaf $browserPath
    Write-Host "      Đang chụp ảnh với trình duyệt: $bName" -ForegroundColor DarkCyan
    
    # Render ra file tạm không có dấu cách trong tên thư mục
    $browserArgs = @(
        "--headless",
        "--screenshot=$tempPngPath",
        "--no-sandbox",
        "--disable-gpu",
        "--window-size=1200,1370",
        "--hide-scrollbars",
        $htmlUri
    )
    
    try {
        $p = Start-Process -FilePath $browserPath -ArgumentList $browserArgs -Wait -PassThru -NoNewWindow
        Start-Sleep -Milliseconds 600
        
        if (Test-Path $tempPngPath) {
            Copy-Item $tempPngPath $outPngPath -Force
            Remove-Item $tempPngPath -Force -ErrorAction SilentlyContinue
            $renderSuccess = $true
            break
        }
    } catch {}
}

# Dọn dẹp file HTML tạm
Remove-Item $tempHtmlPath -Force -ErrorAction SilentlyContinue

if ($renderSuccess -and (Test-Path $outPngPath)) {
    # XÓA TOÀN BỘ FILE THỪA TRONG OUTPUT - CHỈ GIỮ LẠI ĐÚNG 1 FILE PNG THEO YÊU CẦU
    Get-ChildItem -Path $OutputDir -File | Where-Object { $_.Name -ne "KetQua_ThiDua_Sim.png" } | Remove-Item -Force -ErrorAction SilentlyContinue

    # Đồng bộ sang thư mục Brain để hiển thị trực tiếp trong báo cáo
    $brainArtifact = "C:\Users\Hue_Laptop\.gemini\antigravity\brain\b78de785-d3e7-49c6-bc65-bba07a21379c\KetQua_ThiDua_Sim.png"
    Copy-Item $outPngPath $brainArtifact -Force -ErrorAction SilentlyContinue

    $fileSize = [math]::Round(((Get-Item $outPngPath).Length / 1KB), 1)
    Write-Host "`n======================================================================" -ForegroundColor Green
    Write-Host " [THÀNH CÔNG] ĐÃ XUẤT 1 FILE ẢNH PNG DUY NHẤT RÕ NÉT (KHÔNG XUẤT HTML)!" -ForegroundColor Green
    Write-Host " File ảnh duy nhất: $outPngPath ($fileSize KB)" -ForegroundColor Yellow
    Write-Host "======================================================================" -ForegroundColor Green
    
    # Tự động mở ảnh cho người dùng xem
    Start-Process $outPngPath
} else {
    Write-Host "`n[LỖI] Không thể chụp ảnh tự động. Vui lòng kiểm tra trình duyệt Chrome hoặc Microsoft Edge!" -ForegroundColor Red
    pause
    exit 1
}
