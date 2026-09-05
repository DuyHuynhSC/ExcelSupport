# ==============================================================================
# Script Cài Đặt Tự Động 1-Click Cho ExcelSupport Add-In
# ==============================================================================

[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "     CAI DAT EXCEL SUPPORT ADD-IN (1-CLICK INSTALLER)   " -ForegroundColor Yellow
Write-Host "========================================================" -ForegroundColor Cyan

# 1. Kiem tra va dong tien trinh Excel neu dang mo
$excelProcesses = Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue
if ($excelProcesses) {
    Write-Host "`n[!] Phat hien Microsoft Excel dang chay." -ForegroundColor Yellow
    Write-Host "Dang tien hanh dong Excel an toan..." -ForegroundColor Yellow
    $excelProcesses | Stop-Process -Force
    Start-Sleep -Seconds 1
}

# 2. Xac dinh thu muc XLSTART cua Microsoft Excel
$xlStartPath = [System.IO.Path]::Combine($env:APPDATA, "Microsoft", "Excel", "XLSTART")
if (-not (Test-Path $xlStartPath)) {
    New-Item -ItemType Directory -Path $xlStartPath -Force | Out-Null
}

# 3. Tim file Add-in 64-bit hoac 32-bit da dong goi
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$sourceDir = Resolve-Path (Join-Path $scriptDir "..")

$xll64Release = Join-Path $sourceDir "bin\release\net48\publish\ExcelSupport-AddIn64-packed.xll"
$xll64Debug   = Join-Path $sourceDir "bin\Debug\net48\publish\ExcelSupport-AddIn64-packed.xll"

$sourceXll = $null
if (Test-Path $xll64Release) {
    $sourceXll = $xll64Release
} elseif (Test-Path $xll64Debug) {
    $sourceXll = $xll64Debug
}

if (-not $sourceXll) {
    Write-Host "`n[X] Khong tim thay file Add-in da dong goi (.xll)!" -ForegroundColor Red
    Write-Host "Vui long chay 'dotnet build -c release' truoc khi cai dat." -ForegroundColor Red
    Pause
    exit 1
}

# 4. Sao chep file vao XLSTART
$targetXll = Join-Path $xlStartPath "ExcelSupport-AddIn.xll"
Write-Host "`n[+] Dang sao chep Add-in vao thu muc khoi dong Excel:" -ForegroundColor Green
Write-Host "    Nguon: $sourceXll" -ForegroundColor Gray
Write-Host "    Dich : $targetXll" -ForegroundColor Gray

Copy-Item -Path $sourceXll -Destination $targetXll -Force

# 5. Mo khoa bao mat Windows (Unblock file) neu co
try {
    Unblock-File -Path $targetXll -ErrorAction SilentlyContinue
} catch { }

Write-Host "`n========================================================" -ForegroundColor Cyan
Write-Host "  [OK] CAI DAT HOAN TAT THANH CONG!                     " -ForegroundColor Green
Write-Host "  Add-in se tu dong kich hoat moi khi ban mo Excel.     " -ForegroundColor White
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "`nNhan phim bat ky de thoat..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
