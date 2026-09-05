# ==============================================================================
# Script Gỡ Cài Đặt Tự Động Cho ExcelSupport Add-In
# ==============================================================================

[CmdletBinding()]
param()

Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "     GO CAI DAT EXCEL SUPPORT ADD-IN                    " -ForegroundColor Yellow
Write-Host "========================================================" -ForegroundColor Cyan

# 1. Kiem tra va dong tien trinh Excel neu dang mo
$excelProcesses = Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue
if ($excelProcesses) {
    Write-Host "`n[!] Phat hien Microsoft Excel dang chay." -ForegroundColor Yellow
    Write-Host "Dang tien hanh dong Excel an toan..." -ForegroundColor Yellow
    $excelProcesses | Stop-Process -Force
    Start-Sleep -Seconds 1
}

# 2. Xac dinh va xoa file Add-in khoi XLSTART
$xlStartPath = [System.IO.Path]::Combine($env:APPDATA, "Microsoft", "Excel", "XLSTART")
$targetXll = Join-Path $xlStartPath "ExcelSupport-AddIn.xll"

if (Test-Path $targetXll) {
    Remove-Item -Path $targetXll -Force
    Write-Host "`n[+] Da xoa Add-in khoi thu muc XLSTART thanh cong!" -ForegroundColor Green
} else {
    Write-Host "`n[*] Add-in khong ton tai trong thu muc XLSTART." -ForegroundColor Gray
}

Write-Host "`n========================================================" -ForegroundColor Cyan
Write-Host "  [OK] GO CAI DAT HOAN TAT!                             " -ForegroundColor Green
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "`nNhan phim bat ky de thoat..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
