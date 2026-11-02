#requires -Version 5.1
# AUDIT: Da kiem tra script. Khong co lenh xoa thu muc/tep hang loat,
#        khong su dung Remove-Item -Recurse -Force, del /f /s /q.
# SmartFarm: tu dong loc log ERROR/CRITICAL va thong ke su co.
# Duong dan duoc xac dinh tu vi tri file script, khong phu thuoc o dia C:.

$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$logFile    = Join-Path (Join-Path $projectRoot 'data') 'log-he-thong.log'
$backupDir  = Join-Path $projectRoot 'backup'
$reportFile = Join-Path $backupDir 'bao-cao-su-co.txt'
$pattern    = '\[(ERROR|CRITICAL)\]'

if (-not (Test-Path -LiteralPath $logFile -PathType Leaf)) {
    throw "Khong tim thay du lieu log: $logFile"
}
if (-not (Test-Path -LiteralPath $backupDir -PathType Container)) {
    New-Item -Path $backupDir -ItemType Directory | Out-Null
}

# (0) Danh dau file .env la Hidden tren Windows, khong lam mat noi dung.
$envFile = Join-Path $projectRoot '.env'
if (Test-Path -LiteralPath $envFile -PathType Leaf) {
    $envItem = Get-Item -LiteralPath $envFile -Force
    $envItem.Attributes = $envItem.Attributes -bor [System.IO.FileAttributes]::Hidden
}

# (1) Write-Output + > : tao file bao cao moi va ghi tieu de.
Write-Output 'BAO CAO SU CO - SMARTFARM' > $reportFile
Write-Output ("Ngay lap: {0}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')) >> $reportFile
Write-Output 'Nguon du lieu: data/log-he-thong.log' >> $reportFile
Write-Output 'Muc do: ERROR, CRITICAL' >> $reportFile
Write-Output '========================================' >> $reportFile
Write-Output 'CHI TIET CAC SU CO:' >> $reportFile

# (2) Pipeline + Get-Content + Select-String + >> : loc va noi dong su co.
Get-Content -LiteralPath $logFile |
    Select-String -Pattern $pattern |
    ForEach-Object { $_.Line } >> $reportFile

# (3) Pipeline + Measure-Object : thong ke so dong su co.
$totalIncidents = Get-Content -LiteralPath $logFile |
    Select-String -Pattern $pattern |
    Measure-Object |
    Select-Object -ExpandProperty Count

Write-Output '========================================' >> $reportFile
Write-Output ("TONG SO SU CO: {0}" -f $totalIncidents) >> $reportFile

# Thong ke bo sung de doi chieu (khong thay doi ket qua loc).
$errors = Get-Content -LiteralPath $logFile |
    Select-String -Pattern '\[ERROR\]' |
    Measure-Object |
    Select-Object -ExpandProperty Count
$criticals = Get-Content -LiteralPath $logFile |
    Select-String -Pattern '\[CRITICAL\]' |
    Measure-Object |
    Select-Object -ExpandProperty Count
Write-Output ("ERROR: {0}" -f $errors) >> $reportFile
Write-Output ("CRITICAL: {0}" -f $criticals) >> $reportFile

Write-Host ("Da tao bao cao: {0}" -f $reportFile) -ForegroundColor Green
Write-Host ("Tong so su co ERROR/CRITICAL: {0}" -f $totalIncidents) -ForegroundColor Cyan
