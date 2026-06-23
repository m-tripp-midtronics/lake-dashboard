# publish-dashboard.ps1
# Reads PH and Temp Log.xlsx, updates dashboard seed data, pushes to GitHub.

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

function Show-Status($msg) { Write-Host "  $msg" -ForegroundColor Cyan }
function Show-OK($msg)     { Write-Host "  OK  $msg" -ForegroundColor Green }
function Show-Fail($msg)   { Write-Host "  ERR $msg" -ForegroundColor Red }

Write-Host ""
Write-Host "  Lake & Beach Dashboard Publisher" -ForegroundColor White
Write-Host "  ----------------------------------" -ForegroundColor DarkGray
Write-Host ""

# 1. Read Excel
$xlsxPath = Join-Path $scriptDir "PH and Temp Log.xlsx"
if (-not (Test-Path $xlsxPath)) {
    Show-Fail "Cannot find: $xlsxPath"
    pause
    exit 1
}

Show-Status "Reading Excel file..."

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Open($xlsxPath)
    $ws = $wb.Sheets.Item(1)
    $lastRow = $ws.UsedRange.Rows.Count

    $rows = @()
    for ($r = 2; $r -le $lastRow; $r++) {
        $rawDate = $ws.Cells($r, 1).Value2
        $rawTime = $ws.Cells($r, 2).Text
        $brPh    = $ws.Cells($r, 3).Value2
        $brTemp  = $ws.Cells($r, 4).Value2
        $npPh    = $ws.Cells($r, 5).Value2
        $npTemp  = $ws.Cells($r, 6).Value2

        if ($null -eq $brPh -or $null -eq $brTemp -or $null -eq $npPh -or $null -eq $npTemp) { continue }

        if ($rawDate -is [double] -or $rawDate -is [int]) {
            $dateObj = [DateTime]::FromOADate($rawDate)
        } else {
            try { $dateObj = [DateTime]::Parse($rawDate) } catch { continue }
        }
        $dateStr = $dateObj.ToString("yyyy-MM-dd")

        $timeStr = [regex]::Replace($rawTime, "^(\d+:\d+):\d+\s*(AM|PM)$", '$1 $2')
        if (-not $timeStr) { $timeStr = "12:00 PM" }

        $rows += [PSCustomObject]@{
            date    = $dateStr
            time    = $timeStr
            brPh    = [Math]::Round([double]$brPh, 2)
            brTemp  = [Math]::Round([double]$brTemp, 1)
            npPh    = [Math]::Round([double]$npPh, 2)
            npTemp  = [Math]::Round([double]$npTemp, 1)
        }
    }

    $wb.Close($false)
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}

if ($rows.Count -eq 0) {
    Show-Fail "No data rows found in Excel file."
    pause
    exit 1
}

Show-OK "Read $($rows.Count) rows from Excel"

# 2. Build JS SEED array
$seedLines = @()
for ($i = 0; $i -lt $rows.Count; $i++) {
    $r = $rows[$i]
    $comma = if ($i -lt $rows.Count - 1) { "," } else { "" }
    $seedLines += "  {date:`"$($r.date)`",time:`"$($r.time)`", brPh:$($r.brPh), brTemp:$($r.brTemp), npPh:$($r.npPh), npTemp:$($r.npTemp)}$comma"
}
$newSeed = "const SEED = [" + "`n" + ($seedLines -join "`n") + "`n];"

# 3. Update both HTML files
$files = @("index.html", "PH and Temp Dashboard.html")
foreach ($fname in $files) {
    $fpath = Join-Path $scriptDir $fname
    if (-not (Test-Path $fpath)) { continue }

    Show-Status "Updating $fname..."
    $content = [System.IO.File]::ReadAllText($fpath)
    $newContent = [regex]::Replace($content, '(?s)const SEED = \[.*?\];', $newSeed)

    if ($newContent -eq $content) {
        Show-Fail "Could not find SEED block in $fname"
        continue
    }

    [System.IO.File]::WriteAllText($fpath, $newContent, [System.Text.Encoding]::UTF8)
    Show-OK "Updated $fname"
}

# 4. Git commit and push via Git Bash
Show-Status "Pushing to GitHub..."

$gitBash = "C:\Program Files\Git\bin\bash.exe"
if (-not (Test-Path $gitBash)) {
    Show-Fail "Git Bash not found at $gitBash"
    pause
    exit 1
}

$lastDate = $rows[-1].date
$count = $rows.Count
$commitMsg = "Update dashboard data through $lastDate"

$bashCmd = "cd '/c/Users/mtripp/OneDrive/MSC/Lake & Beach' && git add index.html 'PH and Temp Dashboard.html' && git commit -m '$commitMsg - $count readings' && git push 2>&1"
& $gitBash -c $bashCmd | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
$exitCode = $LASTEXITCODE

if ($exitCode -eq 0) {
    Write-Host ""
    Show-OK "Live at: https://m-tripp-midtronics.github.io/lake-dashboard/"
} else {
    Show-Fail "Git push failed. See output above."
}

Write-Host ""
Write-Host "  Press any key to close..." -ForegroundColor DarkGray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
