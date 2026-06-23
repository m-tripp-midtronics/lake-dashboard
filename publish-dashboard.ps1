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

# ── 1. Read Excel ─────────────────────────────────────────────────────────────
$xlsxPath = Join-Path $scriptDir "PH and Temp Log.xlsx"
if (-not (Test-Path $xlsxPath)) {
    Show-Fail "Cannot find: $xlsxPath"
    Read-Host "Press Enter to exit"
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

        # Skip empty or non-numeric rows
        if ($null -eq $brPh -or $null -eq $brTemp -or $null -eq $npPh -or $null -eq $npTemp) { continue }

        # Convert Excel date serial to yyyy-MM-dd
        if ($rawDate -is [double] -or $rawDate -is [int]) {
            $dateObj = [DateTime]::FromOADate($rawDate)
        } else {
            try { $dateObj = [DateTime]::Parse($rawDate) } catch { continue }
        }
        $dateStr = $dateObj.ToString("yyyy-MM-dd")

        # Clean time string: "7:19:00 AM" → "7:19 AM"
        $timeStr = $rawTime -replace "^(\d+:\d+):\d+\s*(AM|PM)$", '$1 $2'
        if (-not $timeStr) { $timeStr = "12:00 PM" }

        $rows += [PSCustomObject]@{
            date    = $dateStr
            time    = $timeStr
            brPh    = [Math]::Round($brPh, 2)
            brTemp  = [Math]::Round($brTemp, 1)
            npPh    = [Math]::Round($npPh, 2)
            npTemp  = [Math]::Round($npTemp, 1)
        }
    }

    $wb.Close($false)
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}

if ($rows.Count -eq 0) {
    Show-Fail "No data rows found in Excel file."
    Read-Host "Press Enter to exit"
    exit 1
}

Show-OK "Read $($rows.Count) rows from Excel"

# ── 2. Build JS SEED array ────────────────────────────────────────────────────
$seedLines = $rows | ForEach-Object {
    "  {date:`"$($_.date)`",time:`"$($_.time)`", brPh:$($_.brPh), brTemp:$($_.brTemp), npPh:$($_.npPh), npTemp:$($_.npTemp)},"
}
# Remove trailing comma from last line
$seedLines[-1] = $seedLines[-1].TrimEnd(',')
$newSeed = "const SEED = [`n" + ($seedLines -join "`n") + "`n];"

# ── 3. Update both HTML files ─────────────────────────────────────────────────
$files = @("index.html", "PH and Temp Dashboard.html")
foreach ($fname in $files) {
    $fpath = Join-Path $scriptDir $fname
    if (-not (Test-Path $fpath)) { continue }

    Show-Status "Updating $fname..."
    $content = Get-Content $fpath -Raw -Encoding UTF8

    # Replace the SEED block (from "const SEED = [" to the closing "];" )
    $newContent = $content -replace '(?s)const SEED = \[.*?\];', $newSeed

    if ($newContent -eq $content) {
        Show-Fail "Could not find SEED block in $fname — skipping"
        continue
    }

    Set-Content $fpath $newContent -Encoding UTF8 -NoNewline
    Show-OK "Updated $fname"
}

# ── 4. Git commit & push ──────────────────────────────────────────────────────
Show-Status "Pushing to GitHub..."

$gitBash = "C:\Program Files\Git\bin\bash.exe"
if (-not (Test-Path $gitBash)) {
    $gitBash = "C:\Program Files\Git\usr\bin\bash.exe"
}

$lastDate = $rows[-1].date
$commitMsg = "Update dashboard data through $lastDate ($($rows.Count) readings)"

& $gitBash -c "cd '/c/Users/mtripp/OneDrive/MSC/Lake & Beach' && git add index.html 'PH and Temp Dashboard.html' && git commit -m '$commitMsg' && git push" 2>&1 | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }

if ($LASTEXITCODE -eq 0) {
    Show-OK "Published! Live at: https://m-tripp-midtronics.github.io/lake-dashboard/"
} else {
    Show-Fail "Git push failed. Check output above."
}

Write-Host ""
Read-Host "Press Enter to close"
