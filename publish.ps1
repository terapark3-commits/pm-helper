# ==============================================================================
# pm+helper Unified Build, Documentation Sync, and Release Script (publish.ps1)
# ==============================================================================
param(
    [string]$NewVersion = "",
    [string]$ReleaseNotes = ""
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
Set-Location $ScriptDir

Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "  pm+helper Release & Documentation Sync Pipeline" -ForegroundColor Cyan
Write-Host "========================================================" -ForegroundColor Cyan

# 1. Read / Update version.json
$versionJsonPath = "$ScriptDir\version.json"
if (-not (Test-Path $versionJsonPath)) {
    Write-Error "version.json not found in $ScriptDir"
}

$jsonText = [System.IO.File]::ReadAllText($versionJsonPath, [System.Text.Encoding]::UTF8)
$vObj = ConvertFrom-Json $jsonText

if ($NewVersion -ne "") {
    $vObj.version = $NewVersion.TrimStart('v', 'V')
}
if ($ReleaseNotes -ne "") {
    $vObj.releaseNotes = $ReleaseNotes
}
$vObj.releaseDate = (Get-Date -Format "yyyy-MM-dd")

$updatedJson = ConvertTo-Json $vObj -Depth 5
[System.IO.File]::WriteAllText($versionJsonPath, $updatedJson, (New-Object System.Text.UTF8Encoding($false)))

$ver = $vObj.version
$date = $vObj.releaseDate
$notes = $vObj.releaseNotes
Write-Host "`n[1/5] Active Version: v$ver (Release Date: $date)" -ForegroundColor Green

# 2. Sync UpdateManager.cs
$updateManagerPath = "$ScriptDir\UpdateManager.cs"
if (Test-Path $updateManagerPath) {
    $umLines = [System.IO.File]::ReadAllLines($updateManagerPath, [System.Text.Encoding]::UTF8)
    for ($i = 0; $i -lt $umLines.Length; $i++) {
        if ($umLines[$i].Contains("public const string CurrentVersion =")) {
            $umLines[$i] = "        public const string CurrentVersion = `"$ver`";"
        }
    }
    [System.IO.File]::WriteAllLines($updateManagerPath, $umLines, (New-Object System.Text.UTF8Encoding($true)))
    Write-Host "[2/5] Synced UpdateManager.cs -> CurrentVersion = `"$ver`"" -ForegroundColor Green
}

# 3. Sync user_manual.md
$manualPath = "$ScriptDir\user_manual.md"
if (Test-Path $manualPath) {
    $manLines = [System.IO.File]::ReadAllLines($manualPath, [System.Text.Encoding]::UTF8)
    for ($i = 0; $i -lt [Math]::Min(15, $manLines.Length); $i++) {
        if ($manLines[$i].StartsWith("- 적용 버전:")) {
            $manLines[$i] = "- 적용 버전: V$ver"
        }
        if ($manLines[$i].StartsWith("- 최종 업데이트:")) {
            $manLines[$i] = "- 최종 업데이트: $date"
        }
    }
    [System.IO.File]::WriteAllLines($manualPath, $manLines, (New-Object System.Text.UTF8Encoding($false)))
    Write-Host "[3/5] Synced user_manual.md -> 적용 버전: V$ver | 최종 업데이트: $date" -ForegroundColor Green
}

# 4. Compile C# Binary
Write-Host "[4/5] Compiling pm+helper.exe (C# with UTF-8 codepage 65001)..." -ForegroundColor Green
$csc64 = "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$csc32 = "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\csc.exe"
$csc = if (Test-Path $csc64) { $csc64 } else { $csc32 }

if (-not (Test-Path $csc)) {
    Write-Error "C# compiler (csc.exe) not found in .NET Framework directory."
}

$compileArgs = @(
    "/target:winexe",
    "/out:pm+helper.exe",
    "/optimize+",
    "/platform:anycpu",
    "/codepage:65001",
    "/win32icon:PatientHelper.ico",
    "/reference:System.dll,System.Data.dll,System.Drawing.dll,System.Windows.Forms.dll,System.Xml.dll",
    "pm+helper.cs",
    "UpdateManager.cs"
)

$proc = Start-Process -FilePath $csc -ArgumentList $compileArgs -Wait -NoNewWindow -PassThru
if ($proc.ExitCode -ne 0) {
    Write-Error "Compilation failed with exit code $($proc.ExitCode)"
}

# Compute SHA256 Hash
$exePath = "$ScriptDir\pm+helper.exe"
$hash = (Get-FileHash -Path $exePath -Algorithm SHA256).Hash
[System.IO.File]::WriteAllText("$ScriptDir\pm+helper.sha256", "SHA256 (pm+helper.exe) = $hash`nVersion = v$ver`nDate = $date`n", [System.Text.Encoding]::UTF8)

Write-Host "      Build SUCCESS! SHA256: $hash" -ForegroundColor Yellow

# 5. Git Commit and Tag
Write-Host "[5/5] Git Commit and Release Tagging..." -ForegroundColor Green
if (Test-Path "$ScriptDir\.git") {
    git add version.json user_manual.md UpdateManager.cs pm+helper.cs CHANGELOG.md build.bat publish.ps1 pm+helper.sha256
    git commit -m "release: v$ver - $date" --allow-empty
    git tag -a "v$ver" -m "Release v$ver ($date)`n`n$notes" -f
    Write-Host "      Git Commit & Tag [v$ver] completed." -ForegroundColor Yellow
}

Write-Host "`n========================================================" -ForegroundColor Cyan
Write-Host "  All Build, Sync, and Tagging Tasks Completed!" -ForegroundColor Cyan
Write-Host "  To push to GitHub, run:" -ForegroundColor White
Write-Host "    git remote add origin https://github.com/terapark3/pm-helper.git" -ForegroundColor Gray
Write-Host "    git push -u origin main --tags" -ForegroundColor Gray
Write-Host "========================================================" -ForegroundColor Cyan
