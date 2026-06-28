$repo = "jare39063124-oss/loan-photo-app"
$apiBase = "https://api.github.com/repos/$repo"
$desktop = [Environment]::GetFolderPath("Desktop")
$tempZip = Join-Path $env:TEMP "loan-photo-apk.zip"
$extractDir = Join-Path $env:TEMP "loan-photo-apk-extract"
$finalApkName = "loan-photo-tool-v3.7.0.apk"

Write-Host "=== GitHub Actions APK Build Monitor ===" -ForegroundColor Cyan

$runId = $null
$waitCount = 0
while (-not $runId -and $waitCount -lt 20) {
    Start-Sleep -Seconds 5
    $waitCount++
    try {
        $runs = Invoke-RestMethod -Uri "$apiBase/actions/runs?branch=master&per_page=3" -UseBasicParsing
        foreach ($run in $runs.workflow_runs) {
            $sha = $run.head_sha
            if ($sha -and $sha.StartsWith("61f39ed")) {
                $runId = $run.id
                $runNum = $run.run_number
                Write-Host "Found build #$runNum (ID: $runId)" -ForegroundColor Green
                break
            }
        }
    } catch {
        Write-Host "  Waiting for build to appear... ($waitCount)"
    }
}

if (-not $runId) {
    $runs = Invoke-RestMethod -Uri "$apiBase/actions/runs?branch=master&per_page=1" -UseBasicParsing
    $runId = $runs.workflow_runs[0].id
    Write-Host "Using latest build ID: $runId" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Monitoring build status (polling every 30s)..." -ForegroundColor Cyan
Write-Host ""

$lastStatus = ""
$startTime = Get-Date
$headers = @{ "User-Agent" = "PowerShell"; "Accept" = "application/vnd.github+json" }

while ($true) {
    try {
        $run = Invoke-RestMethod -Uri "$apiBase/actions/runs/$runId" -UseBasicParsing -Headers $headers
        $status = $run.status
        $conclusion = $run.conclusion
        $elapsed = [math]::Round(((Get-Date) - $startTime).TotalSeconds)

        switch ($status) {
            "completed" {
                if ($conclusion -eq "success") {
                    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] BUILD SUCCESS (${elapsed}s)" -ForegroundColor Green
                } else {
                    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] BUILD FAILED: $conclusion" -ForegroundColor Red
                    Write-Host "Check logs: $($run.html_url)" -ForegroundColor Yellow
                    exit 1
                }
                break
            }
            "in_progress" {
                $msg = "[$(Get-Date -Format 'HH:mm:ss')] Building... (${elapsed}s)"
                if ($msg -ne $lastStatus) { Write-Host $msg; $lastStatus = $msg }
            }
            "queued" {
                $msg = "[$(Get-Date -Format 'HH:mm:ss')] Queued... (${elapsed}s)"
                if ($msg -ne $lastStatus) { Write-Host $msg -ForegroundColor Yellow; $lastStatus = $msg }
            }
            default {
                Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Status: $status"
            }
        }

        if ($status -eq "completed" -and $conclusion -eq "success") { break }
        Start-Sleep -Seconds 30
    } catch {
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Query error, retrying..." -ForegroundColor DarkGray
        Start-Sleep -Seconds 30
    }
}

Write-Host ""
Write-Host "Downloading APK artifact..." -ForegroundColor Cyan

$artifacts = Invoke-RestMethod -Uri "$apiBase/actions/runs/$runId/artifacts" -UseBasicParsing -Headers $headers
if ($artifacts.artifacts.Count -eq 0) {
    Write-Host "No artifacts found!" -ForegroundColor Red
    exit 1
}

$artifact = $artifacts.artifacts[0]
$sizeMB = [math]::Round($artifact.size_in_bytes / 1048576, 2)
Write-Host "Artifact: $($artifact.name) ($sizeMB MB)"

$downloadUrl = $artifact.archive_download_url

if (Test-Path $tempZip) { Remove-Item $tempZip -Force }
if (Test-Path $extractDir) { Remove-Item $extractDir -Recurse -Force }
New-Item -ItemType Directory -Force -Path $extractDir | Out-Null

Invoke-WebRequest -Uri $downloadUrl -OutFile $tempZip -Headers $headers -UseBasicParsing
Write-Host "Download complete, extracting..."

Expand-Archive -Path $tempZip -DestinationPath $extractDir -Force

$apkFiles = @(Get-ChildItem -Path $extractDir -Filter "*.apk" -Recurse)
if ($apkFiles.Count -eq 0) {
    Write-Host "No APK found in zip!" -ForegroundColor Red
    Get-ChildItem -Path $extractDir -Recurse | ForEach-Object { Write-Host "  $_" }
    exit 1
}

$apkFile = $apkFiles[0]
$destPath = Join-Path $desktop $finalApkName
Copy-Item -Path $apkFile.FullName -Destination $destPath -Force
$destSize = [math]::Round((Get-Item $destPath).Length / 1048576, 2)

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "  APK downloaded to Desktop!" -ForegroundColor Green
Write-Host "  File: $finalApkName"
Write-Host "  Size: ${destSize} MB"
Write-Host "  Path: $destPath"
Write-Host "========================================" -ForegroundColor Green

Remove-Item $tempZip -Force -ErrorAction SilentlyContinue
Remove-Item $extractDir -Recurse -Force -ErrorAction SilentlyContinue

explorer.exe "/select,$destPath"
