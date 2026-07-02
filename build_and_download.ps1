$repo = "jare39063124-oss/loan-photo-app"
$apiBase = "https://api.github.com/repos/$repo"
$desktop = [Environment]::GetFolderPath("Desktop")
$projectDir = "D:\hermes\loan_photo_app"
$finalApkName = "loan-photo-tool-v3.22.1.apk"
$destPath = Join-Path $desktop $finalApkName

Write-Host "=== Step 1: Git Push (retry every 60s) ===" -ForegroundColor Cyan
$pushOk = $false
$pushAttempt = 0
while (-not $pushOk) {
    $pushAttempt++
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Push attempt #$pushAttempt"
    Set-Location $projectDir
    git push 2>&1 | Out-Host
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Push SUCCESS!" -ForegroundColor Green
        $pushOk = $true
    } else {
        Write-Host "Push failed, waiting 60s..." -ForegroundColor Yellow
        Start-Sleep -Seconds 60
    }
}

Write-Host ""
Write-Host "=== Step 2: Waiting for build to appear ===" -ForegroundColor Cyan
Start-Sleep -Seconds 15
$runId = $null
$headers = @{ "User-Agent" = "PowerShell"; "Accept" = "application/vnd.github+json" }
$headSha = (git -C $projectDir rev-parse --short HEAD).Trim()
Write-Host "Looking for build with SHA: $headSha"
$waitCount = 0
while (-not $runId -and $waitCount -lt 30) {
    Start-Sleep -Seconds 10
    $waitCount++
    try {
        $runs = Invoke-RestMethod -Uri "$apiBase/actions/runs?branch=master&per_page=5" -UseBasicParsing -Headers $headers
        foreach ($run in $runs.workflow_runs) {
            if ($run.head_sha -and $run.head_sha.StartsWith($headSha)) {
                $runId = $run.id
                Write-Host "Found build #$($run.run_number) (ID: $runId)" -ForegroundColor Green
                break
            }
        }
    } catch {
        Write-Host "  Waiting... ($waitCount)"
    }
}
if (-not $runId) {
    $runs = Invoke-RestMethod -Uri "$apiBase/actions/runs?branch=master&per_page=1" -UseBasicParsing -Headers $headers
    $runId = $runs.workflow_runs[0].id
    Write-Host "Using latest build ID: $runId" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "=== Step 3: Monitoring build (polling every 30s) ===" -ForegroundColor Cyan
$lastStatus = ""
$startTime = Get-Date
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
                    Write-Host "Check: $($run.html_url)" -ForegroundColor Yellow
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
        }
        if ($status -eq "completed" -and $conclusion -eq "success") { break }
        Start-Sleep -Seconds 30
    } catch {
        Write-Host "Query error, retrying..." -ForegroundColor DarkGray
        Start-Sleep -Seconds 30
    }
}

Write-Host ""
Write-Host "=== Step 4: Downloading APK via nightly.link ===" -ForegroundColor Cyan
$tempZip = Join-Path $env:TEMP "loan-apk-v3221.zip"
$extractDir = Join-Path $env:TEMP "loan-apk-v3221-extract"
$downloadUrl = "https://nightly.link/$repo/actions/runs/$runId/loan-photo-apk.zip"
if (Test-Path $tempZip) { Remove-Item $tempZip -Force }
if (Test-Path $extractDir) { Remove-Item $extractDir -Recurse -Force }
New-Item -ItemType Directory -Force -Path $extractDir | Out-Null
Write-Host "Downloading..."
Invoke-WebRequest -Uri $downloadUrl -OutFile $tempZip -UseBasicParsing
$zipSize = [math]::Round((Get-Item $tempZip).Length / 1MB, 2)
Write-Host "Downloaded: ${zipSize} MB"
Expand-Archive -Path $tempZip -DestinationPath $extractDir -Force
$apks = @(Get-ChildItem $extractDir -Filter "*.apk" -Recurse)
Write-Host "Found $($apks.Count) APK(s):"
foreach ($a in $apks) {
    $sz = [math]::Round($a.Length / 1MB, 2)
    Write-Host "  $($a.Name) (${sz} MB)"
}
if ($apks.Count -gt 0) {
    Copy-Item $apks[0].FullName $destPath -Force
    $finalSize = [math]::Round((Get-Item $destPath).Length / 1MB, 2)
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "  v3.22.1 APK READY!" -ForegroundColor Green
    Write-Host "  File: $finalApkName"
    Write-Host "  Size: ${finalSize} MB"
    Write-Host "  Path: $destPath"
    Write-Host "========================================" -ForegroundColor Green
    explorer.exe "/select,$destPath"
}
Remove-Item $tempZip -Force -ErrorAction SilentlyContinue
Remove-Item $extractDir -Recurse -Force -ErrorAction SilentlyContinue
