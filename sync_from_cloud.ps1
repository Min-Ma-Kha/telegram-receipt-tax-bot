# Pulls a backup zip from the cloud bot and unpacks it into .\data\
# Reads CLOUD_URL and BACKUP_KEY from the .env file next to this script.
# Safe to run any time; runs automatically via Task Scheduler (see
# setup_sync_task.ps1).

$ErrorActionPreference = "Stop"
$proj = $PSScriptRoot
$logFile = Join-Path $proj "data\sync.log"

function Log($msg) {
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $msg"
    Write-Host $line
    Add-Content -Path $logFile -Value $line -Encoding utf8
}

# Parse .env
$envVars = @{}
Get-Content (Join-Path $proj ".env") | ForEach-Object {
    if ($_ -match '^\s*([A-Z_]+)\s*=\s*(.*)\s*$') { $envVars[$Matches[1]] = $Matches[2] }
}
$cloudUrl = $envVars["CLOUD_URL"]
$backupKey = $envVars["BACKUP_KEY"]

if (-not $cloudUrl) {
    Log "SKIP: CLOUD_URL is not set in .env yet (deploy to the cloud first)."
    exit 0
}
if (-not $backupKey) {
    Log "ERROR: BACKUP_KEY missing from .env."
    exit 1
}

$tmp = Join-Path $env:TEMP "receipts_backup.zip"
try {
    Invoke-WebRequest -Uri "$cloudUrl/backup?key=$backupKey" -OutFile $tmp -TimeoutSec 120
    $dest = Join-Path $proj "data"
    if (-not (Test-Path $dest)) { New-Item -ItemType Directory -Path $dest | Out-Null }
    Expand-Archive -Path $tmp -DestinationPath $dest -Force
    Remove-Item $tmp -Force
    Log "OK: synced cloud data to $dest"
} catch {
    Log "ERROR: sync failed - $($_.Exception.Message)"
    exit 1
}
