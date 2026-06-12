# Run this ONCE (after deploying to the cloud and setting CLOUD_URL in .env).
# Registers a Windows scheduled task that pulls the cloud data to this PC:
#   - every time you log on (i.e. whenever the PC comes back online)
#   - and every 4 hours while it's on

$script = Join-Path $PSScriptRoot "sync_from_cloud.ps1"
$action = New-ScheduledTaskAction -Execute "powershell.exe" `
    -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$script`""

$atLogon = New-ScheduledTaskTrigger -AtLogOn
$every4h = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(5) `
    -RepetitionInterval (New-TimeSpan -Hours 4)

$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable `
    -DontStopIfGoingOnBatteries -AllowStartIfOnBatteries

Register-ScheduledTask -TaskName "ReceiptTaxBotSync" `
    -Action $action -Trigger $atLogon, $every4h -Settings $settings `
    -Description "Pulls receipt-tax-bot data from the cloud to this PC" -Force

Write-Host "Scheduled task 'ReceiptTaxBotSync' registered."
Write-Host "It runs at every logon and every 4 hours. Test it now with:"
Write-Host "  powershell -ExecutionPolicy Bypass -File `"$script`""
