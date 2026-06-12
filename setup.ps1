# One-time setup for Receipt Tax Bot. Run it by double-clicking setup.bat.
# Installs what's missing, asks for your bot token, and offers auto-start.

$ErrorActionPreference = "Stop"
$proj = $PSScriptRoot

Write-Host ""
Write-Host "=============================================="
Write-Host "   Receipt Tax Bot - one-time setup"
Write-Host "=============================================="
Write-Host ""

# ---------------------------------------------------------- 1. Python
Write-Host "[1/5] Checking Python..."
$python = Get-Command python -ErrorAction SilentlyContinue
$pyOk = $false
if ($python) {
    try { $v = (& python --version) 2>$null; if ($v -match "Python 3\.(\d+)") { $pyOk = [int]$Matches[1] -ge 10 } } catch {}
}
if (-not $pyOk) {
    Write-Host "  Python 3.10+ not found. Installing it now (this can take a few minutes)..."
    winget install --id Python.Python.3.12 --accept-source-agreements --accept-package-agreements --silent
    Write-Host ""
    Write-Host "  Python installed! Please CLOSE this window and double-click setup.bat"
    Write-Host "  AGAIN so Windows picks up the new installation."
    exit 0
}
Write-Host "  OK: $(& python --version)"

# ---------------------------------------------------------- 2. Tesseract OCR
Write-Host "[2/5] Checking Tesseract OCR (reads the receipt photos)..."
$tess = Get-Command tesseract -ErrorAction SilentlyContinue
if (-not $tess -and -not (Test-Path "C:\Program Files\Tesseract-OCR\tesseract.exe")) {
    Write-Host "  Not found. Installing (this can take a few minutes)..."
    winget install --id UB-Mannheim.TesseractOCR --accept-source-agreements --accept-package-agreements --silent
}
Write-Host "  OK: Tesseract is installed."

# ---------------------------------------------------------- 3. Python packages
Write-Host "[3/5] Installing Python packages..."
& python -m pip install -r (Join-Path $proj "requirements.txt") --quiet
Write-Host "  OK: packages installed."

# ---------------------------------------------------------- 4. Bot token
Write-Host "[4/5] Setting up your bot token..."
$envFile = Join-Path $proj ".env"
if (-not (Test-Path $envFile)) {
    Copy-Item (Join-Path $proj ".env.example") $envFile
}
$envText = Get-Content $envFile -Raw
if ($envText -match "PASTE_YOUR_TOKEN_HERE") {
    Write-Host ""
    Write-Host "  You need a Telegram bot token. If you don't have one yet:"
    Write-Host "   1. Open Telegram and search for: BotFather"
    Write-Host "   2. Send it the message: /newbot"
    Write-Host "   3. Give your bot any name, then a username ending in 'bot'"
    Write-Host "   4. BotFather replies with a token like 1234567890:AAH..."
    Write-Host ""
    $token = Read-Host "  Paste your bot token here and press Enter"
    $token = $token.Trim()
    if ($token -notmatch "^\d+:[\w-]+$") {
        Write-Host "  That doesn't look like a token. Run setup.bat again and retry."
        exit 1
    }
    ($envText -replace "PASTE_YOUR_TOKEN_HERE", $token) | Set-Content $envFile -Encoding utf8 -NoNewline
    Write-Host "  OK: token saved to .env"
} else {
    Write-Host "  OK: token already configured."
}

# ---------------------------------------------------------- 5. Auto-start
Write-Host "[5/5] Auto-start with Windows..."
$startupDir = [Environment]::GetFolderPath("Startup")
$lnkPath = Join-Path $startupDir "Receipt Tax Bot.lnk"
if (Test-Path $lnkPath) {
    Write-Host "  OK: already set to start with Windows."
} else {
    $ans = Read-Host "  Start the bot automatically every time Windows starts? (Y/n)"
    if ($ans -eq "" -or $ans -match "^[Yy]") {
        $ws = New-Object -ComObject WScript.Shell
        $sc = $ws.CreateShortcut($lnkPath)
        $sc.TargetPath = Join-Path $proj "run_bot.bat"
        $sc.WorkingDirectory = $proj
        $sc.WindowStyle = 7  # minimized
        $sc.Save()
        Write-Host "  OK: the bot will start (minimized) whenever you log in."
    } else {
        Write-Host "  Skipped. Start it manually by double-clicking run_bot.bat"
    }
}

# ---------------------------------------------------------- done
Write-Host ""
Write-Host "=============================================="
Write-Host "   Setup complete!"
Write-Host "=============================================="
Write-Host ""
$ans = Read-Host "Start the bot now? (Y/n)"
if ($ans -eq "" -or $ans -match "^[Yy]") {
    Start-Process -FilePath (Join-Path $proj "run_bot.bat") -WorkingDirectory $proj
    Write-Host ""
    Write-Host "The bot is starting in its own window. Open Telegram, find YOUR bot,"
    Write-Host "press START, and send it a photo of a receipt!"
}
