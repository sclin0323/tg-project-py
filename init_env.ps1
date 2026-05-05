# Stop on error
$ErrorActionPreference = "Stop"

Write-Host "🚀 Start initializing Python environment..." -ForegroundColor Green

# Remove old venv
if (Test-Path ".venv") {
    Write-Host "🗑 Removing old virtual environment .venv"
    Remove-Item -Recurse -Force ".venv"
}

# Create new venv
Write-Host "📦 Creating new virtual environment .venv"
python -m venv .venv

# Activate venv (for setup)
Write-Host "🔑 Activating virtual environment (for setup)"
& .\.venv\Scripts\Activate.ps1

# Upgrade pip
Write-Host "⬆️ Upgrading pip"
python -m pip install --upgrade pip

# Install requirements
Write-Host "📥 Installing requirements.txt"
pip install -r requirements.txt

# Fix pandas-ta posix import issue (Windows only)
$pta_file = ".\.venv\Lib\site-packages\pandas_ta\overlap\alligator.py"
if (Test-Path $pta_file) {
    Write-Host "🔧 Patching pandas-ta (removing 'from posix import pread')"
    (Get-Content $pta_file) | ForEach-Object { $_ -replace 'from posix import pread', '# from posix import pread (patched for Windows)' } | Set-Content $pta_file
}

# Keep venv active for user session
Write-Host "✅ Environment initialization completed! (.venv is active)" -ForegroundColor Cyan
Write-Host "👉 Please run the following to stay in venv:" -ForegroundColor Yellow
Write-Host "    .\.venv\Scripts\Activate.ps1" -ForegroundColor Yellow
