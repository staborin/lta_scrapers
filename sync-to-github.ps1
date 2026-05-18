# sync-to-github.ps1 - Run from lta_scrapers repo root
Set-Location $PSScriptRoot

# 1. Pull latest Apps Script files
Write-Host ""
Write-Host "--- Pulling Apps Script files ---" -ForegroundColor Cyan
Set-Location apps-script
clasp pull
if ($LASTEXITCODE -ne 0) {
    Write-Host "clasp pull failed!" -ForegroundColor Red
    pause
    exit 1
}
Set-Location ..

# 2. Stage all changes
Write-Host ""
Write-Host "--- Staging changes ---" -ForegroundColor Cyan
git add .

# 3. Check if there is anything to commit
$status = git status --porcelain
if (-not $status) {
    Write-Host ""
    Write-Host "Nothing to commit - all files up to date." -ForegroundColor Green
    pause
    exit 0
}

Write-Host ""
Write-Host "Changed files:" -ForegroundColor Yellow
git status --short

# 4. Commit and push
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm"
git commit -m "Sync all files $timestamp"
git push

Write-Host ""
Write-Host "Done!" -ForegroundColor Green
pause