# Auto-sync and run PowerPoint generator
Write-Host "Syncing with repository..." -ForegroundColor Cyan
git pull origin claude/powerpoint-generation-mvp-HdREb

if ($LASTEXITCODE -eq 0) {
    Write-Host "`nRunning PowerPoint generator...`n" -ForegroundColor Green
    python generate_ppt.py
} else {
    Write-Host "`nGit pull failed. Please resolve conflicts manually." -ForegroundColor Red
    exit 1
}
