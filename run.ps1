# Auto-sync and run PowerPoint generator
Write-Host "Syncing with repository..." -ForegroundColor Cyan
git pull origin claude/powerpoint-generation-mvp-HdREb

if ($LASTEXITCODE -eq 0) {
    Write-Host "`nRunning PowerPoint generation pipeline...`n" -ForegroundColor Green
    python generate.py
} else {
    Write-Host "`nGit pull failed. Please resolve conflicts manually." -ForegroundColor Red
    exit 1
}
