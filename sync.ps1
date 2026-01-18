# Quick sync script - pull latest changes without running generation

# Ensure we're in the script directory
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ScriptDir

Write-Host "Working directory: $ScriptDir" -ForegroundColor Cyan
Write-Host "Syncing with remote repository..." -ForegroundColor Cyan

$currentBranch = git rev-parse --abbrev-ref HEAD
Write-Host "Current branch: $currentBranch" -ForegroundColor Yellow

git fetch origin $currentBranch

$LOCAL = git rev-parse HEAD
$REMOTE = git rev-parse origin/$currentBranch

if ($LOCAL -eq $REMOTE) {
    Write-Host "`n✓ Already up to date - no changes to pull" -ForegroundColor Green
} else {
    Write-Host "`nPulling latest changes..." -ForegroundColor Cyan
    git pull origin $currentBranch

    if ($LASTEXITCODE -eq 0) {
        Write-Host "`n✓ Sync complete - files updated" -ForegroundColor Green
        Write-Host "`nNext steps:" -ForegroundColor Yellow
        Write-Host "  - Review changes: git log -1 --stat"
        Write-Host "  - Generate presentation: python generate.py"
        Write-Host "  - Or run full pipeline: .\run.ps1"
    } else {
        Write-Host "`n✗ Sync failed - please resolve conflicts manually" -ForegroundColor Red
        exit 1
    }
}
