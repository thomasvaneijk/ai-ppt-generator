#!/bin/bash
# Quick sync script - pull latest changes without running generation

echo "Syncing with remote repository..."

CURRENT_BRANCH=$(git rev-parse --abbrev-ref HEAD)
echo "Current branch: $CURRENT_BRANCH"

git fetch origin $CURRENT_BRANCH

LOCAL=$(git rev-parse HEAD)
REMOTE=$(git rev-parse origin/$CURRENT_BRANCH)

if [ "$LOCAL" = "$REMOTE" ]; then
    echo "✓ Already up to date - no changes to pull"
else
    echo "Pulling latest changes..."
    git pull origin $CURRENT_BRANCH

    if [ $? -eq 0 ]; then
        echo "✓ Sync complete - files updated"
        echo ""
        echo "Next steps:"
        echo "  - Review changes: git log -1 --stat"
        echo "  - Generate presentation: python generate.py"
        echo "  - Or run full pipeline: ./run.ps1 (Windows) or python generate.py"
    else
        echo "✗ Sync failed - please resolve conflicts manually"
        exit 1
    fi
fi
