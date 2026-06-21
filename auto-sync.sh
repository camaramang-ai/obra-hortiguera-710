#!/bin/bash
# Auto-sync: commit & push changes to GitHub
cd /Users/PabloMan/Hermes/projects/active/obra-hortiguera-710 || exit 1

# Check if there are any changes
if [[ -z $(git status --porcelain) ]]; then
  exit 0  # Nothing to do
fi

git add -A
git commit -m "auto-sync $(date '+%Y-%m-%d %H:%M')"
git push origin main 2>&1
