#!/usr/bin/env bash

# Exit immediately if a command exits with a non-zero status
set -e

# Change directory to the repository folder where this script is located
REPO_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$REPO_DIR"

echo "==========================================="
echo "   Switchgear Web App - Git Sync Script"
echo "==========================================="

# Check if git is installed
if ! command -v git &> /dev/null; then
    echo "Error: 'git' command not found. Please install git."
    exit 1
fi

# Check if this directory is a git repository
if [ ! -d ".git" ]; then
    echo "Error: Not a git repository. Initialize git first using:"
    echo "  git init"
    echo "  git remote add origin <repo_url>"
    exit 1
fi

# Fetch current branch name (defaults to main)
BRANCH=$(git branch --show-current 2>/dev/null || echo "main")
if [ -z "$BRANCH" ]; then
    BRANCH="main"
fi

echo "Current branch: $BRANCH"

# Check if there are changes to commit
if [ -z "$(git status --porcelain)" ]; then
    echo "No local changes detected."
    
    # Prompt to pull latest changes from GitHub
    read -p "Do you want to pull the latest changes from GitHub? (y/n): " PULL_CHOICE
    if [[ "$PULL_CHOICE" =~ ^[Yy]$ ]]; then
        echo "Pulling latest changes from origin/$BRANCH..."
        git pull origin "$BRANCH"
        echo "Repository is up to date!"
    fi
    exit 0
fi

# Display current changes
echo "Detected changes:"
git status -s
echo "-------------------------------------------"

# Prompt for a commit message
DEFAULT_MSG="Update switchgear app and templates"
read -p "Enter commit message [Default: '$DEFAULT_MSG']: " COMMIT_MSG
COMMIT_MSG="${COMMIT_MSG:-$DEFAULT_MSG}"

# Stage all files
echo "Staging files..."
git add .

# Commit changes
echo "Committing with message: '$COMMIT_MSG'..."
git commit -m "$COMMIT_MSG"

# Push to GitHub
echo "Pushing changes to GitHub (origin/$BRANCH)..."
git push origin "$BRANCH"

echo "==========================================="
echo " Successfully pushed changes to GitHub!"
echo " GitHub Pages will redeploy in 1-2 minutes."
echo "==========================================="
