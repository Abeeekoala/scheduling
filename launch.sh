#!/usr/bin/env bash
set -e

REPO_DIR="$(cd "$(dirname "$0")" && pwd)"

cd "$REPO_DIR"

# fetch latest refs
git fetch origin

LOCAL=$(git rev-parse @)
REMOTE=$(git rev-parse @{u})
BASE=$(git merge-base @ @{u})

if [ "$LOCAL" = "$REMOTE" ]; then
  echo "Already up to date."
elif [ "$LOCAL" = "$BASE" ]; then
  echo "Updating from remote..."
  git pull --ff-only
elif [ "$REMOTE" = "$BASE" ]; then
  echo "Local ahead of remote; push pending changes first."
else
  echo "Diverged from remote; resolve manually."
  exit 1
fi

start index.html  # use `xdg-open` on Linux or `start` on Windows
