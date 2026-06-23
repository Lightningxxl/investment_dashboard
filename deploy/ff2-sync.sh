#!/usr/bin/env bash
set -euo pipefail

APP_DIR="${APP_DIR:-/home/xavierx/www/zy-sj-dashboard}"
REPO_URL="${REPO_URL:-https://github.com/Lightningxxl/investment_dashboard.git}"
BRANCH="${BRANCH:-main}"

if [ ! -d "$APP_DIR" ]; then
  mkdir -p "$(dirname "$APP_DIR")"
  git clone --branch "$BRANCH" "$REPO_URL" "$APP_DIR"
  exit 0
fi

if [ ! -d "$APP_DIR/.git" ]; then
  echo "$APP_DIR exists but is not a git repository. Move it aside or choose another APP_DIR." >&2
  exit 1
fi

git -C "$APP_DIR" fetch origin "$BRANCH"
git -C "$APP_DIR" pull --ff-only origin "$BRANCH"
