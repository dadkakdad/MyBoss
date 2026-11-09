#!/usr/bin/env bash
# Emergency local CapRover deploy for the MyBoss browser app (prefer GitHub Actions).
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
APP_NAME="${1:-myboss}"
SRC="$ROOT/web"

if [ ! -d "$SRC" ]; then
  echo "web/ directory not found" >&2
  exit 1
fi

TEMP_DIR="$(mktemp -d)"
cleanup() { rm -rf "$TEMP_DIR"; }
trap cleanup EXIT

cp -a "$SRC"/. "$TEMP_DIR/"
date +%s > "$TEMP_DIR/.cachebust"
echo "Cache bust: $(cat "$TEMP_DIR/.cachebust")"

if [ ! -f "$TEMP_DIR/captain-definition" ]; then
  echo "captain-definition missing" >&2
  exit 1
fi

TAR=/tmp/myboss-web-deploy.tar
tar --exclude='node_modules' \
    --exclude='test' \
    --exclude='.git' \
    -cf "$TAR" \
    -C "$TEMP_DIR" .

echo "Tarball size: $(du -h "$TAR" | cut -f1)"

deploy_args=(deploy --tarFile "$TAR" --appName "$APP_NAME")
if [ -n "${CAPROVER_URL:-}" ]; then
  deploy_args+=(--caproverUrl "$CAPROVER_URL")
fi
if [ -n "${CAPROVER_APP_TOKEN:-}" ]; then
  deploy_args+=(--appToken "$CAPROVER_APP_TOKEN")
elif [ -n "${CAPROVER_MYBOSS_APP_TOKEN:-}" ]; then
  deploy_args+=(--appToken "$CAPROVER_MYBOSS_APP_TOKEN")
fi

CI=true caprover "${deploy_args[@]}"
echo "Deploy accepted for app $APP_NAME"
