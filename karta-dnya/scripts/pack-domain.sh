#!/usr/bin/env bash
# Собрать ZIP для загрузки на домен.
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
DIST="$ROOT/dist"
STAGE="$(mktemp -d)"
DATE="$(date -u +%Y%m%d)"
ZIP_NAME="karta-dnya-forus-domen-${DATE}.zip"
trap 'rm -rf "$STAGE"' EXIT

mkdir -p "$STAGE/karta-dnya" "$DIST"
cp -a "$ROOT/server.js" "$ROOT/package.json" "$ROOT/package-lock.json" \
  "$ROOT/README.md" "$ROOT/.gitignore" "$ROOT/start.sh" \
  "$ROOT/ЗАГРУЗКА-НА-ДОМЕН.txt" "$STAGE/karta-dnya/"
cp -a "$ROOT/public" "$STAGE/karta-dnya/"
cp -a "$ROOT/deploy" "$STAGE/karta-dnya/"
mkdir -p "$STAGE/karta-dnya/data"
REPO_ROOT="$(cd "$ROOT/.." && pwd)"
if git -C "$REPO_ROOT" show HEAD:karta-dnya/data/db.json >/dev/null 2>&1; then
  git -C "$REPO_ROOT" show HEAD:karta-dnya/data/db.json > "$STAGE/karta-dnya/data/db.json"
else
  cp -a "$ROOT/data/db.json" "$STAGE/karta-dnya/data/db.json"
fi
chmod +x "$STAGE/karta-dnya/start.sh"

(
  cd "$STAGE"
  zip -r -q "$DIST/$ZIP_NAME" karta-dnya
)
(
  cd "$DIST"
  sha256sum "$ZIP_NAME" > SHA256SUMS.txt
  cp -f "$ZIP_NAME" karta-dnya-domen.zip
)
ls -lh "$DIST/$ZIP_NAME"
echo "Готово: $DIST/$ZIP_NAME"
