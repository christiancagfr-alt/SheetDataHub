#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")"
python -m unittest discover -s tests -v
python -m PyInstaller --noconfirm --clean --distpath dist-onedir SheetDataHub.spec
mkdir -p dist/release
if [[ -d dist-onedir/SheetDataHub.app ]]; then
  APP="dist-onedir/SheetDataHub.app"
elif [[ -d dist-onedir/SheetDataHub ]]; then
  APP="dist-onedir/SheetDataHub"
else
  echo "macOS build output missing"
  exit 1
fi
ZIP="dist/release/SheetDataHub-macos-arm64.zip"
ditto -c -k --sequesterRsrc --keepParent "$APP" "$ZIP"
shasum -a 256 "$ZIP" | awk '{print $1}' > "${ZIP}.sha256"
echo "Created $ZIP"
