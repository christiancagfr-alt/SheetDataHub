#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")"
VERSION="$(python -c "from sheet_hub.version import APP_VERSION; print(APP_VERSION)")"
python -m unittest discover -s tests -v
python -m PyInstaller --noconfirm --clean --distpath dist-onedir SheetDataHub.spec
mkdir -p dist/release
if [[ -d dist-onedir/SheetDataHub.app ]]; then
  APP="dist-onedir/SheetDataHub.app"
elif [[ -d dist-onedir/SheetDataHub/SheetDataHub.app ]]; then
  APP="dist-onedir/SheetDataHub/SheetDataHub.app"
elif [[ -d dist-onedir/SheetDataHub ]]; then
  APP="dist-onedir/SheetDataHub"
else
  echo "macOS build output missing"
  exit 1
fi
ZIP="dist/release/SheetDataHub-macos-arm64-v${VERSION}.zip"
ditto -c -k --sequesterRsrc --keepParent "$APP" "$ZIP"
shasum -a 256 "$ZIP" | awk '{print $1}' > "${ZIP}.sha256"
echo "Created $ZIP"
STAGE="dist/dmg_stage"
rm -rf "$STAGE"
mkdir -p "$STAGE"
ditto "$APP" "$STAGE/SheetDataHub.app"
ln -s /Applications "$STAGE/Applications"
DMG="dist/release/SheetDataHub-macos-arm64-v${VERSION}.dmg"
hdiutil create -volname "SheetDataHub ${VERSION}" -srcfolder "$STAGE" -ov -format UDZO "$DMG"
shasum -a 256 "$DMG" | awk '{print $1}' > "${DMG}.sha256"
rm -rf "$STAGE"
echo "Created $DMG"
