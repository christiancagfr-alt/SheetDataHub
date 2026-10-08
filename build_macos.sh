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
else
  echo "macOS .app bundle missing"
  exit 1
fi
bash packaging/macos_sign.sh --app "$APP" --version "$VERSION" --out dist/release
echo "Signed and notarized macOS packages in dist/release"
