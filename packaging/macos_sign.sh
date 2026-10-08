#!/usr/bin/env bash
# Sign a PyInstaller .app with Developer ID, notarize, staple, then emit zip + DMG.
set -euo pipefail

log() { printf '%s\n' "$*"; }
die() { printf 'ERROR: %s\n' "$*" >&2; exit 1; }

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
ENTITLEMENTS="$ROOT/packaging/macos_entitlements.plist"
APP=""
VERSION=""
OUT=""

while [[ $# -gt 0 ]]; do
  case "$1" in
    --app) APP="$2"; shift 2 ;;
    --version) VERSION="$2"; shift 2 ;;
    --out) OUT="$2"; shift 2 ;;
    *) die "unknown argument: $1" ;;
  esac
done

[[ "$(uname -s)" == "Darwin" ]] || die "macOS signing must run on macOS"
[[ -d "$APP" ]] || die "app bundle missing: $APP"
[[ "$APP" == *.app ]] || die "expected a .app bundle, got: $APP"
[[ -n "$VERSION" ]] || die "--version is required"
[[ -f "$ENTITLEMENTS" ]] || die "entitlements missing: $ENTITLEMENTS"
mkdir -p "$OUT"
OUT="$(cd "$OUT" && pwd)"
APP="$(cd "$APP" && pwd)"

WORKDIR="$(mktemp -d "${TMPDIR:-/tmp}/sheetdatahub-sign.XXXXXX")"
KEYCHAIN=""
cleanup() {
  if [[ -n "$KEYCHAIN" ]]; then
    security delete-keychain "$KEYCHAIN" >/dev/null 2>&1 || true
  fi
  rm -rf "$WORKDIR"
}
trap cleanup EXIT

decode_b64() {
  local dest="$1"
  python3 -c 'import base64, os, pathlib, sys; pathlib.Path(sys.argv[1]).write_bytes(base64.b64decode(os.environ[sys.argv[2]]))' "$dest" "$2"
}

import_certificate() {
  local p12_b64="${MACOS_CERTIFICATE_P12_BASE64:-}"
  local p12_pass="${MACOS_CERTIFICATE_PASSWORD:-}"
  if [[ -z "$p12_b64" ]]; then
    return 1
  fi
  [[ -n "$p12_pass" ]] || die "MACOS_CERTIFICATE_PASSWORD is required"
  local p12="$WORKDIR/certificate.p12"
  MACOS_CERTIFICATE_P12_BASE64="$p12_b64" decode_b64 "$p12" MACOS_CERTIFICATE_P12_BASE64
  KEYCHAIN="$WORKDIR/signing.keychain-db"
  local password
  password="$(python3 -c 'import secrets; print(secrets.token_urlsafe(24))')"
  security create-keychain -p "$password" "$KEYCHAIN"
  security set-keychain-settings -lut 21600 "$KEYCHAIN"
  security unlock-keychain -p "$password" "$KEYCHAIN"
  security import "$p12" -k "$KEYCHAIN" -P "$p12_pass" -T /usr/bin/codesign -T /usr/bin/security >/dev/null
  local existing
  existing="$(security list-keychains -d user | sed 's/"//g')"
  # shellcheck disable=SC2086
  security list-keychains -d user -s "$KEYCHAIN" $existing
  security set-key-partition-list -S apple-tool:,apple:,codesign: -s -k "$password" "$KEYCHAIN" >/dev/null
  rm -f "$p12"
  return 0
}

resolve_identity() {
  local identity="${MACOS_CODESIGN_IDENTITY:-}"
  local query=()
  if [[ -n "$KEYCHAIN" ]]; then
    query=("$KEYCHAIN")
  fi
  if [[ -z "$identity" ]]; then
    identity="$(security find-identity -v -p codesigning "${query[@]}" | awk -F'"' '/Developer ID Application/{print $2; exit}')"
  fi
  [[ -n "$identity" ]] || die "no Developer ID Application identity found"
  printf '%s' "$identity"
}

sign_file() {
  local path="$1"
  local with_entitlements="${2:-}"
  if [[ -n "$with_entitlements" ]]; then
    codesign --force --options runtime --timestamp --entitlements "$ENTITLEMENTS" --sign "$IDENTITY" "$path"
  else
    codesign --force --options runtime --timestamp --sign "$IDENTITY" "$path"
  fi
}

sign_app() {
  xattr -cr "$APP" || true
  local path
  while IFS= read -r -d '' path; do
    [[ -L "$path" ]] && continue
    if file -b "$path" | grep -q 'Mach-O'; then
      sign_file "$path"
    fi
  done < <(find "$APP/Contents" -depth -type f ! -path '*/_CodeSignature/*' -print0)

  while IFS= read -r -d '' path; do
    sign_file "$path"
  done < <(find "$APP/Contents" -depth \( -name '*.framework' -o -name '*.app' -o -name '*.xpc' -o -name '*.bundle' \) -print0)

  local executable="$APP/Contents/MacOS/SheetDataHub"
  if [[ -f "$executable" ]]; then
    sign_file "$executable" entitlements
  fi
  sign_file "$APP" entitlements
  codesign --verify --deep --strict --verbose=2 "$APP"
  log "Signed $APP as $IDENTITY"
}

write_api_key() {
  local b64="${APPLE_API_KEY_P8_BASE64:-}"
  local key_id="${APPLE_API_KEY_ID:-}"
  local issuer="${APPLE_API_ISSUER:-}"
  [[ -n "$b64" && -n "$key_id" && -n "$issuer" ]] || die "notarization requires APPLE_API_KEY_P8_BASE64, APPLE_API_KEY_ID, APPLE_API_ISSUER"
  API_KEY_FILE="$WORKDIR/AuthKey_${key_id}.p8"
  APPLE_API_KEY_P8_BASE64="$b64" decode_b64 "$API_KEY_FILE" APPLE_API_KEY_P8_BASE64
  chmod 600 "$API_KEY_FILE"
}

notarize() {
  local file="$1"
  log "Submitting $(basename "$file") for notarization"
  xcrun notarytool submit "$file" \
    --key "$API_KEY_FILE" \
    --key-id "$APPLE_API_KEY_ID" \
    --issuer "$APPLE_API_ISSUER" \
    --wait \
    --timeout 1800
}

sha256_file() {
  shasum -a 256 "$1" | awk '{print $1}' > "$1.sha256"
}

package_zip() {
  local zip_path="$OUT/SheetDataHub-macos-arm64-v${VERSION}.zip"
  rm -f "$zip_path"
  ditto -c -k --sequesterRsrc --keepParent "$APP" "$zip_path" >/dev/null
  printf '%s' "$zip_path"
}

package_dmg() {
  local stage="$WORKDIR/dmg_stage"
  local dmg_path="$OUT/SheetDataHub-macos-arm64-v${VERSION}.dmg"
  rm -rf "$stage"
  mkdir -p "$stage"
  ditto "$APP" "$stage/SheetDataHub.app"
  ln -s /Applications "$stage/Applications"
  rm -f "$dmg_path"
  hdiutil create -volname "SheetDataHub ${VERSION}" -srcfolder "$stage" -ov -format UDZO "$dmg_path" >/dev/null
  codesign --force --timestamp --sign "$IDENTITY" "$dmg_path" >/dev/null
  printf '%s' "$dmg_path"
}

if [[ -n "${MACOS_CERTIFICATE_P12_BASE64:-}" ]]; then
  import_certificate
elif [[ -n "${GITHUB_ACTIONS:-}" ]]; then
  die "MACOS_CERTIFICATE_P12_BASE64 secret is missing"
else
  log "No P12 in environment; using identities already in the keychain"
fi
IDENTITY="$(resolve_identity)"
write_api_key
sign_app

ZIP_PATH="$(package_zip)"
notarize "$ZIP_PATH"
xcrun stapler staple "$APP"
ZIP_PATH="$(package_zip)"
sha256_file "$ZIP_PATH"

DMG_PATH="$(package_dmg)"
notarize "$DMG_PATH"
xcrun stapler staple "$DMG_PATH"
sha256_file "$DMG_PATH"

log "Created $ZIP_PATH"
log "Created $DMG_PATH"
xcrun stapler validate "$APP"
xcrun stapler validate "$DMG_PATH"
codesign --verify --deep --strict "$APP"
