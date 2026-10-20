#!/usr/bin/env bash
# =============================================================================
#  Rust+ Desktop (Avalonia preview) - macOS .app bundle
# =============================================================================
#  Publishes the self-contained osx-arm64 (or osx-x64) build and wraps it in the
#  bundle layout macOS expects: RustPlus Desktop.app/Contents/MacOS/<binary>
#  with an Info.plist beside it.
#
#  Usage: ./make-macos-app.sh [osx-arm64|osx-x64]
# =============================================================================
set -euo pipefail

RUNTIME="${1:-osx-arm64}"
DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ROOT="$(cd "$DIR/.." && pwd)"
APP_PROJECT="$ROOT/src/RustPlusDesk.Avalonia/RustPlusDesk.Avalonia.csproj"
PUBLISH_DIR="$ROOT/artifacts/publish/$RUNTIME"
BUNDLE="$ROOT/artifacts/publish/RustPlus Desktop.app"
VERSION="$(sed -n 's:.*<Version>\(.*\)</Version>.*:\1:p' "$APP_PROJECT" | head -1)"
VERSION="${VERSION:-0.1.0}"

echo "==> Publishing the $RUNTIME build..."
dotnet publish "$APP_PROJECT" --configuration Release --runtime "$RUNTIME" --self-contained true --output "$PUBLISH_DIR"

echo "==> Building the app bundle..."
rm -rf "$BUNDLE"
mkdir -p "$BUNDLE/Contents/MacOS" "$BUNDLE/Contents/Resources"
cp -R "$PUBLISH_DIR/." "$BUNDLE/Contents/MacOS/"
cp "$ROOT/src/RustPlusDesk.Avalonia/Assets/app.icns" "$BUNDLE/Contents/Resources/" 2>/dev/null || true

cat > "$BUNDLE/Contents/Info.plist" <<PLIST
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>CFBundleName</key><string>RustPlus Desktop</string>
  <key>CFBundleDisplayName</key><string>Rust+ Desktop</string>
  <key>CFBundleIdentifier</key><string>com.rustplusdesk.avalonia</string>
  <key>CFBundleExecutable</key><string>RustPlusDesk.Avalonia</string>
  <key>CFBundleIconFile</key><string>app</string>
  <key>CFBundleShortVersionString</key><string>$VERSION</string>
  <key>CFBundleVersion</key><string>$VERSION</string>
  <key>CFBundlePackageType</key><string>APPL</string>
  <key>LSMinimumSystemVersion</key><string>11.0</string>
  <key>NSHighResolutionCapable</key><true/>
  <key>NSRequiresAquaSystemAppearance</key><false/>
</dict>
</plist>
PLIST

# The bundle is unsigned: macOS refuses to open an unsigned app on another
# machine, and signing needs a developer identity this repository does not have.
echo "==> The bundle is in $BUNDLE"
echo "    It is unsigned. On the machine that built it:  open \"$BUNDLE\""
echo "    On another machine, right-click > Open, or run:  xattr -dr com.apple.quarantine \"$BUNDLE\""
