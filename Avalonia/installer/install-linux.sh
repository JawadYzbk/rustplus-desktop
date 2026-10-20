#!/usr/bin/env bash
# =============================================================================
#  Rust+ Desktop (Avalonia preview) - Linux installer
# =============================================================================
#  Publishes the self-contained linux-x64 build and installs it for the current
#  user: the files under ~/.local/opt/rustplus-desktop, a launcher on the PATH,
#  a desktop entry and the XDG autostart entry the app's "launch at sign-in"
#  setting writes.
#
#  Usage: ./install-linux.sh [--prefix DIR]
# =============================================================================
set -euo pipefail

RUNTIME="linux-x64"
DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ROOT="$(cd "$DIR/.." && pwd)"
APP_PROJECT="$ROOT/src/RustPlusDesk.Avalonia/RustPlusDesk.Avalonia.csproj"
PUBLISH_DIR="$ROOT/artifacts/publish/$RUNTIME"
PREFIX="${HOME}/.local/opt/rustplus-desktop"
BIN_DIR="${HOME}/.local/bin"
DESKTOP_DIR="${HOME}/.local/share/applications"
AUTOSTART_DIR="${XDG_CONFIG_HOME:-${HOME}/.config}/autostart"

while [ $# -gt 0 ]; do
  case "$1" in
    --prefix) PREFIX="$2"; shift 2 ;;
    *) echo "Unknown option: $1"; exit 1 ;;
  esac
done

echo "==> Publishing the $RUNTIME build..."
dotnet publish "$APP_PROJECT" --configuration Release --runtime "$RUNTIME" --self-contained true --output "$PUBLISH_DIR"

echo "==> Installing into $PREFIX ..."
mkdir -p "$PREFIX"
cp -R "$PUBLISH_DIR/." "$PREFIX/"

mkdir -p "$BIN_DIR"
cat > "$BIN_DIR/rustplus-desktop" <<LAUNCHER
#!/usr/bin/env bash
exec "$PREFIX/RustPlusDesk.Avalonia" "\$@"
LAUNCHER
chmod +x "$BIN_DIR/rustplus-desktop"

mkdir -p "$DESKTOP_DIR"
cat > "$DESKTOP_DIR/rustplus-desktop.desktop" <<DESKTOP
[Desktop Entry]
Type=Application
Name=Rust+ Desktop
Comment=Monitor Rust+ servers, devices and team chat
Exec=$BIN_DIR/rustplus-desktop
Icon=rustplus-desktop
Terminal=false
Categories=Utility;
DESKTOP

# The XDG autostart entry the app's own "launch at sign-in" switch writes. It is
# installed switched off: the app registers it itself when the switch is turned on.
mkdir -p "$AUTOSTART_DIR"
cat > "$AUTOSTART_DIR/RustPlusDesk.desktop" <<AUTOSTART
[Desktop Entry]
Type=Application
Name=Rust+ Desktop
Exec="$PREFIX/RustPlusDesk.Avalonia" --background
Terminal=false
X-GNOME-Autostart-enabled=true
AUTOSTART

echo "==> Installed."
echo "    Launcher:   $BIN_DIR/rustplus-desktop"
echo "    Desktop:    $DESKTOP_DIR/rustplus-desktop.desktop"
echo "    Autostart:  $AUTOSTART_DIR/RustPlusDesk.desktop (switched off until you turn it on in the app)"
