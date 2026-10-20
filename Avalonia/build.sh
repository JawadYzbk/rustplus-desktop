#!/usr/bin/env bash
set -e

# ==============================================================================
#  Rust+ Desktop — Avalonia Build & Dev Script (Bash / macOS / Linux)
# ==============================================================================

DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SLN_PATH="$DIR/RustPlusDesk.CrossPlatform.sln"
APP_PROJECT="$DIR/src/RustPlusDesk.Avalonia/RustPlusDesk.Avalonia.csproj"
TEST_PROJECT="$DIR/tests/RustPlusDesk.CrossPlatform.Tests/RustPlusDesk.CrossPlatform.Tests.csproj"

ACTION="${1:-build}"
CONFIG="${2:-Debug}"

case "$ACTION" in
    build)
        echo "==> [Avalonia] Building solution ($CONFIG)..."
        dotnet build "$SLN_PATH" --configuration "$CONFIG"
        ;;
    release)
        echo "==> [Avalonia] Building solution (Release)..."
        dotnet build "$SLN_PATH" --configuration Release
        ;;
    run)
        echo "==> [Avalonia] Running application ($CONFIG)..."
        dotnet run --project "$APP_PROJECT" --configuration "$CONFIG"
        ;;
    dev|watch)
        echo "==> [Avalonia] Starting Hot Reload watch loop ($CONFIG)..."
        export DOTNET_WATCH_RESTART_ON_RUDE_EDIT=1
        dotnet watch --project "$APP_PROJECT" run --configuration "$CONFIG"
        ;;
    test)
        echo "==> [Avalonia] Running cross-platform tests ($CONFIG)..."
        dotnet test "$TEST_PROJECT" --configuration "$CONFIG"
        ;;
    clean)
        echo "==> [Avalonia] Cleaning solution..."
        dotnet clean "$SLN_PATH"
        ;;
    help|-h|--help)
        echo "Usage: ./build.sh [build|release|run|dev|test|clean|help] [Debug|Release]"
        ;;
    *)
        echo "[ERROR] Unknown command: $ACTION"
        echo "Usage: ./build.sh [build|release|run|dev|test|clean|help]"
        exit 1
        ;;
esac
