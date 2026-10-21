@echo off
REM ============================================================================
REM  Fast dev loop: Hot Reload for Rust+ Desktop (Avalonia Cross-Platform)
REM  Edits to C# method bodies and AXAML apply to the RUNNING app in milliseconds.
REM
REM  Just run:  dev-avalonia.cmd
REM
REM  While it's running:
REM    - Save a .cs / .axaml file  -> change is hot-applied instantly.
REM    - "Rude edits" (new fields, signatures, etc.) -> auto-restarts the app.
REM    - Press Ctrl+R in this window to force a manual restart.
REM    - Press Ctrl+C to stop.
REM ============================================================================

setlocal
set DOTNET_WATCH_RESTART_ON_RUDE_EDIT=1

dotnet watch --project "%~dp0Avalonia\src\RustPlusDesk.Avalonia\RustPlusDesk.Avalonia.csproj" run --configuration Debug %*