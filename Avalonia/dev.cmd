@echo off
setlocal
set DOTNET_WATCH_RESTART_ON_RUDE_EDIT=1
if not "%~1"=="" shift
dotnet watch --project "%~dp0src\RustPlusDesk.Avalonia\RustPlusDesk.Avalonia.csproj" run --configuration Debug %*
