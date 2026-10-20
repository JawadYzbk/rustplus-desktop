@echo off
setlocal
set "CONFIG=%~1"
if "%CONFIG%"=="" set "CONFIG=Debug"
if not "%~1"=="" shift
dotnet run --project "%~dp0src\RustPlusDesk.Avalonia\RustPlusDesk.Avalonia.csproj" --configuration %CONFIG% %*
