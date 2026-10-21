@echo off
setlocal

REM ============================================================================
REM  Rust+ Desktop - Avalonia Cross-Platform Build & Dev Helper
REM ============================================================================

set "ROOT_DIR=%~dp0"
set "SLN_PATH=%ROOT_DIR%Avalonia\RustPlusDesk.CrossPlatform.sln"
set "APP_PROJECT=%ROOT_DIR%Avalonia\src\RustPlusDesk.Avalonia\RustPlusDesk.Avalonia.csproj"
set "TEST_PROJECT=%ROOT_DIR%Avalonia\tests\RustPlusDesk.CrossPlatform.Tests\RustPlusDesk.CrossPlatform.Tests.csproj"
set "ARTIFACTS_DIR=%ROOT_DIR%Avalonia\artifacts\win-x64"

set "ACTION=%~1"
if "%ACTION%"=="" set "ACTION=build"

if /i "%ACTION%"=="help" goto :do_help
if /i "%ACTION%"=="/?" goto :do_help
if /i "%ACTION%"=="-h" goto :do_help
if /i "%ACTION%"=="--help" goto :do_help

if /i "%ACTION%"=="build" goto :do_build
if /i "%ACTION%"=="debug" goto :do_build_debug
if /i "%ACTION%"=="release" goto :do_build_release
if /i "%ACTION%"=="run" goto :do_run
if /i "%ACTION%"=="watch" goto :do_watch
if /i "%ACTION%"=="dev" goto :do_watch
if /i "%ACTION%"=="test" goto :do_test
if /i "%ACTION%"=="publish" goto :do_publish
if /i "%ACTION%"=="clean" goto :do_clean

if /i "%ACTION%"=="-c" goto :do_custom_config
if /i "%ACTION%"=="--configuration" goto :do_custom_config

echo [ERROR] Unknown command: %ACTION%
echo.
goto :do_help

:do_build
set "CONFIG=%~2"
if "%CONFIG%"=="" set "CONFIG=Debug"
echo [Avalonia] Building solution (%CONFIG%)...
dotnet build "%SLN_PATH%" --configuration %CONFIG%
goto :do_end

:do_build_debug
echo [Avalonia] Building solution (Debug)...
dotnet build "%SLN_PATH%" --configuration Debug
goto :do_end

:do_build_release
echo [Avalonia] Building solution (Release)...
dotnet build "%SLN_PATH%" --configuration Release
goto :do_end

:do_run
set "CONFIG=%~2"
if "%CONFIG%"=="" set "CONFIG=Debug"
echo [Avalonia] Running application (%CONFIG%)...
dotnet run --project "%APP_PROJECT%" --configuration %CONFIG%
goto :do_end

:do_watch
echo [Avalonia] Starting Hot Reload dev loop (dotnet watch)...
echo Press Ctrl+R to restart, Ctrl+C to stop.
set DOTNET_WATCH_RESTART_ON_RUDE_EDIT=1
dotnet watch --project "%APP_PROJECT%" run --configuration Debug
goto :do_end

:do_test
set "CONFIG=%~2"
if "%CONFIG%"=="" set "CONFIG=Debug"
echo [Avalonia] Running test suite (%CONFIG%)...
dotnet test "%TEST_PROJECT%" --configuration %CONFIG%
goto :do_end

:do_publish
echo [Avalonia] Publishing self-contained win-x64 release preview...
dotnet publish "%APP_PROJECT%" --configuration Release --runtime win-x64 --self-contained true --output "%ARTIFACTS_DIR%"
if not errorlevel 1 (
    echo.
    echo [SUCCESS] Published to: %ARTIFACTS_DIR%
)
goto :do_end

:do_clean
echo [Avalonia] Cleaning solution...
dotnet clean "%SLN_PATH%"
goto :do_end

:do_custom_config
set "CONFIG=%~2"
if "%CONFIG%"=="" set "CONFIG=Debug"
echo [Avalonia] Building solution (%CONFIG%)...
dotnet build "%SLN_PATH%" --configuration %CONFIG%
goto :do_end

:do_help
echo ============================================================================
echo  Rust+ Desktop - Avalonia Helper Script
echo ============================================================================
echo  Usage:
echo    build-avalonia.cmd [command] [options]
echo.
echo  Commands:
echo    build [Debug^|Release]  Build the solution (default: Debug)
echo    release                Build the solution in Release configuration
echo    run [Debug^|Release]    Run the Avalonia desktop application
echo    dev / watch            Run with Hot Reload (auto-restarts on rude edits)
echo    test                   Run cross-platform unit tests
echo    publish                Publish self-contained win-x64 preview build
echo    clean                  Clean build artifacts
echo    help                   Show this help message
echo.
echo  Examples:
echo    build-avalonia.cmd             (builds Debug)
echo    build-avalonia.cmd release     (builds Release)
echo    build-avalonia.cmd run         (runs the app)
echo    build-avalonia.cmd dev         (runs with Hot Reload)
echo    build-avalonia.cmd publish     (creates self-contained package)
echo ============================================================================
goto :do_end

:do_end
exit /b %ERRORLEVEL%