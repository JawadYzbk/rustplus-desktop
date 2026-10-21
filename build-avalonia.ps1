<#
.SYNOPSIS
    Build and dev helper script for Rust+ Desktop Avalonia.
.DESCRIPTION
    Builds, runs, tests, or publishes the Avalonia cross-platform solution.
.EXAMPLE
    .\build-avalonia.ps1
.EXAMPLE
    .\build-avalonia.ps1 -Action Run
.EXAMPLE
    .\build-avalonia.ps1 -Action Watch
.EXAMPLE
    .\build-avalonia.ps1 -Configuration Release
.EXAMPLE
    .\build-avalonia.ps1 -Action Publish
#>
[CmdletBinding()]
param(
    [ValidateSet('Build', 'Run', 'Watch', 'Dev', 'Test', 'Publish', 'Clean')]
    [string]$Action = 'Build',

    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Debug',

    [string]$Runtime = 'win-x64'
)

$ErrorActionPreference = 'Stop'
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$SolutionPath   = Join-Path $ScriptDir "Avalonia\RustPlusDesk.CrossPlatform.sln"
$AppProject     = Join-Path $ScriptDir "Avalonia\src\RustPlusDesk.Avalonia\RustPlusDesk.Avalonia.csproj"
$TestProject    = Join-Path $ScriptDir "Avalonia\tests\RustPlusDesk.CrossPlatform.Tests\RustPlusDesk.CrossPlatform.Tests.csproj"
$ArtifactsDir   = Join-Path $ScriptDir "Avalonia\artifacts\$Runtime"

switch ($Action) {
    'Build' {
        Write-Host "==> [Avalonia] Building solution ($Configuration)..." -ForegroundColor Cyan
        dotnet build $SolutionPath --configuration $Configuration
    }
    'Run' {
        Write-Host "==> [Avalonia] Running application ($Configuration)..." -ForegroundColor Cyan
        dotnet run --project $AppProject --configuration $Configuration
    }
    { $_ -in @('Watch', 'Dev') } {
        Write-Host "==> [Avalonia] Starting Hot Reload watch loop ($Configuration)..." -ForegroundColor Cyan
        Write-Host "    Press Ctrl+R to restart, Ctrl+C to stop." -ForegroundColor Yellow
        $env:DOTNET_WATCH_RESTART_ON_RUDE_EDIT = '1'
        dotnet watch --project $AppProject run --configuration $Configuration
    }
    'Test' {
        Write-Host "==> [Avalonia] Running cross-platform tests ($Configuration)..." -ForegroundColor Cyan
        dotnet test $TestProject --configuration $Configuration
    }
    'Publish' {
        Write-Host "==> [Avalonia] Publishing self-contained $Runtime preview build..." -ForegroundColor Cyan
        dotnet publish $AppProject --configuration $Configuration --runtime $Runtime --self-contained true --output $ArtifactsDir
        Write-Host "==> [Avalonia] Published successfully to: $ArtifactsDir" -ForegroundColor Green
    }
    'Clean' {
        Write-Host "==> [Avalonia] Cleaning solution..." -ForegroundColor Cyan
        dotnet clean $SolutionPath
    }
}