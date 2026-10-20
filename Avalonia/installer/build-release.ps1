<#
.SYNOPSIS
    Builds the release assets for the Avalonia preview.
.DESCRIPTION
    Publishes the self-contained Windows build, compiles the Inno Setup installer and copies both
    into a release folder with the names the app's update check looks for:

      RustPlusDesk-Avalonia-Setup.exe   the installer
      RustPlusDesk-Avalonia-win-x64.zip the portable build

    The update card in the app reads the newest published release and offers the installer asset
    with exactly that name, so a release built here is found by an installed app.
.EXAMPLE
    .\build-release.ps1
.EXAMPLE
    .\build-release.ps1 -Version 0.2.0
#>
[CmdletBinding()]
param(
    [string]$Version = '',
    [string]$Runtime = 'win-x64'
)

$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $PSScriptRoot
$AppProject = Join-Path $Root 'src\RustPlusDesk.Avalonia\RustPlusDesk.Avalonia.csproj'
$PublishDir = Join-Path $Root "artifacts\publish\$Runtime"
$ReleaseDir = Join-Path $Root "artifacts\release\$Runtime"

if (-not $Version)
{
    $Version = (Select-String -Path $AppProject -Pattern '<Version>([^<]+)</Version>').Matches[0].Groups[1].Value
}

Write-Host "==> Publishing the $Runtime build..." -ForegroundColor Cyan
dotnet publish $AppProject --configuration Release --runtime $Runtime --self-contained true --output $PublishDir
if ($LASTEXITCODE -ne 0) { throw 'The publish failed.' }

Write-Host "==> Assembling the release folder..." -ForegroundColor Cyan
if (Test-Path $ReleaseDir) { Remove-Item $ReleaseDir -Recurse -Force }
New-Item -ItemType Directory -Path $ReleaseDir -Force | Out-Null

$installer = Join-Path $ReleaseDir 'RustPlusDesk-Avalonia-Setup.exe'
$iscc = Get-Command 'ISCC.exe' -ErrorAction SilentlyContinue
if (-not $iscc)
{
    $candidate = Join-Path ${env:ProgramFiles(x86)} 'Inno Setup 6\ISCC.exe'
    if (Test-Path $candidate) { $iscc = @{ Source = $candidate } }
}
if ($iscc)
{
    Write-Host "==> Compiling the installer..." -ForegroundColor Cyan
    & $iscc.Source "/DMyAppVersion=$Version" (Join-Path $PSScriptRoot 'Setup.iss')
    if ($LASTEXITCODE -ne 0) { throw 'Inno Setup failed.' }
    Copy-Item (Join-Path $Root "artifacts\installer\RustPlusDesk-Avalonia-Setup.exe") $installer -Force
}
else
{
    Write-Warning 'Inno Setup 6 was not found, so the installer asset is missing from this release.'
}

$zip = Join-Path $ReleaseDir 'RustPlusDesk-Avalonia-win-x64.zip'
if (Test-Path $zip) { Remove-Item $zip -Force }
Compress-Archive -Path (Join-Path $PublishDir '*') -DestinationPath $zip -CompressionLevel Optimal

@"
Rust+ Desktop (Avalonia preview) $Version
Runtime: $Runtime

The installer is the Windows build. The zip is the portable build: unpack it and run
RustPlusDesk.Avalonia.exe. The app keeps the user's data in the application-data folder, so
upgrading replaces the program and leaves the servers, the devices and the settings alone.
"@ | Set-Content (Join-Path $ReleaseDir 'RELEASE-NOTES.txt') -Encoding utf8

Write-Host "==> The release assets are in $ReleaseDir" -ForegroundColor Green
Get-ChildItem $ReleaseDir | Select-Object Name, Length
