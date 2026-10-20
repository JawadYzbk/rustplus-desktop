<#
.SYNOPSIS
    Builds the Windows installer for the Avalonia preview.
.DESCRIPTION
    Publishes the self-contained win-x64 build and compiles the Inno Setup script.
    Inno Setup 6 is not bundled: install it from https://jrsoftware.org/isdl.php and
    make sure ISCC.exe is on PATH (or pass -IsccPath).
.EXAMPLE
    .\build-installer.ps1
.EXAMPLE
    .\build-installer.ps1 -Configuration Release -Runtime win-x64
#>
[CmdletBinding()]
param(
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Release',

    [string]$Runtime = 'win-x64',

    [string]$IsccPath = ''
)

$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$AppProject = Join-Path $Root 'src\RustPlusDesk.Avalonia\RustPlusDesk.Avalonia.csproj'
$PublishDir = Join-Path $Root "artifacts\publish\$Runtime"
$InstallerDir = Join-Path $Root 'artifacts\installer'
$SetupScript = Join-Path $PSScriptRoot 'Setup.iss'

Write-Host "==> Publishing the $Runtime build..." -ForegroundColor Cyan
dotnet publish $AppProject --configuration $Configuration --runtime $Runtime --self-contained true --output $PublishDir
if ($LASTEXITCODE -ne 0) { throw "The publish failed." }

if (-not $IsccPath) {
    $found = Get-Command 'ISCC.exe' -ErrorAction SilentlyContinue
    if (-not $found) {
        $candidate = Join-Path ${env:ProgramFiles(x86)} 'Inno Setup 6\ISCC.exe'
        if (Test-Path $candidate) { $IsccPath = $candidate }
        else {
            Write-Warning "Inno Setup 6 was not found. The publish is ready in $PublishDir."
            Write-Warning "Install Inno Setup from https://jrsoftware.org/isdl.php and run this script again."
            exit 0
        }
    }
    else { $IsccPath = $found.Source }
}

New-Item -ItemType Directory -Path $InstallerDir -Force | Out-Null
Write-Host "==> Compiling the installer..." -ForegroundColor Cyan
& $IsccPath "/DMyAppVersion=$((Get-Content $AppProject | Select-String '<Version>([^<]+)</Version>').Matches[0].Groups[1].Value)" $SetupScript
if ($LASTEXITCODE -ne 0) { throw "Inno Setup failed." }

Write-Host "==> The installer is in $InstallerDir" -ForegroundColor Green
