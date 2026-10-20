; =============================================
; Rust+ Desktop (Avalonia preview) - Windows installer
; =============================================
; Build the publish folder first:
;   dotnet publish src/RustPlusDesk.Avalonia -c Release -r win-x64 --self-contained true -o artifacts/publish/win-x64
; Then compile this script with Inno Setup 6 (ISCC.exe):
;   powershell -File installer/build-installer.ps1
; =============================================

#define MyAppName      "RustPlus Desktop"
#define MyAppExeName   "RustPlusDesk.Avalonia.exe"
#define MyAppPublisher "Rust+ Desktop"
#define MyAppURL       "https://github.com/Pronwan/rustplus-desktop"
; A separate id from the Windows app: the preview installs beside it, so the
; shipped app keeps working while the cross-platform build is tried out.
#define MyAppId        "{{3F1C6A24-9C0E-4E7B-9F3A-2B7C6D5E4F30}}"
#ifndef MyAppVersion
  #define MyAppVersion "0.1.0"
#endif

[Setup]
AppId={#MyAppId}
AppName={#MyAppName} (preview)
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
DefaultDirName={autopf}\RustPlusDesk Avalonia
DefaultGroupName={#MyAppName}
UsePreviousAppDir=yes
UsePreviousGroup=yes
CreateUninstallRegKey=yes
OutputDir=..\artifacts\installer
OutputBaseFilename=RustPlusDesk-Avalonia-Setup
Compression=lzma2/max
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64
ArchitecturesAllowed=x64compatible
DisableProgramGroupPage=yes
PrivilegesRequired=admin
UninstallDisplayIcon={app}\{#MyAppExeName}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; The self-contained publish folder: the executable, its assemblies and the native runtimes.
Source: "..\artifacts\publish\win-x64\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#MyAppName}}"; Flags: nowait postinstall skipifsilent
