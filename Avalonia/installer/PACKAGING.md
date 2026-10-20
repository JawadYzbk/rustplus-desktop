# Packaging the Avalonia preview

The Avalonia build is a self-contained .NET application: one executable plus its
assemblies and the native runtimes it needs. Everything below produces the
folder a user installs, and nothing here uploads or publishes anything.

| Platform | What the script produces | Needs |
| --- | --- | --- |
| Windows | `artifacts/publish/win-x64/` and an Inno Setup installer in `artifacts/installer/` | the .NET 10 SDK; Inno Setup 6 for the installer |
| macOS | `artifacts/publish/RustPlus Desktop.app` | the .NET 10 SDK; an Apple silicon or Intel Mac for the matching runtime |
| Linux | `artifacts/publish/linux-x64/` and a per-user installation | the .NET 10 SDK |

## Windows

```powershell
# The publish folder on its own
dotnet publish src/RustPlusDesk.Avalonia -c Release -r win-x64 --self-contained true -o artifacts/publish/win-x64

# The publish folder and the installer
powershell -File installer/build-installer.ps1
```

`build-installer.ps1` publishes first and then compiles `installer/Setup.iss`
with Inno Setup 6. When Inno Setup is not installed it stops after the publish
and says so, rather than failing. The installer installs per machine (it needs
administrator rights), creates a Start-menu entry and an optional desktop icon,
and keeps the user's data: `profiles.json` and the settings live in the
application-data folder, not in the install directory, so an upgrade or an
uninstall leaves them alone.

The installer uses its own application id, so the preview installs beside the
shipped Windows app instead of replacing it. Both can be present.

## macOS

```bash
./installer/make-macos-app.sh osx-arm64   # Apple silicon
./installer/make-macos-app.sh osx-x64     # Intel
```

The script publishes the self-contained build and wraps it in the bundle layout
macOS expects, with an `Info.plist` beside the executable. The bundle is
**unsigned**: signing needs a developer identity that is not in this repository.
On the machine that built it, `open "RustPlus Desktop.app"` works. On another
machine macOS refuses to open an unsigned app until it is cleared:

```bash
xattr -dr com.apple.quarantine "RustPlus Desktop.app"
```

The native macOS appearance (menu conventions, system fonts, the system file
dialog) is still pending: the app runs on macOS with the built-in Fluent theme.

## Linux

```bash
./installer/install-linux.sh
```

The script publishes the self-contained build and installs it for the current
user: the files under `~/.local/opt/rustplus-desktop`, a launcher in
`~/.local/bin`, a desktop entry and an XDG autostart entry. The autostart
entry is installed switched off; the app registers it itself when "launch at
sign-in" is turned on in the settings.

## What is not in these packages

- **Signing.** Neither the Windows installer nor the macOS bundle is signed, so
  Windows SmartScreen and Gatekeeper will warn on a first run from a
  downloaded file.
- **Automatic updates.** The app does not update itself yet. A new build is
  installed by running the installer again; the user's data is kept.
- **Native macOS and Linux appearance.** The app runs on all three platforms
  with the built-in Fluent theme. The system-native controls, fonts and menus
  are still pending.
- **The shipped Windows app's extra files.** The preview has no embedded
  WebView2 payload of its own (it uses the system's), and it has no map-3D
  helper process.

## Where the user's data lives

The preview reads the same `profiles.json` the shipped app writes, from the
operating system's application-data folder, and opens it in place. The settings
file, the chat history, the death log and the local map cache sit beside it.
Copying that file to another machine moves the whole configuration.
