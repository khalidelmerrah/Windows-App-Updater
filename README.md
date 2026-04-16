# Windows App Updater

A GUI tool for batch-updating installed Windows applications using [winget](https://learn.microsoft.com/en-us/windows/package-manager/winget/) (Microsoft's package manager).

> **Original project by [ilukezippo (BoYaqoub)](https://github.com/ilukezippo/Windows-App-Updater)**
> This fork adds security hardening, bug fixes, and UI improvements.

![Python](https://img.shields.io/badge/Python-3.12+-blue)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey)
![License](https://img.shields.io/badge/License-Freeware-green)

---

## Features

### Core
- **One-click update check** - Scans all installed apps via winget and shows available updates
- **Batch update** - Select individual apps or update all at once
- **Skip / Cancel** - Skip the current app or cancel the entire batch mid-update
- **Retry Failed** - One-click to re-select and retry only failed apps
- **Real-time progress** - Dual progress bars (overall + per-app download) with animated spinner
- **Include unknown apps** - Toggle to include apps with unrecognized version numbers
- **Run as Admin** - One-click UAC elevation for silent installs
- **Self-update** - Checks GitHub releases for newer versions and can download/replace itself automatically

### Search & Navigation
- **Search/filter** - Find apps instantly by name or package ID
- **Column sorting** - Click any column header to sort ascending/descending
- **Keyboard shortcuts** - Ctrl+A (select all), Enter (start update), Escape (cancel)

### Customization
- **Dark mode** - Full dark theme with live toggle in Settings
- **Exclude list** - Right-click any app to permanently exclude from updates
- **Portable config** - All settings saved to config.json next to the EXE
- **Settings dialog** - Central place for all preferences

### Automation
- **Scheduled auto-check** - Configurable interval (1/4/8/24 hours)
- **Windows notifications** - Toast notification when updates complete
- **System restore point** - Optional restore point before batch updates (requires admin)

### Tools
- **App info popup** - Right-click or double-click any app to view details via winget
- **Update history** - Last 50 update sessions tracked, viewable in Settings
- **Export/import apps** - Export installed app list to file, bulk install on new machines
- **Update log** - Live log output with export to file, collapsible panel
- **Temp file management** - View, open, or clear temporary installer files
- **Per-app download tracking** - Right-click to open or delete downloaded installers
- **Success sound** - Plays a sound when all updates complete without errors
- **Resizable window** - Supports full-screen / maximize with responsive layout

## Download

[**Download Windows-App-Updater.exe (v2.3.0)**](https://github.com/khalidelmerrah/Windows-App-Updater/releases/download/v2.3.0/Windows-App-Updater.exe)

No installation required. Just download and run the `.exe` file.

See all releases: [Releases page](https://github.com/khalidelmerrah/Windows-App-Updater/releases)

## Requirements

- Windows 10/11
- [winget](https://learn.microsoft.com/en-us/windows/package-manager/winget/) (included with Windows 10 1809+ and Windows 11 via App Installer from Microsoft Store)

### Running from source

```bash
pip install pillow
python App-Updater.py
```

## Building the EXE

```bash
pip install pyinstaller pillow

pyinstaller --noconfirm --onefile --windowed ^
  --name "Windows-App-Updater" ^
  --icon "windows-updater.ico" ^
  --add-data "success.wav;." ^
  --add-data "kuwait.png;." ^
  --add-data "windows-updater.ico;." ^
  App-Updater.py
```

Output: `dist/Windows-App-Updater.exe`

## Security Hardening (This Fork)

This fork includes the following security fixes over the original:

### Critical
- **Batch script injection prevention** - Paths embedded in the self-update batch script are sanitized to strip dangerous characters (`"`, `%`, `^`, `&`, `|`, `<`, `>`)
- **Symlink attack protection** - Temp cleanup now rejects symlinks before changing permissions, preventing attackers from using symlinks to modify files outside the temp directory
- **Download URL validation** - Self-update downloads are restricted to HTTPS URLs from `github.com` and `objects.githubusercontent.com` only. Asset filenames are sanitized to prevent path traversal

### High
- **DLL hijacking prevention** - VC runtime checks now use absolute `%WINDIR%\System32` paths instead of bare DLL names, preventing malicious DLLs in the working directory from being loaded
- **UAC parameter injection fix** - Admin re-launch now uses `subprocess.list2cmdline()` for proper Windows argument escaping instead of manual quoting that could be bypassed with crafted arguments

### Medium
- **Dead code removal** - Removed unused `_download_file_verified` duplicate method
- **Log memory cap** - Log widget is capped at 5,000 lines to prevent unbounded memory growth during long update sessions
- **Symlink-safe temp snapshots** - Temp directory scanning uses `os.lstat()` instead of `os.stat()` to avoid following symlinks
- **Thread safety** - Added `threading.Lock()` for shared state accessed by both the UI thread and worker threads

### UI Fixes
- **Window maximize** - Removed the `maxsize` constraint that prevented maximizing/full-screen. The app list now expands vertically when the window is resized

## Credits

- **Original author:** [ilukezippo (BoYaqoub)](https://github.com/ilukezippo/Windows-App-Updater)
- **Security hardening & fixes:** [khalidelmerrah](https://github.com/khalidelmerrah/Windows-App-Updater)
