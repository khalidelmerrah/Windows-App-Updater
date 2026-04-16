# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Windows App Updater - a single-file Python/Tkinter desktop application that wraps Windows `winget` to provide a GUI for batch-updating installed applications. Packaged as a standalone EXE via PyInstaller.

## Build & Run

```bash
# Run from source (requires Python 3.12+, Windows only)
python App-Updater.py

# Install dependencies
pip install pyinstaller pillow

# Build standalone EXE (mirrors CI)
pyinstaller --noconfirm --onefile --windowed \
  --name "Windows-App-Updater" \
  --icon "windows-updater.ico" \
  --add-data "success.wav;." \
  --add-data "kuwait.png;." \
  --add-data "windows-updater.ico;." \
  App-Updater.py
# Output: dist/Windows-App-Updater.exe
```

No test suite exists. No linter is configured.

## Architecture

The entire application lives in **`App-Updater.py`** (~900 lines, single file). Key structure:

- **Top-level constants**: `APP_VERSION_ONLY`, `APP_NAME_VERSION`, `DATE_APP`, GitHub URLs, window dimensions
- **Utility functions**: `resource_path()` (PyInstaller `_MEIPASS` handling), `is_admin()`/`relaunch_as_admin()` (UAC elevation via `ShellExecuteW`), `run()` (subprocess wrapper with hidden window), `_download_file()` (atomic download with `.part` temp + MZ validation)
- **`get_winget_upgrades()`** / **`parse_table_upgrade_output()`**: Calls `winget upgrade` and parses the fixed-width column table output using header position slicing (not regex splitting)
- **`WingetUpdaterUI`**: The single Tkinter UI class containing all GUI logic:
  - Tree view with custom checkbox images for app selection
  - Threaded update loop (`update_selected_async` -> `worker`) that streams `winget upgrade --id` stdout, parses percent/size progress via regex, and updates dual progress bars
  - Self-update mechanism: checks GitHub API for newer releases, downloads EXE, writes a batch script for atomic self-replacement, handles UAC elevation for protected install directories
  - Temp file management: snapshots `%TEMP%` before/after each update to track downloaded installers per package

## Key Patterns

- **Threading model**: All long operations (check, update, clear temp, self-update) run in `threading.Thread(daemon=True)` with UI updates marshaled through `self.root.after(0, callback)`
- **winget interaction**: Always passes `--accept-source-agreements --disable-interactivity` and forces English output via `DOTNET_CLI_UI_LANGUAGE=en` env var. Hidden console window via `STARTUPINFO` + `CREATE_NO_WINDOW`
- **Resource bundling**: `resource_path()` resolves assets from PyInstaller's `_MEIPASS` temp dir when frozen, or from CWD when running from source. Bundled assets: `windows-updater.ico`, `kuwait.png`, `success.wav`
- **Version strings**: `APP_VERSION_ONLY` (e.g., "v2.2.1") is compared as parsed integer tuples via `_parse_ver_tuple()`

## CI/CD

GitHub Actions workflow (`.github/workflows/release.yml`) triggers on `v*` tags, builds EXE with PyInstaller on `windows-latest`, and creates a GitHub release with the tagged EXE attached.

## Dependencies

- **Runtime**: Python 3.12+, Windows OS, `winget` CLI installed
- **Python packages**: `pillow` (donate button gradient image generation)
- **Standard library**: `tkinter`, `subprocess`, `urllib.request`, `ctypes`, `winsound`, `json`, `re`, `threading`, `tempfile`, `shutil`
