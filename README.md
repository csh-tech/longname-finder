<p align="center">
  <img src="assets/icon.png" width="128" alt="LongName Finder icon">
</p>

<h1 align="center">LongName Finder</h1>

<p align="center">
  Find the files with extremely long names that make Windows Explorer hang or crash,<br>
  and get them in an Excel report with clickable links.
</p>

<p align="center">
  <a href="https://github.com/csh-tech/longname-finder/releases/latest"><img src="https://img.shields.io/github/v/release/csh-tech/longname-finder?label=download" alt="Latest release"></a>
  <img src="https://img.shields.io/badge/platform-Windows-0078D6" alt="Windows">
  <img src="https://img.shields.io/badge/python-3.8%2B-3776AB" alt="Python 3.8+">
  <img src="https://img.shields.io/github/license/csh-tech/longname-finder" alt="MIT License">
</p>

<p align="center"><b>English</b> | <a href="README.he.md">עברית</a></p>

---

## Why?

Files saved from chat apps, social media or web pages often get the whole post text
as their file name. Windows allows up to **255 characters per file name**, and names
that come close to that limit are a common reason for:

- File Explorer freezing or restarting when you open a folder or search
- Files that can't be copied, moved, renamed or deleted
- Backup and sync tools skipping files

Finding them by hand is almost impossible. LongName Finder scans a folder (or an
entire drive) and lists them for you.

## Features

- Scans a folder and all its subfolders, including paths longer than 260 characters
- Checks the **file name only** (with extension), the way Windows counts it
- Excel report, longest names first, with:
  - file name and location
  - a link to open the file and a link to open its folder
  - name length and full path length
  - files at the 255-character limit highlighted
- Folders that couldn't be read are listed on a separate sheet instead of stopping the scan
- English and Hebrew interface, picked automatically from your Windows language
- Read-only: it **never** renames, moves or deletes anything
- Single portable `.exe`: no installation, no Python, nothing else to download

## Download

Get **`LongNameFinder.exe`** from the [latest release](https://github.com/csh-tech/longname-finder/releases/latest).

> **Windows SmartScreen** may warn you because the app isn't code-signed.
> Click **More info → Run anyway**. The full source code is in this repository,
> and you can build the exe yourself (see below).

## Usage

1. Double-click `LongNameFinder.exe`.
2. Choose the folder or drive to scan.
3. When the scan finishes, choose where to save the report
   (default name: `long_names_YYYYMMDD_HHMM.xlsx`). It then opens in Excel.

The interface and the report use your Windows display language (English or
Hebrew; anything else falls back to English). Use `--lang` to choose another.

### Command line

```
LongNameFinder.exe [folder] [--min N] [-o report.xlsx] [--lang en|he]
```

| Option | Description |
|---|---|
| `folder` | Folder to scan. If omitted, a folder picker opens. |
| `--min N` | Minimum file name length, incl. extension. Default: `240`. |
| `-o FILE` | Output Excel file name. |
| `--lang` | Force the interface and report language (`en` or `he`). |
| `--version` | Show the version. |

Examples:

```
LongNameFinder.exe "D:\Downloads"
LongNameFinder.exe "D:\" --min 200 -o report.xlsx
```

## What to do with the files it finds

Shorten the names, but keep the original text if it matters:

1. Select the file, press **F2**, then **Ctrl+A** and **Ctrl+C** to copy the long name.
2. Right-click → **Properties** → **Details** and paste it into **Comments**
   (works for photos and Office files), or paste it into a text file next to it.
3. Rename the file to something short.

If even Explorer can't handle a file, these usually work:
[7-Zip File Manager](https://www.7-zip.org/), or `rename` / `move` in Command Prompt
using the `\\?\` prefix, e.g. `move "\\?\D:\folder\long name.jpg" "D:\folder\short.jpg"`.

## File names vs. long paths

LongName Finder checks **file names only** (the name and its extension), not the
full folder path.

Windows has a separate limit of **260 characters for the full path**
(drive + folders + file name). Newer Windows 10/11 versions can lift that limit
for apps that support it, by enabling long paths in the registry. Run this in
PowerShell **as Administrator**, then restart the computer:

```powershell
New-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem" -Name "LongPathsEnabled" -Value 1 -PropertyType DWORD -Force
```

Or set it manually in the Registry Editor:

1. Press **Win+R**, type `regedit` and press Enter. Approve the permission prompt.
2. Go to this key (you can paste it into the address bar at the top of the window):
   ```
   HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\FileSystem
   ```
3. On the right, look for **`LongPathsEnabled`**:
   - **If it exists:** double-click it, set the value to `1` and click OK.
   - **If it doesn't exist:** right-click an empty area → **New** → **DWORD (32-bit) Value**,
     name it `LongPathsEnabled` and set its value to `1`.
4. Restart the computer.

If the value is already `1`, long paths are already enabled and nothing needs to change.

Note that File Explorer itself may still not fully support long paths even with
this setting enabled.

## Disclaimer

This software and these instructions are provided **"as is", without warranty
of any kind**. Use them at your own risk. The author is not responsible for any
damage, data loss or other problems caused by using the software or by following
these instructions, including changes to the Windows registry. Always back up
important files before renaming, moving or deleting them.

## Run from source

*Only needed if you want to work on the code. The exe already includes everything.*

Requires Python 3.8+.

```
pip install -r requirements.txt
python longname_finder.py
```

## Build the exe yourself

```
build.bat
```

The result is `dist\LongNameFinder.exe`. Releases on GitHub are built automatically
by [GitHub Actions](.github/workflows/release.yml) whenever a release is published.

## Adding a language

All interface text lives in [`locales/`](locales), one JSON file per language.
To add a language, copy `locales/en.json` to a file named with the
two-letter language code (for example `fr.json`) and translate the values.
Keep the `{placeholders}` as they are.

The language is chosen automatically from the Windows display language.
You can also put a `locales` folder next to `LongNameFinder.exe`, and it will
be used without rebuilding. Pull requests with new languages are welcome.

## License

[MIT](LICENSE)
