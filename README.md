# PPT ToolBox

<p align="center">
  <img src="docs/git-readmeheader.png" alt="PPT ToolBox" width="100%"/>
</p>

> A PowerPoint add-in that puts professional formatting controls one click away — without leaving your slide.

[![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20Mac%20%7C%20Web-0078d4?logo=windows)](https://github.com/ravitejachillara/PPT-ToolBox)
[![Office](https://img.shields.io/badge/Office-2016%2B-D83B01?logo=microsoft-office)](https://github.com/ravitejachillara/PPT-ToolBox)
[![.NET](https://img.shields.io/badge/.NET_Framework-4.8-512BD4?logo=dotnet)](https://github.com/ravitejachillara/PPT-ToolBox)
[![License](https://img.shields.io/badge/license-MIT-green)](LICENSE)

---

## What it does ?

PPT ToolBox adds a **"PPT Tools"** toggle button to PowerPoint's **Home tab**. Clicking it opens a side panel with five tabs:

| Tab | Controls |
|-----|----------|
| **Arrange** | Align (L/C/R/T/M/B), Distribute H/V, Z-order (Fwd/Back/Front/Back), Match size, Exact W/H/X/Y in cm |
| **Font** | Family, size, Bold/Italic/Underline, colour picker, colour swatches |
| **Para** | Align L/C/R/J, line spacing, space before/after |
| **Fill** | Fill colour, No Fill, transparency slider, brand swatches, outline colour & width |
| **Shadow** | Soft / Hard / Bottom / Perspective presets, Remove |

Everything works on **multi-selection** — select 10 shapes, hit Align Left, done.

---

## Features

- **Editable colour swatches** — 12 slots, saved to `%APPDATA%\PPTToolbox\swatches.json`. Click *Edit Swatches* to replace any colour; persists across sessions. Same swatch palette shared across Fill, Font colour, and Outline.
- **No-admin install** — registers entirely in `HKCU`. No elevated rights needed, compatible with locked-down corporate machines.
- **Silent deploy** — `PPTToolbox_Setup.exe /SILENT /SUPPRESSMSGBOXES /NORESTART` for Intune/SCCM.
- **Zero dependencies** — ships as a self-contained add-in; no external packages or internet access required.

---

## Requirements

### Windows (VSTO add-in)

| Requirement | Minimum |
|-------------|---------|
| Windows | 10 or 11 |
| Microsoft Office | 2016, 2019, 2021, or Microsoft 365 |
| .NET Framework | 4.8 (pre-installed on all Windows 10+ machines) |
| Architecture | x86 or x64 |

### Mac / Web (Office Web Add-in)

| Requirement | Details |
|-------------|---------|
| macOS | Any version running Microsoft 365 for Mac |
| Microsoft 365 | PowerPoint for Mac, Windows, or Web |
| Internet | Required (add-in hosted on GitHub Pages) |

---

## Installation

### Windows

#### Option 1 — Installer (recommended)

1. Download `PPTToolbox_Setup.exe` from [Releases](https://github.com/ravitejachillara/PPT-ToolBox/releases)
2. Run it — no admin prompt required
3. Open PowerPoint; the **PPT Tools** button appears in the Home tab

> If Windows shows "Unknown Publisher" warning: click **More info → Run anyway** (one-time per machine).

#### Option 2 — Xcopy (no installer)

```bat
xcopy_deploy.bat
```

Run as the target user. Copies files to `%APPDATA%\PPTToolbox` and writes the HKCU registry key.

#### Option 3 — Silent deploy (Intune / SCCM)

```bat
PPTToolbox_Setup.exe /SILENT /SUPPRESSMSGBOXES /NORESTART
```

Uninstall:

```bat
%APPDATA%\PPTToolbox\unins000.exe /SILENT
```

---

## Uninstall

**Via Control Panel** → Programs → PPT Toolbox → Uninstall

**Manual:**

```bat
reg delete "HKCU\Software\Microsoft\Office\PowerPoint\Addins\PPTToolbox" /f
rmdir /s /q "%APPDATA%\PPTToolbox"
```

### Mac / Web (Microsoft 365)

The Web Add-in runs in PowerPoint for Mac, Windows, and Browser — no `.exe` needed.

#### Option 1 — Shell script (recommended for Mac)

1. Download [`manifest.xml`](https://github.com/ravitejachillara/PPT-ToolBox/raw/main/PPTToolbox_WebAddin/manifest.xml) and [`install-mac.sh`](https://github.com/ravitejachillara/PPT-ToolBox/raw/main/install-mac.sh) to your **Downloads** folder
2. Open Terminal and run:

```bash
chmod +x ~/Downloads/install-mac.sh
~/Downloads/install-mac.sh
```

The script will:
- Quit PowerPoint if it is open
- Copy `manifest.xml` to the correct Office add-in folder
- Clear the add-in cache
- Print next steps

3. Open PowerPoint → **Insert → Add-ins → My Add-ins** → select **PPT Tools**

#### Option 2 — Manual (Mac)

```bash
# Create the folder if it doesn't exist
mkdir -p ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef

# Copy the manifest
cp ~/Downloads/manifest.xml \
   ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/
```

Then restart PowerPoint and load via **Insert → Add-ins → My Add-ins**.

#### Option 3 — Microsoft 365 Admin Center (organisation-wide)

Upload `manifest.xml` via the [Microsoft 365 Admin Center](https://admin.microsoft.com) → **Settings → Integrated apps**. The add-in will appear in PowerPoint for all users in your tenant — no per-machine steps required.

---

## Building from source

### Prerequisites

- Visual Studio 2022 (Community or higher)
- Workload: **Office/SharePoint development** (install via VS Installer)
- Inno Setup 6 — [download](https://jrsoftware.org/isinfo.php)

### Build add-in

```bat
# Open PPTToolbox_VSTO\PPTToolbox.sln in Visual Studio 2022
# Switch configuration to Release, then Build > Build Solution
```

Or from a Developer Command Prompt:

```bat
msbuild PPTToolbox_VSTO\PPTToolbox\PPTToolbox.csproj /p:Configuration=Release
```

### Build installer

```bat
cd PPTToolbox_VSTO\Installer
build_installer.bat
```

Output: `PPTToolbox_VSTO\dist\PPTToolbox_Setup.exe`

---

## Project structure

```
PPTToolbox_VSTO/                   ← Windows VSTO add-in
├── PPTToolbox/
│   ├── BrandingConfig.cs          # All brand colours, default swatches, fonts
│   ├── SwatchStore.cs             # Persistent swatch library (%APPDATA%\…\swatches.json)
│   ├── FormattingEngine.cs        # All COM/PowerPoint operations
│   ├── RibbonPPT.cs/.xml          # Home-tab toggle button (IRibbonExtensibility)
│   ├── TaskPaneControl.cs         # Task pane event handlers
│   ├── TaskPaneControl.Designer.cs  # Task pane UI layout
│   ├── ThisAddIn.cs               # Add-in startup, CustomTaskPane wiring
│   └── Resources/
│       ├── ribbon_icon.png        # 32×32 icon for the Home tab button
│       ├── company_logo.png       # Shown in task pane footer — replace with your logo
│       └── app_icon.ico           # Installer icon
├── Installer/
│   ├── setup.iss                  # Inno Setup script
│   └── build_installer.bat        # MSBuild + Inno Setup in one step
└── Deployment/
    ├── xcopy_deploy.bat            # No-installer deployment
    └── README_IT.txt               # IT department deployment guide

PPTToolbox_WebAddin/               ← Mac / Web add-in (Office.js)
├── src/taskpane/
│   ├── taskpane.ts                # All Office.js formatting operations
│   ├── taskpane.html              # 5-tab dark UI
│   └── taskpane.css               # Dark theme
├── assets/                        # Icons (16/32/80 px)
├── manifest.xml                   # Production (GitHub Pages)
├── manifest.dev.xml               # Development (localhost:3000)
└── webpack.config.js

install-mac.sh                     ← One-command Mac installer script
docs/
└── git-readmeheader.png           # README banner image
```

---

## Customising swatches

Swatches are stored at:

```
%APPDATA%\PPTToolbox\swatches.json
```

Example file (12 hex colours):

```json
["#1A1A2E","#E94F37","#FFFFFF","#000000","#2E86AB","#F6C90E","#3DC47E","#A83F9E","#FF6B35","#04A777","#D62246","#C0C0C0"]
```

Delete the file to reset to defaults.

---

## Troubleshooting

**Add-in not visible in ribbon**
1. PowerPoint → File → Options → Add-ins
2. Manage: **COM Add-ins** → Go
3. Check **PPT Tools** is ticked
4. If listed as *inactive*: remove it, then re-run `xcopy_deploy.bat`

**Disabled by Office Trust Center**
File → Options → Trust Center → Trust Center Settings → Trusted Locations → Add `%APPDATA%\PPTToolbox`

---

## Contributing

Pull requests are welcome. For significant changes, open an issue first to discuss what you'd like to change.

---

## Author

Made with ♥ by **Ravi Teja Chillara**

---

## License

[MIT](LICENSE)
