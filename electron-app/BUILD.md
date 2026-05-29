# Vyapar Outstanding Formatter — Build Guide

## Prerequisites (Windows)

1. Install **Node.js 20+** from https://nodejs.org
2. Install **Python** (needed by some native modules): https://python.org
3. Install **Visual Studio Build Tools** (for native modules):
   https://visualstudio.microsoft.com/visual-cpp-build-tools/

## Setup

```bash
# 1. Navigate to the electron-app folder
cd electron-app

# 2. Install dependencies
npm install

# 3. Run in development mode (shows the app live)
npm run dev
```

## Build Windows EXE Installer

```bash
# Build the Windows installer (generates setup.exe)
npm run dist
```

The installer will be generated at:
```
electron-app/dist-installer/Vyapar Outstanding Formatter Setup 1.0.0.exe
```

## What gets created

- `dist-installer/Vyapar Outstanding Formatter Setup 1.0.0.exe` — Windows installer
- Double-click to install, creates desktop shortcut automatically
- Fully offline — works with no internet connection

## Features

- Upload Vyapar .xlsx / .xls export
- Auto-detects Party Name, Mobile, Address/City, Outstanding Amount columns
- Groups all parties by city (normalized: ADHGAM = adhgam = Adhgam)
- Editable report header: Company Name, Address, Mobile numbers, Report Title, Date range
- Settings saved permanently to disk (auto-loads on next open)
- Save as PDF — native file save dialog, generates real PDF
- Print — sends directly to printer via system print dialog
- A4 portrait layout, professional accounting format
- Fully offline — no internet required

## Settings storage location

Settings are saved at:
```
C:\Users\<YourName>\AppData\Roaming\vyapar-outstanding-formatter\vyapar-settings.json
```

## Troubleshooting

**"electron is not recognized"** → Run `npm install` first

**Build fails with native module error** → Install Visual Studio Build Tools (see Prerequisites)

**White screen when running** → Make sure `npm run dev` starts vite first, then electron loads it

**PDF not saving** → Make sure you have write permission to the target folder
