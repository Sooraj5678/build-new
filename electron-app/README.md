# Vyapar Outstanding Formatter

> Convert Vyapar Outstanding Excel exports into professional city-wise printable reports.  
> Fully offline Windows desktop app — no internet required.

![Build Windows EXE](https://github.com/YOUR_USERNAME/YOUR_REPO/actions/workflows/build.yml/badge.svg)

---

## Features

- **Upload** Vyapar `.xlsx` / `.xls` export
- **Auto-detect** columns: Party Name, Mobile, Address/Area/City, Outstanding Amount
- **City-wise grouping** — all parties under one heading (ADHGAM = adhgam = Adhgam)
- **Editable header** — Company Name, Address, Mobile numbers, Report Title, Date range
- **Persistent settings** — saved to disk, auto-loads every time you open the app
- **Save as PDF** — native file dialog, real PDF via Chromium engine
- **Print** — send directly to any installed printer
- **A4 portrait layout** — traditional accounting report style
- **Fully offline** — no internet, no cloud, no data upload

---

## Download (Windows EXE)

Go to the [Releases](../../releases) page and download the latest  
`VyaparOutstandingFormatter-x.x.x-Setup.exe`

Double-click to install — a desktop shortcut is created automatically.

---

## Development Setup

### Prerequisites

| Tool | Version | Link |
|------|---------|------|
| Node.js | 20+ | https://nodejs.org |
| npm | included with Node.js | — |

```bash
# Clone the repository
git clone https://github.com/YOUR_USERNAME/YOUR_REPO.git
cd YOUR_REPO/electron-app

# Install dependencies
npm install

# Run in development mode (hot-reload)
npm run dev
```

---

## Build the Windows EXE

### Option 1 — Automatic via GitHub Actions (recommended)

Push a version tag and GitHub builds the EXE automatically:

```bash
git tag v1.0.0
git push origin v1.0.0
```

The installer appears under **Releases** within ~5 minutes.

### Option 2 — Build locally on Windows

```bash
cd electron-app
npm install
npm run dist:win
```

Output: `electron-app/dist-installer/VyaparOutstandingFormatter-1.0.0-Setup.exe`

---

## GitHub Actions Setup

The workflow at `.github/workflows/build.yml` runs automatically on:

| Trigger | What happens |
|---------|-------------|
| Push to `main` | Builds EXE, uploads as workflow artifact (downloadable for 30 days) |
| Push a `v*` tag | Builds EXE **and** creates a public GitHub Release with the installer |
| Pull Request to `main` | Builds EXE to verify nothing is broken |
| Manual trigger | Run from the Actions tab anytime |

No secrets needed — uses the built-in `GITHUB_TOKEN`.

---

## Project Structure

```
electron-app/
├── electron/
│   ├── main.cjs        — Main process: window, IPC, printToPDF, settings
│   └── preload.cjs     — Context bridge: exposes electronAPI to renderer
├── src/
│   ├── App.tsx         — React UI: Excel parsing, grouping, report view
│   ├── main.tsx        — React entry point
│   ├── index.css       — Tailwind + base styles
│   └── assets/
│       └── logo.png    — App logo (transparent background)
├── build/
│   ├── icon.png        — App icon source (electron-builder converts to .ico)
│   └── icon.ico        — Windows icon (add manually or via CI)
├── index.html
├── vite.config.ts
├── tailwind.config.js
├── tsconfig.json
└── package.json        — Dependencies + electron-builder config
```

---

## Settings Storage

Settings are saved at:
```
C:\Users\<YourName>\AppData\Roaming\vyapar-outstanding-formatter\vyapar-settings.json
```

---

## License

MIT
