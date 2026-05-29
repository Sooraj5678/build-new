# Vyapar Outstanding Formatter

A browser-based tool that converts Vyapar Outstanding Excel exports into professional city-wise printable accounting reports.

## Run & Operate

- `pnpm --filter @workspace/vyapar-formatter run dev` — run the app (frontend only, no backend needed)
- `pnpm run typecheck` — full typecheck across all packages
- `pnpm run build` — typecheck + build all packages

## Stack

- pnpm workspaces, Node.js 24, TypeScript 5.9
- React + Vite (frontend only — no backend)
- TailwindCSS v4
- xlsx (Excel parsing, runs entirely in the browser)
- Browser Print API (PDF + print, no Puppeteer needed in web mode)

## Where things live

- `artifacts/vyapar-formatter/src/App.tsx` — all app logic: Excel parsing, grouping, report rendering

## Architecture decisions

- Fully client-side: Excel files are parsed in the browser using the `xlsx` library — no server upload required, completely offline-capable
- Auto-detects column names with fuzzy matching (handles "Address", "Area", "City" as the same field)
- Normalizes city names: trim + uppercase so "adhgam", "ADHGAM", "Adhgam" all group together
- PDF generation uses the browser's built-in print dialog with `@page { size: A4 portrait }` CSS — no Puppeteer dependency needed in the web version
- Courier New font throughout to match traditional accounting software feel

## Product

- Upload a Vyapar .xlsx/.xls export
- Automatically detects Party Name, Mobile, Address/City, and Outstanding Amount columns
- Groups all parties under a single city heading (no duplicate sections)
- Shows city totals and a grand total
- Print or Save as PDF using the browser's native print dialog (A4 portrait, ready for accounting use)

## User preferences

- Minimal UI — no dashboards, no fancy components, accounting software feel
- Courier New monospace font throughout
- City names normalized to uppercase, duplicates merged

## Gotchas

- Column detection is fuzzy — it searches for keywords in header names. If a column isn't found, a clear error message appears.
- Empty address/city values become "UNKNOWN" group
- The "Save as PDF" button uses browser print — user should select "Save as PDF" in the print dialog

## Pointers

- See the `pnpm-workspace` skill for workspace structure, TypeScript setup, and package details
