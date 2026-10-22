# Asset Manager

A Google Apps Script add-on for a Google Sheet that tracks company devices (laptops, phones, monitors…) and who is using them. It adds a dashboard to the sheet and keeps the staff list and the asset list in step automatically.

## Features

- **Dashboard** (pop-up in the sheet or full-screen browser tab) with light / dark theme
  - Overview: KPIs, charts by type, department and condition, recent activity
  - Assets: search and filter, click a row to see the description and previous holders
  - Staff: what each person holds now and has used before
  - Alerts: available devices, pending checks, repairs, unknown holders, status conflicts, old devices, missing info
  - History: every assign / return / status change, with who did it
- **Assign, return, change status** of a device from the dashboard
- **Add device** (Asset ID suggested automatically) and **Add staff** (inserted under their department)
- **USING / stock areas**: assigned devices move under the `USING` banner, under their type heading; devices set to *In stock* move back to the stock area
- **Previous Holder** keeps everyone who has had a device, oldest first
- **Staffs sheet sync**: the `<Type> ID` columns (Laptop ID, Phone ID…) are filled with linked asset IDs and the checkboxes ticked
- **Type IDs in Staffs**: typing or removing an asset ID in a Staffs `<Type> ID` cell assigns or returns that device

## Sheets

### AllAssetLists
| Asset ID | Using By Name | Asset Type | Brand / Model | Description | Purchase Date | Condition | Previous Holder | Remarks | Status |
|---|---|---|---|---|---|---|---|---|---|

- A **banner row** (only text in the first column, e.g. `LAPTOP`) is a type heading.
- Row 2 can be a `USING` banner: devices held by staff sit under it, grouped by type heading. The stock area follows (optionally started by a `STOCK` banner).
- The serial number can be its own column (`Serial`) or a line `S/N: …` in the Description.

### Staffs
| Name | Position | Email | Contact | Laptop | Laptop ID | Phone | Phone ID | … |
|---|---|---|---|---|---|---|---|---|

- A row with only a name (no position, email or contact) is a **department banner**; staff below it belong to that department.
- Device columns come in pairs: a checkbox (`Laptop`) and an ID column (`Laptop ID`). The header word must match the Asset Type (`LAPTOP`).

### History
Created by **Set up sheets**. One row per change: time, asset, action, from, to, condition, done by, note.

## Statuses

| Status | Meaning |
|---|---|
| In use | Held by a staff member |
| Pending | Returned, waiting to be checked — set it to *In stock* when OK |
| In stock | Free to assign |
| Repair | Damaged / at the repair shop |
| Lost | Missing |
| Retired | No longer used |

## Menu: 🗂 Asset Manager

| Item | What it does |
|---|---|
| Open dashboard | Opens the dashboard as a pop-up |
| Open dashboard in new tab (full screen) | Opens the web-app version in its own tab |
| Set up sheets (run once) | Adds the Status column and dropdowns, creates History |
| Sync Staffs sheet from assets | Rewrites the Staffs ID columns and checkboxes from AllAssetLists |
| Import holders from Staffs ID columns | The reverse: writes holders into AllAssetLists from the IDs typed in Staffs (shows a preview first) |

## Install

Requires [Node.js](https://nodejs.org) and [clasp](https://github.com/google/clasp).

```bash
npm install -g @google/clasp
clasp login
clasp push          # uploads Code.js, Dashboard.html, appsscript.json
```

`.clasp.json` points to the Apps Script project bound to the spreadsheet; change `scriptId` to use your own.

Then in the spreadsheet: reload, and run **🗂 Asset Manager → Set up sheets** once.

### Full-screen tab (web app)
1. Apps Script editor → **Deploy → New deployment → Web app → Deploy** (settings come from `appsscript.json`: runs as the user accessing it, so only people with access to the sheet can see data).
2. After later code changes: `clasp push`, then **Deploy → Manage deployments → Edit → Version: New version**.

## Configuration

At the top of `Code.js`, `CONFIG` holds the sheet names, Staffs column letters, status and condition lists, and how many years make a device "old". `TYPE_ALIASES` maps a Staffs header word to an Asset Type when they differ (e.g. `'PHONE': 'MOBILE'`).

## Files

| File | Purpose |
|---|---|
| `Code.js` | Server code: menu, reading/writing sheets, actions, sync, triggers |
| `Dashboard.html` | Dashboard UI (HTML, CSS, JS, Chart.js) |
| `appsscript.json` | Apps Script manifest (time zone, runtime, web app settings) |
| `.clasp.json` | clasp project link |
