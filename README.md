# RouteDash

RouteDash is a Google Apps Script dashboard optimized for kiosk displays and wide monitors. The app now uses a responsive tile grid, a Sheets-backed data model, and batched edit writes via `google.script.run`.

## What's new in this refactor
- Replaced draggable/free-floating windows with a structured CSS Grid dashboard.
- Added **Display mode** (clean TV output) and **Edit mode** (tile/data management).
- Moved persistence to a dedicated Google Spreadsheet (META + IMAGES + data tabs).
- Added workbook-style tab support and server-side formula handling using Google Sheets formulas.
- Added a calm single-image viewer (fade transition, reduced-motion aware, lazy-load + prefetch).
- Added a migration function to move legacy Script Properties dashboard data to Sheets.

## Apps Script setup
1. Open/create an Apps Script project.
2. Copy `Code.gs` and `index.html` into that project.
3. Deploy as a Web App (execute as the deployment owner; grant kiosk access as needed).
4. Open the web app URL. On first run, `getOrCreateDataSpreadsheet()` creates and stores a spreadsheet ID in Script Properties (`dataSpreadsheetId`).

## Data spreadsheet model
The spreadsheet created by RouteDash is the primary datastore:

- `META` tab
  - Key/value JSON records for dashboard config + layout.
  - `appConfig` includes values like `headerTitle`, `driverOfWeek`, and image rotation timing.
  - `dashboardLayout` includes tile order, type, sheet source, and size preset (`small`, `wide`, `tall`, `big`).
- `IMAGES` tab
  - Metadata index for uploaded Drive images (`imageId`, `fileId`, name, mime type, timestamps).
- Data tabs (e.g. `Sheet1`, `Sales`, `Stops`)
  - Your workbook-like tile datasets.
  - Formulas are authored and evaluated by Google Sheets, including cross-tab references like `=Sheet2!A1`.

## Server API highlights
Primary server functions in `Code.gs`:
- `getOrCreateDataSpreadsheet()`
- `getAppState()`
- `getSheetRange(sheetName, a1Range, options)`
- `applySheetEdits(edits)` (batch write)
- `createSheetTab(name)`, `renameSheetTab(oldName, newName)`, `deleteSheetTab(name)`
- `getImagesIndex()`, `addImages(base64List)`, `deleteImage(imageId)`, `getImageById(imageId)`
- `migrateLegacyData()`

## Running migration from legacy Script Properties
If you have older frame/layout/image data in Script Properties:

1. Open Apps Script editor.
2. Run `migrateLegacyData()` once.
3. Confirm return payload contains `migrated: true`.
4. Reload the web app.

Migration behavior:
- Converts legacy `frames` to grid tile records in `META.dashboardLayout`.
- Pulls legacy `headerTitle` and `driverOfWeek` into `META.appConfig`.
- Converts legacy `driverImageIds` to `IMAGES` index rows (without duplicating files).

## Image management workflow
- In Edit mode, upload files from the Images panel.
- Upload calls `addImages(base64List)` and stores metadata in the `IMAGES` sheet while binaries remain in Drive.
- Image tiles show one image at a time with a gentle fade.
- Current + next images are fetched/lazy-prefetched using `getImageById(imageId)`.

## Local testing
```bash
npm install
npm test
```

Existing Jest coverage validates weather code mapping parity in `weather.js`.
