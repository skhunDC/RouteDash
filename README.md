# RouteDash

## Overview
RouteDash is a Google Apps Script powered operations dashboard designed for delivery and logistics teams. The app delivers a polished, kiosk-friendly experience with live weather, dynamic data visualizations, team recognition panels, and rich customization options that can be managed directly from the published web app.

## Key features
- **Customizable hero header** – Displays the team logo, editable title, and live date/time alongside current conditions pulled from Open-Meteo. Weather codes are translated into readable text so the display is easy to understand at a glance.
- **Movable data frames** – Operators can create, drag, resize, and rename frames. Each frame contains a lightweight spreadsheet editor backed by Apps Script properties and automatically generates a Chart.js visualization (line, bar, or pie) from the entered numbers.
- **Dashboard visibility controls** – Frames remember whether their dashboard is hidden, the preferred chart type, and sheet density. Built-in controls adjust the number of rows/columns and the rendered layout without reloading the page.
- **Driver engagement tools** – A "Driver of the Week" callout, driver spotlight image gallery, and uploader keep the team section fresh. Images are stored in Google Drive and reloaded on subsequent visits.
- **Daily inspiration** – The app requests a quote from ZenQuotes, caches it for 24 hours via script properties, and falls back to curated quotes when the API is unavailable.
- **Smooth loading experience** – A progress overlay communicates background work while dashboards initialize, reflecting recent performance-focused commits.

## Project layout
- `Code.gs` – Google Apps Script backend that serves the UI, fetches weather/quote data, manages Script Properties, and stores driver imagery in Drive.
- `index.html` – Standalone HTML, CSS, and JavaScript for the front-end. It loads Chart.js from a CDN, renders the interface, synchronizes frame data with Apps Script, and contains the spreadsheet + chart logic.
- `weather.js` – Node-friendly copy of `weatherCodeToText` used in automated tests.
- `tests/weather.test.js` – Jest suite validating the weather code mapping.

## Working with the Apps Script project
1. Create or open a Google Apps Script project and replace the default files with the contents of `Code.gs` and `index.html`.
2. Deploy the project as a Web App (Execute as "Me" and allow access to "Anyone with the link" for kiosk displays).
3. The app stores settings (frames, header title, quote cache, etc.) in Script Properties. Images uploaded through the interface are saved to Google Drive; ensure the deployment account has permission to create and manage files.
4. To reset frames or spotlight images, clear the corresponding Script Properties or use the provided UI controls.

## Local development & testing
- Run `npm install` once to install Jest.
- Execute `npm test` to validate the weather mapping helper before shipping Apps Script changes.
- The front-end code lives entirely in `index.html`; consider using a dedicated editor for large HTML/JS files and copy the result back into Apps Script.

## Contributing
Focus on maintainable, accessible UI changes and keep Script Properties migrations backward compatible. When adding server utilities, mirror the logic in the Node helpers if you need automated coverage. Update this README when the dashboard surface or workflows change so downstream deployments remain in sync.
