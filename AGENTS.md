# AGENTS

## Project snapshot
- RouteDash is a Google Apps Script web app with a kiosk-style dashboard for delivery operations.
- `index.html` contains the entire front-end (styling, drag/resizable frames, spreadsheet-to-chart pipeline, loading overlay, driver spotlight, header editor, weather panel, quote widget, etc.).
- `Code.gs` hosts server utilities for rendering the HTML, fetching weather/quote data, persisting Script Properties, and storing uploaded images in Drive. `weather.js` mirrors the weather helper for Node tests.

## Development guidelines
- Keep UI additions accessible (semantic headings, readable contrast, keyboard focus states) and responsive across large displays.
- Preserve backward compatibility for Script Properties and Drive assets. Provide migrations or defensive guards when changing stored data formats.
- Mirror critical server-side utilities in the Node helpers when automated coverage is helpful, and extend the Jest suite as needed.
- Run `npm test` before submitting changes.

## Apps Script tips
- After editing locally, copy `Code.gs` and `index.html` into an Apps Script project and deploy as a Web App.
- Ensure the executing account has Drive access for driver image uploads and network access for Open-Meteo/ZenQuotes requests.
- When introducing new long-running operations, update the loading overlay messaging in `index.html` so kiosk users get feedback while tasks complete.
