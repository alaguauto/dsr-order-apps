# DSR Order Slip Apps — Alagu Auto Agency

Static order-taking apps used by each DSR (Delivery Sales Representative) to record
orders on their phone. All apps share one Google Apps Script backend + one Google
Sheet (scoped per DSR by name), so an order placed in any app shows up in the same
place.

Hosted on GitHub Pages, live at:
**https://alaguauto.github.io/dsr-order-apps/**

That page links out to each DSR's own app:

- https://alaguauto.github.io/dsr-order-apps/prabha/
- https://alaguauto.github.io/dsr-order-apps/arun/
- https://alaguauto.github.io/dsr-order-apps/nagaraj/
- https://alaguauto.github.io/dsr-order-apps/chellamani/
- https://alaguauto.github.io/dsr-order-apps/aaadirect/
- https://alaguauto.github.io/dsr-order-apps/anandh/
- https://alaguauto.github.io/dsr-order-apps/madhu/
- https://alaguauto.github.io/dsr-order-apps/saravanan/

## Layout

- `<name>/index.html` — one folder per DSR; each is a self-contained order-slip app.
- `order-slip-template.html` — the master template every DSR app is generated from.
  When making a change that should apply to every DSR, edit this file first, then
  re-apply the same change to each `<name>/index.html`.
- `backend/Code.gs` — the shared Google Apps Script Web App (paste into
  Extensions > Apps Script in the Google Sheet, Deploy > Web app). All DSR apps
  point at the same deployed `/exec` URL via `ORDERS_API_URL` in their HTML.
- `index.html` — a simple landing page linking to all 8 DSR apps.

## Deploying a change

GitHub Pages auto-publishes whatever is on the `main` branch, so once a change is
pushed/uploaded to this repo it goes live within about a minute — no separate
deploy step needed.
