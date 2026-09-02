# DSR Order Slip Apps — Alagu Auto Agency

Static order-taking apps used by each DSR (Delivery Sales Representative) to record
orders on their phone. All apps share one Google Apps Script backend + one Google
Sheet (scoped per DSR by name), so an order placed in any app shows up in the same
place.

## Layout

- `dsr-apps/<name>/index.html` — one folder per DSR, each deployable as its own
  Netlify site (site's "Base directory" = that folder). Currently live:
  - `dsr-apps/prabha` → https://prabha-order.netlify.app
  - `dsr-apps/arun` → https://arun-order.netlify.app
  - The rest (`nagaraj`, `chellamani`, `aaadirect`, `anandh`, `madhu`, `saravanan`)
    are built and ready to deploy whenever a Netlify site is created for them.
- `order-slip-template.html` — the master template every DSR app is generated from.
  When making a change that should apply to every DSR, edit this file first, then
  re-apply the same change to each `dsr-apps/<name>/index.html`.
- `backend/Code.gs` — the shared Google Apps Script Web App (paste into
  Extensions > Apps Script in the Google Sheet, Deploy > Web app). All DSR apps
  point at the same deployed `/exec` URL via `ORDERS_API_URL` in their HTML.

## Deploying a change

1. Edit `order-slip-template.html` and the relevant `dsr-apps/<name>/index.html`
   file(s).
2. Commit and push.
3. If a DSR's Netlify site is linked to this repo (see below), it redeploys
   automatically. Otherwise, drag-and-drop that folder onto Netlify manually.

## Connecting Netlify to this repo (one-time, per DSR site)

In each Netlify site's dashboard: **Site configuration → Build & deploy → Link repository**,
choose this repo, and set **Base directory** to `dsr-apps/<name>` and leave the build
command empty (publish directory `.` relative to the base directory) — it's a static
HTML file, nothing to build. From then on, every push to `main` auto-deploys that
site.
