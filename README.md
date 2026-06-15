# Statistics

Standalone country statistics reports for Georgia — trade, tourism,
investments (FDI), companies, and a data appendix. The statistics page and
its functionality were ported from the Vector Portal project; everything
else (authentication, accounts, sidebar, the document-workflow app) has been
removed. This is just the statistics.

## Data sources

- **Trade** — proxied live from the Geostat trade API
  (`ex-trade-api.geostat.ge`), called from the browser with a server-side
  proxy fallback (`/api/statistics/...`).
- **Tourism** — GNTA international-visitor workbooks (bundled copies under
  `server/data/`, refreshed from `api.gnta.ge` when reachable).
- **Investments (FDI)** — Geostat FDI-by-country workbook (bundled copy
  under `server/data/`, refreshed from `geostat.ge` when reachable).
- **FDI sectors / companies** — optional datasets uploaded via
  `POST /api/statistics/fdi-sectors/upload` and
  `POST /api/statistics/companies/data`; persisted to a local file store
  (`server/data/store/`). These sections stay empty until data is uploaded.

No database and no accounts are required.

## Running locally

```bash
npm install
npm start          # serves on http://localhost:3000
```

Open <http://localhost:3000> — it redirects to the statistics page.

Set `PORT` to change the port (see `.env.example`).

## Reports

Pick a country, choose the report language (Georgian / English), and click
**Generate**. The report can be exported to **PDF** or **Word**.
