/**
 * Tiny file-backed key/value store.
 *
 * The original Vector Portal persisted two things in Postgres: a cache of
 * Geostat trade-data responses (`trade_cache`) and admin-uploaded datasets
 * (`admin_uploads`). This standalone statistics app has no database — these
 * helpers keep the same behaviour using small JSON files under
 * `server/data/store/`. All failures degrade gracefully so the app keeps
 * working from the live Geostat API / bundled workbooks even if the disk is
 * read-only (e.g. ephemeral hosting).
 */
const fs = require('fs');
const path = require('path');

const DIR = path.join(__dirname, 'data', 'store');
try { fs.mkdirSync(DIR, { recursive: true }); } catch (_) { /* read-only fs */ }

const TRADE_CACHE_FILE = path.join(DIR, 'trade-cache.json');

function readJson(file, fallback) {
  try { return JSON.parse(fs.readFileSync(file, 'utf8')); } catch (_) { return fallback; }
}
function writeJson(file, obj) {
  try { fs.writeFileSync(file, JSON.stringify(obj)); }
  catch (err) { console.warn(`store write(${path.basename(file)}) failed:`, err.message); }
}

module.exports = {
  // ── Trade-data cache (was the `trade_cache` table) ──────────────────────
  tradeCacheGet(key) {
    const all = readJson(TRADE_CACHE_FILE, {});
    return all[key] != null ? all[key] : null;
  },
  tradeCacheSet(key, data) {
    const all = readJson(TRADE_CACHE_FILE, {});
    all[key] = data;
    writeJson(TRADE_CACHE_FILE, all);
  },
  tradeCacheDeleteByPrefix(prefix) {
    const all = readJson(TRADE_CACHE_FILE, {});
    let changed = false;
    for (const k of Object.keys(all)) {
      if (k.startsWith(prefix)) { delete all[k]; changed = true; }
    }
    if (changed) writeJson(TRADE_CACHE_FILE, all);
  },

  // ── Uploaded datasets (was the `admin_uploads` table) ───────────────────
  uploadGet(kind) {
    return readJson(path.join(DIR, `upload-${kind}.json`), null);
  },
  uploadSet(kind, parsed) {
    writeJson(path.join(DIR, `upload-${kind}.json`), parsed);
  },
};
