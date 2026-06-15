/**
 * Shared helpers for dataset (XLSX/CSV) uploads.
 *
 * Uploads are persisted via the file-backed store (server/store.js) so they
 * survive restarts where the disk is writable. This standalone statistics
 * app has no accounts, so uploads are unauthenticated.
 *
 * Intended usage (from a specific route file):
 *
 *   const { upload, adminOnly, saveParsedAndRaw, loadParsed } =
 *     require('./admin-uploads');
 *
 *   router.post('/my-kind/upload', ...adminOnly, upload.single('file'),
 *     async (req, res) => {
 *       const wb = XLSX.read(req.file.buffer, { type: 'buffer' });
 *       const parsed = parseMyKind(wb);
 *       parsed.uploadedAt = new Date().toISOString();
 *       await saveParsedAndRaw('my-kind', parsed, req.file.buffer);
 *       // ...update in-memory cache, return response...
 *     });
 */

const multer = require('multer');
const store = require('../store');

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024 },
});

// No authentication in the standalone statistics app. Spread into route
// definitions for compatibility: router.post('/x', ...adminOnly, handler)
const adminOnly = [];

async function saveParsedAndRaw(kind, parsed /*, buffer */) {
  // Raw bytes are never read back, so only the parsed JSON is persisted.
  store.uploadSet(kind, parsed);
}

async function loadParsed(kind) {
  try {
    return store.uploadGet(kind);
  } catch (err) {
    console.error(`uploads: loadParsed(${kind}) failed:`, err.message);
    return null;
  }
}

module.exports = {
  upload,
  adminOnly,
  saveParsedAndRaw,
  loadParsed,
};
