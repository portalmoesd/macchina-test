/**
 * Statistics app server.
 *
 * Serves the static frontend and the statistics API, which proxies the
 * Geostat trade API and parses bundled GNTA (tourism) / Geostat (FDI)
 * workbooks. No database, no authentication — just plain statistics.
 */
const express = require('express');
const cors = require('cors');
const path = require('path');
const config = require('./config');

const app = express();

app.use(cors());
app.use(express.json({ limit: '10mb' }));

// Serve frontend static files
app.use(express.static(path.join(__dirname, '../frontend')));

// Statistics API
app.use('/api/statistics', require('./routes/statistics'));

// Root → statistics page
app.get('/', (req, res) => res.redirect('/pages/statistics.html'));

app.listen(config.port, () => {
  console.log(`Statistics app listening on http://localhost:${config.port}`);
});
