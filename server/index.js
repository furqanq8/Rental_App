const path = require('path');
const express = require('express');
const morgan = require('morgan');
const cors = require('cors');

const storage = require('./storage');

const app = express();
const port = process.env.PORT || 3000;

const corsOrigins = (process.env.ALLOWED_ORIGINS || '')
  .split(',')
  .map(origin => origin.trim())
  .filter(Boolean);

if (corsOrigins.length) {
  app.use(cors({
    origin(origin, callback) {
      if (!origin) {
        return callback(null, true);
      }
      if (corsOrigins.includes(origin)) {
        return callback(null, true);
      }
      return callback(new Error('Not allowed by CORS'));
    },
    credentials: true
  }));
}

app.use(morgan('tiny'));
app.use(express.json({ limit: '5mb' }));

app.get('/api/health', (req, res) => {
  res.json({ status: 'ok' });
});

app.get('/api/state', async (req, res, next) => {
  try {
    const snapshot = await storage.getState();
    res.json(snapshot);
  } catch (error) {
    next(error);
  }
});

app.put('/api/state', async (req, res, next) => {
  const payload = req.body;
  if (!payload || typeof payload !== 'object') {
    res.status(400).json({ error: 'Invalid payload. Expected a JSON object.' });
    return;
  }
  try {
    const saved = await storage.saveState(payload);
    res.json(saved);
  } catch (error) {
    next(error);
  }
});

const publicDir = path.join(__dirname, '..');
app.use(express.static(publicDir, { extensions: ['html'] }));

app.get('*', (req, res) => {
  res.sendFile(path.join(publicDir, 'index.html'));
});

app.use((err, req, res, next) => { // eslint-disable-line no-unused-vars
  if (err && err.message === 'Not allowed by CORS') {
    console.warn(`Blocked CORS request from origin: ${req.headers.origin || 'unknown'}`);
    res.status(403).json({ error: 'Origin not allowed' });
    return;
  }
  console.error(err);
  res.status(500).json({ error: 'Internal Server Error' });
});

async function start() {
  await storage.initialize();
  app.listen(port, () => {
    console.log(`Azmat server listening on http://localhost:${port}`);
  });
}

start().catch(error => {
  console.error('Failed to start server:', error);
  process.exit(1);
});
