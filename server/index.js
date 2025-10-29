const path = require('path');
const express = require('express');
const morgan = require('morgan');
const cors = require('cors');
const crypto = require('crypto');

const storage = require('./storage');

const app = express();
const port = process.env.PORT || 3000;
const defaultSessionTtl = 1000 * 60 * 60 * 12;
const configuredSessionTtl = Number(process.env.SESSION_TTL_MS);
const sessionTtlMs = Number.isFinite(configuredSessionTtl) && configuredSessionTtl > 0
  ? configuredSessionTtl
  : defaultSessionTtl;

const expectedUsername = process.env.AZMAT_AUTH_USERNAME || 'admin';
const expectedPassword = process.env.AZMAT_AUTH_PASSWORD || 'admin';
const usingDefaultCredentials = !process.env.AZMAT_AUTH_USERNAME || !process.env.AZMAT_AUTH_PASSWORD;

if (usingDefaultCredentials) {
  console.warn('Authentication credentials not configured. Using default admin/admin credentials. Set AZMAT_AUTH_USERNAME and AZMAT_AUTH_PASSWORD for production.');
}

const sessions = new Map();

function createSession(username) {
  const token = crypto.randomBytes(32).toString('hex');
  sessions.set(token, { username, expiresAt: Date.now() + sessionTtlMs });
  return token;
}

function extractToken(req) {
  const header = req.headers.authorization || '';
  if (!header.toLowerCase().startsWith('bearer ')) {
    return null;
  }
  return header.slice(7).trim();
}

function getSession(token) {
  if (!token) {
    return null;
  }
  const entry = sessions.get(token);
  if (!entry) {
    return null;
  }
  if (entry.expiresAt && entry.expiresAt < Date.now()) {
    sessions.delete(token);
    return null;
  }
  return entry;
}

function authenticateRequest(req) {
  const token = extractToken(req);
  if (!token) {
    return null;
  }
  return getSession(token);
}

function requireAuth(req, res, next) {
  if (!expectedUsername || !expectedPassword) {
    return next();
  }
  const session = authenticateRequest(req);
  if (!session) {
    res.status(401).json({ error: 'Unauthorized' });
    return;
  }
  req.user = session.username;
  next();
}

setInterval(() => {
  const now = Date.now();
  sessions.forEach((value, key) => {
    if (value.expiresAt && value.expiresAt < now) {
      sessions.delete(key);
    }
  });
}, Math.max(60_000, Math.floor(sessionTtlMs / 4))).unref();

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
  res.json({ status: 'ok', auth: Boolean(expectedUsername && expectedPassword) ? 'enabled' : 'disabled' });
});

app.post('/api/login', (req, res) => {
  const { username, password } = req.body || {};
  if (!username || !password) {
    res.status(400).json({ error: 'Username and password are required.' });
    return;
  }
  if (username === expectedUsername && password === expectedPassword) {
    const token = createSession(username);
    res.json({ token, user: { username } });
    return;
  }
  res.status(401).json({ error: 'Invalid username or password.' });
});

app.post('/api/logout', requireAuth, (req, res) => {
  const token = extractToken(req);
  if (token) {
    sessions.delete(token);
  }
  res.json({ status: 'ok' });
});

app.get('/api/session', requireAuth, (req, res) => {
  res.json({ user: { username: req.user } });
});

app.get('/api/state', requireAuth, async (req, res, next) => {
  try {
    const snapshot = await storage.getState();
    res.json(snapshot);
  } catch (error) {
    next(error);
  }
});

app.put('/api/state', async (req, res, next) => {
app.put('/api/state', requireAuth, async (req, res, next) => {
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
