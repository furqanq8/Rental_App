const fs = require('fs');
const path = require('path');
const http = require('http');
const { URL } = require('url');
const crypto = require('crypto');

// Lightweight .env loader so we do not depend on external packages
const envPath = path.join(__dirname, '..', '.env');
if (fs.existsSync(envPath)) {
  const lines = fs.readFileSync(envPath, 'utf8').split(/\r?\n/);
  for (const line of lines) {
    if (!line || line.trim().startsWith('#')) {
      continue;
    }
    const idx = line.indexOf('=');
    if (idx === -1) {
      continue;
    }
    const key = line.slice(0, idx).trim();
    if (!key || Object.prototype.hasOwnProperty.call(process.env, key)) {
      continue;
    }
    const value = line.slice(idx + 1).trim();
    process.env[key] = value;
  }
}

const storage = require('./storage');

const port = Number(process.env.PORT) || 3000;
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
const allowedOrigins = (process.env.ALLOWED_ORIGINS || '')
  .split(',')
  .map(origin => origin.trim())
  .filter(Boolean);

const publicDir = path.join(__dirname, '..');

const mimeTypes = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon',
  '.woff': 'font/woff',
  '.woff2': 'font/woff2'
};

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

function requireAuth(req) {
  if (!expectedUsername || !expectedPassword) {
    return { ok: true, user: null };
  }
  const session = authenticateRequest(req);
  if (!session) {
    return { ok: false, status: 401, payload: { error: 'Unauthorized' } };
  }
  return { ok: true, user: session.username };
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

function isOriginAllowed(origin) {
  if (!allowedOrigins.length) {
    return true;
  }
  if (!origin) {
    return true;
  }
  return allowedOrigins.includes(origin);
}

function applyCors(req, res) {
  const origin = req.headers.origin;
  if (!allowedOrigins.length) {
    return true;
  }
  if (!origin || !isOriginAllowed(origin)) {
    return false;
  }
  res.setHeader('Access-Control-Allow-Origin', origin);
  res.setHeader('Access-Control-Allow-Credentials', 'true');
  res.setHeader('Vary', 'Origin');
  return true;
}

function sendJson(res, statusCode, payload, extraHeaders = {}) {
  res.writeHead(statusCode, {
    'Content-Type': 'application/json; charset=utf-8',
    ...extraHeaders
  });
  res.end(JSON.stringify(payload));
}

function sendNoContent(res) {
  res.writeHead(204);
  res.end();
}

function sendText(res, statusCode, body, contentType = 'text/plain; charset=utf-8') {
  res.writeHead(statusCode, { 'Content-Type': contentType });
  res.end(body);
}

async function readRequestBody(req, limitBytes = 5 * 1024 * 1024) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let total = 0;
    req.on('data', chunk => {
      total += chunk.length;
      if (total > limitBytes) {
        reject(new Error('Payload too large'));
        req.destroy();
        return;
      }
      chunks.push(chunk);
    });
    req.on('end', () => {
      resolve(Buffer.concat(chunks).toString('utf8'));
    });
    req.on('error', reject);
  });
}

async function readJson(req) {
  const body = await readRequestBody(req);
  if (!body) {
    return {};
  }
  try {
    return JSON.parse(body);
  } catch (error) {
    throw new Error('Invalid JSON body');
  }
}

async function handleApi(req, res, parsedUrl) {
  if (!applyCors(req, res)) {
    sendJson(res, 403, { error: 'Origin not allowed' });
    return;
  }

  if (req.method === 'OPTIONS') {
    res.setHeader('Access-Control-Allow-Methods', 'GET,POST,PUT,OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
    sendNoContent(res);
    return;
  }

  if (req.method === 'GET' && parsedUrl.pathname === '/api/health') {
    sendJson(res, 200, { status: 'ok', auth: Boolean(expectedUsername && expectedPassword) ? 'enabled' : 'disabled' });
    return;
  }

  if (req.method === 'POST' && parsedUrl.pathname === '/api/login') {
    try {
      const { username, password } = await readJson(req);
      if (!username || !password) {
        sendJson(res, 400, { error: 'Username and password are required.' });
        return;
      }
      if (username === expectedUsername && password === expectedPassword) {
        const token = createSession(username);
        sendJson(res, 200, { token, user: { username } });
        return;
      }
      sendJson(res, 401, { error: 'Invalid username or password.' });
    } catch (error) {
      sendJson(res, 400, { error: error.message || 'Invalid request.' });
    }
    return;
  }

  if (req.method === 'POST' && parsedUrl.pathname === '/api/logout') {
    const authResult = requireAuth(req);
    if (!authResult.ok) {
      sendJson(res, authResult.status, authResult.payload);
      return;
    }
    const token = extractToken(req);
    if (token) {
      sessions.delete(token);
    }
    sendJson(res, 200, { status: 'ok' });
    return;
  }

  if (req.method === 'GET' && parsedUrl.pathname === '/api/session') {
    const authResult = requireAuth(req);
    if (!authResult.ok) {
      sendJson(res, authResult.status, authResult.payload);
      return;
    }
    sendJson(res, 200, { user: { username: authResult.user || expectedUsername } });
    return;
  }

  if (req.method === 'GET' && parsedUrl.pathname === '/api/state') {
    const authResult = requireAuth(req);
    if (!authResult.ok) {
      sendJson(res, authResult.status, authResult.payload);
      return;
    }
    try {
      const snapshot = await storage.getState();
      sendJson(res, 200, snapshot);
    } catch (error) {
      console.error('Failed to read persisted state', error);
      sendJson(res, 500, { error: 'Internal Server Error' });
    }
    return;
  }

  if (req.method === 'PUT' && parsedUrl.pathname === '/api/state') {
    const authResult = requireAuth(req);
    if (!authResult.ok) {
      sendJson(res, authResult.status, authResult.payload);
      return;
    }
    try {
      const payload = await readJson(req);
      if (!payload || typeof payload !== 'object') {
        sendJson(res, 400, { error: 'Invalid payload. Expected a JSON object.' });
        return;
      }
      const saved = await storage.saveState(payload);
      sendJson(res, 200, saved);
    } catch (error) {
      const status = error.message === 'Payload too large' ? 413 : 500;
      console.error('Failed to persist state', error);
      sendJson(res, status, { error: status === 413 ? 'Payload too large' : 'Internal Server Error' });
    }
    return;
  }

  sendJson(res, 404, { error: 'Not Found' });
}

async function serveStatic(req, res, parsedUrl) {
  let pathname = decodeURIComponent(parsedUrl.pathname);
  if (pathname.endsWith('/')) {
    pathname += 'index.html';
  }
  if (pathname === '/api' || pathname.startsWith('/api/')) {
    sendJson(res, 404, { error: 'Not Found' });
    return;
  }

  const normalizedPath = path.normalize(pathname).replace(/^\.\/?/, '');
  let filePath = path.join(publicDir, normalizedPath);

  if (!filePath.startsWith(publicDir)) {
    sendJson(res, 403, { error: 'Forbidden' });
    return;
  }

  try {
    let stats = await fs.promises.stat(filePath);
    if (stats.isDirectory()) {
      filePath = path.join(filePath, 'index.html');
      stats = await fs.promises.stat(filePath);
    }
    const stream = fs.createReadStream(filePath);
    const ext = path.extname(filePath).toLowerCase();
    const headers = {
      'Content-Type': mimeTypes[ext] || 'application/octet-stream',
      'Cache-Control': 'no-cache'
    };
    res.writeHead(200, headers);
    stream.pipe(res);
    stream.on('error', error => {
      console.error('Static file stream error', error);
      if (!res.headersSent) {
        sendJson(res, 500, { error: 'Internal Server Error' });
      } else {
        res.destroy(error);
      }
    });
  } catch (error) {
    try {
      const fallback = path.join(publicDir, 'index.html');
      const html = await fs.promises.readFile(fallback);
      res.writeHead(200, { 'Content-Type': mimeTypes['.html'], 'Cache-Control': 'no-cache' });
      res.end(html);
    } catch (fallbackError) {
      console.error('Failed to serve static asset', fallbackError);
      sendText(res, 404, 'Not Found');
    }
  }
}

async function requestListener(req, res) {
  try {
    const parsedUrl = new URL(req.url, `http://${req.headers.host}`);
    if (parsedUrl.pathname.startsWith('/api/')) {
      await handleApi(req, res, parsedUrl);
      return;
    }
    if (parsedUrl.pathname === '/api') {
      sendJson(res, 404, { error: 'Not Found' });
      return;
    }
    await serveStatic(req, res, parsedUrl);
  } catch (error) {
    console.error('Unhandled server error', error);
    if (!res.headersSent) {
      sendJson(res, 500, { error: 'Internal Server Error' });
    } else {
      res.destroy(error);
    }
  }
}

async function start() {
  await storage.initialize();
  const server = http.createServer(requestListener);
  server.listen(port, () => {
    console.log(`Azmat server listening on http://localhost:${port}`);
  });
  server.on('clientError', (err, socket) => {
    socket.end('HTTP/1.1 400 Bad Request\r\n\r\n');
  });
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
