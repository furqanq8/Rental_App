const fs = require('fs');
const path = require('path');
const fsp = fs.promises;

const defaultState = require('./defaultState.json');

const databasePath = process.env.DATABASE_PATH
  ? path.resolve(process.env.DATABASE_PATH)
  : path.join(__dirname, '..', 'data', 'azmat-state.json');

fs.mkdirSync(path.dirname(databasePath), { recursive: true });

function normalizeSnapshot(raw) {
  const base = JSON.parse(JSON.stringify(defaultState));
  if (!raw || typeof raw !== 'object') {
    return base;
  }

  base.fleet = Array.isArray(raw.fleet) ? raw.fleet.map(record => ({ ...record })) : [];
  base.drivers = Array.isArray(raw.drivers) ? raw.drivers.map(record => ({ ...record })) : [];
  base.trips = Array.isArray(raw.trips) ? raw.trips.map(record => ({ ...record })) : [];
  base.invoices = Array.isArray(raw.invoices) ? raw.invoices.map(record => ({ ...record })) : [];
  base.suppliers = Array.isArray(raw.suppliers) ? raw.suppliers.map(record => ({ ...record })) : [];
  base.supplierDirectory = Array.isArray(raw.supplierDirectory)
    ? raw.supplierDirectory.map(record => ({ ...record }))
    : [];
  base.customers = Array.isArray(raw.customers)
    ? [...new Set(raw.customers.filter(Boolean))].sort((a, b) => a.localeCompare(b))
    : [];

  if (raw.dashboardFilter && typeof raw.dashboardFilter === 'object') {
    base.dashboardFilter = {
      mode: raw.dashboardFilter.mode || defaultState.dashboardFilter.mode,
      start: raw.dashboardFilter.start || '',
      end: raw.dashboardFilter.end || ''
    };
  }

  if (typeof raw.nextTripNumber === 'number' && raw.nextTripNumber > 0) {
    base.nextTripNumber = raw.nextTripNumber;
  }

  return base;
}

async function initialize() {
  try {
    await fsp.access(databasePath, fs.constants.F_OK);
  } catch (error) {
    await fsp.writeFile(databasePath, JSON.stringify(defaultState, null, 2), 'utf8');
  }
}

async function getState() {
  try {
    const serialized = await fsp.readFile(databasePath, 'utf8');
    const parsed = JSON.parse(serialized);
    return normalizeSnapshot(parsed);
  } catch (error) {
    console.warn('Unable to read persisted state. Resetting to defaults.', error);
    await fsp.writeFile(databasePath, JSON.stringify(defaultState, null, 2), 'utf8');
    return JSON.parse(JSON.stringify(defaultState));
  }
}

async function saveState(payload) {
  const normalized = normalizeSnapshot(payload);
  const serialized = JSON.stringify(normalized, null, 2);
  await fsp.writeFile(databasePath, serialized, 'utf8');
  return normalized;
}

module.exports = {
  initialize,
  getState,
  saveState
};
