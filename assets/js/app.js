const authStorageKey = 'azmat-auth-token';
const themeStorageKey = 'azmat-theme-preference';
let authToken = '';
let authUser = null;

const bodyElement = document.body;

try {
  const storedToken = localStorage.getItem(authStorageKey);
  if (storedToken) {
    authToken = storedToken;
  }
} catch (error) {
  console.warn('Unable to restore authentication token from storage.', error);
}

try {
  const storedTheme = localStorage.getItem(themeStorageKey);
  if (storedTheme === 'dark' && bodyElement) {
    bodyElement.classList.add('dark-theme');
  }
} catch (error) {
  console.warn('Unable to restore theme preference from storage.', error);
}

if (!authToken && bodyElement) {
  bodyElement.classList.add('app-locked');
}

const storageKey = 'azmat-fleet-state-v2';
const apiBase = (window.__AZMAT_API_BASE__ || document.documentElement.getAttribute('data-api-base') || '').replace(/\/$/, '');
const storageKey = 'azmat-fleet-state-v3';
const defaultState = {
  fleet: [],
  drivers: [],
  trips: [],
  invoices: [],
  suppliers: [],
  supplierDirectory: [],
  customers: [],
  dashboardFilter: {
    mode: 'thisMonth',
    start: '',
    end: ''
  },
  nextTripNumber: 1
};

const state = loadState();
let stateHydratedFromServer = false;
const persistQueue = { timer: null };
const selectRefs = {};
let tripIdInput;
let tripOwnershipField;
let fleetOwnershipSelect;
let driverAffiliationSelect;
let bootstrapPromise = null;
let authSubmitDefaultLabel = 'Sign In';
const editingContext = { type: null, index: -1 };
const modalTypeMap = {
  fleetModal: 'fleet',
  driverModal: 'driver',
  tripModal: 'trip',
  invoiceModal: 'invoice',
  supplierModal: 'supplierPayment',
  supplierDirectoryModal: 'supplierDirectory'
};
const modalEditTitles = {
  fleetModal: 'Edit Fleet Unit',
  driverModal: 'Edit Driver',
  tripModal: 'Edit Trip',
  invoiceModal: 'Edit Invoice',
  supplierModal: 'Edit Supplier Payment',
  supplierDirectoryModal: 'Edit Supplier'
};
const modalEditSubmitLabels = {
  fleetModal: 'Update Unit',
  driverModal: 'Update Driver',
  tripModal: 'Update Trip',
  invoiceModal: 'Update Invoice',
  supplierModal: 'Update Payment',
  supplierDirectoryModal: 'Update Supplier'
};

const syncBannerElement = document.getElementById('syncStatusBanner');
const syncBannerMessage = document.getElementById('syncStatusMessage');
let hideSyncBannerTimer = null;
let currentSyncStatus = 'idle';

function describeSyncError(error) {
  if (!error) {
    return '';
  }
  if (typeof error === 'string') {
    return error;
  }
  if (error.message) {
    return error.message;
  }
  try {
    return JSON.stringify(error);
  } catch (err) {
    return '';
  }
}

function setSyncStatus(status, error) {
  const previousStatus = currentSyncStatus;
  currentSyncStatus = status;
  if (!syncBannerElement || !syncBannerMessage) {
    return;
  }

  if (hideSyncBannerTimer) {
    window.clearTimeout(hideSyncBannerTimer);
    hideSyncBannerTimer = null;
  }

  syncBannerElement.classList.remove('status-connecting', 'status-connected', 'status-offline', 'hidden');

  let message = '';
  if (status === 'connecting') {
    message = 'Connecting to the Azmat data server…';
  } else if (status === 'connected') {
    message = 'Data is synced to the Azmat server. You can open the app on any device and see the same records.';
  } else {
    const detail = describeSyncError(error);
    let hint = ' Start the server with "npm start" on your host or set window.__AZMAT_API_BASE__ to your deployed API domain.';
    if (detail && /auth|sign\s?in|login/i.test(detail)) {
      hint = ' Sign in with your Azmat credentials to resume syncing.';
    }
    message = 'Unable to reach the Azmat data server. Changes will stay in this browser until the connection is restored.';
    if (detail) {
      message += ` (${detail})`;
    }
    message += hint;
  }

  syncBannerMessage.textContent = message;
  syncBannerElement.classList.add(`status-${status}`);

  if (status === 'connected') {
    hideSyncBannerTimer = window.setTimeout(() => {
      syncBannerElement.classList.add('hidden');
      hideSyncBannerTimer = null;
    }, 4000);
  }

  if (previousStatus === status && status === 'offline') {
    syncBannerElement.classList.remove('hidden');
  }
}

function setAuthToken(token, user) {
  authToken = token || '';
  if (user && typeof user === 'object') {
    authUser = { ...user };
  }
  try {
    if (authToken) {
      localStorage.setItem(authStorageKey, authToken);
    } else {
      localStorage.removeItem(authStorageKey);
    }
  } catch (error) {
    console.warn('Unable to persist authentication token.', error);
  }
  if (bodyElement) {
    if (authToken) {
      bodyElement.classList.remove('app-locked');
    } else {
      bodyElement.classList.add('app-locked');
    }
  }
  if (!authToken) {
    stateHydratedFromServer = false;
  }
}

function getAuthHeaders() {
  if (!authToken) {
    return {};
  }
  return { Authorization: `Bearer ${authToken}` };
}

function normalizeText(value) {
  return value ? String(value).trim().toLowerCase() : '';
}

function normalizeUnitId(value) {
  return normalizeText(value);
}

function findFleetUnit(vehicleId, fleetList = state.fleet) {
  if (!vehicleId) return null;
  const collection = Array.isArray(fleetList) ? fleetList : [];
  const normalized = normalizeUnitId(vehicleId);
  return collection.find(item => item && normalizeUnitId(item.unitId) === normalized) || null;
}

function findVehicleOwnership(vehicleId, fleetList = state.fleet) {
  if (!vehicleId) return '';
  const match = findFleetUnit(vehicleId, fleetList);
  return match ? match.ownership || '' : '';
}

function loadState() {
  try {
    const saved = localStorage.getItem(storageKey);
    if (saved) {
      const parsed = JSON.parse(saved);
      return normalizeState(parsed);
    }
  } catch (error) {
    console.warn('Unable to load saved data. Starting fresh.', error);
function resolveApiPath(path) {
  if (!path.startsWith('/')) {
    return `${apiBase}/${path}`;
  }
  if (!apiBase) {
    return path;
  }
  return `${apiBase}${path}`;
}

function cloneRecordArray(source) {
  if (!Array.isArray(source)) {
    return [];
  }
  return source.map(item => {
    if (!item || typeof item !== 'object') {
      return item;
    }
    return { ...item };
  });
}

function getSerializableState() {
  const snapshot = normalizeState(state);
  return {
    fleet: cloneRecordArray(snapshot.fleet),
    drivers: cloneRecordArray(snapshot.drivers),
    trips: cloneRecordArray(snapshot.trips),
    invoices: cloneRecordArray(snapshot.invoices),
    suppliers: cloneRecordArray(snapshot.suppliers),
    supplierDirectory: cloneRecordArray(snapshot.supplierDirectory),
    customers: Array.isArray(snapshot.customers) ? [...snapshot.customers] : [],
    dashboardFilter: snapshot.dashboardFilter ? { ...snapshot.dashboardFilter } : { ...defaultState.dashboardFilter },
    nextTripNumber: typeof snapshot.nextTripNumber === 'number' && snapshot.nextTripNumber > 0
      ? snapshot.nextTripNumber
      : defaultState.nextTripNumber
  };
}

function loadLocalSnapshot() {
  try {
    const saved = localStorage.getItem(storageKey);
    if (saved) {
      return JSON.parse(saved);
    }
  } catch (error) {
    console.warn('Unable to read local snapshot.', error);
  }
  return null;
}

function saveLocalSnapshot(snapshot) {
  try {
    localStorage.setItem(storageKey, JSON.stringify(snapshot));
  } catch (error) {
    console.warn('Unable to persist snapshot locally.', error);
  }
}

function applyStateSnapshot(snapshot) {
  const normalized = normalizeState(snapshot);
  state.fleet = cloneRecordArray(normalized.fleet);
  state.drivers = cloneRecordArray(normalized.drivers);
  state.trips = cloneRecordArray(normalized.trips);
  state.invoices = cloneRecordArray(normalized.invoices);
  state.suppliers = cloneRecordArray(normalized.suppliers);
  state.supplierDirectory = cloneRecordArray(normalized.supplierDirectory);
  state.customers = Array.isArray(normalized.customers) ? [...normalized.customers] : [];
  state.dashboardFilter = normalized.dashboardFilter
    ? { ...normalized.dashboardFilter }
    : { ...defaultState.dashboardFilter };
  state.nextTripNumber = typeof normalized.nextTripNumber === 'number' && normalized.nextTripNumber > 0
    ? normalized.nextTripNumber
    : defaultState.nextTripNumber;
  return getSerializableState();
}

async function fetchRemoteState() {
  const response = await fetch(resolveApiPath('/api/state'), {
    method: 'GET',
    headers: { Accept: 'application/json', ...getAuthHeaders() },
    credentials: 'include'
  });
  if (response.status === 401) {
    handleUnauthorized('Authentication required. Please sign in to continue.');
    throw new Error('Authentication required');
  }
    headers: { Accept: 'application/json' },
    credentials: 'include'
  });
  if (!response.ok) {
    throw new Error(`Failed to fetch remote state (${response.status})`);
  }
  return response.json();
}

async function pushRemoteState(snapshot) {
  if (!authToken) {
    throw new Error('Authentication required. Please sign in.');
  }
  const payload = snapshot || getSerializableState();
  if (currentSyncStatus !== 'offline') {
    setSyncStatus('connecting');
  }
  try {
    const response = await fetch(resolveApiPath('/api/state'), {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json', Accept: 'application/json', ...getAuthHeaders() },
      credentials: 'include',
      body: JSON.stringify(payload)
    });
    if (response.status === 401) {
      handleUnauthorized('Authentication required. Please sign in to continue.');
      throw new Error('Authentication required');
    }
    if (!response.ok) {
      throw new Error(`Failed to persist remote state (${response.status})`);
    }
    const data = await response.json();
    setSyncStatus('connected');
    return data;
  } catch (error) {
    setSyncStatus('offline', error);
    throw error;
  }
}

function queueRemotePersist() {
  if (!stateHydratedFromServer || !authToken) {
  const payload = snapshot || getSerializableState();
  const response = await fetch(resolveApiPath('/api/state'), {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json', Accept: 'application/json' },
    credentials: 'include',
    body: JSON.stringify(payload)
  });
  if (!response.ok) {
    throw new Error(`Failed to persist remote state (${response.status})`);
  }
  return response.json();
}

function queueRemotePersist() {
  if (!stateHydratedFromServer) {
    return;
  }
  if (persistQueue.timer) {
    clearTimeout(persistQueue.timer);
  }
  persistQueue.timer = setTimeout(() => {
    persistQueue.timer = null;
    if (!authToken) {
      return;
    }
    pushRemoteState().catch(error => {
      console.warn('Unable to sync data with the server. Changes will remain saved locally until the next successful sync.', error);
    });
  }, 300);
}

async function bootstrapState() {
  const localSnapshot = loadLocalSnapshot();
  if (localSnapshot) {
    applyStateSnapshot(localSnapshot);
  }
  if (!authToken) {
    return;
  }
  setSyncStatus('connecting');
  let remoteLoaded = false;
  try {
    const remoteSnapshot = await fetchRemoteState();
    if (remoteSnapshot) {
      const normalized = applyStateSnapshot(remoteSnapshot);
      saveLocalSnapshot(normalized);
    }
    remoteLoaded = true;
    setSyncStatus('connected');
  } catch (error) {
    console.warn('Falling back to locally cached data because the server state could not be loaded.', error);
    setSyncStatus('offline', error);
  } finally {
    if (!authToken) {
      stateHydratedFromServer = false;
      return;
    }
      remoteLoaded = true;
    }
  } catch (error) {
    console.warn('Falling back to locally cached data because the server state could not be loaded.', error);
  } finally {
    stateHydratedFromServer = true;
    if (!remoteLoaded) {
      queueRemotePersist();
    }
  }
}

function loadState() {
  const snapshot = loadLocalSnapshot();
  if (snapshot) {
    return normalizeState(snapshot);
  }
  return JSON.parse(JSON.stringify(defaultState));
}

function normalizeState(parsed) {
  const base = JSON.parse(JSON.stringify(defaultState));
  if (parsed && typeof parsed === 'object') {
    Object.assign(base, parsed);
  }
  if (!base.dashboardFilter || typeof base.dashboardFilter !== 'object') {
    base.dashboardFilter = { ...defaultState.dashboardFilter };
  } else {
    base.dashboardFilter = {
      mode: base.dashboardFilter.mode || defaultState.dashboardFilter.mode,
      start: base.dashboardFilter.start || '',
      end: base.dashboardFilter.end || ''
    };
  }
  if (!Array.isArray(base.supplierDirectory)) {
    base.supplierDirectory = [];
  }
  if (!Array.isArray(base.customers)) {
    base.customers = [];
  }
  base.customers = Array.from(new Set(base.customers.filter(Boolean)));
  if (typeof base.nextTripNumber !== 'number' || base.nextTripNumber < 1) {
    base.nextTripNumber = 1;
  }
  return hydrateDerivedCollections(base);
}

function hydrateDerivedCollections(currentState) {
  if (!Array.isArray(currentState.fleet)) {
    currentState.fleet = [];
  } else {
    const seenUnitIds = new Set();
    currentState.fleet = currentState.fleet.filter(item => {
      if (!item) return false;
      const trimmedId = item.unitId ? String(item.unitId).trim() : '';
      if (!trimmedId) return false;
      item.unitId = trimmedId;
      if (item.supplier) {
        item.supplier = String(item.supplier).trim();
      }
      const normalized = normalizeUnitId(trimmedId);
      if (seenUnitIds.has(normalized)) {
        return false;
      }
      seenUnitIds.add(normalized);
      return true;
    });
  }
  if (!Array.isArray(currentState.drivers)) {
    currentState.drivers = [];
  } else {
    currentState.drivers.forEach(driver => {
      if (driver && driver.supplier) {
        driver.supplier = String(driver.supplier).trim();
      }
    });
  }
  const customers = new Set(currentState.customers || []);
  (currentState.trips || []).forEach(trip => {
    if (trip && trip.customer) customers.add(trip.customer);
  });
  (currentState.invoices || []).forEach(invoice => {
    if (invoice && invoice.customer) customers.add(invoice.customer);
  });
  currentState.customers = Array.from(customers).sort((a, b) => a.localeCompare(b));
  (currentState.trips || []).forEach(trip => {
    if (trip && !trip.unitOwnership) {
      trip.unitOwnership = findVehicleOwnership(trip.vehicle, currentState.fleet);
    }
  });
  const highestTripNumber = (currentState.trips || [])
    .map(trip => {
      const match = String(trip.tripId || '').match(/(\d+)/);
      return match ? Number(match[1]) : NaN;
    })
    .filter(num => Number.isFinite(num))
    .reduce((max, num) => Math.max(max, num), 0);
  if (highestTripNumber >= currentState.nextTripNumber) {
    currentState.nextTripNumber = highestTripNumber + 1;
  }
  return currentState;
}

function persistState() {
  try {
    localStorage.setItem(storageKey, JSON.stringify(state));
  } catch (error) {
    console.warn('Unable to persist data to localStorage.', error);
  }
    saveLocalSnapshot(getSerializableState());
  } catch (error) {
    console.warn('Unable to persist data locally.', error);
  }
  queueRemotePersist();
}

function parseDateValue(value) {
  if (!value) return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return new Date(value.getTime());
  }
  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function startOfDay(date) {
  if (!date) return null;
  const clone = new Date(date.getTime());
  clone.setHours(0, 0, 0, 0);
  return clone;
}

function endOfDay(date) {
  if (!date) return null;
  const clone = new Date(date.getTime());
  clone.setHours(23, 59, 59, 999);
  return clone;
}

function getDashboardDateRange() {
  const filter = state.dashboardFilter || defaultState.dashboardFilter;
  const mode = filter.mode || defaultState.dashboardFilter.mode;
  const now = new Date();
  let start = null;
  let end = null;

  if (mode === 'thisMonth') {
    const firstDay = new Date(now.getFullYear(), now.getMonth(), 1);
    const lastDay = new Date(now.getFullYear(), now.getMonth() + 1, 0);
    start = startOfDay(firstDay);
    end = endOfDay(lastDay);
  } else if (mode === 'lastMonth') {
    const firstDay = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    const lastDay = new Date(now.getFullYear(), now.getMonth(), 0);
    start = startOfDay(firstDay);
    end = endOfDay(lastDay);
  } else if (mode === 'custom') {
    const customStart = parseDateValue(filter.start);
    const customEnd = parseDateValue(filter.end);
    if (customStart) start = startOfDay(customStart);
    if (customEnd) end = endOfDay(customEnd);
  }

  if (start && end && start.getTime() > end.getTime()) {
    const temp = start;
    start = end;
    end = temp;
  }

  return { start, end, mode };
}

function filterRecordsByDateRange(collection, dateAccessor, range) {
  if (!Array.isArray(collection)) return [];
  const validItems = collection.filter(Boolean);
  if (!range || (!range.start && !range.end)) {
    return validItems;
  }
  const startTime = range.start ? range.start.getTime() : null;
  const endTime = range.end ? range.end.getTime() : null;
  return validItems.filter(item => {
    const rawDate = dateAccessor(item);
    const parsed = parseDateValue(rawDate);
    if (!parsed) return false;
    const time = parsed.getTime();
    if (startTime !== null && time < startTime) return false;
    if (endTime !== null && time > endTime) return false;
    return true;
  });
}

function normalizedDateString(value) {
  const parsed = parseDateValue(value);
  return parsed ? parsed.toISOString().slice(0, 10) : '';
}

function defaultDueDate(tripDate) {
  return normalizedDateString(tripDate) || new Date().toISOString().slice(0, 10);
}

function sanitizeIdentifier(value, fallback) {
  const base = String(value || '').trim().replace(/\s+/g, '-').replace(/[^a-zA-Z0-9-]/g, '');
  return base || fallback;
}

function generateInvoiceNumber(tripId) {
  const fallback = `TRIP-${String(state.nextTripNumber || 1).padStart(4, '0')}`;
  const base = sanitizeIdentifier(tripId, fallback);
  let candidate = `INV-${base}`;
  let suffix = 1;
  while (state.invoices.some(invoice => invoice && invoice.invoice === candidate)) {
    candidate = `INV-${base}-${suffix}`;
    suffix += 1;
  }
  return candidate;
}

function generateSupplierPaymentReference(tripId) {
  const fallback = `TRIP-${String(state.nextTripNumber || 1).padStart(4, '0')}`;
  const base = sanitizeIdentifier(tripId, fallback);
  let candidate = `SUP-${base}`;
  let suffix = 1;
  while (state.suppliers.some(payment => payment && payment.reference === candidate)) {
    candidate = `SUP-${base}-${suffix}`;
    suffix += 1;
  }
  return candidate;
}

function getVehicleSupplier(vehicleId) {
  if (!vehicleId) return '';
  const match = findFleetUnit(vehicleId);
  return match && match.supplier ? match.supplier : '';
}

function ensureInvoiceForTrip(trip) {
  if (!trip) return false;
  const existing = state.invoices.find(invoice => invoice && invoice.trip === trip.tripId);
  const amount = numericAmount(trip.rentalCharges);
  const dueDate = defaultDueDate(trip.date);
  const customer = trip.customer || '';

  if (existing) {
    let changed = false;
    if (!existing.invoice) {
      existing.invoice = generateInvoiceNumber(trip.tripId);
      changed = true;
    }
    if (!existing.trip && trip.tripId) {
      existing.trip = trip.tripId;
      changed = true;
    }
    if (!existing.customer && customer) {
      existing.customer = customer;
      ensureCustomer(customer);
      changed = true;
    }
    if (String(existing.status).toLowerCase() === 'pending') {
      if (numericAmount(existing.amount) !== amount) {
        existing.amount = amount;
        changed = true;
      }
      if (!existing.dueDate && dueDate) {
        existing.dueDate = dueDate;
        changed = true;
      }
    }
    return changed;
  }

  const invoiceNumber = generateInvoiceNumber(trip.tripId);
  const newInvoice = {
    invoice: invoiceNumber,
    customer,
    trip: trip.tripId || '',
    amount,
    dueDate,
    status: 'Pending',
    notes: 'Auto-generated from completed trip'
  };
  state.invoices.unshift(newInvoice);
  ensureCustomer(customer);
  return true;
}

function ensureSupplierPaymentForTrip(trip) {
  if (!trip) return false;
  const ownership = String(getTripOwnership(trip) || '').toLowerCase();
  if (ownership !== 'rent-in') return false;
  const supplierName = getVehicleSupplier(trip.vehicle);
  if (!supplierName) return false;
  const amount = numericAmount(trip.rentalCharges);
  const dueDate = defaultDueDate(trip.date);

  const existing = state.suppliers.find(payment => payment && payment.trip === trip.tripId);
  if (existing) {
    let changed = false;
    if (!existing.reference) {
      existing.reference = generateSupplierPaymentReference(trip.tripId);
      changed = true;
    }
    if (!existing.supplier) {
      existing.supplier = supplierName;
      changed = true;
    }
    if (!existing.vehicle && trip.vehicle) {
      existing.vehicle = trip.vehicle;
      changed = true;
    }
    if (String(existing.status).toLowerCase() === 'pending' && numericAmount(existing.amount) !== amount) {
      existing.amount = amount;
      changed = true;
    }
    if (!existing.dueDate && dueDate) {
      existing.dueDate = dueDate;
      changed = true;
    }
    return changed;
  }

  const payment = {
    reference: generateSupplierPaymentReference(trip.tripId),
    supplier: supplierName,
    vehicle: trip.vehicle || '',
    trip: trip.tripId || '',
    amount,
    dueDate,
    status: 'Pending',
    notes: 'Auto-generated from completed rent-in trip'
  };
  state.suppliers.unshift(payment);
  return true;
}

function handleTripFinancialAutomation(trip) {
  const result = { invoice: false, supplier: false };
  if (!trip || String(trip.status).toLowerCase() !== 'completed') {
    return result;
  }
  result.invoice = ensureInvoiceForTrip(trip);
  result.supplier = ensureSupplierPaymentForTrip(trip);
  return result;
}

function reconcileCompletedTripFinancials() {
  let mutated = false;
  (state.trips || []).forEach(trip => {
    const result = handleTripFinancialAutomation(trip);
    if (result.invoice || result.supplier) {
      mutated = true;
    }
  });
  if (mutated) {
    persistState();
  }
}

const dom = {
  year: document.getElementById('year'),
  cards: document.querySelectorAll('.card'),
  modals: document.querySelectorAll('.modal'),
  exportBtn: document.getElementById('exportExcel'),
  fleetTable: document.querySelector('#fleetTable tbody'),
  fleetFilter: document.getElementById('fleetFilter'),
  driverTable: document.querySelector('#driverTable tbody'),
  tripTable: document.querySelector('#tripTable tbody'),
  invoiceTable: document.querySelector('#invoiceTable tbody'),
  invoiceFilter: document.getElementById('invoiceFilter'),
  invoiceReport: document.getElementById('invoiceReport'),
  supplierTable: document.querySelector('#supplierTable tbody'),
  supplierFilter: document.getElementById('supplierFilter'),
  supplierReport: document.getElementById('supplierReport'),
  supplierDirectoryTable: document.querySelector('#supplierDirectoryTable tbody'),
  dashboard: {
    totalTrips: document.getElementById('statTotalTrips'),
    tripBreakdown: document.getElementById('statTripBreakdown'),
    customerReceivedAmount: document.getElementById('statCustomerReceivedAmount'),
    customerReceivedCount: document.getElementById('statCustomerReceivedCount'),
    customerPendingAmount: document.getElementById('statCustomerPendingAmount'),
    customerPendingCount: document.getElementById('statCustomerPendingCount'),
    supplierPaidAmount: document.getElementById('statSupplierPaidAmount'),
    supplierPaidCount: document.getElementById('statSupplierPaidCount'),
    supplierPendingAmount: document.getElementById('statSupplierPendingAmount'),
    supplierPendingCount: document.getElementById('statSupplierPendingCount'),
    netCashFlow: document.getElementById('statNetCashFlow'),
    netOutstanding: document.getElementById('statNetOutstanding')
  },
  dashboardControls: {
    buttons: document.querySelectorAll('#dashboardRangeButtons [data-range]'),
    customRange: document.getElementById('dashboardCustomRange'),
    start: document.getElementById('dashboardRangeStart'),
    end: document.getElementById('dashboardRangeEnd')
  },
  theme: {
    toggle: document.getElementById('dashboardThemeToggle')
  },
  auth: {
    overlay: document.getElementById('authOverlay'),
    form: document.getElementById('authForm'),
    username: document.getElementById('authUsername'),
    password: document.getElementById('authPassword'),
    submit: document.getElementById('authSubmit'),
    message: document.getElementById('authMessage'),
    error: document.getElementById('authError')
  }
};

if (dom.auth && dom.auth.submit && dom.auth.submit.textContent) {
  authSubmitDefaultLabel = dom.auth.submit.textContent;
}

function setThemePreference(mode) {
  const useDark = mode === 'dark';
  if (bodyElement) {
    bodyElement.classList.toggle('dark-theme', useDark);
  }
  try {
    localStorage.setItem(themeStorageKey, useDark ? 'dark' : 'light');
  } catch (error) {
    console.warn('Unable to persist theme preference.', error);
  }
  if (dom.theme && dom.theme.toggle) {
    dom.theme.toggle.checked = useDark;
  }
}

function initializeThemeToggle() {
  if (!dom.theme || !dom.theme.toggle) {
    return;
  }
  const useDark = bodyElement ? bodyElement.classList.contains('dark-theme') : false;
  dom.theme.toggle.checked = useDark;
  dom.theme.toggle.addEventListener('change', event => {
    setThemePreference(event.target.checked ? 'dark' : 'light');
  });
}

function showAuthOverlay(message) {
  if (!dom.auth || !dom.auth.overlay) {
    return;
  }
  dom.auth.overlay.classList.remove('hidden');
  if (bodyElement) {
    bodyElement.classList.add('app-locked');
  }
  if (dom.auth.message && message) {
    dom.auth.message.textContent = message;
  }
  if (dom.auth.error) {
    dom.auth.error.textContent = '';
  }
  if (dom.auth.username) {
    window.setTimeout(() => dom.auth.username.focus(), 60);
  }
}

function hideAuthOverlay() {
  if (!dom.auth || !dom.auth.overlay) {
    return;
  }
  dom.auth.overlay.classList.add('hidden');
  if (dom.auth.error) {
    dom.auth.error.textContent = '';
  }
}

function handleUnauthorized(message) {
  const detail = message || 'Authentication required. Please sign in to continue.';
  setAuthToken('');
  setSyncStatus('offline', detail);
  showAuthOverlay(detail);
  if (dom.auth && dom.auth.error) {
    dom.auth.error.textContent = detail;
  }
}

async function attemptLogin(username, password) {
  const response = await fetch(resolveApiPath('/api/login'), {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Accept: 'application/json'
    },
    body: JSON.stringify({ username, password })
  });
  if (response.status === 401) {
    throw new Error('Invalid username or password.');
  }
  if (!response.ok) {
    throw new Error('Unable to sign in right now. Please try again.');
  }
  return response.json();
}

function initializeAuthUi() {
  if (!dom.auth || !dom.auth.form) {
    return;
  }
  dom.auth.form.addEventListener('submit', async event => {
    event.preventDefault();
    const username = dom.auth.username ? dom.auth.username.value.trim() : '';
    const password = dom.auth.password ? dom.auth.password.value : '';
    if (!username || !password) {
      if (dom.auth.error) {
        dom.auth.error.textContent = 'Enter both username and password to continue.';
      }
      return;
    }
    if (dom.auth.submit) {
      dom.auth.submit.disabled = true;
      dom.auth.submit.textContent = 'Signing In…';
    }
    if (dom.auth.error) {
      dom.auth.error.textContent = '';
    }
    try {
      const result = await attemptLogin(username, password);
      setAuthToken(result.token, result.user || { username });
      hideAuthOverlay();
      if (dom.auth.form) {
        dom.auth.form.reset();
      }
      await bootstrapAfterAuth();
    } catch (error) {
      const message = error && error.message ? error.message : 'Unable to sign in. Please try again.';
      showAuthOverlay(message);
      if (dom.auth.error) {
        dom.auth.error.textContent = message;
      }
    } finally {
      if (dom.auth.submit) {
        dom.auth.submit.disabled = false;
        dom.auth.submit.textContent = authSubmitDefaultLabel;
      }
    }
  });
}

async function bootstrapAfterAuth() {
  if (bootstrapPromise) {
    return bootstrapPromise;
  }
  bootstrapPromise = (async () => {
    await bootstrapState();
    reconcileCompletedTripFinancials();
    renderAll();
    updateAllSelectOptions();
  })();
  try {
    await bootstrapPromise;
  } finally {
    bootstrapPromise = null;
  }
}

  }
};

function initializeModalDefaults() {
  document.querySelectorAll('.modal').forEach(modal => {
    const title = modal.querySelector('header h3');
    const submit = modal.querySelector('button[type="submit"]');
    if (title && !modal.dataset.defaultTitle) {
      modal.dataset.defaultTitle = title.textContent;
    }
    if (submit && !submit.dataset.defaultLabel) {
      submit.dataset.defaultLabel = submit.textContent;
    }
    setModalMode(modal.id, 'create');
  });
}

function setModalMode(modalId, mode = 'create') {
  const modal = document.getElementById(modalId);
  if (!modal) return;
  const title = modal.querySelector('header h3');
  const submit = modal.querySelector('button[type="submit"]');
  if (mode === 'edit') {
    modal.dataset.mode = 'edit';
    if (title) {
      title.textContent = modalEditTitles[modalId] || modal.dataset.defaultTitle || title.textContent;
    }
    if (submit) {
      submit.textContent = modalEditSubmitLabels[modalId] || submit.dataset.defaultLabel || submit.textContent;
    }
  } else {
    modal.dataset.mode = 'create';
    if (title && modal.dataset.defaultTitle) {
      title.textContent = modal.dataset.defaultTitle;
    }
    if (submit && submit.dataset.defaultLabel) {
      submit.textContent = submit.dataset.defaultLabel;
    }
  }
}

function isModalInEditMode(modalId) {
  const modal = document.getElementById(modalId);
  return modal ? modal.dataset.mode === 'edit' : false;
}

function startEditing(type, index) {
  editingContext.type = type;
  editingContext.index = index;
}

function clearEditingContext() {
  editingContext.type = null;
  editingContext.index = -1;
}

function getModalType(modalId) {
  return modalTypeMap[modalId] || null;
}

function isEditing(type) {
  return editingContext.type === type && editingContext.index > -1;
}

document.addEventListener('DOMContentLoaded', async () => {
  if (dom.year) {
    dom.year.textContent = new Date().getFullYear();
  }
document.addEventListener('DOMContentLoaded', () => {
document.addEventListener('DOMContentLoaded', async () => {
  dom.year.textContent = new Date().getFullYear();

  initializeModalDefaults();
  setupModals();
  setupForms();
  setupFilters();
  initializeThemeToggle();
  initializeAuthUi();
  if (dom.exportBtn) {
    dom.exportBtn.addEventListener('click', exportWorkbook);
  }
  reconcileCompletedTripFinancials();
  dom.exportBtn.addEventListener('click', exportWorkbook);

  renderAll();
  updateAllSelectOptions();

  if (authToken) {
    hideAuthOverlay();
    await bootstrapAfterAuth();
  } else {
    showAuthOverlay('Sign in with your Azmat administrator account to manage fleet operations.');
    setSyncStatus('offline', 'Authentication required. Please sign in to sync data.');
  }
  await bootstrapState();
  reconcileCompletedTripFinancials();
  renderAll();
  updateAllSelectOptions();
});

function setupModals() {
  dom.cards.forEach(card => {
    const modalId = card.dataset.modal;
    card.addEventListener('click', () => {
      if (!modalId) return;
      clearEditingContext();
      setModalMode(modalId, 'create');
      openModal(modalId);
    });
    const button = card.querySelector('button');
    if (button) {
      button.addEventListener('click', evt => {
        evt.stopPropagation();
        if (!modalId) return;
        clearEditingContext();
        setModalMode(modalId, 'create');
        openModal(modalId);
      });
    }
  });

  dom.modals.forEach(modal => {
    modal.addEventListener('click', event => {
      if (event.target === modal || event.target.hasAttribute('data-close')) {
        closeModal(modal.id);
      }
    });
  });
}

function openModal(id) {
  const modal = document.getElementById(id);
  if (modal) {
    prepareModal(id);
    modal.setAttribute('aria-hidden', 'false');
    const focusable = modal.querySelector('input, select, textarea');
    if (focusable) {
      setTimeout(() => focusable.focus(), 60);
    }
  }
}

function closeModal(id) {
  const modal = document.getElementById(id);
  if (modal) {
    modal.setAttribute('aria-hidden', 'true');
    const form = modal.querySelector('form');
    if (form) form.reset();
    if (id === 'fleetModal' && selectRefs.fleetSupplier && fleetOwnershipSelect) {
      toggleDependentSelect(selectRefs.fleetSupplier, fleetOwnershipSelect.value === 'rent-in');
    }
    if (id === 'driverModal' && selectRefs.driverSupplier && driverAffiliationSelect) {
      toggleDependentSelect(selectRefs.driverSupplier, driverAffiliationSelect.value === 'supplier');
    }
    if (id === 'tripModal') {
      updateTripOwnershipDisplay('');
      refreshTripDriverOptions();
    }
    if (id === 'supplierModal') {
      updateSupplierPaymentDependencies();
    }
    const type = getModalType(id);
    if (type && editingContext.type === type) {
      clearEditingContext();
    }
    setModalMode(id, 'create');
  }
}

function prepareModal(id) {
  const editing = isModalInEditMode(id);
  if (id === 'tripModal' && tripIdInput && !editing) {
    tripIdInput.value = formatTripId(state.nextTripNumber);
  }
  if (id === 'tripModal' || id === 'invoiceModal' || id === 'supplierModal' || id === 'fleetModal' || id === 'driverModal') {
    updateAllSelectOptions();
  }
  if (id === 'fleetModal' && selectRefs.fleetSupplier && fleetOwnershipSelect) {
    toggleDependentSelect(selectRefs.fleetSupplier, fleetOwnershipSelect.value === 'rent-in');
  }
  if (id === 'driverModal' && selectRefs.driverSupplier && driverAffiliationSelect) {
    toggleDependentSelect(selectRefs.driverSupplier, driverAffiliationSelect.value === 'supplier');
  }
  if (id === 'tripModal') {
    const selectedVehicle = selectRefs.tripVehicle ? selectRefs.tripVehicle.value : '';
    updateTripOwnershipDisplay(selectedVehicle);
    refreshTripDriverOptions();
  }
}

function toggleDependentSelect(select, enabled) {
  if (!select) return;
  select.disabled = !enabled;
  select.required = !!enabled;
  if (!enabled) {
    select.value = '';
  }
}

function handleAllowNewSelection(select) {
  if (!select || select.value !== '__new__') return;
  const label = select.getAttribute('data-label') || select.name || 'entry';
  const response = window.prompt(`Enter new ${label} name`);
  if (response && response.trim()) {
    const value = response.trim();
    ensureCustomer(value);
    select.dataset.pendingValue = value;
    persistState();
    updateAllSelectOptions();
    select.value = value;
  } else {
    select.value = '';
  }
}

function updateAllSelectOptions() {
  const supplierNames = getSupplierNames();
  setSelectOptions(selectRefs.fleetSupplier, supplierNames, 'Select supplier');
  setSelectOptions(selectRefs.driverSupplier, supplierNames, 'Select supplier');
  setSelectOptions(selectRefs.supplierPaymentSupplier, supplierNames, 'Select supplier');

  const vehicleIds = state.fleet.map(item => item.unitId).filter(Boolean);
  setSelectOptions(selectRefs.tripVehicle, vehicleIds, 'Select vehicle');

  const customers = [...state.customers].sort((a, b) => a.localeCompare(b));
  setSelectOptions(selectRefs.tripCustomer, customers, 'Select customer', { allowNew: true });
  setSelectOptions(selectRefs.invoiceCustomer, customers, 'Select customer', { allowNew: true });

  const tripReferences = state.trips.map(item => item.tripId).filter(Boolean);
  setSelectOptions(selectRefs.invoiceTrip, tripReferences, 'Select trip');

  updateSupplierPaymentDependencies();

  refreshTripDriverOptions();
  const activeVehicleId = selectRefs.tripVehicle ? selectRefs.tripVehicle.value : '';
  updateTripOwnershipDisplay(activeVehicleId);
}

function setSelectOptions(select, values, placeholder, options = {}) {
  if (!select) return;
  const { allowNew = false } = options;
  const previousValue = select.dataset.pendingValue || select.value;
  select.innerHTML = '';

  const placeholderOption = document.createElement('option');
  placeholderOption.value = '';
  placeholderOption.textContent = placeholder;
  select.appendChild(placeholderOption);

  const uniqueValues = Array.from(new Set(values || [])).filter(Boolean);
  uniqueValues.forEach(value => {
    const option = document.createElement('option');
    option.value = value;
    option.textContent = value;
    select.appendChild(option);
  });

  if (allowNew) {
    const addNew = document.createElement('option');
    addNew.value = '__new__';
    addNew.textContent = '+ Add new';
    select.appendChild(addNew);
  }

  if (previousValue && Array.from(select.options).some(option => option.value === previousValue)) {
    select.value = previousValue;
  } else {
    select.value = '';
  }

  delete select.dataset.pendingValue;
}

function getVehicleOwnership(vehicleId) {
  return findVehicleOwnership(vehicleId) || '';
}

function describeOwnership(value) {
  if (!value) return '';
  const normalized = String(value).toLowerCase();
  if (normalized === 'rent-in' || normalized === 'rentin' || normalized === 'rent in') {
    return 'Rent-In';
  }
  if (normalized === 'owned' || normalized === 'company owned' || normalized === 'company') {
    return 'Company Owned';
  }
  return value;
}

function getTripOwnership(trip) {
  if (!trip) return '';
  if (trip.unitOwnership) {
    return trip.unitOwnership;
  }
  return getVehicleOwnership(trip.vehicle);
}

function updateTripOwnershipDisplay(vehicleId, presetOwnership = '') {
  if (!tripOwnershipField) return;
  if (presetOwnership) {
    tripOwnershipField.dataset.fallbackOwnership = presetOwnership;
  } else if (!vehicleId) {
    delete tripOwnershipField.dataset.fallbackOwnership;
  }
  const fallback = presetOwnership || tripOwnershipField.dataset.fallbackOwnership || '';
  const ownership = getVehicleOwnership(vehicleId) || fallback;
  tripOwnershipField.value = describeOwnership(ownership) || '';
}

function refreshTripDriverOptions(preservedDriver) {
  const driverSelect = selectRefs.tripDriver;
  const vehicleSelect = selectRefs.tripVehicle;
  if (!driverSelect) return;
  const vehicleId = vehicleSelect ? vehicleSelect.value : '';
  const vehicleEntry = findFleetUnit(vehicleId);
  const ownership = vehicleEntry ? vehicleEntry.ownership : getVehicleOwnership(vehicleId);
  const vehicleSupplier = vehicleEntry && vehicleEntry.supplier ? vehicleEntry.supplier : '';
  const normalizedVehicleSupplier = normalizeText(vehicleSupplier);
  let drivers = Array.isArray(state.drivers) ? state.drivers.filter(Boolean) : [];
  if (ownership === 'rent-in') {
    drivers = drivers.filter(driver => {
      if (!driver || driver.affiliation !== 'supplier') return false;
      if (!normalizedVehicleSupplier) return !!driver.supplier;
      return normalizeText(driver.supplier) === normalizedVehicleSupplier;
    });
  } else if (ownership === 'owned') {
    drivers = drivers.filter(driver => driver && driver.affiliation !== 'supplier');
  }
  const driverNames = drivers.map(driver => driver && driver.name).filter(Boolean);
  const desiredValue = preservedDriver || driverSelect.dataset.pendingValue || driverSelect.value;
  if (desiredValue) {
    driverSelect.dataset.pendingValue = desiredValue;
  }
  setSelectOptions(driverSelect, driverNames, 'Select driver');
  if (vehicleId) {
    driverSelect.disabled = driverNames.length === 0;
  } else {
    driverSelect.disabled = false;
  }
}

function getSupplierVehicles(supplierName) {
  if (!supplierName) return [];
  const normalizedSupplier = normalizeText(supplierName);
  return state.fleet
    .filter(item => {
      if (!item || item.ownership !== 'rent-in') return false;
      return normalizeText(item.supplier) === normalizedSupplier;
    })
    .map(item => item.unitId)
    .filter(Boolean)
    .sort((a, b) => a.localeCompare(b));
}

function getSupplierTripsForPayment(supplierName, vehicleId) {
  if (!supplierName) return [];
  const vehicles = vehicleId ? [vehicleId] : getSupplierVehicles(supplierName);
  const vehicleSet = new Set(vehicles);
  return state.trips
    .filter(trip => trip && vehicleSet.has(trip.vehicle))
    .map(trip => trip.tripId)
    .filter(Boolean)
    .sort((a, b) => a.localeCompare(b));
}

function updateSupplierPaymentVehicleOptions(supplierName) {
  const select = selectRefs.supplierPaymentVehicle;
  if (!select) return;
  const previous = select.dataset.pendingValue || select.value;
  const vehicles = supplierName ? getSupplierVehicles(supplierName) : [];
  const values = [...vehicles];
  if (supplierName && previous && !values.includes(previous)) {
    values.unshift(previous);
  }
  setSelectOptions(select, values, 'Select vehicle');
  select.disabled = !supplierName || values.length === 0;
}

function updateSupplierPaymentTripOptions(supplierName, vehicleId) {
  const select = selectRefs.supplierPaymentTrip;
  if (!select) return;
  const previous = select.dataset.pendingValue || select.value;
  const trips = supplierName ? getSupplierTripsForPayment(supplierName, vehicleId) : [];
  const values = [...trips];
  if (supplierName && previous && !values.includes(previous)) {
    values.unshift(previous);
  }
  setSelectOptions(select, values, 'Select trip');
  select.disabled = !supplierName || values.length === 0;
}

function updateSupplierPaymentDependencies() {
  const supplier = selectRefs.supplierPaymentSupplier ? selectRefs.supplierPaymentSupplier.value : '';
  updateSupplierPaymentVehicleOptions(supplier);
  const vehicleValue = selectRefs.supplierPaymentVehicle ? selectRefs.supplierPaymentVehicle.value : '';
  updateSupplierPaymentTripOptions(supplier, vehicleValue);
}

function getSupplierNames() {
  const names = new Set();
  (state.supplierDirectory || []).forEach(supplier => {
    if (supplier && supplier.name) names.add(supplier.name);
  });
  (state.fleet || []).forEach(item => {
    if (item && item.supplier) names.add(item.supplier);
  });
  (state.drivers || []).forEach(driver => {
    if (driver && driver.supplier) names.add(driver.supplier);
  });
  (state.suppliers || []).forEach(payment => {
    if (payment && payment.supplier) names.add(payment.supplier);
  });
  return Array.from(names).sort((a, b) => a.localeCompare(b));
}

function ensureCustomer(customerName) {
  if (!customerName) return;
  const trimmed = String(customerName).trim();
  if (!trimmed) return;
  if (!state.customers.includes(trimmed)) {
    state.customers.push(trimmed);
    state.customers.sort((a, b) => a.localeCompare(b));
  }
}

function formatTripId(sequenceNumber) {
  const number = Number(sequenceNumber) || 0;
  return `TRIP-${String(Math.max(number, 1)).padStart(4, '0')}`;
}

function setupForms() {
  const fleetForm = document.getElementById('fleetForm');
  const driverForm = document.getElementById('driverForm');
  const tripForm = document.getElementById('tripForm');
  const invoiceForm = document.getElementById('invoiceForm');
  const supplierForm = document.getElementById('supplierForm');
  const supplierDirectoryForm = document.getElementById('supplierDirectoryForm');

  if (fleetForm) {
    selectRefs.fleetSupplier = fleetForm.querySelector('select[name="supplier"]');
    fleetOwnershipSelect = fleetForm.querySelector('select[name="ownership"]');
    fleetForm.addEventListener('submit', event => {
      event.preventDefault();
      const formData = new FormData(fleetForm);
      const ownership = formData.get('ownership');
      const unitId = String(formData.get('unitId') || '').trim();
      if (!unitId) {
        window.alert('Unit ID is required.');
        return;
      }
      const normalizedUnitId = normalizeUnitId(unitId);
      const duplicateIndex = state.fleet.findIndex((item, index) => {
        if (!item || !item.unitId) return false;
        if (normalizeUnitId(item.unitId) !== normalizedUnitId) return false;
        if (isEditing('fleet') && index === editingContext.index) return false;
        return true;
      });
      if (duplicateIndex !== -1) {
        window.alert('A fleet unit with this ID already exists. Please use a unique Unit ID.');
        return;
      }
      const entry = {
        unitId,
        fleetType: formData.get('fleetType'),
        ownership,
        model: formData.get('model'),
        capacity: formData.get('capacity'),
        status: formData.get('status'),
        supplier: ownership === 'rent-in' ? String(formData.get('supplier') || '').trim() : ''
      };
      if (isEditing('fleet')) {
        state.fleet.splice(editingContext.index, 1, entry);
      } else {
        state.fleet.unshift(entry);
      }
      persistState();
      renderFleet();
      updateAllSelectOptions();
      closeModal('fleetModal');
    });
    if (fleetOwnershipSelect && selectRefs.fleetSupplier) {
      fleetOwnershipSelect.addEventListener('change', () => {
        toggleDependentSelect(selectRefs.fleetSupplier, fleetOwnershipSelect.value === 'rent-in');
      });
      toggleDependentSelect(selectRefs.fleetSupplier, fleetOwnershipSelect.value === 'rent-in');
    }
  }

  if (driverForm) {
    selectRefs.driverSupplier = driverForm.querySelector('select[name="supplier"]');
    driverAffiliationSelect = driverForm.querySelector('select[name="affiliation"]');
    driverForm.addEventListener('submit', event => {
      event.preventDefault();
      const formData = new FormData(driverForm);
      const affiliation = formData.get('affiliation');
      const supplierName = affiliation === 'supplier' ? String(formData.get('supplier') || '').trim() : '';
      const entry = {
        name: formData.get('name'),
        license: formData.get('license'),
        phone: formData.get('phone'),
        availability: formData.get('availability'),
        affiliation,
        supplier: supplierName
      };
      if (isEditing('driver')) {
        state.drivers.splice(editingContext.index, 1, entry);
      } else {
        state.drivers.unshift(entry);
      }
      persistState();
      renderDrivers();
      updateAllSelectOptions();
      closeModal('driverModal');
    });
    if (driverAffiliationSelect && selectRefs.driverSupplier) {
      driverAffiliationSelect.addEventListener('change', () => {
        const isSupplier = driverAffiliationSelect.value === 'supplier';
        toggleDependentSelect(selectRefs.driverSupplier, isSupplier);
      });
      toggleDependentSelect(selectRefs.driverSupplier, driverAffiliationSelect.value === 'supplier');
    }
  }

  if (tripForm) {
    tripIdInput = tripForm.querySelector('input[name="tripId"]');
    tripOwnershipField = tripForm.querySelector('input[name="unitOwnership"]');
    selectRefs.tripCustomer = tripForm.querySelector('select[name="customer"]');
    selectRefs.tripVehicle = tripForm.querySelector('select[name="vehicle"]');
    selectRefs.tripDriver = tripForm.querySelector('select[name="driver"]');
    if (selectRefs.tripCustomer) {
      selectRefs.tripCustomer.addEventListener('change', () => handleAllowNewSelection(selectRefs.tripCustomer));
    }
    if (selectRefs.tripVehicle) {
      selectRefs.tripVehicle.addEventListener('change', () => {
        const vehicleId = selectRefs.tripVehicle.value;
        updateTripOwnershipDisplay(vehicleId);
        refreshTripDriverOptions();
      });
    }
    tripForm.addEventListener('submit', event => {
      event.preventDefault();
      const formData = new FormData(tripForm);
      const customer = formData.get('customer');
      const vehicle = formData.get('vehicle');
      const rentalCharges = Number(formData.get('rentalCharges'));
      const editingTrip = isEditing('trip');
      let tripId = formData.get('tripId');
      if (!editingTrip) {
        tripId = formatTripId(state.nextTripNumber);
        state.nextTripNumber += 1;
      } else if (!tripId && state.trips[editingContext.index]) {
        tripId = state.trips[editingContext.index].tripId;
      }
      const ownership = getVehicleOwnership(vehicle);
      const entry = {
        tripId,
        date: formData.get('date'),
        customer,
        vehicle,
        driver: formData.get('driver'),
        route: formData.get('route'),
        rentalCharges,
        status: formData.get('status'),
        unitOwnership: ownership
      };
      ensureCustomer(customer);
      if (editingTrip) {
        state.trips.splice(editingContext.index, 1, entry);
      } else {
        state.trips.unshift(entry);
      }
      const automationResult = handleTripFinancialAutomation(entry);
      persistState();
      renderTrips();
      if (automationResult.invoice) {
        renderInvoices();
      }
      if (automationResult.supplier) {
        renderSuppliers();
      }
      updateAllSelectOptions();
      closeModal('tripModal');
    });
  }

  if (invoiceForm) {
    selectRefs.invoiceCustomer = invoiceForm.querySelector('select[name="customer"]');
    selectRefs.invoiceTrip = invoiceForm.querySelector('select[name="trip"]');
    if (selectRefs.invoiceCustomer) {
      selectRefs.invoiceCustomer.addEventListener('change', () => handleAllowNewSelection(selectRefs.invoiceCustomer));
    }
    invoiceForm.addEventListener('submit', event => {
      event.preventDefault();
      const formData = new FormData(invoiceForm);
      const customer = formData.get('customer');
      const entry = {
        invoice: formData.get('invoice'),
        customer,
        trip: formData.get('trip'),
        amount: Number(formData.get('amount')),
        dueDate: formData.get('dueDate'),
        status: formData.get('status'),
        notes: formData.get('notes')
      };
      ensureCustomer(customer);
      if (isEditing('invoice')) {
        state.invoices.splice(editingContext.index, 1, entry);
      } else {
        state.invoices.unshift(entry);
      }
      persistState();
      renderInvoices();
      updateAllSelectOptions();
      closeModal('invoiceModal');
    });
  }

  if (supplierForm) {
    selectRefs.supplierPaymentSupplier = supplierForm.querySelector('select[name="supplier"]');
    selectRefs.supplierPaymentVehicle = supplierForm.querySelector('select[name="vehicle"]');
    selectRefs.supplierPaymentTrip = supplierForm.querySelector('select[name="trip"]');
    if (selectRefs.supplierPaymentSupplier) {
      selectRefs.supplierPaymentSupplier.addEventListener('change', () => {
        if (selectRefs.supplierPaymentVehicle) {
          selectRefs.supplierPaymentVehicle.dataset.pendingValue = '';
        }
        if (selectRefs.supplierPaymentTrip) {
          selectRefs.supplierPaymentTrip.dataset.pendingValue = '';
        }
        updateSupplierPaymentDependencies();
      });
    }
    if (selectRefs.supplierPaymentVehicle) {
      selectRefs.supplierPaymentVehicle.addEventListener('change', () => {
        if (selectRefs.supplierPaymentTrip) {
          selectRefs.supplierPaymentTrip.dataset.pendingValue = '';
        }
        const supplier = selectRefs.supplierPaymentSupplier ? selectRefs.supplierPaymentSupplier.value : '';
        updateSupplierPaymentTripOptions(supplier, selectRefs.supplierPaymentVehicle.value);
      });
    }
    supplierForm.addEventListener('submit', event => {
      event.preventDefault();
      const formData = new FormData(supplierForm);
      const entry = {
        reference: formData.get('reference'),
        supplier: formData.get('supplier'),
        vehicle: formData.get('vehicle') || '',
        trip: formData.get('trip') || '',
        amount: Number(formData.get('amount')),
        dueDate: formData.get('dueDate'),
        status: formData.get('status'),
        notes: formData.get('notes')
      };
      if (isEditing('supplierPayment')) {
        state.suppliers.splice(editingContext.index, 1, entry);
      } else {
        state.suppliers.unshift(entry);
      }
      persistState();
      renderSuppliers();
      updateAllSelectOptions();
      closeModal('supplierModal');
    });
  }

  if (supplierDirectoryForm) {
    supplierDirectoryForm.addEventListener('submit', event => {
      event.preventDefault();
      const formData = new FormData(supplierDirectoryForm);
      const entry = {
        name: formData.get('name'),
        contact: formData.get('contact'),
        phone: formData.get('phone'),
        email: formData.get('email')
      };
      if (isEditing('supplierDirectory')) {
        state.supplierDirectory.splice(editingContext.index, 1, entry);
      } else {
        state.supplierDirectory.unshift(entry);
      }
      persistState();
      renderSupplierDirectory();
      updateAllSelectOptions();
      closeModal('supplierDirectoryModal');
    });
  }
}

function setupFilters() {
  if (dom.fleetFilter) {
    dom.fleetFilter.addEventListener('change', renderFleet);
  }
  if (dom.invoiceFilter) {
    dom.invoiceFilter.addEventListener('change', renderInvoices);
  }
  if (dom.supplierFilter) {
    dom.supplierFilter.addEventListener('change', renderSuppliers);
  }
  setupDashboardRangeControls();
}

function setupDashboardRangeControls() {
  const controls = dom.dashboardControls || {};
  const buttons = controls.buttons ? Array.from(controls.buttons) : [];
  buttons.forEach(button => {
    button.addEventListener('click', () => {
      const mode = button.dataset.range;
      if (mode) {
        setDashboardRangeMode(mode);
      }
    });
  });
  if (controls.start) {
    controls.start.addEventListener('change', handleCustomRangeInputChange);
  }
  if (controls.end) {
    controls.end.addEventListener('change', handleCustomRangeInputChange);
  }
  syncDashboardFilterControls();
}

function setDashboardRangeMode(mode) {
  if (!mode) return;
  if (!state.dashboardFilter || typeof state.dashboardFilter !== 'object') {
    state.dashboardFilter = { ...defaultState.dashboardFilter };
  }
  if (state.dashboardFilter.mode === mode) {
    toggleCustomRangeVisibility(mode === 'custom');
    syncDashboardFilterControls();
    updateDashboard();
    return;
  }
  state.dashboardFilter.mode = mode;
  if (mode !== 'custom') {
    state.dashboardFilter.start = '';
    state.dashboardFilter.end = '';
  }
  persistState();
  syncDashboardFilterControls();
  updateDashboard();
}

function handleCustomRangeInputChange() {
  if (!state.dashboardFilter || typeof state.dashboardFilter !== 'object') {
    state.dashboardFilter = { ...defaultState.dashboardFilter };
  }
  const controls = dom.dashboardControls || {};
  if (controls.start) {
    state.dashboardFilter.start = controls.start.value || '';
  }
  if (controls.end) {
    state.dashboardFilter.end = controls.end.value || '';
  }
  state.dashboardFilter.mode = 'custom';
  persistState();
  syncDashboardFilterControls();
  updateDashboard();
}

function syncDashboardFilterControls() {
  const controls = dom.dashboardControls || {};
  const filter = state.dashboardFilter || defaultState.dashboardFilter;
  const mode = filter.mode || defaultState.dashboardFilter.mode;
  const buttons = controls.buttons ? Array.from(controls.buttons) : [];
  buttons.forEach(button => {
    const isActive = button.dataset.range === mode;
    button.setAttribute('aria-pressed', isActive ? 'true' : 'false');
  });
  if (controls.start) {
    controls.start.value = filter.start || '';
  }
  if (controls.end) {
    controls.end.value = filter.end || '';
  }
  toggleCustomRangeVisibility(mode === 'custom');
}

function toggleCustomRangeVisibility(shouldShow) {
  const controls = dom.dashboardControls || {};
  if (!controls.customRange) return;
  controls.customRange.classList.toggle('active', Boolean(shouldShow));
}

function renderAll() {
  renderFleet();
  renderDrivers();
  renderSupplierDirectory();
  renderTrips();
  renderInvoices();
  renderSuppliers();
}

function renderFleet() {
  const filter = dom.fleetFilter.value;
  const rows = state.fleet.filter(item => filter === 'all' || item.ownership === filter);
  dom.fleetTable.innerHTML = rows
    .map(item => {
      const statusClass = statusClassName(item.status);
      const index = state.fleet.indexOf(item);
      return `<tr>
        <td data-label="Unit ID">${escapeHtml(item.unitId)}</td>
        <td data-label="Fleet Type">${escapeHtml(item.fleetType || '')}</td>
        <td data-label="Type"><span class="status-pill ${item.ownership === 'owned' ? 'status-Available' : 'status-Scheduled'}">${item.ownership === 'owned' ? 'Company Owned' : 'Rent-In'}</span></td>
        <td data-label="Make & Model">${escapeHtml(item.model)}</td>
        <td data-label="Capacity">${escapeHtml(item.capacity || '')}</td>
        <td data-label="Status"><span class="status-pill ${statusClass}">${escapeHtml(item.status)}</span></td>
        <td data-label="Supplier">${escapeHtml(item.supplier || '-')}</td>
        <td data-label="Actions"><div class="table-actions"><button class="btn secondary" data-edit="fleet" data-index="${index}">Edit</button></div></td>
      </tr>`;
    })
    .join('');

  dom.fleetTable.querySelectorAll('[data-edit="fleet"]').forEach(button => {
    button.addEventListener('click', () => {
      const index = Number(button.dataset.index);
      if (!Number.isNaN(index)) {
        openFleetEditor(index);
      }
    });
  });
}

function renderDrivers() {
  dom.driverTable.innerHTML = state.drivers
    .map(item => {
      const affiliationLabel = item.affiliation === 'supplier'
        ? `Supplier${item.supplier ? ` (${item.supplier})` : ''}`
        : 'Company';
      const index = state.drivers.indexOf(item);
      return `<tr>
        <td data-label="Name">${escapeHtml(item.name)}</td>
        <td data-label="License No.">${escapeHtml(item.license)}</td>
        <td data-label="Phone"><a href="tel:${escapeHtml(item.phone)}">${escapeHtml(item.phone)}</a></td>
        <td data-label="Availability"><span class="status-pill ${statusClassName(item.availability)}">${escapeHtml(item.availability)}</span></td>
        <td data-label="Affiliation">${escapeHtml(affiliationLabel)}</td>
        <td data-label="Actions"><div class="table-actions"><button class="btn secondary" data-edit="driver" data-index="${index}">Edit</button></div></td>
      </tr>`;
    })
    .join('');

  dom.driverTable.querySelectorAll('[data-edit="driver"]').forEach(button => {
    button.addEventListener('click', () => {
      const index = Number(button.dataset.index);
      if (!Number.isNaN(index)) {
        openDriverEditor(index);
      }
    });
  });
}

function renderSupplierDirectory() {
  if (!dom.supplierDirectoryTable) return;
  dom.supplierDirectoryTable.innerHTML = (state.supplierDirectory || [])
    .map(item => {
      const index = state.supplierDirectory.indexOf(item);
      return `<tr>
      <td data-label="Name">${escapeHtml(item.name)}</td>
      <td data-label="Contact">${escapeHtml(item.contact || '')}</td>
      <td data-label="Phone">${escapeHtml(item.phone || '')}</td>
      <td data-label="Email">${escapeHtml(item.email || '')}</td>
      <td data-label="Actions"><div class="table-actions"><button class="btn secondary" data-edit="supplierDirectory" data-index="${index}">Edit</button></div></td>
    </tr>`;
    })
    .join('');

  dom.supplierDirectoryTable.querySelectorAll('[data-edit="supplierDirectory"]').forEach(button => {
    button.addEventListener('click', () => {
      const index = Number(button.dataset.index);
      if (!Number.isNaN(index)) {
        openSupplierDirectoryEditor(index);
      }
    });
  });
}

function renderTrips() {
  dom.tripTable.innerHTML = state.trips
    .map(item => {
      const index = state.trips.indexOf(item);
      const ownershipLabel = describeOwnership(getTripOwnership(item)) || '-';
      return `<tr>
      <td data-label="Trip ID">${escapeHtml(item.tripId)}</td>
      <td data-label="Date">${formatDate(item.date)}</td>
      <td data-label="Customer">${escapeHtml(item.customer)}</td>
      <td data-label="Vehicle">${escapeHtml(item.vehicle)}</td>
      <td data-label="Driver">${escapeHtml(item.driver)}</td>
      <td data-label="Unit Ownership">${escapeHtml(ownershipLabel)}</td>
      <td data-label="Route">${escapeHtml(item.route || '')}</td>
      <td data-label="Rental Charges">${formatCurrency(typeof item.rentalCharges === 'number' ? item.rentalCharges : Number(item.rentalCharges))}</td>
      <td data-label="Status"><span class="status-pill ${statusClassName(item.status)}">${escapeHtml(item.status)}</span></td>
      <td data-label="Actions"><div class="table-actions"><button class="btn secondary" data-edit="trip" data-index="${index}">Edit</button></div></td>
    </tr>`;
    })
    .join('');

  dom.tripTable.querySelectorAll('[data-edit="trip"]').forEach(button => {
    button.addEventListener('click', () => {
      const index = Number(button.dataset.index);
      if (!Number.isNaN(index)) {
        openTripEditor(index);
      }
    });
  });

  updateDashboard();
}

function renderInvoices() {
  const filter = dom.invoiceFilter.value;
  const rows = state.invoices.filter(item => filter === 'all' || item.status === filter);
  dom.invoiceTable.innerHTML = rows
    .map(item => {
      const index = state.invoices.indexOf(item);
      return `<tr>
      <td data-label="Invoice #">${escapeHtml(item.invoice)}</td>
      <td data-label="Customer">${escapeHtml(item.customer)}</td>
      <td data-label="Trip">${escapeHtml(item.trip || '-')}</td>
      <td data-label="Amount">${formatCurrency(item.amount)}</td>
      <td data-label="Due Date">${formatDate(item.dueDate)}</td>
      <td data-label="Status"><span class="status-pill ${statusClassName(item.status)}">${escapeHtml(item.status)}</span></td>
      <td data-label="Actions"><div class="table-actions"><button class="btn secondary" data-edit="invoice" data-index="${index}">Edit</button><button class="btn secondary" data-print="${encodeURIComponent(item.invoice)}">Print PDF</button></div></td>
    </tr>`;
    })
    .join('');

  dom.invoiceTable.querySelectorAll('[data-edit="invoice"]').forEach(button => {
    button.addEventListener('click', () => {
      const index = Number(button.dataset.index);
      if (!Number.isNaN(index)) {
        openInvoiceEditor(index);
      }
    });
  });

  dom.invoiceTable.querySelectorAll('[data-print]').forEach(button => {
    button.addEventListener('click', () => {
      const invoiceId = decodeURIComponent(button.dataset.print);
      printInvoice(invoiceId);
    });
  });

  dom.invoiceReport.innerHTML = buildInvoiceReport();
  updateDashboard();
}

function renderSuppliers() {
  const filter = dom.supplierFilter.value;
  const rows = state.suppliers.filter(item => filter === 'all' || item.status === filter);
  dom.supplierTable.innerHTML = rows
    .map(item => {
      const index = state.suppliers.indexOf(item);
      return `<tr>
      <td data-label="Reference">${escapeHtml(item.reference)}</td>
      <td data-label="Supplier">${escapeHtml(item.supplier)}</td>
      <td data-label="Vehicle">${escapeHtml(item.vehicle || '-')}</td>
      <td data-label="Trip Reference">${escapeHtml(item.trip || '-')}</td>
      <td data-label="Amount">${formatCurrency(item.amount)}</td>
      <td data-label="Due Date">${formatDate(item.dueDate)}</td>
      <td data-label="Status"><span class="status-pill ${statusClassName(item.status)}">${escapeHtml(item.status)}</span></td>
      <td data-label="Actions"><div class="table-actions"><button class="btn secondary" data-edit="supplierPayment" data-index="${index}">Edit</button></div></td>
    </tr>`;
    })
    .join('');

  dom.supplierTable.querySelectorAll('[data-edit="supplierPayment"]').forEach(button => {
    button.addEventListener('click', () => {
      const index = Number(button.dataset.index);
      if (!Number.isNaN(index)) {
        openSupplierPaymentEditor(index);
      }
    });
  });

  dom.supplierReport.innerHTML = buildSupplierReport();
  updateDashboard();
}

function updateDashboard() {
  if (!dom.dashboard) return;
  const {
    totalTrips,
    tripBreakdown,
    customerReceivedAmount,
    customerReceivedCount,
    customerPendingAmount,
    customerPendingCount,
    supplierPaidAmount,
    supplierPaidCount,
    supplierPendingAmount,
    supplierPendingCount,
    netCashFlow,
    netOutstanding
  } = dom.dashboard;

  const range = getDashboardDateRange();
  const tripsInRange = filterRecordsByDateRange(state.trips, trip => trip && trip.date, range);
  const invoicesInRange = filterRecordsByDateRange(state.invoices, invoice => invoice && invoice.dueDate, range);
  const suppliersInRange = filterRecordsByDateRange(state.suppliers, payment => payment && payment.dueDate, range);

  const totalTripsCount = tripsInRange.length;
  const completedTrips = tripsInRange.filter(trip => trip && String(trip.status).toLowerCase() === 'completed').length;
  const activeTrips = Math.max(totalTripsCount - completedTrips, 0);

  setTextContent(totalTrips, String(totalTripsCount));
  setTextContent(tripBreakdown, `Active: ${activeTrips} • Completed: ${completedTrips}`);

  const invoicesPaid = invoicesInRange.filter(invoice => invoice && String(invoice.status).toLowerCase() === 'paid');
  const invoicesPending = invoicesInRange.filter(invoice => invoice && String(invoice.status).toLowerCase() !== 'paid');
  const suppliersPaid = suppliersInRange.filter(payment => payment && String(payment.status).toLowerCase() === 'paid');
  const suppliersPending = suppliersInRange.filter(payment => payment && String(payment.status).toLowerCase() !== 'paid');

  const paidInAmount = invoicesPaid.reduce((sum, invoice) => sum + numericAmount(invoice.amount), 0);
  const pendingInAmount = invoicesPending.reduce((sum, invoice) => sum + numericAmount(invoice.amount), 0);
  const paidOutAmount = suppliersPaid.reduce((sum, payment) => sum + numericAmount(payment.amount), 0);
  const pendingOutAmount = suppliersPending.reduce((sum, payment) => sum + numericAmount(payment.amount), 0);

  setTextContent(customerReceivedAmount, formatCurrency(paidInAmount));
  setTextContent(customerReceivedCount, `${invoicesPaid.length} invoice${invoicesPaid.length === 1 ? '' : 's'} paid`);
  setTextContent(customerPendingAmount, formatCurrency(pendingInAmount));
  setTextContent(customerPendingCount, `${invoicesPending.length} invoice${invoicesPending.length === 1 ? '' : 's'} open`);

  setTextContent(supplierPaidAmount, formatCurrency(paidOutAmount));
  setTextContent(supplierPaidCount, `${suppliersPaid.length} payment${suppliersPaid.length === 1 ? '' : 's'} sent`);
  setTextContent(supplierPendingAmount, formatCurrency(pendingOutAmount));
  setTextContent(supplierPendingCount, `${suppliersPending.length} payment${suppliersPending.length === 1 ? '' : 's'} pending`);

  const netCash = paidInAmount - paidOutAmount;
  const outstandingImpact = pendingInAmount - pendingOutAmount;
  setTextContent(netCashFlow, formatCurrency(netCash));
  setTextContent(netOutstanding, `Outstanding receivable impact: ${formatCurrency(outstandingImpact)}`);
}

function openFleetEditor(index) {
  if (!state.fleet[index]) return;
  startEditing('fleet', index);
  setModalMode('fleetModal', 'edit');
  populateFleetForm(index);
  openModal('fleetModal');
}

function openDriverEditor(index) {
  if (!state.drivers[index]) return;
  startEditing('driver', index);
  setModalMode('driverModal', 'edit');
  populateDriverForm(index);
  openModal('driverModal');
}

function openTripEditor(index) {
  if (!state.trips[index]) return;
  startEditing('trip', index);
  setModalMode('tripModal', 'edit');
  populateTripForm(index);
  openModal('tripModal');
}

function openInvoiceEditor(index) {
  if (!state.invoices[index]) return;
  startEditing('invoice', index);
  setModalMode('invoiceModal', 'edit');
  populateInvoiceForm(index);
  openModal('invoiceModal');
}

function openSupplierPaymentEditor(index) {
  if (!state.suppliers[index]) return;
  startEditing('supplierPayment', index);
  setModalMode('supplierModal', 'edit');
  populateSupplierPaymentForm(index);
  openModal('supplierModal');
}

function openSupplierDirectoryEditor(index) {
  if (!state.supplierDirectory[index]) return;
  startEditing('supplierDirectory', index);
  setModalMode('supplierDirectoryModal', 'edit');
  populateSupplierDirectoryForm(index);
  openModal('supplierDirectoryModal');
}

function setFormValue(form, selector, value) {
  if (!form) return;
  const field = form.querySelector(selector);
  if (field) {
    field.value = value ?? '';
  }
}

function setTextContent(element, value) {
  if (!element) return;
  element.textContent = value;
}

function setSelectValue(select, value) {
  if (!select) return;
  const sanitized = value ?? '';
  select.dataset.pendingValue = sanitized;
  select.value = sanitized;
}

function populateFleetForm(index) {
  const entry = state.fleet[index];
  const modal = document.getElementById('fleetModal');
  if (!entry || !modal) return;
  const form = modal.querySelector('form');
  setFormValue(form, 'input[name="unitId"]', entry.unitId);
  setFormValue(form, 'select[name="fleetType"]', entry.fleetType);
  if (fleetOwnershipSelect) {
    fleetOwnershipSelect.value = entry.ownership || 'owned';
  }
  setFormValue(form, 'input[name="model"]', entry.model);
  setFormValue(form, 'input[name="capacity"]', entry.capacity);
  setFormValue(form, 'select[name="status"]', entry.status);
  if (selectRefs.fleetSupplier) {
    const supplierValue = entry.ownership === 'rent-in' ? entry.supplier : '';
    setSelectValue(selectRefs.fleetSupplier, supplierValue);
    toggleDependentSelect(selectRefs.fleetSupplier, entry.ownership === 'rent-in');
  }
}

function populateDriverForm(index) {
  const entry = state.drivers[index];
  const modal = document.getElementById('driverModal');
  if (!entry || !modal) return;
  const form = modal.querySelector('form');
  setFormValue(form, 'input[name="name"]', entry.name);
  setFormValue(form, 'input[name="license"]', entry.license);
  setFormValue(form, 'input[name="phone"]', entry.phone);
  setFormValue(form, 'select[name="availability"]', entry.availability);
  if (driverAffiliationSelect) {
    driverAffiliationSelect.value = entry.affiliation || 'company';
  }
  if (selectRefs.driverSupplier) {
    setSelectValue(selectRefs.driverSupplier, entry.supplier);
    toggleDependentSelect(selectRefs.driverSupplier, entry.affiliation === 'supplier');
  }
}

function populateTripForm(index) {
  const entry = state.trips[index];
  const modal = document.getElementById('tripModal');
  if (!entry || !modal) return;
  const form = modal.querySelector('form');
  setFormValue(form, 'input[name="tripId"]', entry.tripId);
  setFormValue(form, 'input[name="date"]', entry.date);
  if (selectRefs.tripCustomer) {
    setSelectValue(selectRefs.tripCustomer, entry.customer);
  }
  if (selectRefs.tripVehicle) {
    setSelectValue(selectRefs.tripVehicle, entry.vehicle);
    updateTripOwnershipDisplay(entry.vehicle, entry.unitOwnership);
  }
  if (selectRefs.tripDriver) {
    setSelectValue(selectRefs.tripDriver, entry.driver);
  }
  refreshTripDriverOptions(entry.driver);
  setFormValue(form, 'input[name="route"]', entry.route);
  const rentalValue = typeof entry.rentalCharges === 'number' ? entry.rentalCharges : Number(entry.rentalCharges) || entry.rentalCharges;
  setFormValue(form, 'input[name="rentalCharges"]', rentalValue);
  setFormValue(form, 'select[name="status"]', entry.status);
}

function populateInvoiceForm(index) {
  const entry = state.invoices[index];
  const modal = document.getElementById('invoiceModal');
  if (!entry || !modal) return;
  const form = modal.querySelector('form');
  setFormValue(form, 'input[name="invoice"]', entry.invoice);
  if (selectRefs.invoiceCustomer) {
    setSelectValue(selectRefs.invoiceCustomer, entry.customer);
  }
  if (selectRefs.invoiceTrip) {
    setSelectValue(selectRefs.invoiceTrip, entry.trip);
  }
  setFormValue(form, 'input[name="amount"]', typeof entry.amount === 'number' ? entry.amount : Number(entry.amount) || entry.amount);
  setFormValue(form, 'input[name="dueDate"]', entry.dueDate);
  setFormValue(form, 'select[name="status"]', entry.status);
  setFormValue(form, 'textarea[name="notes"]', entry.notes);
}

function populateSupplierPaymentForm(index) {
  const entry = state.suppliers[index];
  const modal = document.getElementById('supplierModal');
  if (!entry || !modal) return;
  const form = modal.querySelector('form');
  setFormValue(form, 'input[name="reference"]', entry.reference);
  if (selectRefs.supplierPaymentSupplier) {
    setSelectValue(selectRefs.supplierPaymentSupplier, entry.supplier);
  }
  if (selectRefs.supplierPaymentVehicle) {
    selectRefs.supplierPaymentVehicle.dataset.pendingValue = entry.vehicle || '';
  }
  if (selectRefs.supplierPaymentTrip) {
    selectRefs.supplierPaymentTrip.dataset.pendingValue = entry.trip || '';
  }
  updateSupplierPaymentDependencies();
  setFormValue(form, 'input[name="amount"]', typeof entry.amount === 'number' ? entry.amount : Number(entry.amount) || entry.amount);
  setFormValue(form, 'input[name="dueDate"]', entry.dueDate);
  setFormValue(form, 'select[name="status"]', entry.status);
  setFormValue(form, 'textarea[name="notes"]', entry.notes);
}

function populateSupplierDirectoryForm(index) {
  const entry = state.supplierDirectory[index];
  const modal = document.getElementById('supplierDirectoryModal');
  if (!entry || !modal) return;
  const form = modal.querySelector('form');
  setFormValue(form, 'input[name="name"]', entry.name);
  setFormValue(form, 'input[name="contact"]', entry.contact);
  setFormValue(form, 'input[name="phone"]', entry.phone);
  setFormValue(form, 'input[name="email"]', entry.email);
}

function buildInvoiceReport() {
  if (!state.invoices.length) {
    return '<p>No invoices recorded yet.</p>';
  }
  const summary = aggregateByStatus(state.invoices);
  const total = state.invoices.reduce((sum, item) => sum + (item.amount || 0), 0);
  const outstanding = state.invoices
    .filter(item => item.status === 'Pending' || item.status === 'Overdue' || item.status === 'Partial')
    .reduce((sum, item) => sum + (item.amount || 0), 0);

  const summaryLines = Object.entries(summary)
    .map(([status, { count, amount }]) => `<li><strong>${status}:</strong> ${count} invoice(s) &mdash; ${formatCurrency(amount)}</li>`)
    .join('');

  return `<div>
    <p><strong>Total billed:</strong> ${formatCurrency(total)}</p>
    <p><strong>Outstanding (Pending / Partial / Overdue):</strong> ${formatCurrency(outstanding)}</p>
    <ul>${summaryLines}</ul>
  </div>`;
}

function buildSupplierReport() {
  if (!state.suppliers.length) {
    return '<p>No supplier payments tracked yet.</p>';
  }
  const summary = aggregateByStatus(state.suppliers);
  const total = state.suppliers.reduce((sum, item) => sum + (item.amount || 0), 0);
  const dueSoon = state.suppliers
    .filter(item => item.status === 'Pending' || item.status === 'Scheduled')
    .reduce((sum, item) => sum + (item.amount || 0), 0);

  const summaryLines = Object.entries(summary)
    .map(([status, { count, amount }]) => `<li><strong>${status}:</strong> ${count} payment(s) &mdash; ${formatCurrency(amount)}</li>`)
    .join('');

  return `<div>
    <p><strong>Total supplier spend:</strong> ${formatCurrency(total)}</p>
    <p><strong>Pending / Scheduled:</strong> ${formatCurrency(dueSoon)}</p>
    <ul>${summaryLines}</ul>
  </div>`;
}

function aggregateByStatus(items) {
  return items.reduce((acc, item) => {
    const status = item.status || 'Unknown';
    if (!acc[status]) {
      acc[status] = { count: 0, amount: 0 };
    }
    acc[status].count += 1;
    acc[status].amount += item.amount || 0;
    return acc;
  }, {});
}

function numericAmount(value) {
  const number = Number(value);
  return Number.isFinite(number) ? number : 0;
}

function exportWorkbook() {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(state.fleet), 'Fleet');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(state.drivers), 'Drivers');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(state.trips), 'Trips');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(state.invoices), 'Customer Invoices');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(state.suppliers), 'Supplier Payments');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(state.supplierDirectory), 'Suppliers');
  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.json_to_sheet(state.customers.map(name => ({ Customer: name }))),
    'Customers'
  );
  XLSX.writeFile(wb, `Azmat-fleet-data-${new Date().toISOString().slice(0, 10)}.xlsx`);
}

function printInvoice(invoiceId) {
  const invoice = state.invoices.find(item => item.invoice === invoiceId);
  if (!invoice) return;

  const trip = state.trips.find(t => t.tripId === invoice.trip);
  const jsPdfNamespace = window.jspdf;
  const JsPDFConstructor = jsPdfNamespace && typeof jsPdfNamespace.jsPDF === 'function' ? jsPdfNamespace.jsPDF : window.jsPDF;
  if (typeof JsPDFConstructor === 'function') {
    const doc = new JsPDFConstructor({ unit: 'pt', format: 'a4' });
    const lineHeight = 22;
    let cursorY = 60;

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(20);
    doc.text('Azmat Rental Fleet Management', 40, cursorY);
    cursorY += lineHeight * 1.5;

    doc.setFont('helvetica', 'normal');
    doc.setFontSize(12);
    doc.text(`Invoice #: ${invoice.invoice}`, 40, cursorY);
    cursorY += lineHeight;
    doc.text(`Customer: ${invoice.customer}`, 40, cursorY);
    cursorY += lineHeight;
    doc.text(`Trip Reference: ${invoice.trip || 'N/A'}`, 40, cursorY);
    cursorY += lineHeight;
    doc.text(`Due Date: ${formatDate(invoice.dueDate)}`, 40, cursorY);
    cursorY += lineHeight;
    doc.text(`Status: ${invoice.status}`, 40, cursorY);
    cursorY += lineHeight * 1.5;

    doc.setFont('helvetica', 'bold');
    doc.text('Invoice Breakdown', 40, cursorY);
    cursorY += lineHeight;
    doc.setFont('helvetica', 'normal');
    const invoiceAmountValue = numericAmount(invoice.amount);
    doc.text(`Amount Due: ${formatCurrency(invoiceAmountValue)}`, 40, cursorY);
    cursorY += lineHeight;
    if (invoice.notes) {
      const textLines = doc.splitTextToSize(`Notes: ${invoice.notes}`, 500);
      doc.text(textLines, 40, cursorY);
      cursorY += lineHeight * textLines.length;
    }

    if (trip) {
      const tripChargeValue = numericAmount(trip.rentalCharges);
      if (tripChargeValue > 0) {
        doc.text(`Trip Rental Charges: ${formatCurrency(tripChargeValue)}`, 40, cursorY);
        cursorY += lineHeight;
      }
      cursorY += lineHeight;
      doc.setFont('helvetica', 'bold');
      doc.text('Trip Details', 40, cursorY);
      cursorY += lineHeight;
      doc.setFont('helvetica', 'normal');
      doc.text(`Trip Date: ${formatDate(trip.date)}`, 40, cursorY);
      cursorY += lineHeight;
      doc.text(`Vehicle: ${trip.vehicle}`, 40, cursorY);
      cursorY += lineHeight;
      doc.text(`Driver: ${trip.driver}`, 40, cursorY);
      cursorY += lineHeight;
      if (trip.route) {
        const routeLines = doc.splitTextToSize(`Route: ${trip.route}`, 500);
        doc.text(routeLines, 40, cursorY);
        cursorY += lineHeight * routeLines.length;
      }
    }

    cursorY += lineHeight * 1.5;
    doc.setFont('helvetica', 'italic');
    doc.text('Thank you for choosing Azmat for your transportation needs.', 40, cursorY);

    doc.save(`Azmat-Invoice-${invoice.invoice}.pdf`);
    return;
  }

  openInvoicePrintPreview(invoice, trip);
}

function openInvoicePrintPreview(invoice, trip) {
  const printWindow = window.open('', '_blank', 'width=900,height=700');
  if (!printWindow) {
    alert('Unable to open the print preview. Please allow pop-ups for this site.');
    return;
  }
  const markup = buildInvoicePrintMarkup(invoice, trip);
  printWindow.document.write(markup);
  printWindow.document.close();
  printWindow.focus();
  printWindow.print();
}

function buildInvoicePrintMarkup(invoice, trip) {
  const invoiceNumber = escapeHtml(invoice.invoice || '');
  const customerName = escapeHtml(invoice.customer || '');
  const dueDate = escapeHtml(formatDate(invoice.dueDate));
  const status = escapeHtml(invoice.status || '');
  const tripReference = invoice.trip || (trip ? trip.tripId : '');
  const safeTripReference = tripReference ? escapeHtml(tripReference) : '—';
  const invoiceAmountValue = numericAmount(invoice.amount);
  const invoiceAmount = escapeHtml(formatCurrency(invoiceAmountValue));
  const ownershipLabel = trip ? describeOwnership(getTripOwnership(trip)) : '';
  const ownershipMarkup = ownershipLabel ? `<p><strong>Unit Ownership</strong><br>${escapeHtml(ownershipLabel)}</p>` : '';
  const rentalChargeValue = trip ? numericAmount(trip.rentalCharges) : 0;
  const rentalMarkup = trip && rentalChargeValue > 0
    ? `<p><strong>Trip Rental Charges</strong><br>${escapeHtml(formatCurrency(rentalChargeValue))}</p>`
    : '';
  const routeMarkup = trip && trip.route
    ? `<p><strong>Route</strong><br>${escapeHtml(trip.route)}</p>`
    : '';
  const notesMarkup = invoice.notes
    ? `<section class="notes"><h2>Notes</h2><p>${escapeHtml(invoice.notes).replace(/\n/g, '<br>')}</p></section>`
    : '';
  const tripDetails = trip
    ? `<section class="details">
        <h2>Trip Details</h2>
        <p><strong>Date</strong><br>${escapeHtml(formatDate(trip.date))}</p>
        <p><strong>Vehicle</strong><br>${escapeHtml(trip.vehicle || '—')}</p>
        <p><strong>Driver</strong><br>${escapeHtml(trip.driver || '—')}</p>
        ${ownershipMarkup}
        ${rentalMarkup}
        ${routeMarkup}
      </section>`
    : '';
  const generatedOn = escapeHtml(new Date().toLocaleString());

  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <title>Invoice ${invoiceNumber}</title>
  <style>
    body { font-family: 'Inter', Arial, sans-serif; margin: 0; padding: 0; color: #1f2933; background: #f2f5fb; }
    .invoice { max-width: 720px; margin: 0 auto; padding: 40px 32px; background: #ffffff; }
    header { text-align: center; margin-bottom: 32px; }
    header h1 { margin-bottom: 8px; font-size: 26px; }
    .summary { display: flex; flex-wrap: wrap; gap: 24px; margin-bottom: 32px; }
    .summary > div { flex: 1 1 260px; background: #f4f7ff; border-radius: 12px; padding: 18px; }
    .summary h2 { margin-bottom: 12px; font-size: 16px; color: #0d47a1; }
    .summary p { margin: 0 0 10px; }
    .details { background: #f9fafc; border-radius: 12px; padding: 20px; margin-bottom: 24px; }
    .details h2 { margin-bottom: 12px; font-size: 16px; color: #0d47a1; }
    .details p { margin: 0 0 10px; }
    .notes { background: #fff9e6; border-radius: 12px; padding: 18px; margin-top: 16px; }
    .notes h2 { margin-bottom: 10px; font-size: 16px; color: #b7791f; }
    footer { margin-top: 32px; text-align: center; color: #5f6c7b; font-size: 0.95rem; }
    strong { color: #0d47a1; }
    @media print {
      body { background: #ffffff; }
      .invoice { box-shadow: none; }
    }
  </style>
</head>
<body>
  <div class="invoice">
    <header>
      <h1>Azmat Rental Fleet Management</h1>
      <p class="subtitle">Customer Invoice Summary</p>
    </header>
    <section class="summary">
      <div>
        <h2>Invoice Details</h2>
        <p><strong>Invoice #</strong><br>${invoiceNumber || '—'}</p>
        <p><strong>Customer</strong><br>${customerName || '—'}</p>
        <p><strong>Due Date</strong><br>${dueDate || '—'}</p>
        <p><strong>Status</strong><br>${status || '—'}</p>
      </div>
      <div>
        <h2>Financial Summary</h2>
        <p><strong>Trip Reference</strong><br>${safeTripReference}</p>
        <p><strong>Invoice Amount</strong><br>${invoiceAmount}</p>
      </div>
    </section>
    ${tripDetails}
    ${notesMarkup}
    <footer>
      <p>Generated on ${generatedOn}</p>
      <p>Thank you for choosing Azmat for your transportation needs.</p>
    </footer>
  </div>
</body>
</html>`;
}

function formatCurrency(value) {
  if (typeof value !== 'number' || Number.isNaN(value)) return '—';
  return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' }).format(value);
}

function formatDate(value) {
  if (!value) return '—';
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return date.toLocaleDateString();
}

function statusClassName(status) {
  if (!status) return '';
  const trimmed = String(status).trim();
  if (!trimmed) return '';
  return `status-${trimmed.replace(/\s+/g, '-')}`;
}

function escapeHtml(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
