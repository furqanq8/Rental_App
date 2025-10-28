const storageKey = 'azmat-fleet-state-v2';
const defaultState = {
  fleet: [],
  drivers: [],
  trips: [],
  invoices: [],
  suppliers: [],
  supplierDirectory: [],
  customers: [],
  nextTripNumber: 1
};

const state = loadState();
const selectRefs = {};
let tripIdInput;
let tripOwnershipField;
let fleetOwnershipSelect;
let driverAffiliationSelect;
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

function findVehicleOwnership(vehicleId, fleetList = state.fleet) {
  if (!vehicleId) return '';
  const source = Array.isArray(fleetList) ? fleetList : [];
  const match = source.find(item => item && item.unitId === vehicleId);
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
  }
  return JSON.parse(JSON.stringify(defaultState));
}

function normalizeState(parsed) {
  const base = JSON.parse(JSON.stringify(defaultState));
  if (parsed && typeof parsed === 'object') {
    Object.assign(base, parsed);
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

document.addEventListener('DOMContentLoaded', () => {
  dom.year.textContent = new Date().getFullYear();

  initializeModalDefaults();
  setupModals();
  setupForms();
  setupFilters();
  dom.exportBtn.addEventListener('click', exportWorkbook);

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
  const rentInVehicles = state.fleet.filter(item => item.ownership === 'rent-in').map(item => item.unitId).filter(Boolean);
  setSelectOptions(selectRefs.supplierPaymentVehicle, rentInVehicles, 'Select vehicle');

  const customers = [...state.customers].sort((a, b) => a.localeCompare(b));
  setSelectOptions(selectRefs.tripCustomer, customers, 'Select customer', { allowNew: true });
  setSelectOptions(selectRefs.invoiceCustomer, customers, 'Select customer', { allowNew: true });

  const tripReferences = state.trips.map(item => item.tripId).filter(Boolean);
  setSelectOptions(selectRefs.invoiceTrip, tripReferences, 'Select trip');
  setSelectOptions(selectRefs.supplierPaymentTrip, tripReferences, 'Select trip');

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
  const ownership = getVehicleOwnership(vehicleId);
  let drivers = state.drivers;
  if (ownership === 'rent-in') {
    drivers = drivers.filter(driver => driver && driver.affiliation === 'supplier');
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
      const entry = {
        unitId: formData.get('unitId'),
        fleetType: formData.get('fleetType'),
        ownership,
        model: formData.get('model'),
        capacity: formData.get('capacity'),
        status: formData.get('status'),
        supplier: ownership === 'rent-in' ? formData.get('supplier') : ''
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
      const supplierName = affiliation === 'supplier' ? formData.get('supplier') : '';
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
      persistState();
      renderTrips();
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
    supplierForm.addEventListener('submit', event => {
      event.preventDefault();
      const formData = new FormData(supplierForm);
      const entry = {
        reference: formData.get('reference'),
        supplier: formData.get('supplier'),
        vehicle: formData.get('vehicle'),
        trip: formData.get('trip'),
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
  dom.fleetFilter.addEventListener('change', renderFleet);
  dom.invoiceFilter.addEventListener('change', renderInvoices);
  dom.supplierFilter.addEventListener('change', renderSuppliers);
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
    setSelectValue(selectRefs.supplierPaymentVehicle, entry.vehicle);
  }
  if (selectRefs.supplierPaymentTrip) {
    setSelectValue(selectRefs.supplierPaymentTrip, entry.trip);
  }
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
  if (typeof JsPDFConstructor !== 'function') {
    alert('PDF library failed to load. Please check your internet connection and try again.');
    return;
  }
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
  doc.text(`Amount Due: ${formatCurrency(invoice.amount)}`, 40, cursorY);
  cursorY += lineHeight;
  if (invoice.notes) {
    const textLines = doc.splitTextToSize(`Notes: ${invoice.notes}`, 500);
    doc.text(textLines, 40, cursorY);
    cursorY += lineHeight * textLines.length;
  }

  if (trip) {
    const tripChargeValue = typeof trip.rentalCharges === 'number' ? trip.rentalCharges : Number(trip.rentalCharges);
    if (!Number.isNaN(tripChargeValue)) {
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
