const storageKey = 'azmat-fleet-state-v1';
const defaultState = {
  fleet: [],
  drivers: [],
  trips: [],
  invoices: [],
  suppliers: []
};

const state = loadState();

function loadState() {
  try {
    const saved = localStorage.getItem(storageKey);
    if (saved) {
      const parsed = JSON.parse(saved);
      return { ...defaultState, ...parsed };
    }
  } catch (error) {
    console.warn('Unable to load saved data. Starting fresh.', error);
  }
  return JSON.parse(JSON.stringify(defaultState));
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
};

document.addEventListener('DOMContentLoaded', () => {
  dom.year.textContent = new Date().getFullYear();

  setupModals();
  setupForms();
  setupFilters();
  dom.exportBtn.addEventListener('click', exportWorkbook);

  renderAll();
});

function setupModals() {
  dom.cards.forEach(card => {
    card.addEventListener('click', () => openModal(card.dataset.modal));
    const button = card.querySelector('button');
    if (button) {
      button.addEventListener('click', evt => {
        evt.stopPropagation();
        openModal(card.dataset.modal);
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
  }
}

function setupForms() {
  document.getElementById('fleetForm').addEventListener('submit', event => {
    event.preventDefault();
    const formData = new FormData(event.target);
    const entry = {
      unitId: formData.get('unitId'),
      ownership: formData.get('ownership'),
      model: formData.get('model'),
      capacity: formData.get('capacity'),
      status: formData.get('status'),
      supplier: formData.get('ownership') === 'rent-in' ? formData.get('supplier') : ''
    };
    state.fleet.unshift(entry);
    persistState();
    renderFleet();
    closeModal('fleetModal');
  });

  document.getElementById('driverForm').addEventListener('submit', event => {
    event.preventDefault();
    const formData = new FormData(event.target);
    const entry = {
      name: formData.get('name'),
      license: formData.get('license'),
      phone: formData.get('phone'),
      availability: formData.get('availability')
    };
    state.drivers.unshift(entry);
    persistState();
    renderDrivers();
    closeModal('driverModal');
  });

  document.getElementById('tripForm').addEventListener('submit', event => {
    event.preventDefault();
    const formData = new FormData(event.target);
    const entry = {
      tripId: formData.get('tripId'),
      date: formData.get('date'),
      customer: formData.get('customer'),
      vehicle: formData.get('vehicle'),
      driver: formData.get('driver'),
      route: formData.get('route'),
      status: formData.get('status')
    };
    state.trips.unshift(entry);
    persistState();
    renderTrips();
    closeModal('tripModal');
  });

  document.getElementById('invoiceForm').addEventListener('submit', event => {
    event.preventDefault();
    const formData = new FormData(event.target);
    const entry = {
      invoice: formData.get('invoice'),
      customer: formData.get('customer'),
      trip: formData.get('trip'),
      amount: Number(formData.get('amount')),
      dueDate: formData.get('dueDate'),
      status: formData.get('status'),
      notes: formData.get('notes')
    };
    state.invoices.unshift(entry);
    persistState();
    renderInvoices();
    closeModal('invoiceModal');
  });

  document.getElementById('supplierForm').addEventListener('submit', event => {
    event.preventDefault();
    const formData = new FormData(event.target);
    const entry = {
      reference: formData.get('reference'),
      supplier: formData.get('supplier'),
      vehicle: formData.get('vehicle'),
      amount: Number(formData.get('amount')),
      dueDate: formData.get('dueDate'),
      status: formData.get('status'),
      notes: formData.get('notes')
    };
    state.suppliers.unshift(entry);
    persistState();
    renderSuppliers();
    closeModal('supplierModal');
  });
}

function setupFilters() {
  dom.fleetFilter.addEventListener('change', renderFleet);
  dom.invoiceFilter.addEventListener('change', renderInvoices);
  dom.supplierFilter.addEventListener('change', renderSuppliers);
}

function renderAll() {
  renderFleet();
  renderDrivers();
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
      return `<tr>
        <td data-label="Unit ID">${escapeHtml(item.unitId)}</td>
        <td data-label="Type"><span class="status-pill ${item.ownership === 'owned' ? 'status-Available' : 'status-Scheduled'}">${item.ownership === 'owned' ? 'Company Owned' : 'Rent-In'}</span></td>
        <td data-label="Make & Model">${escapeHtml(item.model)}</td>
        <td data-label="Capacity">${escapeHtml(item.capacity || '')}</td>
        <td data-label="Status"><span class="status-pill ${statusClass}">${escapeHtml(item.status)}</span></td>
        <td data-label="Supplier">${escapeHtml(item.supplier || '-')}</td>
      </tr>`;
    })
    .join('');
}

function renderDrivers() {
  dom.driverTable.innerHTML = state.drivers
    .map(item => `<tr>
      <td data-label="Name">${escapeHtml(item.name)}</td>
      <td data-label="License No.">${escapeHtml(item.license)}</td>
      <td data-label="Phone"><a href="tel:${escapeHtml(item.phone)}">${escapeHtml(item.phone)}</a></td>
      <td data-label="Availability"><span class="status-pill ${statusClassName(item.availability)}">${escapeHtml(item.availability)}</span></td>
    </tr>`)
    .join('');
}

function renderTrips() {
  dom.tripTable.innerHTML = state.trips
    .map(item => `<tr>
      <td data-label="Trip ID">${escapeHtml(item.tripId)}</td>
      <td data-label="Date">${formatDate(item.date)}</td>
      <td data-label="Customer">${escapeHtml(item.customer)}</td>
      <td data-label="Vehicle">${escapeHtml(item.vehicle)}</td>
      <td data-label="Driver">${escapeHtml(item.driver)}</td>
      <td data-label="Route">${escapeHtml(item.route || '')}</td>
      <td data-label="Status"><span class="status-pill ${statusClassName(item.status)}">${escapeHtml(item.status)}</span></td>
    </tr>`)
    .join('');
}

function renderInvoices() {
  const filter = dom.invoiceFilter.value;
  const rows = state.invoices.filter(item => filter === 'all' || item.status === filter);
  dom.invoiceTable.innerHTML = rows
    .map(item => `<tr>
      <td data-label="Invoice #">${escapeHtml(item.invoice)}</td>
      <td data-label="Customer">${escapeHtml(item.customer)}</td>
      <td data-label="Trip">${escapeHtml(item.trip || '-')}</td>
      <td data-label="Amount">${formatCurrency(item.amount)}</td>
      <td data-label="Due Date">${formatDate(item.dueDate)}</td>
      <td data-label="Status"><span class="status-pill ${statusClassName(item.status)}">${escapeHtml(item.status)}</span></td>
      <td data-label="Actions"><button class="btn secondary" data-print="${encodeURIComponent(item.invoice)}">Print PDF</button></td>
    </tr>`)
    .join('');

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
    .map(item => `<tr>
      <td data-label="Reference">${escapeHtml(item.reference)}</td>
      <td data-label="Supplier">${escapeHtml(item.supplier)}</td>
      <td data-label="Vehicle">${escapeHtml(item.vehicle || '-')}</td>
      <td data-label="Amount">${formatCurrency(item.amount)}</td>
      <td data-label="Due Date">${formatDate(item.dueDate)}</td>
      <td data-label="Status"><span class="status-pill ${statusClassName(item.status)}">${escapeHtml(item.status)}</span></td>
    </tr>`)
    .join('');

  dom.supplierReport.innerHTML = buildSupplierReport();
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
  XLSX.writeFile(wb, `Azmat-fleet-data-${new Date().toISOString().slice(0, 10)}.xlsx`);
}

function printInvoice(invoiceId) {
  const invoice = state.invoices.find(item => item.invoice === invoiceId);
  if (!invoice) return;

  const trip = state.trips.find(t => t.tripId === invoice.trip);
  const jsPdfNamespace = window.jspdf;
  if (!jsPdfNamespace) {
    alert('PDF library failed to load. Please check your internet connection and try again.');
    return;
  }
  const { jsPDF } = jsPdfNamespace;
  const doc = new jsPDF({ unit: 'pt', format: 'a4' });
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
