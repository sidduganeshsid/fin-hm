const STORAGE_KEY = "management.finance.v7";

const DEFAULT_SHEET = {
  id: "sheet-main",
  name: "Main",
  created_at: new Date().toISOString(),
  entries: []
};

let workbook = loadWorkbook();
let currentSheetId = workbook.activeSheetId;
let entries = getCurrentSheet().entries;
let currentFilter = "ALL";
let currentPeriod = "ALL";
let currentArchiveScope = "ACTIVE";
let inlineMode = null;
let deferredInstallPrompt = null;

const $ = selector => document.querySelector(selector);

function nowISO() {
  return new Date().toISOString();
}

function today() {
  const d = new Date();
  const offset = d.getTimezoneOffset();
  return new Date(d.getTime() - offset * 60000).toISOString().slice(0, 10);
}

function createId() {
  return crypto.randomUUID
    ? crypto.randomUUID()
    : `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function loadWorkbook() {
  try {
    const current = JSON.parse(localStorage.getItem(STORAGE_KEY));
    if (current?.sheets?.length) {
      current.sheets.forEach(sheet => {
        sheet.entries = (sheet.entries || []).map(normalizeEntry);
      });
      current.activeSheetId ||= current.sheets[0].id;
      current.services ||= [];
      current.services = current.services.map(normalizeService).filter(Boolean);
      return current;
    }

    // Migrate the single-sheet prototype data.
    const old = JSON.parse(localStorage.getItem("home-management.finance.v1"));
    if (Array.isArray(old)) {
      const migrated = {
        version: 14,
        activeSheetId: DEFAULT_SHEET.id,
        services: [],
        sheets: [{
          ...DEFAULT_SHEET,
          entries: old.map(normalizeEntry)
        }]
      };
      localStorage.setItem(STORAGE_KEY, JSON.stringify(migrated, null, 2));
      return migrated;
    }
  } catch {}

  return {
    version: 14,
    activeSheetId: DEFAULT_SHEET.id,
    services: [],
    sheets: [{ ...DEFAULT_SHEET, entries: [] }]
  };
}

function normalizeEntry(entry) {
  const planned = entry.planned || entry.date || today();
  let spent = Object.prototype.hasOwnProperty.call(entry, "spent")
    ? entry.spent
    : (entry.status === "SPENT" ? (entry.date || today()) : null);

  return {
    id: entry.id || createId(),
    amount: Number(entry.amount) || 0,
    for_what: entry.for_what || "Miscellaneous",
    status: entry.status === "SPENT" ? "SPENT" : "PLANNED",
    payment_mode: entry.payment_mode || "ONLINE",
    archived: entry.archived === true,
    archived_at: entry.archived_at || null,
    planned,
    spent: spent || null,
    created_at: entry.created_at || nowISO(),
    updated_at: entry.updated_at || nowISO()
  };
}

function normalizeService(service) {
  if (!service) return null;

  const name = String(service.name || service.service_name || "").trim();
  const amount = Number(service.amount ?? service.service_amount);

  if (!name || !Number.isFinite(amount) || amount < 0) return null;

  return {
    id: service.id || createId(),
    name,
    amount,
    created_at: service.created_at || nowISO(),
    updated_at: service.updated_at || nowISO()
  };
}

function renderServices() {
  const container = $("#serviceButtons");
  if (!container) return;

  const services = workbook.services || [];

  const strip = container.closest(".services-strip");

  if (!services.length) {
    container.innerHTML = "";
    strip?.classList.add("hidden");
    return;
  }

  strip?.classList.remove("hidden");

  container.innerHTML = services.map(service => `
    <button class="service-button" data-service-id="${service.id}" title="Add ${escapeHtml(service.name)} as spent">
      ${escapeHtml(service.name)}
    </button>
  `).join("");
}

function addService(name, amount) {
  const normalized = normalizeService({ name, amount });
  if (!normalized) {
    showToast("Enter a service name and valid amount");
    return false;
  }

  if ((workbook.services || []).some(s => s.name.toLowerCase() === normalized.name.toLowerCase())) {
    showToast("That service already exists");
    return false;
  }

  workbook.services ||= [];
  workbook.services.push(normalized);
  persist();
  renderServices();
  return true;
}

function updateService(id, name, amount) {
  const service = workbook.services?.find(s => s.id === id);
  const normalized = normalizeService({ ...service, name, amount, id });

  if (!service || !normalized) {
    showToast("Enter a service name and valid amount");
    return false;
  }

  if (workbook.services.some(s =>
    s.id !== id && s.name.toLowerCase() === normalized.name.toLowerCase()
  )) {
    showToast("That service already exists");
    return false;
  }

  service.name = normalized.name;
  service.amount = normalized.amount;
  service.updated_at = nowISO();

  persist();
  renderServices();
  return true;
}

function deleteService(id) {
  const service = workbook.services?.find(s => s.id === id);
  if (!service) return;

  if (!confirm(`Delete service "${service.name}"? Existing finance records will stay unchanged.`)) {
    return;
  }

  workbook.services = workbook.services.filter(s => s.id !== id);
  persist();
  renderServices();
  renderServiceList();
  showToast("Service deleted");
}

function renderServiceList() {
  const list = $("#serviceList");
  if (!list) return;

  const services = workbook.services || [];

  if (!services.length) {
    list.innerHTML = '<div class="service-list-empty">No services yet. Add one above.</div>';
    return;
  }

  list.innerHTML = services.map(service => `
    <div class="service-row" data-service-row="${service.id}">
      <div class="service-row-name">${escapeHtml(service.name)}</div>
      <div class="service-row-amount">${formatAmount(service.amount)}</div>
      <div class="service-row-actions">
        <button class="service-action" data-service-edit="${service.id}" title="Edit service">✎</button>
        <button class="service-action danger" data-service-delete="${service.id}" title="Delete service">×</button>
      </div>
    </div>
  `).join("");
}

function openServiceDialog() {
  closeSheetMenus();
  $("#serviceDialog").classList.remove("hidden");
  renderServiceList();
  $("#serviceNameInput").value = "";
  $("#serviceAmountInput").value = "";
  requestAnimationFrame(() => $("#serviceNameInput").focus());
}

function closeServiceDialog() {
  $("#serviceDialog").classList.add("hidden");
}

function saveServiceFromDialog() {
  const name = $("#serviceNameInput").value.trim();
  const amount = Number($("#serviceAmountInput").value);

  if (addService(name, amount)) {
    $("#serviceNameInput").value = "";
    $("#serviceAmountInput").value = "";
    renderServiceList();
    $("#serviceNameInput").focus();
    showToast("Service added");
  }
}

function editServiceFromDialog(id) {
  const service = workbook.services?.find(s => s.id === id);
  if (!service) return;

  const name = prompt("Service name", service.name);
  if (name === null) return;

  const amountInput = prompt("Service amount", String(service.amount));
  if (amountInput === null) return;

  if (updateService(id, name.trim(), Number(amountInput))) {
    renderServiceList();
    showToast("Service updated");
  }
}

function createFinanceEntryFromService(id) {
  const service = workbook.services?.find(s => s.id === id);
  if (!service) return;

  const timestamp = nowISO();
  const todayDate = today();

  entries.push({
    id: createId(),
    amount: Number(service.amount),
    for_what: service.name,
    status: "SPENT",
    payment_mode: "ONLINE",
    planned: todayDate,
    spent: todayDate,
    created_at: timestamp,
    updated_at: timestamp
  });

  persist();

  // Always return to All/Today so the newly created record is immediately visible
  // regardless of the current combination of filters.
  currentFilter = "ALL";
  currentPeriod = "DAY";

  document.querySelectorAll(".filter").forEach(b => b.classList.remove("active"));
  document.querySelector('[data-filter="ALL"]')?.classList.add("active");

  document.querySelectorAll(".period-filter").forEach(b => b.classList.remove("active"));
  document.querySelector('[data-period="DAY"]')?.classList.add("active");

  inlineMode = null;
  render();
  showToast(`${service.name} added to finance`);
}

function getCurrentSheet() {
  let sheet = workbook.sheets.find(s => s.id === currentSheetId);
  if (!sheet) {
    sheet = workbook.sheets[0];
    currentSheetId = sheet.id;
    workbook.activeSheetId = sheet.id;
  }
  sheet.entries ||= [];
  return sheet;
}

function persist() {
  getCurrentSheet().entries = entries;
  workbook.activeSheetId = currentSheetId;
  localStorage.setItem(STORAGE_KEY, JSON.stringify(workbook, null, 2));
}

function formatAmount(value) {
  return new Intl.NumberFormat("en-IN", {
    style: "currency",
    currency: "INR",
    maximumFractionDigits: 2
  }).format(Number(value));
}

function formatDate(value) {
  if (!value) return "—";
  return new Intl.DateTimeFormat("en-IN", {
    day: "2-digit",
    month: "short",
    year: "numeric"
  }).format(new Date(value + (value.length === 10 ? "T00:00:00" : "")));
}

function formatCsvDate(value) {
  if (!value) return "";
  const [year, month, day] = String(value).slice(0, 10).split("-");
  return `${day}-${month}-${year}`;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function displayDate(entry) {
  return entry.status === "SPENT" ? entry.spent : entry.planned;
}

function formatPaymentMode(value) {
  const labels = {
    ONLINE: "Online",
    CASH: "Cash",
    UPI: "UPI",
    DEBIT_CARD: "Debit card",
    CREDIT_CARD: "Credit card",
    BANK_TRANSFER: "Bank transfer",
    OTHER: "Other"
  };
  return labels[value] || "Online";
}

function paymentOptions(selected = "ONLINE") {
  return [
    ["ONLINE", "Online"],
    ["CASH", "Cash"],
    ["UPI", "UPI"],
    ["DEBIT_CARD", "Debit card"],
    ["CREDIT_CARD", "Credit card"],
    ["BANK_TRANSFER", "Bank transfer"],
    ["OTHER", "Other"]
  ].map(([value, label]) =>
    `<option value="${value}" ${value === selected ? "selected" : ""}>${label}</option>`
  ).join("");
}

function localDateKey(date = new Date()) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function startOfWeek(date = new Date()) {
  const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  const day = d.getDay();
  d.setDate(d.getDate() + (day === 0 ? -6 : 1 - day));
  return d;
}

function endOfWeek(date = new Date()) {
  const d = startOfWeek(date);
  d.setDate(d.getDate() + 6);
  return d;
}

function getPeriodRange(period, reference = new Date()) {
  const d = new Date(reference.getFullYear(), reference.getMonth(), reference.getDate());

  if (period === "DAY") return { start: d, end: d };

  if (period === "WEEK") return { start: startOfWeek(d), end: endOfWeek(d) };

  if (period === "MONTH") {
    return {
      start: new Date(d.getFullYear(), d.getMonth(), 1),
      end: new Date(d.getFullYear(), d.getMonth() + 1, 0)
    };
  }

  if (period === "YEAR") {
    return {
      start: new Date(d.getFullYear(), 0, 1),
      end: new Date(d.getFullYear(), 11, 31)
    };
  }

  return null;
}

function matchesPeriod(entry, period) {
  if (period === "ALL") return true;
  const value = displayDate(entry);
  if (!value) return false;

  const [y, m, d] = value.slice(0, 10).split("-").map(Number);
  const entryDate = new Date(y, m - 1, d);
  const range = getPeriodRange(period);

  return entryDate >= range.start && entryDate <= range.end;
}

function formatPeriodContext(period) {
  const range = getPeriodRange(period);
  if (!range) return "";

  if (period === "DAY") return `Today · ${formatDate(localDateKey())}`;

  if (period === "WEEK") {
    return `This week · ${formatDate(localDateKey(range.start))} – ${formatDate(localDateKey(range.end))}`;
  }

  if (period === "MONTH") {
    return `This month · ${new Intl.DateTimeFormat("en-IN", { month: "long", year: "numeric" }).format(range.start)}`;
  }

  return `Annual · ${range.start.getFullYear()}`;
}

function updatePeriodContext() {
  const el = $("#periodContext");
  if (!el) return;
  el.textContent = formatPeriodContext(currentPeriod);
  el.classList.toggle("hidden", currentPeriod === "ALL");
}


function getVisibleActiveEntries() {
  return entries.filter(e => e.archived !== true)
    .filter(e => currentFilter === "ALL" || e.status === currentFilter)
    .filter(e => matchesPeriod(e, currentPeriod));
}

function archiveEntries(ids) {
  const set = new Set(ids);
  const timestamp = nowISO();
  let count = 0;
  entries.forEach(e => {
    if (set.has(e.id) && !e.archived) {
      e.archived = true;
      e.archived_at = timestamp;
      e.updated_at = timestamp;
      count++;
    }
  });
  if (!count) { showToast("No active records to archive"); return; }
  persist();
  render();
  showToast(`${count} record${count === 1 ? "" : "s"} archived`);
}

function archiveOne(id) {
  const e = entries.find(x => x.id === id);
  if (!e || e.archived) return;
  if (!confirm(`Archive "${e.for_what || "Miscellaneous"}" (${formatAmount(e.amount)})?`)) return;
  archiveEntries([id]);
}

function restoreEntry(id) {
  const e = entries.find(x => x.id === id);
  if (!e || !e.archived) return;
  e.archived = false;
  e.archived_at = null;
  e.updated_at = nowISO();
  persist();
  render();
  showToast("Record restored");
}

function openArchiveDialog() {
  if (currentArchiveScope === "ARCHIVED") return;
  const visible = getVisibleActiveEntries();
  if (!visible.length) { showToast("No active records in this view"); return; }
  const period = currentPeriod === "ALL" ? "all dates" :
    currentPeriod === "DAY" ? "today" :
    currentPeriod === "WEEK" ? "this week" :
    currentPeriod === "MONTH" ? "this month" : "this year";
  $("#archiveDialogText").textContent =
    `${visible.length} active record${visible.length === 1 ? "" : "s"} match the current view (${period}).`;
  $("#archiveDialog").classList.remove("hidden");
}

function closeArchiveDialog() { $("#archiveDialog").classList.add("hidden"); }

function archiveVisible() {
  const visible = getVisibleActiveEntries();
  closeArchiveDialog();
  archiveEntries(visible.map(e => e.id));
}

function archivePeriod() {
  if (currentPeriod === "ALL") {
    showToast("Choose Today, Week, Month, or Annual first");
    return;
  }
  const selected = getVisibleActiveEntries();
  closeArchiveDialog();
  archiveEntries(selected.map(e => e.id));
}

function updateConnectionStatus() {
  const el = $("#connectionStatus");
  if (!el) return;

  if (navigator.onLine) {
    el.classList.remove("offline", "syncing");
    el.textContent = "Online";
  } else {
    el.classList.remove("syncing");
    el.classList.add("offline");
    el.textContent = "Offline • Saved locally";
  }
}

window.addEventListener("online", () => {
  updateConnectionStatus();
  showToast("Back online");
});

window.addEventListener("offline", () => {
  updateConnectionStatus();
  showToast("Offline mode • changes are saved locally");
});

function render() {
  const body = $("#entryBody");

  const filtered = entries
    .filter(e => currentArchiveScope === "ARCHIVED" ? e.archived === true : e.archived !== true)
    .filter(e => currentFilter === "ALL" || e.status === currentFilter)
    .filter(e => matchesPeriod(e, currentPeriod))
    .sort((a, b) => new Date(displayDate(b) || 0) - new Date(displayDate(a) || 0));

  body.innerHTML = filtered.map(renderEntry).join("");

  if (inlineMode?.type === "add") {
    body.insertAdjacentHTML("afterbegin", renderAddRow());
  }

  if (inlineMode?.type === "edit") {
    const row = body.querySelector(`[data-entry-row="${inlineMode.id}"]`);
    if (row) row.outerHTML = renderEditRow(entries.find(e => e.id === inlineMode.id));
  }

  $("#emptyState").classList.toggle(
    "hidden",
    filtered.length > 0 || inlineMode?.type === "add"
  );

  renderSheetTabs();
  renderServices();
  updateSummary();
  updatePeriodContext();
  const archiveBtn = $("#archiveVisibleBtn");
  if (archiveBtn) {
    archiveBtn.disabled = currentArchiveScope === "ARCHIVED" || getVisibleActiveEntries().length === 0;
    archiveBtn.textContent = currentArchiveScope === "ARCHIVED" ? "Archived view" : "Archive visible";
  }
}


function archiveIcon() {
  return `<svg class="entry-action-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
    <path d="M4 7h16v13H4z"/><path d="M3 4h18v3H3z"/><path d="M9 11h6"/><path d="M9 15h6"/>
  </svg>`;
}

function restoreIcon() {
  return `<svg class="entry-action-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
    <path d="M9 7H4v5"/><path d="M4 12a8 8 0 1 0 2-5"/><path d="M4 12l4-4"/>
  </svg>`;
}

function archiveAllActiveRecords() {
  const timestamp = nowISO();
  let count = 0;

  workbook.sheets.forEach(sheet => {
    (sheet.entries || []).forEach(entry => {
      if (entry.archived !== true) {
        entry.archived = true;
        entry.archived_at = timestamp;
        entry.updated_at = timestamp;
        count++;
      }
    });
  });

  if (!count) {
    showToast("No active finance records to archive");
    return;
  }

  // Current `entries` references the active sheet, so its mutations are already
  // reflected. Persist writes every sheet from the workbook.
  persist();
  closeArchiveAllDialog();
  currentArchiveScope = "ARCHIVED";
  document.querySelectorAll(".archive-filter").forEach(b => b.classList.remove("active"));
  document.querySelector('[data-archive-scope="ARCHIVED"]')?.classList.add("active");
  render();
  showToast(`${count} finance record${count === 1 ? "" : "s"} archived across the app`);
}

function openArchiveAllDialog() {
  const total = workbook.sheets.reduce(
    (sum, sheet) => sum + (sheet.entries || []).filter(e => e.archived !== true).length,
    0
  );

  if (!total) {
    showToast("No active finance records to archive");
    return;
  }

  $("#archiveAllDialogText").textContent =
    `This will archive ${total} active finance record${total === 1 ? "" : "s"} across all ${workbook.sheets.length} sheet${workbook.sheets.length === 1 ? "" : "s"}.`;
  $("#archiveAllDialog").classList.remove("hidden");
}

function closeArchiveAllDialog() {
  $("#archiveAllDialog").classList.add("hidden");
}

function renderEntry(entry) {
  const spent = entry.status === "SPENT";
  const description = entry.for_what || "Miscellaneous";
  const archived = entry.archived === true;

  return `
    <tr data-entry-row="${entry.id}" class="${archived ? "is-archived" : ""}">
      <td class="amount">${formatAmount(entry.amount)}</td>
      <td class="description">
        ${escapeHtml(description)}
        ${archived ? '<span class="archive-badge">Archived</span>' : ""}
      </td>
      <td class="status-cell">
        <button class="toggle ${spent ? "on" : ""}"
          data-toggle="${entry.id}"
          title="${spent ? "Mark as planned" : "Mark as spent"}"
          aria-label="${spent ? "Mark as planned" : "Mark as spent"}"></button>
      </td>
      <td class="payment-cell">${escapeHtml(formatPaymentMode(entry.payment_mode))}</td>
      <td class="date">${formatDate(entry.planned)}</td>
      <td class="date">${formatDate(entry.spent)}</td>
      <td class="actions">
        ${archived
          ? `<button class="restore-entry-btn" data-restore-entry="${entry.id}" title="Restore record" aria-label="Restore record">${restoreIcon()}</button>`
          : `<button class="archive-entry-btn" data-archive-entry="${entry.id}" title="Archive record" aria-label="Archive record">${archiveIcon()}</button>`}
        <button class="menu-btn" data-menu="${entry.id}" aria-label="More options">⋮</button>
        <div class="menu hidden" id="menu-${entry.id}">
          <button data-action="edit" data-id="${entry.id}">Edit</button>
          <button data-action="toggle" data-id="${entry.id}">
            ${spent ? "Mark as planned" : "Mark as spent"}
          </button>
          ${archived
            ? `<button data-action="restore" data-id="${entry.id}">Restore</button>`
            : `<button data-action="archive" data-id="${entry.id}">Archive</button>`}
          <button class="danger" data-action="delete" data-id="${entry.id}">Delete</button>
        </div>
      </td>
    </tr>
  `;
}

function renderAddRow() {
  return `
    <tr class="inline-row" id="inline-add-row">
      <td><input class="inline-input inline-amount" id="newAmount" type="number" min="0" step="0.01" placeholder="Amount" autofocus></td>
      <td><input class="inline-input" id="newForWhat" type="text" maxlength="120" placeholder="For what? (optional)"></td>
      <td class="status-cell"><button class="toggle" id="newToggle" title="Planned"></button></td>
      <td><select class="inline-select" id="newPayment">${paymentOptions()}</select></td>
      <td><input class="inline-input inline-date" id="newPlanned" type="date" value="${today()}" title="Planned date"></td>
      <td>—</td>
      <td>
        <div class="inline-actions">
          <button class="save-mini" id="saveNew">Save</button>
          <button class="cancel-mini" id="cancelNew">Cancel</button>
        </div>
      </td>
    </tr>
  `;
}

function renderEditRow(entry) {
  if (!entry) return "";

  return `
    <tr class="inline-row" data-entry-row="${entry.id}">
      <td><input class="inline-input inline-amount" id="editAmount" type="number" min="0" step="0.01" value="${entry.amount}"></td>
      <td><input class="inline-input" id="editForWhat" type="text" maxlength="120" value="${escapeHtml(entry.for_what)}" placeholder="For what? (optional)"></td>
      <td class="status-cell">
        <button class="toggle ${entry.status === "SPENT" ? "on" : ""}" id="editToggle" title="Toggle spent"></button>
      </td>
      <td><select class="inline-select" id="editPayment">${paymentOptions(entry.payment_mode)}</select></td>
      <td><input class="inline-input inline-date" id="editPlanned" type="date" value="${entry.planned}"></td>
      <td>${formatDate(entry.spent)}</td>
      <td>
        <div class="inline-actions">
          <button class="save-mini" id="saveEdit">Save</button>
          <button class="cancel-mini" id="cancelEdit">Cancel</button>
        </div>
      </td>
    </tr>
  `;
}

function updateSummary() {
  const activeEntries = entries.filter(e => e.archived !== true);

  const planned = activeEntries
    .filter(e => e.status === "PLANNED")
    .reduce((sum, e) => sum + Number(e.amount), 0);

  const spent = activeEntries
    .filter(e => e.status === "SPENT")
    .reduce((sum, e) => sum + Number(e.amount), 0);

  $("#plannedTotal").textContent = formatAmount(planned);
  $("#spentTotal").textContent = formatAmount(spent);
}

function renderSheetTabs() {
  const tabs = $("#sheetTabs");
  tabs.innerHTML = "";

  workbook.sheets.forEach(sheet => {
    const tab = document.createElement("div");
    tab.className = `sheet-tab ${sheet.id === currentSheetId ? "active" : ""}`;
    tab.dataset.sheetId = sheet.id;

    const label = document.createElement("span");
    label.className = "sheet-tab-label";
    label.textContent = sheet.name;

    const menu = document.createElement("button");
    menu.className = "sheet-tab-menu";
    menu.dataset.sheetMenu = sheet.id;
    menu.title = "Sheet options";
    menu.textContent = "⋮";

    tab.append(label, menu);
    tabs.appendChild(tab);
  });
}

function startAdd() {
  if (inlineMode) return;
  closeRecordMenus();
  inlineMode = { type: "add" };
  render();
  requestAnimationFrame(() => $("#newAmount")?.focus());
}

function saveNew() {
  const amount = $("#newAmount").value;
  const forWhat = $("#newForWhat").value.trim() || "Miscellaneous";
  const planned = $("#newPlanned").value;
  const paymentMode = $("#newPayment").value;
  const spentNow = $("#newToggle").classList.contains("on");

  if (!amount || Number(amount) < 0 || !planned) {
    showToast("Enter amount and planned date");
    return;
  }

  const timestamp = nowISO();
  entries.push({
    id: createId(),
    amount: Number(amount),
    for_what: forWhat,
    status: spentNow ? "SPENT" : "PLANNED",
    payment_mode: paymentMode || "ONLINE",
    planned,
    spent: spentNow ? today() : null,
    created_at: timestamp,
    updated_at: timestamp
  });

  persist();
  inlineMode = null;
  render();
  showToast("Entry saved");
}

function startEdit(id) {
  closeRecordMenus();
  inlineMode = { type: "edit", id };
  render();
  requestAnimationFrame(() => $("#editAmount")?.focus());
}

function saveEdit() {
  const entry = entries.find(e => e.id === inlineMode?.id);
  if (!entry) return;

  const amount = $("#editAmount").value;
  const forWhat = $("#editForWhat").value.trim() || "Miscellaneous";
  const planned = $("#editPlanned").value;
  const paymentMode = $("#editPayment").value;
  const spentNow = $("#editToggle").classList.contains("on");
  const wasSpent = entry.status === "SPENT";

  if (!amount || Number(amount) < 0 || !planned) {
    showToast("Enter amount and planned date");
    return;
  }

  entry.amount = Number(amount);
  entry.for_what = forWhat;
  entry.payment_mode = paymentMode || "ONLINE";
  entry.planned = planned;

  if (spentNow && !wasSpent) {
    entry.status = "SPENT";
    entry.spent = today();
  } else if (!spentNow && wasSpent) {
    entry.status = "PLANNED";
    entry.spent = null;
  } else {
    entry.status = spentNow ? "SPENT" : "PLANNED";
  }

  entry.updated_at = nowISO();

  persist();
  inlineMode = null;
  render();
  showToast("Changes saved");
}

function toggleStatus(id) {
  if (inlineMode) return;

  const entry = entries.find(e => e.id === id);
  if (!entry) return;

  if (entry.status === "SPENT") {
    entry.status = "PLANNED";
    entry.spent = null;
    showToast("Marked as planned");
  } else {
    entry.status = "SPENT";
    entry.spent = today();
    showToast("Marked as spent");
  }

  entry.updated_at = nowISO();
  persist();
  render();
}

function deleteEntry(id) {
  closeRecordMenus();
  const entry = entries.find(e => e.id === id);
  if (!entry) return;

  if (!confirm(`Delete "${entry.for_what || "Miscellaneous"}"?`)) return;

  entries = entries.filter(e => e.id !== id);
  persist();
  render();
  showToast("Entry deleted");
}

function closeInline() {
  inlineMode = null;
  render();
}

function positionFixedMenu(menu, anchor) {
  const rect = anchor.getBoundingClientRect();
  menu.classList.remove("hidden");
  const menuRect = menu.getBoundingClientRect();
  const margin = 8;

  let top = rect.bottom + 6;
  if (top + menuRect.height > window.innerHeight - margin) {
    top = rect.top - menuRect.height - 6;
  }

  let left = rect.right - menuRect.width;
  top = Math.max(margin, Math.min(top, window.innerHeight - menuRect.height - margin));
  left = Math.max(margin, Math.min(left, window.innerWidth - menuRect.width - margin));

  menu.style.top = `${top}px`;
  menu.style.left = `${left}px`;
}

function closeRecordMenus() {
  document.querySelectorAll(".menu").forEach(menu => {
    menu.classList.add("hidden");
    menu.style.top = "";
    menu.style.left = "";
  });
}

function closeSheetMenus() {
  document.querySelectorAll(".sheet-menu").forEach(menu => menu.remove());
}

function openSheetMenu(button, sheetId) {
  closeSheetMenus();

  const menu = document.createElement("div");
  menu.className = "sheet-menu";
  menu.innerHTML = `
    <button data-sheet-action="rename" data-sheet-id="${sheetId}">Rename</button>
    <button data-sheet-action="duplicate" data-sheet-id="${sheetId}">Duplicate</button>
    <button class="danger" data-sheet-action="delete" data-sheet-id="${sheetId}">Delete</button>
  `;
  document.body.appendChild(menu);
  positionFixedMenu(menu, button);
}

function switchSheet(id) {
  const sheet = workbook.sheets.find(s => s.id === id);
  if (!sheet) return;

  closeRecordMenus();
  closeSheetMenus();
  inlineMode = null;
  currentSheetId = id;
  entries = sheet.entries || [];
  currentFilter = "ALL";
  currentPeriod = "ALL";
  currentArchiveScope = "ACTIVE";

  document.querySelectorAll(".filter").forEach(b => b.classList.remove("active"));
  document.querySelector('[data-filter="ALL"]')?.classList.add("active");

  persist();
  render();
}

function createSheet() {
  const name = $("#sheetNameInput").value.trim();
  if (!name) {
    showToast("Enter a sheet name");
    return;
  }

  if (workbook.sheets.some(s => s.name.toLowerCase() === name.toLowerCase())) {
    showToast("A sheet with that name already exists");
    return;
  }

  const sheet = {
    id: createId(),
    name,
    created_at: nowISO(),
    entries: []
  };

  workbook.sheets.push(sheet);
  currentSheetId = sheet.id;
  entries = sheet.entries;
  inlineMode = null;
  currentFilter = "ALL";
  persist();
  closeSheetDialog();
  render();
  showToast("Sheet created");
}

function renameSheet(id) {
  const sheet = workbook.sheets.find(s => s.id === id);
  if (!sheet) return;

  const name = prompt("Rename sheet", sheet.name);
  if (name === null) return;

  const trimmed = name.trim();
  if (!trimmed) return;

  if (workbook.sheets.some(s => s.id !== id && s.name.toLowerCase() === trimmed.toLowerCase())) {
    showToast("A sheet with that name already exists");
    return;
  }

  sheet.name = trimmed;
  persist();
  closeSheetMenus();
  render();
  showToast("Sheet renamed");
}

function duplicateSheet(id) {
  const source = workbook.sheets.find(s => s.id === id);
  if (!source) return;

  let name = `${source.name} Copy`;
  let n = 2;
  while (workbook.sheets.some(s => s.name.toLowerCase() === name.toLowerCase())) {
    name = `${source.name} Copy ${n++}`;
  }

  const copy = {
    id: createId(),
    name,
    created_at: nowISO(),
    entries: structuredClone(source.entries || []).map(e => ({ ...e, id: createId() }))
  };

  workbook.sheets.push(copy);
  currentSheetId = copy.id;
  entries = copy.entries;
  inlineMode = null;
  currentFilter = "ALL";
  persist();
  closeSheetMenus();
  render();
  showToast("Sheet duplicated");
}

function deleteSheet(id) {
  if (workbook.sheets.length === 1) {
    showToast("Keep at least one sheet");
    return;
  }

  const sheet = workbook.sheets.find(s => s.id === id);
  if (!sheet) return;

  if (!confirm(`Delete sheet "${sheet.name}" and its records?`)) return;

  workbook.sheets = workbook.sheets.filter(s => s.id !== id);

  if (currentSheetId === id) {
    currentSheetId = workbook.sheets[0].id;
    entries = workbook.sheets[0].entries;
  }

  persist();
  closeSheetMenus();
  render();
  showToast("Sheet deleted");
}

function openSheetDialog() {
  $("#sheetNameInput").value = "";
  $("#sheetDialog").classList.remove("hidden");
  requestAnimationFrame(() => $("#sheetNameInput").focus());
}

function closeSheetDialog() {
  $("#sheetDialog").classList.add("hidden");
}

function csvEscape(value) {
  const text = String(value ?? "");
  return `"${text.replaceAll('"', '""')}"`;
}

function exportCsv() {
  const headers = ["amount", "for_what", "status", "payment_mode", "planned", "spent", "archived", "archived_at"];
  const rows = entries.map(e => [
    e.amount,
    e.for_what,
    e.status.toLowerCase(),
    formatPaymentMode(e.payment_mode),
    formatCsvDate(e.planned),
    formatCsvDate(e.spent),
    e.archived ? "yes" : "no",
    e.archived_at || ""
  ]);

  const csv = [
    headers.join(","),
    ...rows.map(row => row.map(csvEscape).join(","))
  ].join("\r\n");

  downloadBlob("\uFEFF" + csv, "finance_entries.csv", "text/csv;charset=utf-8;");
  showToast("CSV exported");
}

function exportJson() {
  downloadBlob(
    JSON.stringify(workbook, null, 2),
    "management_finance.json",
    "application/json"
  );
  showToast("JSON exported");
}

function downloadBlob(content, filename, type) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

function showToast(message) {
  const toast = $("#toast");
  toast.textContent = message;
  toast.classList.add("show");
  clearTimeout(showToast.timer);
  showToast.timer = setTimeout(() => toast.classList.remove("show"), 1600);
}

// Record actions.
document.addEventListener("click", event => {
  if (event.target.closest("#addBtn")) {
    startAdd();
    return;
  }

  if (event.target.closest("#saveNew")) {
    saveNew();
    return;
  }

  if (event.target.closest("#cancelNew")) {
    closeInline();
    return;
  }

  if (event.target.closest("#saveEdit")) {
    saveEdit();
    return;
  }

  if (event.target.closest("#cancelEdit")) {
    closeInline();
    return;
  }

  const toggle = event.target.closest("[data-toggle]");
  if (toggle) {
    toggleStatus(toggle.dataset.toggle);
    return;
  }

  if (event.target.closest("#newToggle")) {
    $("#newToggle").classList.toggle("on");
    return;
  }

  if (event.target.closest("#editToggle")) {
    $("#editToggle").classList.toggle("on");
    return;
  }

  const menuButton = event.target.closest("[data-menu]");
  if (menuButton) {
    const menu = $(`#menu-${menuButton.dataset.menu}`);
    const wasHidden = menu.classList.contains("hidden");
    closeRecordMenus();
    if (wasHidden) positionFixedMenu(menu, menuButton);
    return;
  }

  const archiveButton = event.target.closest("[data-archive-entry]");
  if (archiveButton) {
    closeRecordMenus();
    archiveOne(archiveButton.dataset.archiveEntry);
    return;
  }

  const restoreButton = event.target.closest("[data-restore-entry]");
  if (restoreButton) {
    closeRecordMenus();
    restoreEntry(restoreButton.dataset.restoreEntry);
    return;
  }

  const action = event.target.closest("[data-action]");
  if (action) {
    const id = action.dataset.id;
    if (action.dataset.action === "edit") startEdit(id);
    if (action.dataset.action === "toggle") toggleStatus(id);
    if (action.dataset.action === "archive") archiveOne(id);
    if (action.dataset.action === "restore") restoreEntry(id);
    if (action.dataset.action === "delete") deleteEntry(id);
    return;
  }

  // Sheet actions.
  const sheetMenuButton = event.target.closest("[data-sheet-menu]");
  if (sheetMenuButton) {
    event.stopPropagation();
    openSheetMenu(sheetMenuButton, sheetMenuButton.dataset.sheetMenu);
    return;
  }

  const sheetTab = event.target.closest(".sheet-tab");
  if (sheetTab) {
    switchSheet(sheetTab.dataset.sheetId);
    return;
  }

  const sheetAction = event.target.closest("[data-sheet-action]");
  if (sheetAction) {
    const id = sheetAction.dataset.sheetId;
    const action = sheetAction.dataset.sheetAction;
    if (action === "rename") renameSheet(id);
    if (action === "duplicate") duplicateSheet(id);
    if (action === "delete") deleteSheet(id);
    return;
  }

  if (!event.target.closest(".menu")) closeRecordMenus();
  if (!event.target.closest(".sheet-menu")) closeSheetMenus();
});

// Filters.
document.querySelectorAll(".filter").forEach(button => {
  button.addEventListener("click", () => {
    currentFilter = button.dataset.filter;
    document.querySelectorAll(".filter").forEach(b => b.classList.remove("active"));
    button.classList.add("active");
    closeRecordMenus();
    render();
  });
});


document.querySelectorAll(".archive-filter").forEach(button => {
  button.addEventListener("click", () => {
    currentArchiveScope = button.dataset.archiveScope;
    document.querySelectorAll(".archive-filter").forEach(b => b.classList.remove("active"));
    button.classList.add("active");
    closeRecordMenus();
    render();
  });
});

$("#archiveVisibleBtn").addEventListener("click", openArchiveDialog);
$("#closeArchiveDialog").addEventListener("click", closeArchiveDialog);
$("#archiveCancel").addEventListener("click", closeArchiveDialog);
$("#archiveVisibleConfirm").addEventListener("click", archiveVisible);
$("#archivePeriodConfirm").addEventListener("click", archivePeriod);
$("#archiveDialog").addEventListener("click", e => {
  if (e.target === $("#archiveDialog")) closeArchiveDialog();
});

// Period filters: Today, current week, current month, current year.
document.querySelectorAll(".period-filter").forEach(button => {
  button.addEventListener("click", () => {
    currentPeriod = button.dataset.period;
    document.querySelectorAll(".period-filter").forEach(b => b.classList.remove("active"));
    button.classList.add("active");
    closeRecordMenus();
    render();
  });
});

// Services.
$("#settingsBtn").addEventListener("click", openServiceDialog);
$("#closeServiceDialog").addEventListener("click", closeServiceDialog);
$("#saveServiceBtn").addEventListener("click", saveServiceFromDialog);

$("#serviceDialog").addEventListener("click", event => {
  if (event.target === $("#serviceDialog")) closeServiceDialog();
});

$("#serviceButtons").addEventListener("click", event => {
  const button = event.target.closest("[data-service-id]");
  if (!button) return;
  createFinanceEntryFromService(button.dataset.serviceId);
});

$("#serviceList").addEventListener("click", event => {
  const edit = event.target.closest("[data-service-edit]");
  const remove = event.target.closest("[data-service-delete]");

  if (edit) {
    editServiceFromDialog(edit.dataset.serviceEdit);
    return;
  }

  if (remove) {
    deleteService(remove.dataset.serviceDelete);
  }
});

document.addEventListener("keydown", event => {
  if ($("#serviceDialog") && !$("#serviceDialog").classList.contains("hidden")) {
    if (event.key === "Escape") closeServiceDialog();

    if (event.key === "Enter" &&
        (event.target.id === "serviceNameInput" || event.target.id === "serviceAmountInput")) {
      saveServiceFromDialog();
    }
  }
});


$("#archiveAllBtn").addEventListener("click", openArchiveAllDialog);
$("#closeArchiveAllDialog").addEventListener("click", closeArchiveAllDialog);
$("#archiveAllCancel").addEventListener("click", closeArchiveAllDialog);
$("#archiveAllConfirm").addEventListener("click", archiveAllActiveRecords);

$("#archiveAllDialog").addEventListener("click", event => {
  if (event.target === $("#archiveAllDialog")) closeArchiveAllDialog();
});

// Exports.
$("#exportCsvBtn").addEventListener("click", exportCsv);
$("#exportJsonBtn").addEventListener("click", exportJson);

// Sheets.
$("#addSheetBtn").addEventListener("click", openSheetDialog);
$("#closeSheetDialog").addEventListener("click", closeSheetDialog);
$("#cancelSheetDialog").addEventListener("click", closeSheetDialog);
$("#saveSheetDialog").addEventListener("click", createSheet);

$("#sheetDialog").addEventListener("click", event => {
  if (event.target === $("#sheetDialog")) closeSheetDialog();
});

// Keyboard shortcuts for inline records.
document.addEventListener("keydown", event => {
  if (event.key === "Escape") {
    if (!$("#sheetDialog").classList.contains("hidden")) closeSheetDialog();
    else if (inlineMode) closeInline();
  }

  if (event.key === "Enter") {
    if (event.target.id === "sheetNameInput") createSheet();

    if (["newAmount", "newForWhat", "newPayment", "newPlanned"].includes(event.target.id)) {
      saveNew();
    }

    if (["editAmount", "editForWhat", "editPayment", "editPlanned"].includes(event.target.id)) {
      saveEdit();
    }
  }
});

// PWA installation.
window.addEventListener("beforeinstallprompt", event => {
  event.preventDefault();
  deferredInstallPrompt = event;
  $("#installBtn").classList.remove("hidden");
});

$("#installBtn").addEventListener("click", async () => {
  if (!deferredInstallPrompt) return;
  deferredInstallPrompt.prompt();
  await deferredInstallPrompt.userChoice;
  deferredInstallPrompt = null;
  $("#installBtn").classList.add("hidden");
});

window.addEventListener("appinstalled", () => {
  deferredInstallPrompt = null;
  $("#installBtn").classList.add("hidden");
  showToast("App installed");
});

window.addEventListener("scroll", () => {
  closeRecordMenus();
  closeSheetMenus();
}, { passive: true });

window.addEventListener("resize", () => {
  closeRecordMenus();
  closeSheetMenus();
});

updateConnectionStatus();
render();
