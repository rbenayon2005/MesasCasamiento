const SHEET_NAME = "Invitados Ariel Casamiento";
const START_ROW = 8;

const state = {
  guests: [],
  tables: [],
  tableOrder: [],
  dragGuestId: null,
  dragTableId: null,
  dragType: null,
  filter: {
    search: "",
    gender: "all",
  },
  remote: {
    available: false,
    saveTimer: null,
    saveInFlight: false,
    revision: 0,
    poller: null,
  },
  auth: {
    user: "",
    pass: "",
  },
};

const refs = {
  fileInput: document.getElementById("excelFile"),
  csvInput: document.getElementById("csvFile"),
  addGuestBtn: document.getElementById("addGuestBtn"),
  addOneTableBtn: document.getElementById("addOneTableBtn"),
  exportBtn: document.getElementById("exportBtn"),
  searchInput: document.getElementById("searchInput"),
  genderFilter: document.getElementById("genderFilter"),
  unassignedList: document.getElementById("unassignedList"),
  tablesGrid: document.getElementById("tablesGrid"),
  stats: document.getElementById("stats"),
  toast: document.getElementById("toast"),
  authModal: document.getElementById("authModal"),
  authForm: document.getElementById("authForm"),
  authUser: document.getElementById("authUser"),
  authPass: document.getElementById("authPass"),
  authError: document.getElementById("authError"),
  guestModal: document.getElementById("guestModal"),
  guestForm: document.getElementById("guestForm"),
  guestFirstName: document.getElementById("guestFirstName"),
  guestLastName: document.getElementById("guestLastName"),
  guestGender: document.getElementById("guestGender"),
  guestError: document.getElementById("guestError"),
  guestCancel: document.getElementById("guestCancel"),
};

async function apiRequest(path, options = {}) {
  const authHeaders = {};
  if (state.auth.user && state.auth.pass) {
    authHeaders["x-auth-user"] = state.auth.user;
    authHeaders["x-auth-pass"] = state.auth.pass;
  }
  const response = await fetch(path, {
    headers: { "Content-Type": "application/json", ...authHeaders, ...(options.headers || {}) },
    ...options,
  });
  if (!response.ok) {
    const text = await response.text();
    const err = new Error(text || `HTTP ${response.status}`);
    err.status = response.status;
    throw err;
  }
  const contentType = response.headers.get("content-type") || "";
  if (contentType.includes("application/json")) {
    return response.json();
  }
  return null;
}

function loadAuthFromStorage() {
  state.auth.user = localStorage.getItem("mesas_auth_user") || "";
  state.auth.pass = localStorage.getItem("mesas_auth_pass") || "";
}

function saveAuthToStorage() {
  localStorage.setItem("mesas_auth_user", state.auth.user);
  localStorage.setItem("mesas_auth_pass", state.auth.pass);
}

function clearAuthInStorage() {
  state.auth.user = "";
  state.auth.pass = "";
  localStorage.removeItem("mesas_auth_user");
  localStorage.removeItem("mesas_auth_pass");
}

function askCredentials() {
  return new Promise((resolve) => {
    refs.authError.classList.add("hidden");
    refs.authUser.value = state.auth.user || "";
    refs.authPass.value = "";
    refs.authModal.classList.remove("hidden");
    refs.authUser.focus();

    const submitHandler = (e) => {
      e.preventDefault();
      const user = refs.authUser.value.trim();
      const pass = refs.authPass.value.trim();
      if (!user || !pass) {
        refs.authError.textContent = "Usuario y password requeridos.";
        refs.authError.classList.remove("hidden");
        return;
      }
      refs.authModal.classList.add("hidden");
      refs.authForm.removeEventListener("submit", submitHandler);
      state.auth.user = user;
      state.auth.pass = pass;
      saveAuthToStorage();
      resolve(true);
    };

    refs.authForm.addEventListener("submit", submitHandler);
  });
}

async function ensureAuthSession() {
  loadAuthFromStorage();
  for (let i = 0; i < 3; i += 1) {
    if (!state.auth.user || !state.auth.pass) {
      const entered = await askCredentials();
      if (!entered) return false;
    }
    try {
      await apiRequest("/api/state?revision=0");
      return true;
    } catch (err) {
      if (err.status !== 401) return true;
      clearAuthInStorage();
      refs.authError.textContent = "Credenciales invalidas.";
      refs.authError.classList.remove("hidden");
    }
  }
  return false;
}

function serializeStateForRemote() {
  const tables = state.tables
    .slice()
    .sort((a, b) => a.number - b.number)
    .map((t) => ({
      number: t.number,
      name: t.name,
      type: t.type || null,
      capacity: t.capacity,
    }));

  const guests = state.guests.map((g) => ({
    id: g.id,
    name: g.name,
    gender: g.gender,
    confirmed: !!g.confirmed,
    sourceRow: Number.isFinite(g.sourceRow) ? g.sourceRow : null,
    tableId: g.tableId || null,
  }));

  return {
    guests,
    tables,
    tableOrder: state.tableOrder.slice(),
  };
}

function applyRemoteSnapshot(snapshot) {
  if (!snapshot || !Array.isArray(snapshot.guests) || !Array.isArray(snapshot.tables)) return false;
  state.guests = snapshot.guests.map((g) => ({
    id: normalize(g.id) || crypto.randomUUID(),
    name: normalize(g.name),
    gender: normalize(g.gender).toUpperCase() === "M" ? "M" : "H",
    confirmed: !!g.confirmed,
    sourceRow: Number.isFinite(Number(g.sourceRow)) ? Number(g.sourceRow) : null,
    tableId: normalize(g.tableId) || null,
  }));
  state.tables = snapshot.tables
    .map((t) => ({
      id: `t-${Number(t.number)}`,
      number: Number(t.number),
      name: normalize(t.name) || `Mesa ${Number(t.number)}`,
      type: ["men", "women"].includes(t.type) ? t.type : null,
      capacity: Number.isFinite(Number(t.capacity)) ? Number(t.capacity) : Number(t.number) === 1 ? 20 : 10,
    }))
    .filter((t) => Number.isFinite(t.number) && t.number > 0)
    .sort((a, b) => a.number - b.number);

  state.tableOrder = Array.isArray(snapshot.tableOrder) ? snapshot.tableOrder.map((id) => normalize(id)).filter(Boolean) : [];
  syncTableOrder();
  return true;
}

async function saveSnapshotNow() {
  if (!state.remote.available || state.remote.saveInFlight) return;
  state.remote.saveInFlight = true;
  try {
    const response = await apiRequest("/api/state", {
      method: "POST",
      body: JSON.stringify(serializeStateForRemote()),
    });
    if (response && Number.isFinite(Number(response.revision))) {
      state.remote.revision = Number(response.revision);
    }
  } catch (err) {
    showToast("No se pudo guardar en la nube.");
  } finally {
    state.remote.saveInFlight = false;
  }
}

function scheduleRemoteSave(delayMs = 450) {
  if (!state.remote.available) return;
  clearTimeout(state.remote.saveTimer);
  state.remote.saveTimer = setTimeout(() => {
    saveSnapshotNow();
  }, delayMs);
}

async function loadRemoteSnapshot() {
  try {
    const data = await apiRequest("/api/state");
    state.remote.available = true;
    if (Number.isFinite(Number(data?.revision))) {
      state.remote.revision = Number(data.revision);
    }
    if (data && (data.guests?.length || data.tables?.length)) {
      applyRemoteSnapshot(data);
      showToast(`Datos cargados de la nube (${data.guests.length} invitados).`);
    }
    startRemotePolling();
  } catch (err) {
    if (err.status === 401) {
      clearAuthInStorage();
      const ok = await ensureAuthSession();
      if (ok) return loadRemoteSnapshot();
    }
    state.remote.available = false;
  }
}

function startRemotePolling() {
  if (!state.remote.available || state.remote.poller) return;
  state.remote.poller = setInterval(async () => {
    if (state.remote.saveInFlight) return;
    try {
      const data = await apiRequest(`/api/state?revision=${state.remote.revision}`);
      if (!data?.changed) return;
      if (Number.isFinite(Number(data.revision))) {
        state.remote.revision = Number(data.revision);
      }
      if (data && (Array.isArray(data.guests) || Array.isArray(data.tables))) {
        applyRemoteSnapshot(data);
        render();
      }
    } catch (err) {
      if (err.status === 401) {
        clearInterval(state.remote.poller);
        state.remote.poller = null;
        clearAuthInStorage();
      }
    }
  }, 4000);
}

function showToast(message) {
  refs.toast.textContent = message;
  refs.toast.classList.remove("hidden");
  setTimeout(() => refs.toast.classList.add("hidden"), 2000);
}

function normalize(value) {
  return (value ?? "").toString().trim();
}

function addGuest({ name, gender, confirmed, sourceRow, initialTable }) {
  if (!name) return;
  state.guests.push({
    id: crypto.randomUUID(),
    name,
    gender,
    confirmed,
    sourceRow,
    tableId: initialTable ?? null,
  });
}

function parseWorkbook(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const ws = wb.Sheets[SHEET_NAME];
  if (!ws) {
    throw new Error(`No se encontro la hoja "${SHEET_NAME}"`);
  }

  state.guests = [];
  state.tables = [];
  state.tableOrder = [];

  let row = START_ROW;
  while (row < 5000) {
    const hName = normalize(ws[`A${row}`]?.v);
    const hLast = normalize(ws[`B${row}`]?.v);
    const mName = normalize(ws[`C${row}`]?.v);
    const mLast = normalize(ws[`D${row}`]?.v);
    const confirmed = normalize(ws[`F${row}`]?.v).toLowerCase() === "ok";
    const hTable = ws[`H${row}`]?.v;
    const mTable = ws[`I${row}`]?.v;

    const hasCore = [hName, hLast, mName, mLast, ws[`F${row}`]?.v, hTable, mTable].some(
      (v) => normalize(v) !== "",
    );
    if (!hasCore) {
      if (row > 220) break;
      row += 1;
      continue;
    }

    addGuest({
      name: `${hName} ${hLast}`.trim(),
      gender: "H",
      confirmed,
      sourceRow: row,
      initialTable: Number.isFinite(Number(hTable)) ? `t-${Number(hTable)}` : null,
    });
    addGuest({
      name: `${mName} ${mLast}`.trim(),
      gender: "M",
      confirmed,
      sourceRow: row,
      initialTable: Number.isFinite(Number(mTable)) ? `t-${Number(mTable)}` : null,
    });
    row += 1;
  }

  if (!state.guests.length) {
    throw new Error("No se detectaron invitados en el formato esperado.");
  }

  showToast(`Excel cargado: ${state.guests.length} invitados detectados`);
}

function inferTableType(guests) {
  const males = guests.filter((g) => g.gender === "H").length;
  const females = guests.filter((g) => g.gender === "M").length;
  if (males === 0 && females === 0) return null;
  if (males > 0 && females === 0) return "men";
  if (females > 0 && males === 0) return "women";
  return "mixed";
}

function tableLabel(type) {
  if (type === "men") return "Solo hombres";
  if (type === "women") return "Solo mujeres";
  return "Mixta";
}

function syncTableOrder() {
  const idToNumber = new Map(state.tables.map((t) => [t.id, t.number]));
  const validIds = new Set(idToNumber.keys());
  const kept = state.tableOrder.filter((id) => validIds.has(id));
  const missing = state.tables
    .map((t) => t.id)
    .filter((id) => !kept.includes(id))
    .sort((a, b) => (idToNumber.get(a) || 0) - (idToNumber.get(b) || 0));
  state.tableOrder = [...kept, ...missing];
}

function setConsecutiveTables(maxNumber, capacityMap = new Map()) {
  const max = Math.max(1, Number(maxNumber) || 1);
  state.tables = [];
  for (let n = 1; n <= max; n += 1) {
    const cap = capacityMap.has(n) ? Number(capacityMap.get(n)) : n === 1 ? 20 : 10;
    state.tables.push({ id: `t-${n}`, number: n, name: `Mesa ${n}`, type: null, capacity: cap });
  }
  syncTableOrder();
}

function rebuildTablesFromCurrentAssignments() {
  const nums = state.guests
    .map((g) => (g.tableId ? Number(g.tableId.split("-")[1]) : null))
    .filter((n) => Number.isFinite(n) && n > 0);
  const maxTable = nums.length ? Math.max(...nums) : 1;
  setConsecutiveTables(maxTable);
}

function addMoreTables(count = 3) {
  const lastNumber = state.tables.length ? Math.max(...state.tables.map((t) => t.number)) : 0;
  for (let i = 1; i <= count; i += 1) {
    const number = lastNumber + i;
    state.tables.push({ id: `t-${number}`, number, name: `Mesa ${number}`, type: null, capacity: 10 });
  }
  syncTableOrder();
}

function ensureTable(tableNumber, capacity = null) {
  const existing = state.tables.find((t) => t.number === tableNumber);
  if (existing) {
    if (capacity && Number.isFinite(capacity)) existing.capacity = Number(capacity);
    return existing;
  }
  const cap = Number.isFinite(capacity) ? Number(capacity) : tableNumber === 1 ? 20 : 10;
  const table = { id: `t-${tableNumber}`, number: tableNumber, name: `Mesa ${tableNumber}`, type: null, capacity: cap };
  state.tables.push(table);
  syncTableOrder();
  return table;
}

function fitsTable(guest, table) {
  if (!table) return true;
  const current = state.guests.filter((g) => g.tableId === table.id).length;
  if (current >= table.capacity) return false;
  if (table.type === "men" && guest.gender !== "H") return false;
  if (table.type === "women" && guest.gender !== "M") return false;
  return true;
}

function moveGuest(guestId, tableId) {
  const guest = state.guests.find((g) => g.id === guestId);
  const table = state.tables.find((t) => t.id === tableId);
  if (!guest) return;

  if (tableId && !table) return;
  if (table && !fitsTable(guest, table)) {
    showToast("No entra por capacidad o restriccion de genero.");
    return;
  }

  guest.tableId = tableId || null;
  render();
  scheduleRemoteSave();
}

function deleteGuest(guestId) {
  const before = state.guests.length;
  state.guests = state.guests.filter((g) => g.id !== guestId);
  if (state.guests.length !== before) {
    showToast("Invitado descartado.");
    render();
    scheduleRemoteSave();
  }
}

function deleteTable(tableId) {
  const table = state.tables.find((t) => t.id === tableId);
  if (!table) return;

  const assignedCount = state.guests.filter((g) => g.tableId === tableId).length;
  if (assignedCount > 0) {
    showToast("Solo se puede borrar una mesa vacia.");
    return;
  }

  state.tables = state.tables.filter((t) => t.id !== tableId);
  state.tableOrder = state.tableOrder.filter((id) => id !== tableId);
  render();
  showToast(`${table.name} eliminada.`);
  scheduleRemoteSave();
}

function askGuestData() {
  return new Promise((resolve) => {
    refs.guestError.classList.add("hidden");
    refs.guestFirstName.value = "";
    refs.guestLastName.value = "";
    refs.guestGender.value = "";
    refs.guestModal.classList.remove("hidden");
    refs.guestFirstName.focus();

    const cleanup = () => {
      refs.guestModal.classList.add("hidden");
      refs.guestForm.removeEventListener("submit", submitHandler);
      refs.guestCancel.removeEventListener("click", cancelHandler);
    };

    const cancelHandler = () => {
      cleanup();
      resolve(null);
    };

    const submitHandler = (e) => {
      e.preventDefault();
      const firstName = normalize(refs.guestFirstName.value);
      const lastName = normalize(refs.guestLastName.value);
      const gender = normalize(refs.guestGender.value).toUpperCase();
      if (!firstName || !lastName || !["H", "M"].includes(gender)) {
        refs.guestError.textContent = "Completa nombre, apellido y genero.";
        refs.guestError.classList.remove("hidden");
        return;
      }
      cleanup();
      resolve({ fullName: `${firstName} ${lastName}`.trim(), gender });
    };

    refs.guestForm.addEventListener("submit", submitHandler);
    refs.guestCancel.addEventListener("click", cancelHandler);
  });
}

async function addGuestManually() {
  const data = await askGuestData();
  if (!data) return;

  state.guests.push({
    id: crypto.randomUUID(),
    name: data.fullName,
    gender: data.gender,
    confirmed: true,
    sourceRow: null,
    tableId: null,
  });
  showToast("Invitado agregado.");
  render();
  scheduleRemoteSave();
}

function filteredGuests() {
  return state.guests.filter((g) => {
    if (state.filter.gender !== "all" && g.gender !== state.filter.gender) return false;
    if (state.filter.search) {
      const q = state.filter.search.toLowerCase();
      if (!g.name.toLowerCase().includes(q)) return false;
    }
    return true;
  });
}

function guestCard(guest, options = {}) {
  const { allowDelete = false } = options;
  const el = document.createElement("div");
  el.className = `guest ${guest.gender === "H" ? "male" : "female"}`;
  el.draggable = true;
  el.dataset.guestId = guest.id;
  const meta = `${guest.gender === "H" ? "Hombre" : "Mujer"}${guest.confirmed ? "" : " - no confirmado"}`;
  el.innerHTML = `
    <div class="guest-head">
      <strong>${guest.name || "(sin nombre)"}</strong>
      ${allowDelete ? '<button class="guest-remove" type="button" aria-label="Descartar invitado" title="Descartar invitado">🗑</button>' : ""}
    </div>
    <small>${meta}</small>
  `;
  el.addEventListener("dragstart", () => {
    state.dragGuestId = guest.id;
    state.dragType = "guest";
  });
  el.addEventListener("dragend", () => {
    state.dragGuestId = null;
    state.dragType = null;
  });
  if (allowDelete) {
    const removeBtn = el.querySelector(".guest-remove");
    removeBtn.addEventListener("mousedown", (e) => {
      e.preventDefault();
      e.stopPropagation();
    });
    removeBtn.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
      deleteGuest(guest.id);
    });
  }
  return el;
}

function applyDropzoneBehavior(element, tableId) {
  element.addEventListener("dragover", (e) => {
    e.preventDefault();
    element.classList.add("drag-over");
  });
  element.addEventListener("dragleave", () => element.classList.remove("drag-over"));
  element.addEventListener("drop", (e) => {
    e.preventDefault();
    element.classList.remove("drag-over");
    if (state.dragType !== "guest" || !state.dragGuestId) return;
    moveGuest(state.dragGuestId, tableId || null);
    state.dragGuestId = null;
    state.dragType = null;
  });
}

function moveTableBefore(draggedId, targetId) {
  if (!draggedId || !targetId || draggedId === targetId) return;
  const base = state.tableOrder.slice();
  const from = base.indexOf(draggedId);
  const to = base.indexOf(targetId);
  if (from < 0 || to < 0) return;
  base.splice(from, 1);
  const insertAt = from < to ? to - 1 : to;
  base.splice(insertAt, 0, draggedId);
  state.tableOrder = base;
}

function applyTableReorderBehavior(card, tableId, handle) {
  handle.draggable = true;
  handle.addEventListener("dragstart", () => {
    state.dragTableId = tableId;
    state.dragType = "table";
    card.classList.add("table-dragging");
  });
  handle.addEventListener("dragend", () => {
    state.dragTableId = null;
    state.dragType = null;
    card.classList.remove("table-dragging");
  });

  card.addEventListener("dragover", (e) => {
    if (state.dragType !== "table") return;
    e.preventDefault();
    card.classList.add("table-drop-over");
  });
  card.addEventListener("dragleave", () => {
    card.classList.remove("table-drop-over");
  });
  card.addEventListener("drop", (e) => {
    if (state.dragType !== "table" || !state.dragTableId) return;
    e.preventDefault();
    card.classList.remove("table-drop-over");
    moveTableBefore(state.dragTableId, tableId);
    render();
    scheduleRemoteSave();
  });
}

function renderStats() {
  const visible = filteredGuests();
  const total = visible.length;
  const assigned = visible.filter((g) => g.tableId).length;
  const unassigned = total - assigned;
  const hm = visible.filter((g) => g.gender === "H").length;
  const wm = visible.filter((g) => g.gender === "M").length;
  refs.stats.innerHTML = [
    `<strong>Visibles:</strong> ${total}`,
    `<strong>Asignados:</strong> ${assigned}`,
    `<strong>Sin asignar:</strong> ${unassigned}`,
    `<strong>H/M:</strong> ${hm}/${wm}`,
  ].join("<br>");
}

function renderUnassigned(visibleGuests) {
  refs.unassignedList.innerHTML = "";
  const list = visibleGuests.filter((g) => !g.tableId);
  list.sort((a, b) => a.name.localeCompare(b.name, "es"));
  list.forEach((g) => refs.unassignedList.appendChild(guestCard(g, { allowDelete: true })));
}

function renderTables(visibleGuests) {
  refs.tablesGrid.innerHTML = "";
  syncTableOrder();
  state.tableOrder.forEach((tableId) => {
      const table = state.tables.find((t) => t.id === tableId);
      if (!table) return;
      const assigned = visibleGuests
        .filter((g) => g.tableId === table.id)
        .sort((a, b) => a.name.localeCompare(b.name, "es"));
      const count = assigned.length;
      const cls = count > table.capacity ? "bad" : count === table.capacity ? "ok" : count >= table.capacity - 1 ? "warn" : "";
      const menCount = assigned.filter((g) => g.gender === "H").length;
      const womenCount = assigned.filter((g) => g.gender === "M").length;
      const displayType = inferTableType(assigned);
      const typePill = displayType
        ? `<span class="type-pill ${displayType}">${tableLabel(displayType)}</span>`
        : "";
      const bigPill = table.capacity === 20 ? '<span class="type-pill mixed">Mesa grande (20)</span>' : "";

      const card = document.createElement("article");
      card.className = `table-card ${cls}`;
      card.dataset.tableId = table.id;
      card.innerHTML = `
        <div class="table-head">
          <strong>${table.name}</strong>
          <div class="table-actions">
            <button class="table-remove" type="button" aria-label="Eliminar mesa" title="Eliminar mesa">Eliminar</button>
            <span class="table-drag-handle" title="Mover mesa en el layout">Mover</span>
          </div>
          ${typePill || bigPill}
        </div>
        <div class="table-meta">${count}/${table.capacity} | H:${menCount} M:${womenCount}</div>
      `;
      const handle = card.querySelector(".table-drag-handle");
      const removeBtn = card.querySelector(".table-remove");
      applyTableReorderBehavior(card, table.id, handle);
      removeBtn.addEventListener("click", (e) => {
        e.preventDefault();
        e.stopPropagation();
        deleteTable(table.id);
      });

      const zone = document.createElement("div");
      zone.className = "guest-list dropzone";
      zone.dataset.tableId = table.id;
      assigned.forEach((g) => zone.appendChild(guestCard(g)));
      applyDropzoneBehavior(zone, table.id);
      card.appendChild(zone);
      refs.tablesGrid.appendChild(card);
    });
}

function render() {
  const visible = filteredGuests();
  renderStats();
  renderUnassigned(visible);
  renderTables(visible);
  applyDropzoneBehavior(refs.unassignedList, null);
}

function exportCsv() {
  const rows = [["TipoRegistro", "Nombre", "Genero", "Confirmado", "Mesa", "Fila Excel", "Capacidad"]];
  state.guests
    .slice()
    .sort((a, b) => a.name.localeCompare(b.name, "es"))
    .forEach((g) => {
      const table = state.tables.find((t) => t.id === g.tableId);
      rows.push(["INVITADO", g.name, g.gender, g.confirmed ? "ok" : "", table ? table.number : "", g.sourceRow, ""]);
    });
  state.tables
    .slice()
    .sort((a, b) => a.number - b.number)
    .forEach((t) => {
      rows.push(["MESA", "", "", "", t.number, "", t.capacity]);
    });
  const csv = rows.map((r) => r.map((v) => `"${String(v ?? "").replaceAll("\"", "\"\"")}"`).join(",")).join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "mesas_asignacion.csv";
  a.click();
  URL.revokeObjectURL(a.href);
}

function normalizeHeaderKey(key) {
  return normalize(key)
    .replace(/^\ufeff/, "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]/g, "");
}

function getRowValue(row, aliases) {
  const normalized = new Map(
    Object.entries(row).map(([k, v]) => [normalizeHeaderKey(k), v]),
  );
  for (const alias of aliases) {
    if (normalized.has(alias)) return normalized.get(alias);
  }
  return "";
}

function readAssignmentRows(text) {
  const parseWith = (options = {}) => {
    const wb = XLSX.read(text, { type: "string", ...options });
    const ws = wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_json(ws, { defval: "" });
  };

  const rows = parseWith();
  if (!rows.length) return rows;

  const firstKeys = Object.keys(rows[0]);
  if (firstKeys.length === 1 && /[;\t]/.test(firstKeys[0])) {
    if (firstKeys[0].includes(";")) return parseWith({ FS: ";" });
    if (firstKeys[0].includes("\t")) return parseWith({ FS: "\t" });
  }

  return rows;
}

function toNumberOrNull(value) {
  const raw = normalize(value).replace(",", ".");
  if (!raw) return null;
  const num = Number(raw);
  return Number.isFinite(num) ? num : null;
}

function normalizeGender(value) {
  const raw = normalize(value).toLowerCase();
  if (!raw) return "";
  if (raw === "h" || raw.startsWith("hom") || raw === "male") return "H";
  if (raw === "m" || raw === "f" || raw.startsWith("muj") || raw.startsWith("fem")) return "M";
  return "";
}

function isConfirmedValue(value) {
  const raw = normalize(value).toLowerCase();
  if (!raw) return false;
  return ["ok", "si", "sí", "true", "1", "x", "confirmado"].includes(raw);
}

function bootstrapGuestsFromAssignmentRows(rows) {
  state.guests = [];
  rows.forEach((row) => {
    const tipo = normalize(getRowValue(row, ["tiporegistro", "tipo"])).toUpperCase();
    if (tipo && tipo !== "INVITADO") return;
    const genero = normalizeGender(getRowValue(row, ["genero"]));
    const nombre = normalize(getRowValue(row, ["nombre"]));
    if (!nombre || !genero) return;

    addGuest({
      name: nombre,
      gender: genero,
      confirmed: isConfirmedValue(getRowValue(row, ["confirmado"])),
      sourceRow: toNumberOrNull(getRowValue(row, ["filaexcel", "fila"])),
      initialTable: null,
    });
  });
}

function importAssignmentCsv(text) {
  const rows = readAssignmentRows(text);
  if (!rows.length) throw new Error("CSV vacio o invalido.");
  if (!state.guests.length) {
    bootstrapGuestsFromAssignmentRows(rows);
    if (!state.guests.length) {
      throw new Error("CSV invalido: no se pudieron leer invitados.");
    }
  }

  state.guests.forEach((g) => {
    g.tableId = null;
  });
  state.tables = [];

  const byRowGender = new Map();
  const byNameGender = new Map();
  state.guests.forEach((g) => {
    if (Number.isFinite(g.sourceRow)) {
      byRowGender.set(`${g.sourceRow}|${g.gender}`, g);
    }
    const key = `${g.name.toLowerCase()}|${g.gender}`;
    if (!byNameGender.has(key)) byNameGender.set(key, []);
    byNameGender.get(key).push(g);
  });

  let assignedCount = 0;
  const mesaDefs = new Map();
  const usedTableNumbers = new Set();
  rows.forEach((row) => {
    const tipo = normalize(getRowValue(row, ["tiporegistro", "tipo"])).toUpperCase();
    const genero = normalizeGender(getRowValue(row, ["genero"]));
    const nombre = normalize(getRowValue(row, ["nombre"]));
    const fila = toNumberOrNull(getRowValue(row, ["filaexcel", "fila"]));
    const mesa = toNumberOrNull(getRowValue(row, ["mesa"]));
    const capacidad = toNumberOrNull(getRowValue(row, ["capacidad"]));

    if (tipo === "MESA") {
      if (mesa !== null && mesa > 0) {
        mesaDefs.set(mesa, capacidad !== null ? capacidad : mesa === 1 ? 20 : 10);
      }
      return;
    }

    if (!(mesa !== null && mesa > 0) && fila === null && !nombre) return;
    if (!["H", "M"].includes(genero)) return;

    let guest = null;
    if (fila !== null) {
      guest = byRowGender.get(`${fila}|${genero}`) || null;
    }
    if (!guest && nombre) {
      const key = `${nombre.toLowerCase()}|${genero}`;
      const candidates = byNameGender.get(key) || [];
      guest = candidates.find((c) => !c.tableId) || candidates[0] || null;
    }
    if (!guest) return;

    if (mesa !== null && mesa > 0) {
      guest.tableId = `t-${mesa}`;
      usedTableNumbers.add(mesa);
    } else {
      guest.tableId = null;
    }
    assignedCount += 1;
  });

  if (mesaDefs.size > 0) {
    const maxFromDefs = Math.max(...mesaDefs.keys());
    const maxFromUsed = usedTableNumbers.size ? Math.max(...usedTableNumbers) : 1;
    const max = Math.max(maxFromDefs, maxFromUsed, 1);
    setConsecutiveTables(max, mesaDefs);
  } else {
    const max = usedTableNumbers.size ? Math.max(...usedTableNumbers) : 1;
    setConsecutiveTables(max);
  }

  showToast(`CSV cargado: ${assignedCount} asignaciones aplicadas.`);
  render();
  scheduleRemoteSave();
}

refs.fileInput.addEventListener("change", async (e) => {
  const file = e.target.files?.[0];
  if (!file) return;
  const data = await file.arrayBuffer();
  try {
    parseWorkbook(data);
    rebuildTablesFromCurrentAssignments();
    render();
    scheduleRemoteSave();
  } catch (err) {
    showToast(err.message || "Error leyendo Excel");
  } finally {
    e.target.value = "";
  }
});

refs.csvInput.addEventListener("change", async (e) => {
  const file = e.target.files?.[0];
  if (!file) return;
  try {
    const text = await file.text();
    importAssignmentCsv(text);
  } catch (err) {
    showToast(err.message || "Error leyendo CSV");
  } finally {
    e.target.value = "";
  }
});

refs.addGuestBtn.addEventListener("click", async () => {
  await addGuestManually();
});

refs.addOneTableBtn.addEventListener("click", () => {
  if (!state.guests.length) return showToast("Primero carga el Excel.");
  addMoreTables(1);
  render();
  showToast("Se agrego 1 mesa mixta de 10.");
  scheduleRemoteSave();
});

refs.exportBtn.addEventListener("click", () => {
  if (!state.guests.length) return showToast("No hay datos para exportar.");
  exportCsv();
});

refs.searchInput.addEventListener("input", (e) => {
  state.filter.search = e.target.value.trim();
  render();
});
refs.genderFilter.addEventListener("change", (e) => {
  state.filter.gender = e.target.value;
  render();
});

async function init() {
  render();
  const authed = await ensureAuthSession();
  if (!authed) {
    showToast("Sesion cancelada.");
    return;
  }
  await loadRemoteSnapshot();
  render();
}

init();
