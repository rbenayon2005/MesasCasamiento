const SHEET_NAME = "Invitados Ariel Casamiento";
const START_ROW = 8;

const STORAGE_KEYS = {
  sessionToken: "mesas_session_token",
  sessionUser: "mesas_session_user",
  currentEventId: "mesas_current_event_id",
};

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
  session: {
    token: "",
    user: null,
  },
  authMode: "login",
  events: [],
  currentEventId: "",
  expandedEventId: "",
};

const refs = {
  appTitle: document.getElementById("appTitle"),
  topbarEventMeta: document.getElementById("topbarEventMeta"),
  fileInput: document.getElementById("excelFile"),
  assignmentFileInput: document.getElementById("assignmentFile"),
  openCreateEventBtn: document.getElementById("openCreateEventBtn"),
  addGuestBtn: document.getElementById("addGuestBtn"),
  exportBtn: document.getElementById("exportBtn"),
  excelHelpBtn: document.getElementById("excelHelpBtn"),
  searchInput: document.getElementById("searchInput"),
  genderFilter: document.getElementById("genderFilter"),
  unassignedList: document.getElementById("unassignedList"),
  tablesGrid: document.getElementById("tablesGrid"),
  stats: document.getElementById("stats"),
  toast: document.getElementById("toast"),
  authModal: document.getElementById("authModal"),
  authForm: document.getElementById("authForm"),
  authModeLogin: document.getElementById("authModeLogin"),
  authModeRegister: document.getElementById("authModeRegister"),
  authTitle: document.getElementById("authTitle"),
  authSubtitle: document.getElementById("authSubtitle"),
  authNameWrap: document.getElementById("authNameWrap"),
  authName: document.getElementById("authName"),
  authUser: document.getElementById("authUser"),
  authPass: document.getElementById("authPass"),
  authError: document.getElementById("authError"),
  authSubmit: document.getElementById("authSubmit"),
  sessionBadge: document.getElementById("sessionBadge"),
  sessionBadgeName: document.getElementById("sessionBadgeName"),
  sessionBadgeEmail: document.getElementById("sessionBadgeEmail"),
  logoutBtn: document.getElementById("logoutBtn"),
  eventCountBadge: document.getElementById("eventCountBadge"),
  eventList: document.getElementById("eventList"),
  createEventForm: document.getElementById("createEventForm"),
  createEventModal: document.getElementById("createEventModal"),
  createEventClose: document.getElementById("createEventClose"),
  importConfirmModal: document.getElementById("importConfirmModal"),
  importConfirmMessage: document.getElementById("importConfirmMessage"),
  importCancelBtn: document.getElementById("importCancelBtn"),
  importConfirmBtn: document.getElementById("importConfirmBtn"),
  excelHelpModal: document.getElementById("excelHelpModal"),
  excelHelpClose: document.getElementById("excelHelpClose"),
  downloadTemplateBtn: document.getElementById("downloadTemplateBtn"),
  newEventName: document.getElementById("newEventName"),
  newEventDate: document.getElementById("newEventDate"),
  eventSettingsForm: document.getElementById("eventSettingsForm"),
  eventNameInput: document.getElementById("eventNameInput"),
  eventDateInput: document.getElementById("eventDateInput"),
  eventSummary: document.getElementById("eventSummary"),
  eventLockMessage: document.getElementById("eventLockMessage"),
  addTableForm: document.getElementById("addTableForm"),
  newTableCapacity: document.getElementById("newTableCapacity"),
  bulkTablesForm: document.getElementById("bulkTablesForm"),
  bulkTableCount: document.getElementById("bulkTableCount"),
  bulkTableCapacity: document.getElementById("bulkTableCapacity"),
  managementModal: document.getElementById("managementModal"),
  managementClose: document.getElementById("managementClose"),
  guestModal: document.getElementById("guestModal"),
  guestForm: document.getElementById("guestForm"),
  guestTitle: document.getElementById("guestTitle"),
  guestFirstName: document.getElementById("guestFirstName"),
  guestLastName: document.getElementById("guestLastName"),
  guestGender: document.getElementById("guestGender"),
  guestConfirmed: document.getElementById("guestConfirmed"),
  guestSubmit: document.getElementById("guestSubmit"),
  guestError: document.getElementById("guestError"),
  guestCancel: document.getElementById("guestCancel"),
};

async function apiRequest(path, options = {}) {
  const { skipAuth = false, headers = {}, ...fetchOptions } = options;
  const authHeaders = {};
  if (!skipAuth && state.session.token) {
    authHeaders["x-session-token"] = state.session.token;
  }
  const response = await fetch(path, {
    headers: { "Content-Type": "application/json", ...authHeaders, ...headers },
    ...fetchOptions,
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

function normalize(value) {
  return (value ?? "").toString().trim();
}

function currentEvent() {
  return state.events.find((event) => event.id === state.currentEventId) || null;
}

function isValidEventDate(value) {
  return /^\d{4}-\d{2}-\d{2}$/.test(normalize(value));
}

function isCurrentEventLocked() {
  return !!currentEvent()?.isLocked;
}

function lockedEventMessage() {
  const event = currentEvent();
  if (!event?.eventDate) return "El evento ya paso y quedo solo lectura.";
  return `El evento fue el ${event.eventDate} y ahora esta en solo lectura.`;
}

function requireEditableCurrentEvent() {
  if (!state.currentEventId) {
    showToast("Primero crea o selecciona un evento.");
    return false;
  }
  if (!isCurrentEventLocked()) return true;
  showToast(lockedEventMessage());
  return false;
}

function syncEventRecord(patch) {
  if (!patch?.id) return;
  state.events = state.events.map((event) => (event.id === patch.id ? { ...event, ...patch } : event));
}

function syncCurrentEventSummary() {
  const event = currentEvent();
  if (!event) return;
  event.tableCount = state.tables.length;
  event.guestCount = state.guests.length;
}

function syncExpandedEventWithCurrent() {
  state.expandedEventId = state.currentEventId || "";
}

function showToast(message) {
  refs.toast.textContent = message;
  refs.toast.classList.remove("hidden");
  setTimeout(() => refs.toast.classList.add("hidden"), 2200);
}

function openManagementModal() {
  if (!state.currentEventId) {
    showToast("Primero crea o selecciona un evento.");
    return;
  }
  refs.managementModal.classList.remove("hidden");
}

function closeManagementModal() {
  refs.managementModal.classList.add("hidden");
}

function openCreateEventModal() {
  refs.newEventName.value = "";
  refs.newEventDate.value = "";
  refs.createEventModal.classList.remove("hidden");
  refs.newEventName.focus();
}

function closeCreateEventModal() {
  refs.createEventModal.classList.add("hidden");
}

function askImportConfirmation(fileName) {
  const event = currentEvent();
  if (!event) return Promise.resolve(false);
  refs.importConfirmMessage.textContent = `El archivo "${fileName}" va a cargar o actualizar invitados y mesas en el evento activo: "${event.name}".`;
  refs.importConfirmModal.classList.remove("hidden");
  refs.importConfirmBtn.focus();

  return new Promise((resolve) => {
    const cleanup = () => {
      refs.importConfirmModal.classList.add("hidden");
      refs.importCancelBtn.removeEventListener("click", onCancel);
      refs.importConfirmBtn.removeEventListener("click", onConfirm);
    };
    const onCancel = () => {
      cleanup();
      resolve(false);
    };
    const onConfirm = () => {
      cleanup();
      resolve(true);
    };
    refs.importCancelBtn.addEventListener("click", onCancel);
    refs.importConfirmBtn.addEventListener("click", onConfirm);
  });
}

function loadSessionFromStorage() {
  state.session.token = localStorage.getItem(STORAGE_KEYS.sessionToken) || "";
  try {
    state.session.user = JSON.parse(localStorage.getItem(STORAGE_KEYS.sessionUser) || "null");
  } catch {
    state.session.user = null;
  }
  state.currentEventId = localStorage.getItem(STORAGE_KEYS.currentEventId) || "";
}

function saveSessionToStorage() {
  if (state.session.token) {
    localStorage.setItem(STORAGE_KEYS.sessionToken, state.session.token);
  } else {
    localStorage.removeItem(STORAGE_KEYS.sessionToken);
  }
  if (state.session.user) {
    localStorage.setItem(STORAGE_KEYS.sessionUser, JSON.stringify(state.session.user));
  } else {
    localStorage.removeItem(STORAGE_KEYS.sessionUser);
  }
  if (state.currentEventId) {
    localStorage.setItem(STORAGE_KEYS.currentEventId, state.currentEventId);
  } else {
    localStorage.removeItem(STORAGE_KEYS.currentEventId);
  }
}

function clearSession() {
  state.session.token = "";
  state.session.user = null;
  state.events = [];
  state.currentEventId = "";
  state.expandedEventId = "";
  clearRemoteState();
  resetLocalEventState();
  saveSessionToStorage();
  render();
}

function clearRemoteState() {
  clearTimeout(state.remote.saveTimer);
  if (state.remote.poller) {
    clearInterval(state.remote.poller);
  }
  state.remote.available = false;
  state.remote.saveTimer = null;
  state.remote.saveInFlight = false;
  state.remote.revision = 0;
  state.remote.poller = null;
}

function resetLocalEventState() {
  state.guests = [];
  state.tables = [];
  state.tableOrder = [];
}

function parseApiError(err, fallback) {
  try {
    const data = JSON.parse(err.message);
    return data.error || data.detail || fallback;
  } catch {
    return fallback;
  }
}

function setAuthMode(mode) {
  state.authMode = mode === "register" ? "register" : "login";
  const register = state.authMode === "register";
  refs.authModeLogin.classList.toggle("active", !register);
  refs.authModeRegister.classList.toggle("active", register);
  refs.authNameWrap.classList.toggle("hidden", !register);
  refs.authName.required = register;
  refs.authTitle.textContent = register ? "Crea tu cuenta" : "Ingresar a tus eventos";
  refs.authSubtitle.textContent = register
    ? "Registra una cuenta para crear eventos y administrar mesas."
    : "Entra con tu cuenta para administrar mesas y eventos.";
  refs.authSubmit.textContent = register ? "Crear cuenta" : "Entrar";
  refs.authPass.autocomplete = register ? "new-password" : "current-password";
}

function showAuthModal() {
  return new Promise((resolve) => {
    refs.authError.classList.add("hidden");
    refs.authError.textContent = "";
    refs.authName.value = "";
    refs.authUser.value = state.session.user?.email || "";
    refs.authPass.value = "";
    setAuthMode(state.authMode || "login");
    refs.authModal.classList.remove("hidden");
    (state.authMode === "register" ? refs.authName : refs.authUser).focus();

    const cleanup = () => {
      refs.authModal.classList.add("hidden");
      refs.authForm.removeEventListener("submit", submitHandler);
      refs.authModeLogin.removeEventListener("click", loginModeHandler);
      refs.authModeRegister.removeEventListener("click", registerModeHandler);
    };

    const loginModeHandler = () => setAuthMode("login");
    const registerModeHandler = () => setAuthMode("register");

    const submitHandler = async (e) => {
      e.preventDefault();
      const email = normalize(refs.authUser.value).toLowerCase();
      const password = normalize(refs.authPass.value);
      const name = normalize(refs.authName.value);

      if (!email || !password || (state.authMode === "register" && !name)) {
        refs.authError.textContent =
          state.authMode === "register"
            ? "Completa nombre, email y password."
            : "Completa email y password.";
        refs.authError.classList.remove("hidden");
        return;
      }

      try {
        const session = await apiRequest("/api/session", {
          method: "POST",
          skipAuth: true,
          body: JSON.stringify({
            action: state.authMode,
            name,
            email,
            password,
          }),
        });
        cleanup();
        resolve(session);
      } catch (err) {
        refs.authError.textContent = parseApiError(
          err,
          state.authMode === "register" ? "No se pudo crear la cuenta." : "No se pudo iniciar sesion.",
        );
        refs.authError.classList.remove("hidden");
      }
    };

    refs.authForm.addEventListener("submit", submitHandler);
    refs.authModeLogin.addEventListener("click", loginModeHandler);
    refs.authModeRegister.addEventListener("click", registerModeHandler);
  });
}

async function ensureSession() {
  loadSessionFromStorage();
  if (state.session.token) {
    try {
      await refreshEvents();
      return true;
    } catch (err) {
      if (err.status !== 401) throw err;
      clearSession();
    }
  }

  const session = await showAuthModal();
  state.session.token = session.token;
  state.session.user = session.user;
  saveSessionToStorage();
  await refreshEvents();
  return true;
}

async function logout() {
  try {
    if (state.session.token) {
      await apiRequest("/api/session", { method: "DELETE" });
    }
  } catch {
    // no-op
  }
  clearSession();
  setAuthMode("login");
  await ensureSession();
}

async function refreshEvents() {
  const data = await apiRequest("/api/events");
  state.events = Array.isArray(data?.events) ? data.events : [];
  const exists = state.events.some((event) => event.id === state.currentEventId);
  if (!exists) {
    state.currentEventId = state.events[0]?.id || "";
  }
  syncExpandedEventWithCurrent();
  saveSessionToStorage();
  renderAccountPanel();
  renderEventsPanel();
}

function serializeStateForRemote() {
  return {
    guests: state.guests.map((guest) => ({
      id: guest.id,
      name: guest.name,
      gender: guest.gender,
      confirmed: !!guest.confirmed,
      sourceRow: Number.isFinite(guest.sourceRow) ? guest.sourceRow : null,
      tableId: guest.tableId || null,
    })),
    tables: state.tables
      .slice()
      .sort((a, b) => a.number - b.number)
      .map((table) => ({
        number: table.number,
        name: table.name,
        type: table.type || null,
        capacity: table.capacity,
      })),
    tableOrder: state.tableOrder.slice(),
  };
}

function applyRemoteSnapshot(snapshot) {
  if (!snapshot) return false;
  if (snapshot.event?.id) {
    syncEventRecord(snapshot.event);
  }
  const guests = Array.isArray(snapshot.guests) ? snapshot.guests : [];
  const tables = Array.isArray(snapshot.tables) ? snapshot.tables : [];

  state.guests = guests.map((guest) => ({
    id: normalize(guest.id) || crypto.randomUUID(),
    name: normalize(guest.name),
    gender: normalize(guest.gender).toUpperCase() === "M" ? "M" : "H",
    confirmed: !!guest.confirmed,
    sourceRow: Number.isFinite(Number(guest.sourceRow)) ? Number(guest.sourceRow) : null,
    tableId: normalize(guest.tableId) || null,
  }));

  state.tables = tables
    .map((table) => ({
      id: `t-${Number(table.number)}`,
      number: Number(table.number),
      name: normalize(table.name) || `Mesa ${Number(table.number)}`,
      type: ["men", "women"].includes(table.type) ? table.type : null,
      capacity: Number.isFinite(Number(table.capacity)) ? Number(table.capacity) : 10,
    }))
    .filter((table) => Number.isFinite(table.number) && table.number > 0)
    .sort((a, b) => a.number - b.number);

  state.tableOrder = Array.isArray(snapshot.tableOrder)
    ? snapshot.tableOrder.map((id) => normalize(id)).filter(Boolean)
    : [];
  syncTableOrder();
  return true;
}

async function saveSnapshotNow() {
  if (!state.remote.available || state.remote.saveInFlight || !state.currentEventId) return;
  state.remote.saveInFlight = true;
  try {
    const response = await apiRequest(`/api/state?eventId=${encodeURIComponent(state.currentEventId)}`, {
      method: "POST",
      body: JSON.stringify(serializeStateForRemote()),
    });
    if (response && Number.isFinite(Number(response.revision))) {
      state.remote.revision = Number(response.revision);
    }
  } catch (err) {
    if (err.status === 401) {
      await logout();
      return;
    }
    showToast("No se pudo guardar el evento.");
  } finally {
    state.remote.saveInFlight = false;
  }
}

function scheduleRemoteSave(delayMs = 450) {
  if (!state.remote.available || !state.currentEventId) return;
  clearTimeout(state.remote.saveTimer);
  state.remote.saveTimer = setTimeout(() => {
    saveSnapshotNow();
  }, delayMs);
}

async function loadRemoteSnapshot() {
  if (!state.currentEventId) {
    clearRemoteState();
    resetLocalEventState();
    render();
    return;
  }

  try {
    const data = await apiRequest(`/api/state?eventId=${encodeURIComponent(state.currentEventId)}`);
    state.remote.available = true;
    state.remote.revision = Number.isFinite(Number(data?.revision)) ? Number(data.revision) : 0;
    applyRemoteSnapshot(data);
    startRemotePolling();
    render();
  } catch (err) {
    if (err.status === 401) {
      await logout();
      return;
    }
    showToast(parseApiError(err, "No se pudo cargar el evento."));
    state.remote.available = false;
  }
}

function startRemotePolling() {
  if (!state.remote.available || state.remote.poller || !state.currentEventId) return;
  state.remote.poller = setInterval(async () => {
    if (state.remote.saveInFlight || !state.currentEventId) return;
    try {
      const data = await apiRequest(
        `/api/state?eventId=${encodeURIComponent(state.currentEventId)}&revision=${state.remote.revision}`,
      );
      if (!data?.changed) return;
      if (Number.isFinite(Number(data.revision))) {
        state.remote.revision = Number(data.revision);
      }
      applyRemoteSnapshot(data);
      render();
    } catch (err) {
      if (err.status === 401) {
        clearRemoteState();
        await logout();
      }
    }
  }, 4000);
}

async function selectEvent(eventId) {
  if (eventId === state.currentEventId) return;
  clearRemoteState();
  state.currentEventId = eventId || "";
  syncExpandedEventWithCurrent();
  saveSessionToStorage();
  resetLocalEventState();
  render();
  renderEventsPanel();
  renderEventConfig();
  await loadRemoteSnapshot();
}

async function createEvent(name) {
  const eventName = normalize(name);
  const eventDate = normalize(refs.newEventDate.value);
  if (!eventName) return;
  if (!isValidEventDate(eventDate)) return showToast("Carga una fecha valida para el evento.");
  const response = await apiRequest("/api/events", {
    method: "POST",
    body: JSON.stringify({ action: "create", name: eventName, eventDate }),
  });
  if (Array.isArray(response?.events)) {
    state.events = response.events;
  } else {
    await refreshEvents();
  }
  const created = response?.event || state.events[state.events.length - 1];
  refs.newEventName.value = "";
  refs.newEventDate.value = "";
  closeCreateEventModal();
  if (created?.id) {
    state.currentEventId = created.id;
  }
  syncExpandedEventWithCurrent();
  saveSessionToStorage();
  renderEventsPanel();
  renderEventConfig();
  await loadRemoteSnapshot();
  showToast("Evento creado.");
}

async function renameCurrentEvent(name) {
  const event = currentEvent();
  if (!event) return;
  if (!requireEditableCurrentEvent()) return;
  const nextName = normalize(name);
  const nextDate = normalize(refs.eventDateInput.value);
  if (!nextName) return showToast("El evento necesita un nombre.");
  if (!isValidEventDate(nextDate)) return showToast("Carga una fecha valida para el evento.");
  const response = await apiRequest("/api/events", {
    method: "POST",
    body: JSON.stringify({ action: "update", eventId: event.id, name: nextName, eventDate: nextDate }),
  });
  state.events = Array.isArray(response?.events) ? response.events : state.events;
  renderEventsPanel();
  renderEventConfig();
  showToast("Evento actualizado.");
}

async function deleteEventById(eventId) {
  const event = state.events.find((item) => item.id === eventId);
  if (!event) return;
  if (event.isLocked) {
    showToast(`El evento "${event.name}" quedo solo lectura y no se puede modificar.`);
    return;
  }

  const response = await apiRequest("/api/events", {
    method: "POST",
    body: JSON.stringify({ action: "delete", eventId }),
  });

  state.events = Array.isArray(response?.events) ? response.events : [];
  if (state.expandedEventId === eventId) {
    state.expandedEventId = "";
  }
  if (state.currentEventId === eventId) {
    clearRemoteState();
    state.currentEventId = state.events[0]?.id || "";
    syncExpandedEventWithCurrent();
    if (!state.currentEventId) {
      resetLocalEventState();
    }
  }
  saveSessionToStorage();
  render();
  if (state.currentEventId) {
    await loadRemoteSnapshot();
  }
  showToast(`Evento "${event.name}" eliminado.`);
}

function renderAccountPanel() {
  const user = state.session.user;
  refs.sessionBadgeName.textContent = user?.name || "Sin sesion";
  refs.sessionBadgeEmail.textContent = user?.email || "";
  refs.sessionBadge.classList.toggle("hidden", !user);
  refs.logoutBtn.classList.toggle("hidden", !user);
}

function renderEventsPanel() {
  refs.eventCountBadge.textContent = String(state.events.length);
  refs.eventList.innerHTML = "";
  if (!state.events.length) {
    refs.eventList.className = "event-list empty-state";
    refs.eventList.textContent = "Todavia no creaste eventos.";
    return;
  }

  syncExpandedEventWithCurrent();
  refs.eventList.className = "event-list";
  state.events.forEach((event) => {
    const isExpanded = event.id === state.currentEventId;
    const card = document.createElement("div");
    card.className = `event-item${isExpanded ? " active" : ""}`;
    card.innerHTML = `
      <button type="button" class="event-item-head" aria-expanded="${isExpanded}">
        <strong>${event.name}</strong>
      </button>
      ${
        isExpanded
          ? `<div class="event-item-body">
              <small>${event.tableCount} mesas · ${event.guestCount} invitados${event.eventDate ? ` · ${event.eventDate}` : ""}</small>
              ${event.isLocked ? '<span class="event-item-status">Solo lectura</span>' : ""}
              <div class="event-item-actions">
                <button type="button" class="event-manage-btn">Gestionar</button>
                <button type="button" class="event-delete-btn" ${event.isLocked ? "disabled" : ""}>Borrar</button>
              </div>
            </div>`
          : ""
      }
    `;
    const headBtn = card.querySelector(".event-item-head");
    const manageBtn = card.querySelector(".event-manage-btn");
    const deleteBtn = card.querySelector(".event-delete-btn");
    headBtn.addEventListener("click", async () => {
      state.expandedEventId = event.id;
      if (event.id === state.currentEventId) {
        renderEventsPanel();
        return;
      }
      await selectEvent(event.id);
    });
    if (manageBtn) {
      manageBtn.addEventListener("click", async () => {
        await selectEvent(event.id);
        state.expandedEventId = event.id;
        renderEventsPanel();
        openManagementModal();
      });
    }
    if (deleteBtn) {
      deleteBtn.addEventListener("click", async () => {
        try {
          await deleteEventById(event.id);
        } catch (err) {
          showToast(parseApiError(err, "No se pudo borrar el evento."));
        }
      });
    }
    refs.eventList.appendChild(card);
  });
}

function renderEventConfig() {
  const event = currentEvent();
  refs.topbarEventMeta.textContent = event
    ? `${event.name}${event.eventDate ? ` · ${event.eventDate}` : ""} · ${state.tables.length} mesas`
    : "Sin evento seleccionado";
  refs.appTitle.textContent = event ? `Organizador de Mesas · ${event.name}` : "Organizador de Mesas";

  if (!event) {
    refs.eventNameInput.value = "";
    refs.eventDateInput.value = "";
    refs.eventSummary.textContent = "Selecciona un evento para configurarlo.";
    refs.eventLockMessage.textContent = "";
    refs.eventLockMessage.classList.add("hidden");
    closeManagementModal();
    return;
  }

  refs.eventNameInput.value = event.name;
  refs.eventDateInput.value = event.eventDate || "";
  const capacityTotal = state.tables.reduce((sum, table) => sum + table.capacity, 0);
  refs.eventSummary.textContent = `${state.tables.length} mesas configuradas · capacidad total ${capacityTotal} personas.${event.isLocked ? " Evento en solo lectura." : ""}`;
  refs.eventLockMessage.textContent = event.isLocked
    ? `${lockedEventMessage()} Se puede ver el evento y exportar asignaciones, pero no editarlo.`
    : "";
  refs.eventLockMessage.classList.toggle("hidden", !event.isLocked);
}

function updateActionAvailability() {
  const hasEvent = Boolean(state.currentEventId);
  const isLocked = isCurrentEventLocked();
  const canEdit = hasEvent && !isLocked;
  if (refs.fileInput) refs.fileInput.disabled = !canEdit;
  refs.assignmentFileInput.disabled = !canEdit;
  refs.addGuestBtn.disabled = !canEdit;
  refs.exportBtn.disabled = !hasEvent || !state.guests.length;
  refs.openCreateEventBtn.disabled = !state.session.user;
  refs.eventNameInput.disabled = !canEdit;
  refs.eventDateInput.disabled = !canEdit;
  refs.newTableCapacity.disabled = !canEdit;
  refs.bulkTableCount.disabled = !canEdit;
  refs.bulkTableCapacity.disabled = !canEdit;
  document.getElementById("saveEventSettingsBtn").disabled = !canEdit;
  document.getElementById("addTableBtn").disabled = !canEdit;
  document.getElementById("bulkAddTablesBtn").disabled = !canEdit;
  refs.fileInput?.closest(".file-input")?.classList.toggle("disabled", !canEdit);
  refs.assignmentFileInput.closest(".file-input").classList.toggle("disabled", !canEdit);
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

  let row = START_ROW;
  while (row < 5000) {
    const hName = normalize(ws[`A${row}`]?.v);
    const hLast = normalize(ws[`B${row}`]?.v);
    const mName = normalize(ws[`C${row}`]?.v);
    const mLast = normalize(ws[`D${row}`]?.v);
    const confirmed = isConfirmedValue(ws[`F${row}`]?.v);
    const hTable = ws[`H${row}`]?.v;
    const mTable = ws[`I${row}`]?.v;

    const hasCore = [hName, hLast, mName, mLast, ws[`F${row}`]?.v, hTable, mTable].some(
      (value) => normalize(value) !== "",
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
}

function syncTableOrder() {
  const idToNumber = new Map(state.tables.map((table) => [table.id, table.number]));
  const validIds = new Set(idToNumber.keys());
  const kept = state.tableOrder.filter((id) => validIds.has(id));
  const missing = state.tables
    .map((table) => table.id)
    .filter((id) => !kept.includes(id))
    .sort((a, b) => (idToNumber.get(a) || 0) - (idToNumber.get(b) || 0));
  state.tableOrder = [...kept, ...missing];
}

function inferTableType(guests) {
  const men = guests.filter((guest) => guest.gender === "H").length;
  const women = guests.filter((guest) => guest.gender === "M").length;
  if (!men && !women) return null;
  if (men && !women) return "men";
  if (women && !men) return "women";
  return "mixed";
}

function tableLabel(type) {
  if (type === "men") return "Solo hombres";
  if (type === "women") return "Solo mujeres";
  return "Mixta";
}

function ensureTable(number, capacity = 10) {
  const existing = state.tables.find((table) => table.number === number);
  if (existing) {
    existing.capacity = Number.isFinite(Number(capacity)) ? Number(capacity) : existing.capacity;
    return existing;
  }
  const table = {
    id: `t-${number}`,
    number,
    name: `Mesa ${number}`,
    type: null,
    capacity: Number.isFinite(Number(capacity)) ? Number(capacity) : 10,
  };
  state.tables.push(table);
  syncTableOrder();
  return table;
}

function addTables(count, capacity) {
  if (!requireEditableCurrentEvent()) return;
  const nextCount = Math.max(1, Math.trunc(Number(count) || 1));
  const nextCapacity = Math.max(1, Math.trunc(Number(capacity) || 10));
  const lastNumber = state.tables.length ? Math.max(...state.tables.map((table) => table.number)) : 0;
  for (let index = 1; index <= nextCount; index += 1) {
    const number = lastNumber + index;
    state.tables.push({
      id: `t-${number}`,
      number,
      name: `Mesa ${number}`,
      type: null,
      capacity: nextCapacity,
    });
  }
  syncTableOrder();
  render();
  scheduleRemoteSave();
}

function rebuildTablesFromCurrentAssignments() {
  const usedNumbers = state.guests
    .map((guest) => (guest.tableId ? Number(guest.tableId.split("-")[1]) : null))
    .filter((number) => Number.isFinite(number) && number > 0);
  if (!usedNumbers.length) {
    if (!state.tables.length) {
      addTables(1, 10);
    }
    return;
  }
  const max = Math.max(...usedNumbers);
  const newTables = [];
  for (let number = 1; number <= max; number += 1) {
    const existing = state.tables.find((table) => table.number === number);
    newTables.push({
      id: `t-${number}`,
      number,
      name: existing?.name || `Mesa ${number}`,
      type: existing?.type || null,
      capacity: existing?.capacity || 10,
    });
  }
  state.tables = newTables;
  syncTableOrder();
}

function canDropGuestOnTable(guest, tableId) {
  if (!tableId) return true;
  const table = state.tables.find((item) => item.id === tableId);
  if (!guest || !table) return false;
  const current = state.guests.filter((candidate) => candidate.tableId === table.id && candidate.id !== guest.id).length;
  if (current >= table.capacity) return false;
  if (table.type === "men" && guest.gender !== "H") return false;
  if (table.type === "women" && guest.gender !== "M") return false;
  return true;
}

function clearGuestDropFeedback() {
  document.querySelectorAll(".dropzone.drag-over, .dropzone.drag-allowed, .dropzone.drag-blocked").forEach((zone) => {
    zone.classList.remove("drag-over", "drag-allowed", "drag-blocked");
  });
  document.querySelectorAll(".table-card.guest-drop-allowed, .table-card.guest-drop-blocked").forEach((card) => {
    card.classList.remove("guest-drop-allowed", "guest-drop-blocked");
  });
}

function moveGuest(guestId, tableId) {
  if (!requireEditableCurrentEvent()) return;
  const guest = state.guests.find((item) => item.id === guestId);
  const table = state.tables.find((item) => item.id === tableId);
  if (!guest) return;
  if (tableId && !table) return;
  if (!canDropGuestOnTable(guest, tableId || null)) {
    showToast("No entra por capacidad o restriccion de genero.");
    return;
  }
  guest.tableId = tableId || null;
  render();
  scheduleRemoteSave();
}

function deleteGuest(guestId) {
  if (!requireEditableCurrentEvent()) return;
  const nextGuests = state.guests.filter((guest) => guest.id !== guestId);
  if (nextGuests.length === state.guests.length) return;
  state.guests = nextGuests;
  render();
  scheduleRemoteSave();
  showToast("Invitado descartado.");
}

function deleteTable(tableId) {
  if (!requireEditableCurrentEvent()) return;
  const table = state.tables.find((item) => item.id === tableId);
  if (!table) return;
  const assigned = state.guests.filter((guest) => guest.tableId === tableId).length;
  if (assigned > 0) {
    showToast("Solo se puede borrar una mesa vacia.");
    return;
  }
  state.tables = state.tables.filter((item) => item.id !== tableId);
  state.tableOrder = state.tableOrder.filter((id) => id !== tableId);
  render();
  scheduleRemoteSave();
  showToast(`${table.name} eliminada.`);
}

function renameTableNumber(tableId, nextNumberRaw) {
  if (!requireEditableCurrentEvent()) return false;
  const table = state.tables.find((item) => item.id === tableId);
  if (!table) return false;
  const nextNumber = Math.trunc(Number(nextNumberRaw));
  if (!Number.isFinite(nextNumber) || nextNumber <= 0) {
    showToast("El numero de mesa debe ser un entero mayor a 0.");
    return false;
  }
  if (table.number === nextNumber) return true;
  if (state.tables.some((item) => item.id !== tableId && item.number === nextNumber)) {
    showToast(`La mesa ${nextNumber} ya existe.`);
    return false;
  }
  const oldId = table.id;
  const oldNumber = table.number;
  const hadDefaultName = normalize(table.name) === `Mesa ${oldNumber}`;
  table.number = nextNumber;
  table.id = `t-${nextNumber}`;
  if (hadDefaultName) {
    table.name = `Mesa ${nextNumber}`;
  }
  state.guests.forEach((guest) => {
    if (guest.tableId === oldId) guest.tableId = table.id;
  });
  state.tableOrder = state.tableOrder.map((id) => (id === oldId ? table.id : id));
  if (state.dragTableId === oldId) state.dragTableId = table.id;
  syncTableOrder();
  return true;
}

function updateTableSettings(tableId, payload) {
  if (!requireEditableCurrentEvent()) return;
  const table = state.tables.find((item) => item.id === tableId);
  if (!table) return;
  const nextCapacity = Math.max(1, Math.trunc(Number(payload.capacity) || 0));
  if (!nextCapacity) {
    showToast("La capacidad debe ser mayor a 0.");
    return;
  }
  table.name = normalize(payload.name) || `Mesa ${table.number}`;
  table.capacity = nextCapacity;
  render();
  scheduleRemoteSave();
  showToast("Mesa actualizada.");
}

function enableTableNumberInlineEdit(titleEl, tableId) {
  let editing = false;

  const startEdit = () => {
    if (editing) return;
    if (!requireEditableCurrentEvent()) return;
    const table = state.tables.find((item) => item.id === tableId);
    if (!table) return;
    editing = true;
    titleEl.classList.add("editing");
    const numberSpan = titleEl.querySelector(".table-number-display");
    const input = document.createElement("input");
    input.type = "number";
    input.min = "1";
    input.step = "1";
    input.value = String(table.number);
    input.className = "table-number-inline-editor";
    input.setAttribute("aria-label", "Numero de mesa");

    const finish = (mode) => {
      if (!editing) return;
      editing = false;
      titleEl.classList.remove("editing");
      if (mode === "save") {
        const ok = renameTableNumber(tableId, input.value);
        if (ok) {
          render();
          scheduleRemoteSave();
          return;
        }
      }
      const restoreSpan = document.createElement("span");
      restoreSpan.className = "table-number-display";
      restoreSpan.textContent = String(table.number);
      if (input.isConnected) input.replaceWith(restoreSpan);
    };

    input.addEventListener("blur", () => finish("save"));
    input.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        input.blur();
      }
      if (event.key === "Escape") {
        event.preventDefault();
        finish("cancel");
      }
    });

    numberSpan.replaceWith(input);
    input.focus();
    input.select();
  };

  titleEl.addEventListener("click", (event) => {
    event.preventDefault();
    startEdit();
  });
  titleEl.addEventListener("keydown", (event) => {
    if (event.key !== "Enter" && event.key !== " ") return;
    event.preventDefault();
    startEdit();
  });
}

function splitGuestName(fullName) {
  const parts = normalize(fullName).split(/\s+/).filter(Boolean);
  if (!parts.length) return { firstName: "", lastName: "" };
  if (parts.length === 1) return { firstName: parts[0], lastName: "" };
  return { firstName: parts[0], lastName: parts.slice(1).join(" ") };
}

function askGuestData(initialData = null, options = {}) {
  const { title = "Agregar invitado", submitLabel = "Agregar" } = options;
  return new Promise((resolve) => {
    const initial = initialData || {};
    const split = splitGuestName(initial.fullName || "");
    refs.guestError.classList.add("hidden");
    refs.guestTitle.textContent = title;
    refs.guestSubmit.textContent = submitLabel;
    refs.guestFirstName.value = split.firstName || "";
    refs.guestLastName.value = split.lastName || "";
    refs.guestGender.value = ["H", "M"].includes(initial.gender) ? initial.gender : "";
    refs.guestConfirmed.value = initial.confirmed ? "yes" : "no";
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

    const submitHandler = (event) => {
      event.preventDefault();
      const firstName = normalize(refs.guestFirstName.value);
      const lastName = normalize(refs.guestLastName.value);
      const gender = normalize(refs.guestGender.value).toUpperCase();
      const confirmed = refs.guestConfirmed.value === "yes";
      if (!firstName || !lastName || !["H", "M"].includes(gender)) {
        refs.guestError.textContent = "Completa nombre, apellido, genero y confirmacion.";
        refs.guestError.classList.remove("hidden");
        return;
      }
      cleanup();
      resolve({ fullName: `${firstName} ${lastName}`.trim(), gender, confirmed });
    };

    refs.guestForm.addEventListener("submit", submitHandler);
    refs.guestCancel.addEventListener("click", cancelHandler);
  });
}

async function addGuestManually() {
  if (!requireEditableCurrentEvent()) return;
  const data = await askGuestData({ confirmed: true }, { title: "Agregar invitado", submitLabel: "Agregar" });
  if (!data) return;
  addGuest({
    name: data.fullName,
    gender: data.gender,
    confirmed: data.confirmed,
    sourceRow: null,
    initialTable: null,
  });
  render();
  scheduleRemoteSave();
  showToast("Invitado agregado.");
}

async function editGuest(guestId) {
  if (!requireEditableCurrentEvent()) return;
  const guest = state.guests.find((item) => item.id === guestId);
  if (!guest) return;
  const data = await askGuestData(
    { fullName: guest.name, gender: guest.gender, confirmed: guest.confirmed },
    { title: "Editar invitado", submitLabel: "Guardar" },
  );
  if (!data) return;
  guest.name = data.fullName;
  guest.gender = data.gender;
  guest.confirmed = data.confirmed;
  render();
  scheduleRemoteSave();
  showToast("Invitado actualizado.");
}

function filteredGuests() {
  return state.guests.filter((guest) => {
    if (state.filter.gender !== "all" && guest.gender !== state.filter.gender) return false;
    if (state.filter.search && !guest.name.toLowerCase().includes(state.filter.search.toLowerCase())) return false;
    return true;
  });
}

function guestCard(guest, options = {}) {
  const { allowDelete = false, allowEdit = true } = options;
  const isLocked = isCurrentEventLocked();
  const el = document.createElement("div");
  el.className = `guest ${guest.gender === "H" ? "male" : "female"}${isLocked ? " read-only" : ""}`;
  el.draggable = !isLocked;
  el.dataset.guestId = guest.id;
  const meta = `${guest.gender === "H" ? "Hombre" : "Mujer"}${guest.confirmed ? "" : " · no confirmado"}`;
  const editButton = allowEdit && !isLocked
    ? '<button class="guest-edit" type="button" aria-label="Editar invitado" title="Editar invitado">Editar</button>'
    : "";
  const removeButton = allowDelete && !isLocked
    ? '<button class="guest-remove" type="button" aria-label="Descartar invitado" title="Descartar invitado">Borrar</button>'
    : "";
  el.innerHTML = `
    <div class="guest-head">
      <strong>${guest.name || "(sin nombre)"}</strong>
      <div class="guest-actions">
        ${editButton}
        ${removeButton}
      </div>
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
    clearGuestDropFeedback();
  });
  if (allowEdit) {
    const editBtn = el.querySelector(".guest-edit");
    editBtn.addEventListener("mousedown", (event) => {
      event.preventDefault();
      event.stopPropagation();
    });
    editBtn.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      editGuest(guest.id);
    });
  }
  if (allowDelete) {
    const removeBtn = el.querySelector(".guest-remove");
    removeBtn.addEventListener("mousedown", (event) => {
      event.preventDefault();
      event.stopPropagation();
    });
    removeBtn.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      deleteGuest(guest.id);
    });
  }
  return el;
}

function applyDropzoneBehavior(element, tableId) {
  const card = element.closest(".table-card");

  function setGuestDropFeedback(allowed) {
    element.classList.toggle("drag-allowed", allowed);
    element.classList.toggle("drag-blocked", !allowed);
    card?.classList.toggle("guest-drop-allowed", allowed);
    card?.classList.toggle("guest-drop-blocked", !allowed);
  }

  element.addEventListener("dragover", (event) => {
    if (isCurrentEventLocked()) return;
    if (state.dragType !== "guest" || !state.dragGuestId) return;
    event.preventDefault();
    const guest = state.guests.find((item) => item.id === state.dragGuestId);
    const allowed = canDropGuestOnTable(guest, tableId || null);
    event.dataTransfer.dropEffect = allowed ? "move" : "none";
    element.classList.add("drag-over");
    setGuestDropFeedback(allowed);
  });
  element.addEventListener("dragleave", () => {
    element.classList.remove("drag-over", "drag-allowed", "drag-blocked");
    card?.classList.remove("guest-drop-allowed", "guest-drop-blocked");
  });
  element.addEventListener("drop", (event) => {
    event.preventDefault();
    if (isCurrentEventLocked()) return;
    if (state.dragType !== "guest" || !state.dragGuestId) return;
    moveGuest(state.dragGuestId, tableId || null);
    state.dragGuestId = null;
    state.dragType = null;
    clearGuestDropFeedback();
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
  handle.draggable = !isCurrentEventLocked();
  handle.addEventListener("dragstart", () => {
    if (isCurrentEventLocked()) return;
    state.dragTableId = tableId;
    state.dragType = "table";
    card.classList.add("table-dragging");
  });
  handle.addEventListener("dragend", () => {
    state.dragTableId = null;
    state.dragType = null;
    card.classList.remove("table-dragging");
  });
  card.addEventListener("dragover", (event) => {
    if (isCurrentEventLocked()) return;
    if (state.dragType !== "table") return;
    event.preventDefault();
    card.classList.add("table-drop-over");
  });
  card.addEventListener("dragleave", () => {
    card.classList.remove("table-drop-over");
  });
  card.addEventListener("drop", (event) => {
    if (isCurrentEventLocked()) return;
    if (state.dragType !== "table" || !state.dragTableId) return;
    event.preventDefault();
    card.classList.remove("table-drop-over");
    moveTableBefore(state.dragTableId, tableId);
    render();
    scheduleRemoteSave();
  });
}

function renderStats() {
  const visible = filteredGuests();
  const total = visible.length;
  const assigned = visible.filter((guest) => guest.tableId).length;
  const unassigned = total - assigned;
  const men = visible.filter((guest) => guest.gender === "H").length;
  const women = visible.filter((guest) => guest.gender === "M").length;
  const confirmed = visible.filter((guest) => guest.confirmed).length;
  refs.stats.innerHTML = `
    <div class="stats-grid">
      <div class="stat-card">
        <strong>${total}</strong>
        <span>Invitados visibles</span>
      </div>
      <div class="stat-card">
        <strong>${assigned}</strong>
        <span>Asignados</span>
      </div>
      <div class="stat-card">
        <strong>${unassigned}</strong>
        <span>Sin asignar</span>
      </div>
      <div class="stat-card">
        <strong>${state.tables.length}</strong>
        <span>Mesas</span>
      </div>
      <div class="stat-card">
        <strong>${men}/${women}</strong>
        <span>Hombres / Mujeres</span>
      </div>
      <div class="stat-card">
        <strong>${confirmed}</strong>
        <span>Confirmados</span>
      </div>
    </div>
  `;
}

function renderUnassigned(visibleGuests) {
  refs.unassignedList.innerHTML = "";
  const list = visibleGuests.filter((guest) => !guest.tableId).sort((a, b) => a.name.localeCompare(b.name, "es"));
  list.forEach((guest) => refs.unassignedList.appendChild(guestCard(guest, { allowDelete: true })));
}

function renderTableSettings(card, table) {
  const isLocked = isCurrentEventLocked();
  const form = document.createElement("form");
  form.className = "table-settings";
  form.innerHTML = `
    <div class="table-settings-top">
      <label>
        Nombre
        <input name="name" type="text" value="${table.name}" ${isLocked ? "disabled" : ""} />
      </label>
      <label>
        Capacidad
        <input name="capacity" type="number" min="1" step="1" value="${table.capacity}" ${isLocked ? "disabled" : ""} />
      </label>
    </div>
    <div class="table-settings-actions">
      <button type="submit" class="save-table-settings" ${isLocked ? "disabled" : ""}>Guardar mesa</button>
    </div>
  `;
  form.addEventListener("submit", (event) => {
    event.preventDefault();
    const fd = new FormData(form);
    updateTableSettings(table.id, {
      name: fd.get("name"),
      capacity: fd.get("capacity"),
    });
  });
  card.appendChild(form);
}

function renderTables(visibleGuests) {
  refs.tablesGrid.innerHTML = "";
  syncTableOrder();

  if (!state.currentEventId) {
    refs.tablesGrid.innerHTML = '<div class="empty-board">Crea o selecciona un evento para empezar.</div>';
    return;
  }

  if (!state.tables.length) {
    refs.tablesGrid.innerHTML = '<div class="empty-board">Todavia no hay mesas. Agrega una desde el panel de configuracion.</div>';
    return;
  }

  state.tableOrder.forEach((tableId) => {
    const table = state.tables.find((item) => item.id === tableId);
    if (!table) return;
    const assigned = visibleGuests
      .filter((guest) => guest.tableId === table.id)
      .sort((a, b) => a.name.localeCompare(b.name, "es"));
    const count = assigned.length;
    const cls =
      count > table.capacity ? "bad" : count === table.capacity ? "ok" : count >= table.capacity - 1 ? "warn" : "";
    const menCount = assigned.filter((guest) => guest.gender === "H").length;
    const womenCount = assigned.filter((guest) => guest.gender === "M").length;
    const displayType = inferTableType(assigned);
    const typePill = displayType ? `<span class="type-pill ${displayType}">${tableLabel(displayType)}</span>` : "";

    const card = document.createElement("article");
    card.className = `table-card ${cls}`;
    card.dataset.tableId = table.id;
    card.innerHTML = `
      <div class="table-head">
        <strong class="table-title-trigger" role="button" tabindex="0" title="Click para editar numero de mesa">
          Mesa <span class="table-number-display">${table.number}</span>
        </strong>
        ${typePill}
        <div class="table-actions">
          <button class="table-remove" type="button" aria-label="Eliminar mesa" title="Eliminar mesa" ${currentEvent()?.isLocked ? "disabled" : ""}>Eliminar</button>
          <span class="table-drag-handle" title="Mover mesa en el layout">${currentEvent()?.isLocked ? "Fija" : "Mover"}</span>
        </div>
      </div>
      <div class="table-meta">${count}/${table.capacity} · H:${menCount} · M:${womenCount}</div>
    `;

    const handle = card.querySelector(".table-drag-handle");
    const removeBtn = card.querySelector(".table-remove");
    const titleTrigger = card.querySelector(".table-title-trigger");
    applyTableReorderBehavior(card, table.id, handle);
    enableTableNumberInlineEdit(titleTrigger, table.id);
    removeBtn.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      deleteTable(table.id);
    });

    renderTableSettings(card, table);

    const zone = document.createElement("div");
    zone.className = "guest-list dropzone";
    zone.dataset.tableId = table.id;
    assigned.forEach((guest) => zone.appendChild(guestCard(guest)));
    applyDropzoneBehavior(zone, table.id);
    card.appendChild(zone);
    refs.tablesGrid.appendChild(card);
  });
}

function render() {
  syncCurrentEventSummary();
  const visible = filteredGuests();
  renderAccountPanel();
  renderEventsPanel();
  renderEventConfig();
  renderStats();
  renderUnassigned(visible);
  renderTables(visible);
  applyDropzoneBehavior(refs.unassignedList, null);
  updateActionAvailability();
}

function buildAssignmentRows() {
  const rows = [["TipoRegistro", "Nombre", "Genero", "Confirmado", "Mesa", "Fila Excel", "Capacidad"]];
  state.guests
    .slice()
    .sort((a, b) => a.name.localeCompare(b.name, "es"))
    .forEach((guest) => {
      const table = state.tables.find((item) => item.id === guest.tableId);
      rows.push([
        "INVITADO",
        guest.name,
        guest.gender,
        guest.confirmed ? "ok" : "",
        table ? table.number : "",
        guest.sourceRow,
        "",
      ]);
    });
  state.tables
    .slice()
    .sort((a, b) => a.number - b.number)
    .forEach((table) => {
      rows.push(["MESA", "", "", "", table.number, "", table.capacity]);
    });

  return rows;
}

function exportAssignmentsExcel() {
  const ws = XLSX.utils.aoa_to_sheet(buildAssignmentRows());
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Asignaciones");
  XLSX.writeFile(wb, "mesas_asignaciones.xlsx");
}

function downloadEmptyExcelTemplate() {
  const rows = [["TipoRegistro", "Nombre", "Genero", "Confirmado", "Mesa", "Fila Excel", "Capacidad"]];
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!cols"] = [{ wch: 14 }, { wch: 28 }, { wch: 12 }, { wch: 14 }, { wch: 10 }, { wch: 12 }, { wch: 12 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Asignaciones");
  XLSX.writeFile(wb, "template_invitados_mesas.xlsx");
}

function openExcelHelp() {
  refs.excelHelpModal.classList.remove("hidden");
  refs.downloadTemplateBtn.focus();
}

function closeExcelHelp() {
  refs.excelHelpModal.classList.add("hidden");
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
  const normalized = new Map(Object.entries(row).map(([key, value]) => [normalizeHeaderKey(key), value]));
  for (const alias of aliases) {
    if (normalized.has(alias)) return normalized.get(alias);
  }
  return "";
}

function readAssignmentRowsFromWorkbook(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { defval: "" });
}

function toNumberOrNull(value) {
  const raw = normalize(value).replace(",", ".");
  if (!raw) return null;
  const number = Number(raw);
  return Number.isFinite(number) ? number : null;
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

function upsertGuestsFromAssignmentRows(rows) {
  let changedCount = 0;
  rows.forEach((row) => {
    const tipo = normalize(getRowValue(row, ["tiporegistro", "tipo"])).toUpperCase();
    if (tipo && tipo !== "INVITADO") return;
    const genero = normalizeGender(getRowValue(row, ["genero"]));
    const nombre = normalize(getRowValue(row, ["nombre"]));
    if (!nombre || !genero) return;
    const sourceRow = toNumberOrNull(getRowValue(row, ["filaexcel", "fila"]));
    const confirmed = isConfirmedValue(getRowValue(row, ["confirmado"]));
    let guest = null;
    if (sourceRow !== null) {
      guest = state.guests.find((candidate) => candidate.sourceRow === sourceRow && candidate.gender === genero) || null;
    }
    if (!guest) {
      guest =
        state.guests.find((candidate) => candidate.name.toLowerCase() === nombre.toLowerCase() && candidate.gender === genero) ||
        null;
    }
    if (guest) {
      guest.name = nombre;
      guest.confirmed = confirmed;
      guest.sourceRow = sourceRow;
    } else {
      addGuest({
        name: nombre,
        gender: genero,
        confirmed,
        sourceRow,
        initialTable: null,
      });
    }
    changedCount += 1;
  });
  return changedCount;
}

function importAssignmentRows(rows) {
  if (!requireEditableCurrentEvent()) return;
  if (!rows.length) throw new Error("Excel vacio o invalido.");
  const importedGuestCount = upsertGuestsFromAssignmentRows(rows);
  if (!state.guests.length) {
    throw new Error("Excel invalido: no se pudieron leer invitados.");
  }

  state.guests.forEach((guest) => {
    guest.tableId = null;
  });

  const byRowGender = new Map();
  const byNameGender = new Map();
  state.guests.forEach((guest) => {
    if (Number.isFinite(guest.sourceRow)) {
      byRowGender.set(`${guest.sourceRow}|${guest.gender}`, guest);
    }
    const key = `${guest.name.toLowerCase()}|${guest.gender}`;
    if (!byNameGender.has(key)) byNameGender.set(key, []);
    byNameGender.get(key).push(guest);
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
        mesaDefs.set(mesa, capacidad !== null ? capacidad : 10);
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
      guest = candidates.find((candidate) => !candidate.tableId) || candidates[0] || null;
    }
    if (!guest) return;

    if (mesa !== null && mesa > 0) {
      guest.tableId = `t-${mesa}`;
      usedTableNumbers.add(mesa);
    }
    assignedCount += 1;
  });

  if (mesaDefs.size) {
    mesaDefs.forEach((capacity, mesa) => {
      ensureTable(mesa, capacity);
    });
  }
  if (usedTableNumbers.size) {
    [...usedTableNumbers].forEach((mesa) => ensureTable(mesa, 10));
  }
  syncTableOrder();
  render();
  scheduleRemoteSave();
  showToast(`Excel cargado: ${importedGuestCount} invitados, ${assignedCount} asignaciones.`);
}

refs.fileInput?.addEventListener("change", async (event) => {
  const file = event.target.files?.[0];
  if (!file) return;
  if (!requireEditableCurrentEvent()) {
    event.target.value = "";
    return;
  }
  try {
    const data = await file.arrayBuffer();
    parseWorkbook(data);
    rebuildTablesFromCurrentAssignments();
    render();
    scheduleRemoteSave();
    showToast(`Excel cargado: ${state.guests.length} invitados detectados.`);
  } catch (err) {
    showToast(err.message || "Error leyendo Excel.");
  } finally {
    event.target.value = "";
  }
});

refs.assignmentFileInput.addEventListener("change", async (event) => {
  const file = event.target.files?.[0];
  if (!file) return;
  if (!requireEditableCurrentEvent()) {
    event.target.value = "";
    return;
  }
  try {
    const confirmed = await askImportConfirmation(file.name);
    if (!confirmed) return;
    const data = await file.arrayBuffer();
    importAssignmentRows(readAssignmentRowsFromWorkbook(data));
  } catch (err) {
    showToast(err.message || "Error leyendo Excel.");
  } finally {
    event.target.value = "";
  }
});

refs.addGuestBtn.addEventListener("click", () => {
  addGuestManually();
});

refs.exportBtn.addEventListener("click", () => {
  if (!state.guests.length) return showToast("No hay datos para exportar.");
  exportAssignmentsExcel();
});

refs.excelHelpBtn.addEventListener("click", openExcelHelp);
refs.excelHelpClose.addEventListener("click", closeExcelHelp);
refs.downloadTemplateBtn.addEventListener("click", downloadEmptyExcelTemplate);

refs.searchInput.addEventListener("input", (event) => {
  state.filter.search = event.target.value.trim();
  render();
});

refs.genderFilter.addEventListener("change", (event) => {
  state.filter.gender = event.target.value;
  render();
});

refs.logoutBtn.addEventListener("click", () => {
  logout();
});

refs.openCreateEventBtn.addEventListener("click", () => {
  openCreateEventModal();
});

refs.createEventClose.addEventListener("click", () => {
  closeCreateEventModal();
});

refs.managementClose.addEventListener("click", () => {
  closeManagementModal();
});

refs.createEventForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  try {
    await createEvent(refs.newEventName.value);
  } catch (err) {
    showToast(parseApiError(err, "No se pudo crear el evento."));
  }
});

refs.eventSettingsForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  try {
    await renameCurrentEvent(refs.eventNameInput.value);
  } catch (err) {
    showToast(parseApiError(err, "No se pudo actualizar el evento."));
  }
});

refs.addTableForm.addEventListener("submit", (event) => {
  event.preventDefault();
  if (!requireEditableCurrentEvent()) return;
  addTables(1, refs.newTableCapacity.value);
  showToast("Mesa agregada.");
});

refs.bulkTablesForm.addEventListener("submit", (event) => {
  event.preventDefault();
  if (!requireEditableCurrentEvent()) return;
  addTables(refs.bulkTableCount.value, refs.bulkTableCapacity.value);
  showToast("Bloque de mesas agregado.");
});

async function init() {
  render();
  try {
    await ensureSession();
    render();
    if (state.currentEventId) {
      await loadRemoteSnapshot();
    }
  } catch (err) {
    showToast(parseApiError(err, "No se pudo iniciar la aplicacion."));
  }
}

init();
