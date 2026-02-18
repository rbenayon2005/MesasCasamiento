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
    onlyConfirmed: true,
  },
};

const refs = {
  fileInput: document.getElementById("excelFile"),
  csvInput: document.getElementById("csvFile"),
  loadCurrentBtn: document.getElementById("loadCurrentBtn"),
  newLayoutBtn: document.getElementById("newLayoutBtn"),
  addOneTableBtn: document.getElementById("addOneTableBtn"),
  exportBtn: document.getElementById("exportBtn"),
  menTables: document.getElementById("menTables"),
  womenTables: document.getElementById("womenTables"),
  mixedTables: document.getElementById("mixedTables"),
  searchInput: document.getElementById("searchInput"),
  genderFilter: document.getElementById("genderFilter"),
  onlyConfirmed: document.getElementById("onlyConfirmed"),
  unassignedList: document.getElementById("unassignedList"),
  tablesGrid: document.getElementById("tablesGrid"),
  stats: document.getElementById("stats"),
  toast: document.getElementById("toast"),
};

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

function buildNewLayout() {
  state.guests.forEach((g) => {
    g.tableId = null;
  });

  const menTables = Number(refs.menTables.value || 0);
  const womenTables = Number(refs.womenTables.value || 0);
  const mixedTables = Number(refs.mixedTables.value || 0);

  let n = 1;
  state.tables = [{ id: `t-${n}`, number: n, name: `Mesa ${n}`, type: null, capacity: 20 }];
  n += 1;

  for (let i = 0; i < menTables; i += 1, n += 1) {
    state.tables.push({ id: `t-${n}`, number: n, name: `Mesa ${n}`, type: "men", capacity: 10 });
  }
  for (let i = 0; i < womenTables; i += 1, n += 1) {
    state.tables.push({ id: `t-${n}`, number: n, name: `Mesa ${n}`, type: "women", capacity: 10 });
  }
  for (let i = 0; i < mixedTables; i += 1, n += 1) {
    state.tables.push({ id: `t-${n}`, number: n, name: `Mesa ${n}`, type: null, capacity: 10 });
  }
  syncTableOrder();
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
}

function filteredGuests() {
  return state.guests.filter((g) => {
    if (state.filter.onlyConfirmed && !g.confirmed) return false;
    if (state.filter.gender !== "all" && g.gender !== state.filter.gender) return false;
    if (state.filter.search) {
      const q = state.filter.search.toLowerCase();
      if (!g.name.toLowerCase().includes(q)) return false;
    }
    return true;
  });
}

function guestCard(guest) {
  const el = document.createElement("div");
  el.className = `guest ${guest.gender === "H" ? "male" : "female"}`;
  el.draggable = true;
  el.dataset.guestId = guest.id;
  el.innerHTML = `<strong>${guest.name || "(sin nombre)"}</strong><br><small>${guest.gender === "H" ? "Hombre" : "Mujer"}${guest.confirmed ? "" : " - no confirmado"}</small>`;
  el.addEventListener("dragstart", () => {
    state.dragGuestId = guest.id;
    state.dragType = "guest";
  });
  el.addEventListener("dragend", () => {
    state.dragGuestId = null;
    state.dragType = null;
  });
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
  list.forEach((g) => refs.unassignedList.appendChild(guestCard(g)));
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
          <span class="table-drag-handle" title="Mover mesa en el layout">Mover</span>
          ${typePill || bigPill}
        </div>
        <div class="table-meta">${count}/${table.capacity} | H:${menCount} M:${womenCount}</div>
      `;
      const handle = card.querySelector(".table-drag-handle");
      applyTableReorderBehavior(card, table.id, handle);

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

function importAssignmentCsv(text) {
  const wb = XLSX.read(text, { type: "string" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
  if (!rows.length) throw new Error("CSV vacio o invalido.");
  if (!state.guests.length) throw new Error("Primero carga el Excel.");

  state.guests.forEach((g) => {
    g.tableId = null;
  });
  state.tables = [];

  const byRowGender = new Map();
  const byNameGender = new Map();
  state.guests.forEach((g) => {
    byRowGender.set(`${g.sourceRow}|${g.gender}`, g);
    const key = `${g.name.toLowerCase()}|${g.gender}`;
    if (!byNameGender.has(key)) byNameGender.set(key, []);
    byNameGender.get(key).push(g);
  });

  let assignedCount = 0;
  const mesaDefs = new Map();
  const usedTableNumbers = new Set();
  rows.forEach((row) => {
    const tipo = normalize(row.TipoRegistro || row.tipo || "");
    const genero = normalize(row.Genero || row.genero || "").toUpperCase();
    const nombre = normalize(row.Nombre || row.nombre || "");
    const fila = Number(row["Fila Excel"] || row.fila_excel || row.Fila || "");
    const mesa = Number(row.Mesa || row.mesa || "");
    const capacidad = Number(row.Capacidad || row.capacidad || "");

    if (tipo === "MESA") {
      if (Number.isFinite(mesa) && mesa > 0) {
        mesaDefs.set(mesa, Number.isFinite(capacidad) ? capacidad : mesa === 1 ? 20 : 10);
      }
      return;
    }

    if (!(Number.isFinite(mesa) && mesa > 0) && !Number.isFinite(fila) && !nombre) return;
    if (!["H", "M"].includes(genero)) return;

    let guest = null;
    if (Number.isFinite(fila)) {
      guest = byRowGender.get(`${fila}|${genero}`) || null;
    }
    if (!guest && nombre) {
      const key = `${nombre.toLowerCase()}|${genero}`;
      const candidates = byNameGender.get(key) || [];
      guest = candidates.find((c) => !c.tableId) || candidates[0] || null;
    }
    if (!guest) return;

    if (Number.isFinite(mesa) && mesa > 0) {
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
}

refs.fileInput.addEventListener("change", async (e) => {
  const file = e.target.files?.[0];
  if (!file) return;
  const data = await file.arrayBuffer();
  try {
    parseWorkbook(data);
    rebuildTablesFromCurrentAssignments();
    render();
  } catch (err) {
    showToast(err.message || "Error leyendo Excel");
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
  }
});

refs.loadCurrentBtn.addEventListener("click", () => {
  if (!state.guests.length) return showToast("Primero carga el Excel.");
  rebuildTablesFromCurrentAssignments();
  render();
});

refs.newLayoutBtn.addEventListener("click", () => {
  if (!state.guests.length) return showToast("Primero carga el Excel.");
  buildNewLayout();
  render();
});

refs.addOneTableBtn.addEventListener("click", () => {
  if (!state.guests.length) return showToast("Primero carga el Excel.");
  addMoreTables(1);
  render();
  showToast("Se agrego 1 mesa mixta de 10.");
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
refs.onlyConfirmed.addEventListener("change", (e) => {
  state.filter.onlyConfirmed = e.target.checked;
  render();
});

render();
