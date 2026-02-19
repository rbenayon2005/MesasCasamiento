const SCHEMA_SQL = `
CREATE TABLE IF NOT EXISTS app_meta (
  key TEXT PRIMARY KEY,
  value TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS tables (
  number INTEGER PRIMARY KEY,
  name TEXT NOT NULL,
  type TEXT,
  capacity INTEGER NOT NULL,
  position INTEGER NOT NULL
);

CREATE TABLE IF NOT EXISTS guests (
  id TEXT PRIMARY KEY,
  name TEXT NOT NULL,
  gender TEXT NOT NULL,
  confirmed INTEGER NOT NULL DEFAULT 0,
  source_row INTEGER,
  table_number INTEGER
);
`;

function toIntBool(value) {
  return value ? 1 : 0;
}

function toGuest(row) {
  const tableNumber = Number.isFinite(Number(row.table_number)) ? Number(row.table_number) : null;
  return {
    id: String(row.id),
    name: String(row.name || ""),
    gender: String(row.gender || "").toUpperCase() === "M" ? "M" : "H",
    confirmed: Boolean(row.confirmed),
    sourceRow: Number.isFinite(Number(row.source_row)) ? Number(row.source_row) : null,
    tableId: tableNumber ? `t-${tableNumber}` : null,
  };
}

function toTable(row) {
  return {
    id: `t-${Number(row.number)}`,
    number: Number(row.number),
    name: String(row.name || `Mesa ${Number(row.number)}`),
    type: row.type === "men" || row.type === "women" ? row.type : null,
    capacity: Number.isFinite(Number(row.capacity)) ? Number(row.capacity) : Number(row.number) === 1 ? 20 : 10,
    position: Number.isFinite(Number(row.position)) ? Number(row.position) : Number(row.number),
  };
}

function parseTableNumber(tableId) {
  if (!tableId) return null;
  const value = Number(String(tableId).split("-")[1]);
  return Number.isFinite(value) && value > 0 ? value : null;
}

function normalizePayload(payload) {
  const tablesInput = Array.isArray(payload?.tables) ? payload.tables : [];
  const guestsInput = Array.isArray(payload?.guests) ? payload.guests : [];
  const orderInput = Array.isArray(payload?.tableOrder) ? payload.tableOrder : [];

  const byNumber = new Map();
  tablesInput.forEach((table) => {
    const number = Number(table?.number);
    if (!Number.isFinite(number) || number <= 0) return;
    byNumber.set(number, {
      number,
      name: String(table?.name || `Mesa ${number}`),
      type: table?.type === "men" || table?.type === "women" ? table.type : null,
      capacity: Number.isFinite(Number(table?.capacity)) ? Number(table.capacity) : number === 1 ? 20 : 10,
      position: number,
    });
  });

  orderInput.forEach((tableId, idx) => {
    const number = parseTableNumber(tableId);
    if (!number || !byNumber.has(number)) return;
    byNumber.get(number).position = idx + 1;
  });

  const orderedTables = [...byNumber.values()].sort((a, b) => a.position - b.position || a.number - b.number);
  if (!orderedTables.length) {
    orderedTables.push({ number: 1, name: "Mesa 1", type: null, capacity: 20, position: 1 });
  }
  const tableNumbers = new Set(orderedTables.map((t) => t.number));

  const guests = guestsInput
    .map((guest) => {
      const id = String(guest?.id || "").trim();
      if (!id) return null;
      const name = String(guest?.name || "").trim();
      if (!name) return null;
      const gender = String(guest?.gender || "").toUpperCase() === "M" ? "M" : "H";
      const tableNumber = parseTableNumber(guest?.tableId);
      return {
        id,
        name,
        gender,
        confirmed: toIntBool(Boolean(guest?.confirmed)),
        sourceRow: Number.isFinite(Number(guest?.sourceRow)) ? Number(guest.sourceRow) : null,
        tableNumber: tableNumber && tableNumbers.has(tableNumber) ? tableNumber : null,
      };
    })
    .filter(Boolean);

  return { tables: orderedTables, guests };
}

export async function ensureSchema(db) {
  const statements = SCHEMA_SQL.trim()
    .split(";")
    .map((s) => s.trim())
    .filter(Boolean);
  for (const statement of statements) {
    await db.prepare(statement).run();
  }
}

export async function readRevision(db) {
  const row = await db.prepare("SELECT value FROM app_meta WHERE key = 'revision'").first();
  const revision = Number(row?.value || 0);
  return Number.isFinite(revision) ? revision : 0;
}

export async function bumpRevision(db) {
  const current = await readRevision(db);
  const next = current + 1;
  await db
    .prepare(
      `INSERT INTO app_meta (key, value) VALUES ('revision', ?)
       ON CONFLICT(key) DO UPDATE SET value = excluded.value`,
    )
    .bind(String(next))
    .run();
  return next;
}

export async function loadState(db) {
  await ensureSchema(db);
  const revision = await readRevision(db);

  const tablesResult = await db
    .prepare("SELECT number, name, type, capacity, position FROM tables ORDER BY position ASC, number ASC")
    .all();
  const guestsResult = await db
    .prepare("SELECT id, name, gender, confirmed, source_row, table_number FROM guests ORDER BY name COLLATE NOCASE ASC")
    .all();

  const tables = (tablesResult.results || []).map(toTable);
  const guests = (guestsResult.results || []).map(toGuest);

  return {
    revision,
    guests,
    tables: tables.map(({ position, ...table }) => table),
    tableOrder: tables.map((t) => t.id),
  };
}

export async function saveState(db, payload) {
  await ensureSchema(db);
  const normalized = normalizePayload(payload);

  await db.prepare("DELETE FROM tables").run();
  await db.prepare("DELETE FROM guests").run();

  if (normalized.tables.length) {
    const tableStatements = normalized.tables.map((table) =>
      db
        .prepare("INSERT INTO tables (number, name, type, capacity, position) VALUES (?, ?, ?, ?, ?)")
        .bind(table.number, table.name, table.type, table.capacity, table.position),
    );
    await db.batch(tableStatements);
  }

  if (normalized.guests.length) {
    const guestStatements = normalized.guests.map((guest) =>
      db
        .prepare(
          "INSERT INTO guests (id, name, gender, confirmed, source_row, table_number) VALUES (?, ?, ?, ?, ?, ?)",
        )
        .bind(guest.id, guest.name, guest.gender, guest.confirmed, guest.sourceRow, guest.tableNumber),
    );
    await db.batch(guestStatements);
  }

  const revision = await bumpRevision(db);
  return {
    revision,
    saved: {
      guests: normalized.guests.length,
      tables: normalized.tables.length,
    },
  };
}
