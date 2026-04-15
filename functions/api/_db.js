const SCHEMA_SQL = `
CREATE TABLE IF NOT EXISTS users (
  id TEXT PRIMARY KEY,
  name TEXT NOT NULL,
  email TEXT NOT NULL UNIQUE,
  password_hash TEXT NOT NULL,
  created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS sessions (
  token TEXT PRIMARY KEY,
  user_id TEXT NOT NULL,
  created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS events (
  id TEXT PRIMARY KEY,
  user_id TEXT NOT NULL,
  name TEXT NOT NULL,
  created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS event_meta (
  event_id TEXT NOT NULL,
  key TEXT NOT NULL,
  value TEXT NOT NULL,
  PRIMARY KEY (event_id, key)
);

CREATE TABLE IF NOT EXISTS tables (
  event_id TEXT NOT NULL,
  number INTEGER NOT NULL,
  name TEXT NOT NULL,
  type TEXT,
  capacity INTEGER NOT NULL,
  position INTEGER NOT NULL,
  PRIMARY KEY (event_id, number)
);

CREATE TABLE IF NOT EXISTS guests (
  id TEXT PRIMARY KEY,
  event_id TEXT NOT NULL,
  name TEXT NOT NULL,
  gender TEXT NOT NULL,
  confirmed INTEGER NOT NULL DEFAULT 0,
  source_row INTEGER,
  table_number INTEGER
);
`;

async function tableExists(db, tableName) {
  const row = await db
    .prepare("SELECT name FROM sqlite_master WHERE type = 'table' AND name = ?")
    .bind(tableName)
    .first();
  return Boolean(row?.name);
}

async function tableHasColumn(db, tableName, columnName) {
  const result = await db.prepare(`PRAGMA table_info(${tableName})`).all();
  return (result.results || []).some((row) => String(row.name) === columnName);
}

async function recreateLegacyStateTables(db) {
  const hasTables = await tableExists(db, "tables");
  const hasGuests = await tableExists(db, "guests");

  if (hasTables) {
    const tablesHasEventId = await tableHasColumn(db, "tables", "event_id");
    if (!tablesHasEventId) {
      await db.prepare("DROP TABLE tables").run();
    }
  }

  if (hasGuests) {
    const guestsHasEventId = await tableHasColumn(db, "guests", "event_id");
    if (!guestsHasEventId) {
      await db.prepare("DROP TABLE guests").run();
    }
  }
}

function toIntBool(value) {
  return value ? 1 : 0;
}

function normalizeEmail(value) {
  return String(value || "").trim().toLowerCase();
}

function parseTableNumber(tableId) {
  if (!tableId) return null;
  const value = Number(String(tableId).split("-")[1]);
  return Number.isFinite(value) && value > 0 ? value : null;
}

async function sha256(input) {
  const buffer = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(input));
  return [...new Uint8Array(buffer)].map((part) => part.toString(16).padStart(2, "0")).join("");
}

function nowIso() {
  return new Date().toISOString();
}

function jsonClone(value) {
  return JSON.parse(JSON.stringify(value));
}

export async function ensureSchema(db) {
  await recreateLegacyStateTables(db);
  const statements = SCHEMA_SQL.trim()
    .split(";")
    .map((part) => part.trim())
    .filter(Boolean);
  for (const statement of statements) {
    await db.prepare(statement).run();
  }
}

async function getUserByEmail(db, email) {
  return db.prepare("SELECT id, name, email, password_hash FROM users WHERE email = ?").bind(email).first();
}

async function getSession(db, token) {
  return db
    .prepare(
      `SELECT s.token, u.id, u.name, u.email
       FROM sessions s
       JOIN users u ON u.id = s.user_id
       WHERE s.token = ?`,
    )
    .bind(token)
    .first();
}

async function readEventRevision(db, eventId) {
  const row = await db
    .prepare("SELECT value FROM event_meta WHERE event_id = ? AND key = 'revision'")
    .bind(eventId)
    .first();
  const revision = Number(row?.value || 0);
  return Number.isFinite(revision) ? revision : 0;
}

async function bumpEventRevision(db, eventId) {
  const current = await readEventRevision(db, eventId);
  const next = current + 1;
  await db
    .prepare(
      `INSERT INTO event_meta (event_id, key, value) VALUES (?, 'revision', ?)
       ON CONFLICT(event_id, key) DO UPDATE SET value = excluded.value`,
    )
    .bind(eventId, String(next))
    .run();
  return next;
}

async function getOwnedEvent(db, userId, eventId) {
  return db
    .prepare("SELECT id, user_id, name, created_at FROM events WHERE id = ? AND user_id = ?")
    .bind(eventId, userId)
    .first();
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
      capacity: Math.max(1, Number(table?.capacity) || 10),
      position: number,
    });
  });

  orderInput.forEach((tableId, index) => {
    const number = parseTableNumber(tableId);
    if (!number || !byNumber.has(number)) return;
    byNumber.get(number).position = index + 1;
  });

  const orderedTables = [...byNumber.values()].sort((a, b) => a.position - b.position || a.number - b.number);
  const tableNumbers = new Set(orderedTables.map((table) => table.number));

  const guests = guestsInput
    .map((guest) => {
      const id = String(guest?.id || "").trim();
      const name = String(guest?.name || "").trim();
      if (!id || !name) return null;
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
    capacity: Number.isFinite(Number(row.capacity)) ? Number(row.capacity) : 10,
    position: Number.isFinite(Number(row.position)) ? Number(row.position) : Number(row.number),
  };
}

export async function registerUser(db, payload) {
  await ensureSchema(db);
  const name = String(payload?.name || "").trim();
  const email = normalizeEmail(payload?.email);
  const password = String(payload?.password || "");
  if (!name || !email || !password) {
    throw new Error("Nombre, email y password son requeridos.");
  }

  const existing = await getUserByEmail(db, email);
  if (existing) {
    throw new Error("Ya existe una cuenta con ese email.");
  }

  const userId = crypto.randomUUID();
  const token = `${crypto.randomUUID()}${crypto.randomUUID()}`;
  const createdAt = nowIso();
  const passwordHash = await sha256(password);

  await db
    .prepare("INSERT INTO users (id, name, email, password_hash, created_at) VALUES (?, ?, ?, ?, ?)")
    .bind(userId, name, email, passwordHash, createdAt)
    .run();
  await db
    .prepare("INSERT INTO sessions (token, user_id, created_at) VALUES (?, ?, ?)")
    .bind(token, userId, createdAt)
    .run();

  return {
    token,
    user: { id: userId, name, email },
  };
}

export async function loginUser(db, payload) {
  await ensureSchema(db);
  const email = normalizeEmail(payload?.email);
  const password = String(payload?.password || "");
  if (!email || !password) {
    throw new Error("Email y password son requeridos.");
  }

  const user = await getUserByEmail(db, email);
  if (!user) {
    throw new Error("Credenciales invalidas.");
  }

  const passwordHash = await sha256(password);
  if (passwordHash !== user.password_hash) {
    throw new Error("Credenciales invalidas.");
  }

  const token = `${crypto.randomUUID()}${crypto.randomUUID()}`;
  await db
    .prepare("INSERT INTO sessions (token, user_id, created_at) VALUES (?, ?, ?)")
    .bind(token, user.id, nowIso())
    .run();

  return {
    token,
    user: { id: user.id, name: user.name, email: user.email },
  };
}

export async function deleteSession(db, token) {
  await ensureSchema(db);
  if (!token) return;
  await db.prepare("DELETE FROM sessions WHERE token = ?").bind(token).run();
}

export async function requireSessionUser(db, request) {
  await ensureSchema(db);
  const token = request.headers.get("x-session-token") || "";
  if (!token) return null;
  const session = await getSession(db, token);
  if (!session) return null;
  return {
    token,
    user: {
      id: session.id,
      name: session.name,
      email: session.email,
    },
  };
}

export async function listUserEvents(db, userId) {
  await ensureSchema(db);
  const eventsResult = await db
    .prepare("SELECT id, name, created_at FROM events WHERE user_id = ? ORDER BY created_at DESC")
    .bind(userId)
    .all();

  const events = eventsResult.results || [];
  const enriched = [];

  for (const event of events) {
    const tableRow = await db.prepare("SELECT COUNT(*) AS count FROM tables WHERE event_id = ?").bind(event.id).first();
    const guestRow = await db.prepare("SELECT COUNT(*) AS count FROM guests WHERE event_id = ?").bind(event.id).first();

    enriched.push({
      id: String(event.id),
      name: String(event.name),
      createdAt: String(event.created_at),
      tableCount: Number(tableRow?.count || 0),
      guestCount: Number(guestRow?.count || 0),
    });
  }

  return enriched;
}

export async function createEvent(db, userId, payload) {
  await ensureSchema(db);
  const name = String(payload?.name || "").trim();
  if (!name) {
    throw new Error("El evento necesita un nombre.");
  }

  const eventId = crypto.randomUUID();
  const createdAt = nowIso();
  await db
    .prepare("INSERT INTO events (id, user_id, name, created_at) VALUES (?, ?, ?, ?)")
    .bind(eventId, userId, name, createdAt)
    .run();
  await db
    .prepare("INSERT INTO event_meta (event_id, key, value) VALUES (?, 'revision', '0')")
    .bind(eventId)
    .run();

  return {
    id: eventId,
    name,
    createdAt,
    tableCount: 0,
    guestCount: 0,
  };
}

export async function updateEvent(db, userId, payload) {
  await ensureSchema(db);
  const eventId = String(payload?.eventId || "").trim();
  const name = String(payload?.name || "").trim();
  if (!eventId || !name) {
    throw new Error("Evento y nombre son requeridos.");
  }

  const event = await getOwnedEvent(db, userId, eventId);
  if (!event) {
    throw new Error("Evento no encontrado.");
  }

  await db.prepare("UPDATE events SET name = ? WHERE id = ?").bind(name, eventId).run();
  return { ...jsonClone(event), name };
}

export async function deleteEvent(db, userId, eventId) {
  await ensureSchema(db);
  const normalizedEventId = String(eventId || "").trim();
  if (!normalizedEventId) {
    throw new Error("Falta eventId.");
  }

  const event = await getOwnedEvent(db, userId, normalizedEventId);
  if (!event) {
    throw new Error("Evento no encontrado.");
  }

  await db.prepare("DELETE FROM guests WHERE event_id = ?").bind(normalizedEventId).run();
  await db.prepare("DELETE FROM tables WHERE event_id = ?").bind(normalizedEventId).run();
  await db.prepare("DELETE FROM event_meta WHERE event_id = ?").bind(normalizedEventId).run();
  await db.prepare("DELETE FROM events WHERE id = ? AND user_id = ?").bind(normalizedEventId, userId).run();
}

export async function loadEventState(db, userId, eventId) {
  await ensureSchema(db);
  const event = await getOwnedEvent(db, userId, eventId);
  if (!event) {
    throw new Error("Evento no encontrado.");
  }

  const revision = await readEventRevision(db, eventId);
  const tablesResult = await db
    .prepare(
      "SELECT number, name, type, capacity, position FROM tables WHERE event_id = ? ORDER BY position ASC, number ASC",
    )
    .bind(eventId)
    .all();
  const guestsResult = await db
    .prepare(
      "SELECT id, name, gender, confirmed, source_row, table_number FROM guests WHERE event_id = ? ORDER BY name COLLATE NOCASE ASC",
    )
    .bind(eventId)
    .all();

  const tables = (tablesResult.results || []).map(toTable);
  const guests = (guestsResult.results || []).map(toGuest);

  return {
    revision,
    guests,
    tables: tables.map(({ position, ...table }) => table),
    tableOrder: tables.map((table) => table.id),
  };
}

export async function saveEventState(db, userId, eventId, payload) {
  await ensureSchema(db);
  const event = await getOwnedEvent(db, userId, eventId);
  if (!event) {
    throw new Error("Evento no encontrado.");
  }

  const normalized = normalizePayload(payload);
  await db.prepare("DELETE FROM tables WHERE event_id = ?").bind(eventId).run();
  await db.prepare("DELETE FROM guests WHERE event_id = ?").bind(eventId).run();

  if (normalized.tables.length) {
    const tableStatements = normalized.tables.map((table) =>
      db
        .prepare("INSERT INTO tables (event_id, number, name, type, capacity, position) VALUES (?, ?, ?, ?, ?, ?)")
        .bind(eventId, table.number, table.name, table.type, table.capacity, table.position),
    );
    await db.batch(tableStatements);
  }

  if (normalized.guests.length) {
    const guestStatements = normalized.guests.map((guest) =>
      db
        .prepare(
          `INSERT INTO guests (id, event_id, name, gender, confirmed, source_row, table_number)
           VALUES (?, ?, ?, ?, ?, ?, ?)`,
        )
        .bind(guest.id, eventId, guest.name, guest.gender, guest.confirmed, guest.sourceRow, guest.tableNumber),
    );
    await db.batch(guestStatements);
  }

  const revision = await bumpEventRevision(db, eventId);
  return {
    revision,
    saved: {
      guests: normalized.guests.length,
      tables: normalized.tables.length,
    },
  };
}
