import { createEvent, deleteEvent, listUserEvents, requireSessionUser, updateEvent } from "./_db.js";

function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      "Cache-Control": "no-store",
    },
  });
}

function unauthorized() {
  return json({ error: "Sesion invalida o vencida." }, 401);
}

export async function onRequestGet(context) {
  const session = await requireSessionUser(context.env.DB, context.request);
  if (!session) return unauthorized();

  try {
    const events = await listUserEvents(context.env.DB, session.user.id);
    return json({ events });
  } catch (err) {
    return json({ error: String(err?.message || err), detail: String(err?.stack || "") }, 400);
  }
}

export async function onRequestPost(context) {
  const session = await requireSessionUser(context.env.DB, context.request);
  if (!session) return unauthorized();

  try {
    const payload = await context.request.json();
    const action = String(payload?.action || "").trim();
    if (action === "create") {
      const event = await createEvent(context.env.DB, session.user.id, payload);
      const events = await listUserEvents(context.env.DB, session.user.id);
      return json({ event, events }, 201);
    }
    if (action === "update") {
      await updateEvent(context.env.DB, session.user.id, payload);
      const events = await listUserEvents(context.env.DB, session.user.id);
      return json({ events });
    }
    if (action === "delete") {
      await deleteEvent(context.env.DB, session.user.id, payload?.eventId);
      const events = await listUserEvents(context.env.DB, session.user.id);
      return json({ events });
    }
    return json({ error: "Accion invalida." }, 400);
  } catch (err) {
    return json({ error: String(err?.message || err), detail: String(err?.stack || "") }, 400);
  }
}
