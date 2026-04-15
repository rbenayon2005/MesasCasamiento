import { loadEventState, requireSessionUser, saveEventState } from "./_db.js";

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

function getEventId(request) {
  const url = new URL(request.url);
  return String(url.searchParams.get("eventId") || "").trim();
}

export async function onRequestGet(context) {
  const session = await requireSessionUser(context.env.DB, context.request);
  if (!session) return unauthorized();

  try {
    const eventId = getEventId(context.request);
    if (!eventId) return json({ error: "Falta eventId." }, 400);
    const state = await loadEventState(context.env.DB, session.user.id, eventId);
    const requestedRevision = Number(new URL(context.request.url).searchParams.get("revision") || 0);
    if (Number.isFinite(requestedRevision) && requestedRevision > 0 && requestedRevision === state.revision) {
      return json({ changed: false, revision: state.revision });
    }
    return json({ changed: true, ...state });
  } catch (err) {
    return json({ error: String(err?.message || err) }, 400);
  }
}

export async function onRequestPost(context) {
  const session = await requireSessionUser(context.env.DB, context.request);
  if (!session) return unauthorized();

  try {
    const eventId = getEventId(context.request);
    if (!eventId) return json({ error: "Falta eventId." }, 400);
    const payload = await context.request.json();
    const result = await saveEventState(context.env.DB, session.user.id, eventId, payload);
    return json(result);
  } catch (err) {
    return json({ error: String(err?.message || err) }, 400);
  }
}
