import { loadState, saveState } from "./_db";

const AUTH_USER = "adminmesas";
const AUTH_PASS = "mesas2026";

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
  return json({ error: "No autorizado" }, 401);
}

function isAuthorized(request) {
  const user = request.headers.get("x-auth-user") || "";
  const pass = request.headers.get("x-auth-pass") || "";
  return user === AUTH_USER && pass === AUTH_PASS;
}

export async function onRequestGet(context) {
  if (!isAuthorized(context.request)) return unauthorized();
  try {
    const state = await loadState(context.env.DB);
    const requestedRevision = Number(context.request.url ? new URL(context.request.url).searchParams.get("revision") : 0);
    if (Number.isFinite(requestedRevision) && requestedRevision > 0 && requestedRevision === state.revision) {
      return json({ changed: false, revision: state.revision });
    }
    return json({ changed: true, ...state });
  } catch (err) {
    return json({ error: "Error leyendo estado D1", detail: String(err?.message || err) }, 500);
  }
}

export async function onRequestPost(context) {
  if (!isAuthorized(context.request)) return unauthorized();
  try {
    const payload = await context.request.json();
    const result = await saveState(context.env.DB, payload);
    return json(result, 200);
  } catch (err) {
    return json({ error: "Error guardando estado D1", detail: String(err?.message || err) }, 500);
  }
}
