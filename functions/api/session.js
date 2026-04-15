import { deleteSession, loginUser, registerUser, requireSessionUser } from "./_db.js";

function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      "Cache-Control": "no-store",
    },
  });
}

export async function onRequestPost(context) {
  try {
    const payload = await context.request.json();
    const action = String(payload?.action || "").trim();
    if (action === "register") {
      const session = await registerUser(context.env.DB, payload);
      return json(session, 201);
    }
    if (action === "login") {
      const session = await loginUser(context.env.DB, payload);
      return json(session, 200);
    }
    return json({ error: "Accion invalida." }, 400);
  } catch (err) {
    return json({ error: String(err?.message || err) }, 400);
  }
}

export async function onRequestGet(context) {
  const session = await requireSessionUser(context.env.DB, context.request);
  if (!session) {
    return json({ error: "Sesion invalida o vencida." }, 401);
  }
  return json({ user: session.user });
}

export async function onRequestDelete(context) {
  const token = context.request.headers.get("x-session-token") || "";
  await deleteSession(context.env.DB, token);
  return json({ ok: true });
}
