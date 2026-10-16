// Stores the invigilation timetable in Netlify Blobs and checks passwords.
//
// Passwords are set in Netlify under Site configuration > Environment variables:
//   ADMIN_PASSWORD   - required, for the exams officer's admin page
//   VIEWER_PASSWORD  - optional, shared with invigilators. Leave unset to make
//                      the timetable viewable by anyone with the link.

import { getStore } from "@netlify/blobs";
import { createHash, timingSafeEqual } from "node:crypto";

const CURRENT_KEY = "current";
const BACKUP_PREFIX = "backups/";
const MAX_BACKUPS = 30;
const MAX_BODY_BYTES = 2_000_000;

export const config = {
  path: ["/api/timetable", "/api/versions", "/api/restore"],
};

export default async (req) => {
  const store = getStore({ name: "invigilation", consistency: "strong" });
  return handle(req, store, process.env);
};

export async function handle(req, store, env) {
  const url = new URL(req.url);
  const role = getRole(req, env);

  if (!role) {
    // Slow down password guessing a little.
    await new Promise((r) => setTimeout(r, 400));
    return json(401, { error: "Incorrect password." });
  }

  try {
    if (url.pathname === "/api/timetable" && req.method === "GET") {
      const data = await store.get(CURRENT_KEY, { type: "json" });
      return json(200, { role, data: data ?? null });
    }

    // Everything below needs the admin password.
    if (role !== "admin") return json(403, { error: "Admin password required." });

    if (url.pathname === "/api/timetable" && req.method === "PUT") {
      const body = await readBody(req);
      const problem = validate(body?.data);
      if (problem) return json(400, { error: problem });

      const current = await store.get(CURRENT_KEY, { type: "json" });
      if (current && !body.force && current.updatedAt !== body.baseUpdatedAt) {
        return json(409, {
          error:
            "Someone else saved changes since you opened this page. Reload to see their version, or save anyway to overwrite it.",
          updatedAt: current.updatedAt,
        });
      }

      const saved = await saveWithBackup(store, current, body.data);
      return json(200, { data: saved });
    }

    if (url.pathname === "/api/versions" && req.method === "GET") {
      const { blobs } = await store.list({ prefix: BACKUP_PREFIX });
      const versions = blobs
        .map((b) => b.key)
        .sort()
        .reverse()
        .map((key) => ({ key, savedAt: key.slice(BACKUP_PREFIX.length) }));
      return json(200, { versions });
    }

    if (url.pathname === "/api/restore" && req.method === "POST") {
      const body = await readBody(req);
      const key = String(body?.key ?? "");
      if (!key.startsWith(BACKUP_PREFIX)) return json(400, { error: "Unknown version." });
      const old = await store.get(key, { type: "json" });
      if (!old) return json(404, { error: "That version no longer exists." });
      const current = await store.get(CURRENT_KEY, { type: "json" });
      const saved = await saveWithBackup(store, current, old);
      return json(200, { data: saved });
    }

    return json(404, { error: "Not found." });
  } catch (err) {
    console.error(err);
    return json(500, { error: "Something went wrong on the server. Please try again." });
  }
}

async function saveWithBackup(store, current, data) {
  if (current?.updatedAt) {
    await store.setJSON(BACKUP_PREFIX + current.updatedAt, current);
    await pruneBackups(store);
  }
  const saved = { ...data, updatedAt: new Date().toISOString() };
  await store.setJSON(CURRENT_KEY, saved);
  return saved;
}

async function pruneBackups(store) {
  const { blobs } = await store.list({ prefix: BACKUP_PREFIX });
  const keys = blobs.map((b) => b.key).sort();
  for (const key of keys.slice(0, Math.max(0, keys.length - MAX_BACKUPS))) {
    await store.delete(key);
  }
}

function getRole(req, env) {
  const given = req.headers.get("x-password") ?? "";
  if (env.ADMIN_PASSWORD && same(given, env.ADMIN_PASSWORD)) return "admin";
  if (!env.VIEWER_PASSWORD) return "viewer";
  if (same(given, env.VIEWER_PASSWORD)) return "viewer";
  return null;
}

function same(a, b) {
  const ha = createHash("sha256").update(String(a)).digest();
  const hb = createHash("sha256").update(String(b)).digest();
  return timingSafeEqual(ha, hb);
}

async function readBody(req) {
  const text = await req.text();
  if (text.length > MAX_BODY_BYTES) throw new Error("Body too large");
  return JSON.parse(text);
}

function validate(data) {
  if (!data || typeof data !== "object") return "No timetable data was sent.";
  if (!Array.isArray(data.invigilators) || data.invigilators.length > 300) return "The list of invigilators is missing or too long.";
  if (!data.invigilators.every((n) => typeof n === "string" && n.trim())) return "Every invigilator needs a name.";
  if (!Array.isArray(data.sessions) || data.sessions.length > 2000) return "The list of sessions is missing or too long.";
  if (!data.sessions.every((s) => s && /^\d{4}-\d{2}-\d{2}$/.test(s.date) && typeof s.period === "string")) return "A session has a missing or invalid date.";
  if (typeof data.duties !== "object" || data.duties === null) return "The duties are missing.";
  return null;
}

function json(status, body) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "content-type": "application/json", "cache-control": "no-store" },
  });
}
