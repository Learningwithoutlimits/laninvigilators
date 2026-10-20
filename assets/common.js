// Shared helpers for the invigilator page and the admin page.

const LWLAT = (() => {
  // The schools, matching the list in netlify/functions/timetable.mjs.
  const SCHOOLS = [
    { id: "lancaster", name: "Lancaster Academy", colour: "#182d70" },
    { id: "babington", name: "Babington Academy", colour: "#009f9a" },
    { id: "south-wigston", name: "South Wigston High School", colour: "#1c6a9b" },
  ];
  const schoolById = (id) => SCHOOLS.find((s) => s.id === id) ?? null;

  const DAY_NAMES = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
  const MONTH_NAMES = ["January", "February", "March", "April", "May", "June", "July",
    "August", "September", "October", "November", "December"];

  // ---- Talking to the server ----

  async function api(path, { method = "GET", body, password } = {}) {
    const res = await fetch(path, {
      method,
      headers: {
        "content-type": "application/json",
        "x-password": password ?? "",
      },
      body: body === undefined ? undefined : JSON.stringify(body),
    });
    let payload = {};
    try { payload = await res.json(); } catch { /* empty or not JSON */ }
    if (!res.ok) {
      const err = new Error(payload.error || `The server replied with error ${res.status}.`);
      err.status = res.status;
      err.payload = payload;
      throw err;
    }
    return payload;
  }

  // Storage can be blocked (private windows, strict settings), so never rely on it.
  function remember(storage, key, value) {
    try {
      if (value === null) storage.removeItem(key);
      else storage.setItem(key, value);
    } catch { /* ignore */ }
  }
  function recall(storage, key) {
    try { return storage.getItem(key); } catch { return null; }
  }

  // ---- Dates ----

  function parseISO(iso) {
    const [y, m, d] = iso.split("-").map(Number);
    return new Date(y, m - 1, d);
  }
  function toISO(date) {
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, "0");
    const d = String(date.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }
  function mondayOf(iso) {
    const d = parseISO(iso);
    const shift = (d.getDay() + 6) % 7; // Monday = 0
    d.setDate(d.getDate() - shift);
    return toISO(d);
  }
  function formatLong(iso) {
    const d = parseISO(iso);
    return `${DAY_NAMES[d.getDay()]} ${d.getDate()} ${MONTH_NAMES[d.getMonth()]}`;
  }
  function formatShort(iso) {
    const d = parseISO(iso);
    return `${DAY_NAMES[d.getDay()].slice(0, 3)} ${d.getDate()} ${MONTH_NAMES[d.getMonth()].slice(0, 3)}`;
  }
  function formatWeek(mondayIso) {
    const d = parseISO(mondayIso);
    return `Week beginning Monday ${d.getDate()} ${MONTH_NAMES[d.getMonth()]} ${d.getFullYear()}`;
  }
  function formatStamp(isoStamp) {
    const d = new Date(isoStamp);
    if (isNaN(d)) return "";
    const time = d.toLocaleTimeString("en-GB", { hour: "2-digit", minute: "2-digit" });
    return `${d.getDate()} ${MONTH_NAMES[d.getMonth()]} ${d.getFullYear()} at ${time}`;
  }

  // ---- Timetable data ----

  function sessionId(date, period) { return `${date}|${period}`; }

  // Each row of the timetable is one exam slot: a date, a session (AM or PM), and its own
  // start time, duration and duties. A session can hold several rows. Rows saved before this
  // existed have no id, and use "date|period", which their duties and exam details are keyed by.
  function idOf(s) { return s.id || sessionId(s.date, s.period); }
  function newId() { return "x" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6); }

  // "AM", or "AM 2", "AM 3" for the second and third exam rows in the same session.
  function slotLabels(data) {
    const labels = new Map(), counts = {}, totals = {};
    for (const s of data.sessions) totals[sessionId(s.date, s.period)] = (totals[sessionId(s.date, s.period)] ?? 0) + 1;
    for (const s of sortSessions(data.sessions, periodOrderOf(data))) {
      const key = sessionId(s.date, s.period);
      counts[key] = (counts[key] ?? 0) + 1;
      labels.set(idOf(s), totals[key] > 1 && counts[key] > 1 ? `${s.period} ${counts[key]}` : s.period);
    }
    return labels;
  }

  function normalise(text) {
    return String(text ?? "").trim().replace(/\s+/g, " ").toLowerCase();
  }

  // Values that mean "this person is not available", whether typed or imported.
  function isUnavailable(value) {
    return /^(n\/?a|x|unavailable|not available|absent|off)$/.test(normalise(value));
  }
  function isClash(value) { return normalise(value) === "clash"; }

  function examKey(sid, location) { return `${sid}|${normalise(location)}`; }

  function sortSessions(sessions, periodOrder) {
    const order = (p) => {
      const i = periodOrder.indexOf(p);
      return i === -1 ? 99 : i;
    };
    return [...sessions].sort((a, b) =>
      a.date === b.date ? order(a.period) - order(b.period) : a.date < b.date ? -1 : 1);
  }

  function periodOrderOf(data) {
    const seen = [];
    for (const s of data.sessions) if (!seen.includes(s.period)) seen.push(s.period);
    const known = ["AM", "PM"];
    return [...known.filter((p) => seen.includes(p)), ...seen.filter((p) => !known.includes(p))];
  }

  // ---- Start times and durations ----

  // "9:00", "09.00", "9:00 am" or "2pm" -> "09:00" / "14:00". Anything else -> null.
  function parseTime(text) {
    const t = normalise(text).replace(/\s+/g, "");
    const m = t.match(/^(\d{1,2})(?:[:.](\d{2}))?(am|pm)?$/);
    if (!m || (!m[2] && !m[3])) return null;
    let h = Number(m[1]);
    const min = Number(m[2] ?? 0);
    if (m[3] === "pm" && h < 12) h += 12;
    if (m[3] === "am" && h === 12) h = 0;
    if (h > 23 || min > 59) return null;
    return `${String(h).padStart(2, "0")}:${String(min).padStart(2, "0")}`;
  }

  // A duration in minutes, from "1h 30m", "1 hour 30 minutes", "90 mins", "1:30", "90" or "1.5".
  // A bare number of 8 or less is taken as hours, anything larger as minutes.
  function parseDuration(value) {
    if (typeof value === "number") {
      if (!(value > 0)) return null;
      return Math.round(value <= 8 ? value * 60 : value);
    }
    const t = normalise(value);
    if (!t) return null;
    let m = t.match(/^(\d{1,2})[:.](\d{2})$/);
    if (m) return Number(m[1]) * 60 + Number(m[2]) || null;
    if (/^\d+(\.\d+)?$/.test(t)) return parseDuration(Number(t));
    const hours = t.match(/(\d+(?:\.\d+)?)\s*(?:h|hr|hrs|hour|hours)\b/);
    const mins = t.match(/(\d+)\s*(?:m|min|mins|minute|minutes)\b/);
    if (!hours && !mins) return null;
    const total = Math.round((hours ? Number(hours[1]) * 60 : 0) + (mins ? Number(mins[1]) : 0));
    return total > 0 ? total : null;
  }

  function formatDuration(minutes) {
    if (!(minutes > 0)) return "";
    const h = Math.floor(minutes / 60), m = minutes % 60;
    const parts = [];
    if (h) parts.push(`${h} ${h === 1 ? "hour" : "hours"}`);
    if (m) parts.push(`${m} ${m === 1 ? "minute" : "minutes"}`);
    return parts.join(" ");
  }

  function minutesBetween(start, end) {
    const a = parseTime(start), b = parseTime(end);
    if (!a || !b) return null;
    const diff = (Number(b.slice(0, 2)) * 60 + Number(b.slice(3))) - (Number(a.slice(0, 2)) * 60 + Number(a.slice(3)));
    return diff > 0 ? diff : null;
  }

  // Bring timetables saved by earlier versions (start and finish times) up to date.
  function upgrade(data) {
    data.sessionTimes ??= {};
    for (const s of data.sessions ?? []) s.id ??= sessionId(s.date, s.period);
    for (const details of Object.values(data.exams ?? {})) {
      if (details.end !== undefined) {
        if (details.duration == null) details.duration = minutesBetween(details.start, details.end);
        delete details.end;
      }
    }
    for (const [period, t] of Object.entries(data.periodTimes ?? {})) {
      // The old "usual session times" become each session's start time. Their finish time was
      // the end of the whole session, not the exam, so it is not turned into a duration.
      if (!t.start) continue;
      for (const s of data.sessions ?? []) {
        if (s.period !== period) continue;
        const st = (data.sessionTimes[idOf(s)] ??= {});
        if (!st.start) st.start = t.start;
      }
    }
    delete data.periodTimes;
    return data;
  }

  // The start time and duration that apply to one room: its own, or else its session's.
  function timesFor(data, sid, location) {
    const exam = data.exams?.[examKey(sid, location)] ?? {};
    const session = data.sessionTimes?.[sid] ?? {};
    return {
      start: exam.start || session.start || "",
      duration: exam.duration ?? session.duration ?? null,
    };
  }

  // Everything one person needs to see for one session.
  function entryFor(data, session, name) {
    const sid = idOf(session);
    const value = (data.duties?.[sid]?.[name] ?? "").trim();
    const entry = { session, value, status: "free", where: "", exam: "", start: "", duration: "" };
    if (!value) return entry;
    if (isUnavailable(value)) { entry.status = "unavailable"; return entry; }
    entry.status = isClash(value) ? "clash" : "duty";
    entry.where = value;
    if (entry.status === "clash") return entry;
    const times = timesFor(data, sid, value);
    const details = data.exams?.[examKey(sid, value)] ?? {};
    const slot = data.sessionTimes?.[sid] ?? {};
    entry.exam = details.exam || slot.exam || "";
    entry.start = times.start;
    entry.duration = formatDuration(times.duration);
    // Further exams in the same room, each with its own name and times. Blank times fall
    // back to the session's, as they do for the room's first exam.
    entry.extras = (details.more ?? []).filter((x) => x.exam || x.start || x.duration != null).map((x) => ({
      exam: x.exam || "",
      start: x.start || slot.start || "",
      duration: formatDuration(x.duration ?? slot.duration),
    }));
    return entry;
  }

  function emptyData() {
    return {
      version: 1,
      title: "",
      notice: "",
      invigilators: [],
      sessions: [],
      duties: {},
      sessionTimes: {},
      exams: {},
    };
  }

  // ---- Small DOM helper. Always uses textContent, so typed text can never inject HTML. ----

  function el(tag, attrs = {}, ...children) {
    const node = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs)) {
      if (v === undefined || v === null || v === false) continue;
      if (k === "class") node.className = v;
      else if (k === "text") node.textContent = v;
      else if (k.startsWith("on")) node.addEventListener(k.slice(2), v);
      else if (k in node && typeof v !== "string") node[k] = v;
      else node.setAttribute(k, v === true ? "" : v);
    }
    for (const child of children.flat()) {
      if (child === null || child === undefined || child === false) continue;
      node.append(child instanceof Node ? child : document.createTextNode(String(child)));
    }
    return node;
  }

  async function imageAsDataUrl(src) {
    const blob = await (await fetch(src)).blob();
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
  }

  return {
    SCHOOLS, schoolById,
    api, remember, recall,
    parseISO, toISO, mondayOf, formatLong, formatShort, formatWeek, formatStamp,
    sessionId, idOf, newId, slotLabels, normalise, isUnavailable, isClash, examKey, sortSessions, periodOrderOf,
    entryFor, emptyData, el, imageAsDataUrl,
    parseTime, parseDuration, formatDuration, minutesBetween, upgrade, timesFor,
    DAY_NAMES, MONTH_NAMES,
  };
})();
