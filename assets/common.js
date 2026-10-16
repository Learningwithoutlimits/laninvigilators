// Shared helpers for the invigilator page and the admin page.

const LWLAT = (() => {
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

  // Everything one person needs to see for one session.
  function entryFor(data, session, name) {
    const sid = sessionId(session.date, session.period);
    const value = (data.duties?.[sid]?.[name] ?? "").trim();
    const entry = { session, value, status: "free", where: "", exam: "", time: "" };
    if (!value) return entry;
    if (isUnavailable(value)) { entry.status = "unavailable"; return entry; }
    entry.status = isClash(value) ? "clash" : "duty";
    entry.where = value;
    const exam = data.exams?.[examKey(sid, value)] ?? {};
    const fallback = data.periodTimes?.[session.period] ?? {};
    entry.exam = exam.exam ?? "";
    const start = exam.start || fallback.start || "";
    const end = exam.end || fallback.end || "";
    entry.time = start && end ? `${start} to ${end}` : start ? `From ${start}` : "";
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
      exams: {},
      periodTimes: {},
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
    api, remember, recall,
    parseISO, toISO, mondayOf, formatLong, formatShort, formatWeek, formatStamp,
    sessionId, normalise, isUnavailable, isClash, examKey, sortSessions, periodOrderOf,
    entryFor, emptyData, el, imageAsDataUrl,
    DAY_NAMES, MONTH_NAMES,
  };
})();
