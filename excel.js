// Reads the exams office's invigilation spreadsheet, and writes it back out.
//
// The expected layout (as in "Overview of invigilation - Year 11 mocks"):
//   - One header row with invigilator names across the top.
//   - A column of day names (Mon, Tue...), a column of dates ("9th Nov" or real dates)
//     and a column of sessions (AM / PM). Day and date cells are often merged.
//   - Each cell under a name holds where that person is (Hall, Gym, Ma 5, Roving...).
//   - Grey-shaded empty cells mean the person is not available.
//   - Optional columns headed "Start time" and "Duration" give each session's exam start
//     time and its duration, not including extra time. They can sit anywhere in the row.
// An optional sheet called "Exam details" holds exam names and times per room and session.

const LWLAT_EXCEL = (() => {
  const MONTHS = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
  const DAYS = ["sun", "mon", "tue", "wed", "thu", "fri", "sat"];
  const EXAM_SHEET = "Exam details";
  const ALL_ROOMS = "All rooms";

  function cellText(cell) {
    if (!cell) return "";
    const v = cell.value;
    if (v === null || v === undefined) return "";
    if (v instanceof Date) return v;
    if (typeof v === "object") {
      if (v.richText) return v.richText.map((r) => r.text).join("").trim();
      if (v.result !== undefined) return v.result instanceof Date ? v.result : String(v.result).trim();
      if (v.text !== undefined) return String(v.text).trim();
      return "";
    }
    return String(v).trim();
  }

  // Is this an empty cell shaded grey or black (meaning "not available")?
  function isGreyFill(cell) {
    const fill = cell?.fill;
    if (!fill || fill.type !== "pattern" || !fill.pattern || fill.pattern === "none") return false;
    const c = fill.fgColor || {};
    if (c.argb) {
      const hex = c.argb.slice(-6);
      const r = parseInt(hex.slice(0, 2), 16), g = parseInt(hex.slice(2, 4), 16), b = parseInt(hex.slice(4, 6), 16);
      return Math.max(r, g, b) - Math.min(r, g, b) < 20 && Math.max(r, g, b) < 0xc8;
    }
    if (c.theme !== undefined) {
      const tint = c.tint ?? 0;
      if (c.theme === 1 || c.theme === 3) return tint < 0.7;   // dark text colours
      if (c.theme === 0 || c.theme === 2) return tint <= -0.15; // light colours, darkened
    }
    return false;
  }

  function dayIndex(text) {
    if (typeof text !== "string") return -1;
    return DAYS.indexOf(text.trim().slice(0, 3).toLowerCase());
  }

  // "9th Nov", "9 November", "09/11", "9/11/2026" or a real date.
  function parseDayMonth(value) {
    if (value instanceof Date) {
      return { day: value.getUTCDate(), month: value.getUTCMonth(), year: value.getUTCFullYear() };
    }
    if (typeof value !== "string") return null;
    let m = value.match(/(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]{3,})\.?(?:\s+(\d{4}))?/);
    if (m) {
      const month = MONTHS.indexOf(m[2].slice(0, 3).toLowerCase());
      if (month !== -1) return { day: +m[1], month, year: m[3] ? +m[3] : null };
    }
    m = value.match(/^(\d{1,2})[/.-](\d{1,2})(?:[/.-](\d{2,4}))?$/);
    if (m) {
      let year = m[3] ? +m[3] : null;
      if (year !== null && year < 100) year += 2000;
      if (+m[1] < 1 || +m[1] > 31 || +m[2] < 1 || +m[2] > 12) return null;
      return { day: +m[1], month: +m[2] - 1, year };
    }
    return null;
  }

  const START_HEADER = /^(exam\s*)?start(\s*time)?$/i;
  const DURATION_HEADER = /duration|length/i;

  // Columns headed "Start time" and "Duration", looked for in the rows above the first session.
  function findTimeColumns(ws, firstDataRow) {
    let startCol = null, durationCol = null;
    for (let r = 1; r < firstDataRow; r++) {
      ws.getRow(r).eachCell((cell, c) => {
        const t = cellText(cell);
        if (typeof t !== "string") return;
        if (!startCol && START_HEADER.test(t)) startCol = c;
        else if (!durationCol && DURATION_HEADER.test(t)) durationCol = c;
      });
    }
    return { startCol, durationCol };
  }

  // A duration cell: an Excel time such as 1:30, a number, or text such as "1h 30m" or "90 mins".
  function durationMinutes(cell) {
    let v = cell?.value;
    if (v && typeof v === "object" && "result" in v) v = v.result;
    if (v === null || v === undefined || v === "") return null;
    if (v instanceof Date) {
      const mins = v.getUTCHours() * 60 + v.getUTCMinutes();
      return mins > 0 ? mins : null;
    }
    if (typeof v === "number") return v < 1 ? Math.round(v * 24 * 60) || null : LWLAT.parseDuration(v);
    return LWLAT.parseDuration(cellText(cell));
  }

  // Dates like "9th Nov" have no year. Pick the year that matches the day name
  // (Mon, Tue...) if there is one, otherwise the one closest to today.
  function chooseYear(dm, dayName) {
    if (dm.year) return dm.year;
    const now = new Date();
    const wanted = dayIndex(dayName);
    const candidates = [now.getFullYear() - 1, now.getFullYear(), now.getFullYear() + 1];
    const score = (y) => {
      const d = new Date(y, dm.month, dm.day);
      const dayMismatch = wanted !== -1 && d.getDay() !== wanted ? 1e12 : 0;
      return dayMismatch + Math.abs(d - now);
    };
    return candidates.sort((a, b) => score(a) - score(b))[0];
  }

  function iso(y, m, d) {
    return `${y}-${String(m + 1).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
  }

  // Find the column that holds AM / PM.
  function findPeriodColumn(ws) {
    let best = null;
    for (let c = 1; c <= Math.min(ws.columnCount, 10); c++) {
      let hits = 0;
      for (let r = 1; r <= ws.rowCount; r++) {
        const t = cellText(ws.getRow(r).getCell(c));
        if (typeof t === "string" && /^(am|pm)$/i.test(t)) hits++;
      }
      if (hits >= 2 && (!best || hits > best.hits)) best = { col: c, hits };
    }
    return best?.col ?? null;
  }

  function findDateColumn(ws, periodCol, skip) {
    for (let c = periodCol - 1; c >= 1; c--) {
      if (skip.includes(c)) continue;
      for (let r = 1; r <= ws.rowCount; r++) {
        if (parseDayMonth(cellText(ws.getRow(r).getCell(c)))) return c;
      }
    }
    return null;
  }

  function findDayColumn(ws, periodCol) {
    for (let c = 1; c < periodCol; c++) {
      for (let r = 1; r <= ws.rowCount; r++) {
        const t = cellText(ws.getRow(r).getCell(c));
        if (typeof t === "string" && /^(mon|tue|wed|thu|fri|sat|sun)/i.test(t)) return c;
      }
    }
    return null;
  }

  async function read(file) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(await file.arrayBuffer());
    const sheets = wb.worksheets.filter((s) => s.name !== EXAM_SHEET);
    const warnings = [];

    let ws = null, periodCol = null;
    for (const s of sheets) {
      const col = findPeriodColumn(s);
      if (col) { ws = s; periodCol = col; break; }
    }
    if (!ws) {
      throw new Error("I could not find a column of AM and PM sessions in this spreadsheet. Please check it uses the usual layout.");
    }
    if (sheets.length > 1) warnings.push(`This file has more than one sheet. I used the sheet called "${ws.name}".`);

    // Header row: the first row above the first AM/PM row with names to the right.
    let firstDataRow = null;
    for (let r = 1; r <= ws.rowCount && !firstDataRow; r++) {
      const t = cellText(ws.getRow(r).getCell(periodCol));
      if (typeof t === "string" && /^(am|pm)$/i.test(t)) firstDataRow = r;
    }

    const { startCol, durationCol } = findTimeColumns(ws, firstDataRow);
    const dateCol = findDateColumn(ws, periodCol, [startCol, durationCol]);
    const dayCol = findDayColumn(ws, periodCol);
    if (!dateCol) throw new Error("I could not find a column of dates to the left of the AM and PM column.");
    let headerRow = null, nameCols = [];
    for (let r = firstDataRow - 1; r >= 1 && !headerRow; r--) {
      const cols = [];
      ws.getRow(r).eachCell((cell, c) => {
        const t = cellText(cell);
        if (c > periodCol && c !== startCol && c !== durationCol && typeof t === "string" && t) {
          cols.push({ col: c, name: t.replace(/\s+/g, " ") });
        }
      });
      if (cols.length) { headerRow = r; nameCols = cols; }
    }
    if (!nameCols.length) throw new Error("I could not find the row of invigilator names above the timetable.");

    const seen = new Set();
    for (const nc of nameCols) {
      if (seen.has(nc.name)) warnings.push(`The name "${nc.name}" appears twice. Only the first column will be used.`);
      seen.add(nc.name);
    }
    const uniqueCols = nameCols.filter((nc, i) => nameCols.findIndex((x) => x.name === nc.name) === i);

    const sessions = [];
    const duties = {};
    const sessionTimes = {};
    const slotCount = {};    // rows seen so far for each date and session
    const extraRows = [];
    let badTimes = 0;
    let current = null;      // { date, period }
    let currentDate = null;
    let unavailableCount = 0, dutyCount = 0;

    for (let r = headerRow + 1; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      const dateRaw = cellText(row.getCell(dateCol));
      const dayRaw = dayCol ? cellText(row.getCell(dayCol)) : "";
      const periodRaw = cellText(row.getCell(periodCol));
      const values = uniqueCols.map((nc) => ({ nc, cell: row.getCell(nc.col), text: cellText(row.getCell(nc.col)) }));
      const hasTimes = (startCol && cellText(row.getCell(startCol))) || (durationCol && cellText(row.getCell(durationCol)));
      const hasAnything = dateRaw || periodRaw || hasTimes || values.some((v) => v.text);

      if (!hasAnything) { current = null; currentDate = null; continue; }

      const dm = parseDayMonth(dateRaw);
      if (dm) {
        const year = chooseYear(dm, dayRaw);
        currentDate = iso(year, dm.month, dm.day);
      }
      if (!currentDate) continue;

      // Every row is its own exam row. A row with no AM/PM label is an extra exam in the
      // session above it, as long as it has something in it.
      let period = typeof periodRaw === "string" ? periodRaw.toUpperCase().replace(/\s+/g, "") : "";
      if (!period) {
        if (!current || current.date !== currentDate) continue;
        if (!hasTimes && !values.some((v) => v.text || isGreyFill(v.cell))) continue;
        period = current.period;
      }

      const key = `${currentDate}|${period}`;
      slotCount[key] = (slotCount[key] ?? 0) + 1;
      const n = slotCount[key];
      const sid = n === 1 ? key : `${key}|${n}`;
      if (n === 2) extraRows.push(`${LWLAT.formatShort(currentDate)} ${period}`);
      duties[sid] = {};
      sessions.push({ id: sid, date: currentDate, period });
      current = { date: currentDate, period };

      const times = (sessionTimes[sid] = {});
      if (startCol) {
        const raw = row.getCell(startCol);
        const start = LWLAT.parseTime(timeText(raw));
        if (start) times.start = start; else if (cellText(raw)) badTimes++;
      }
      if (durationCol) {
        const raw = row.getCell(durationCol);
        const minutes = durationMinutes(raw);
        if (minutes) times.duration = minutes; else if (cellText(raw)) badTimes++;
      }

      for (const { nc, cell, text } of values) {
        if (typeof text === "string" && text) duties[sid][nc.name] = text.replace(/\s+/g, " ");
        else if (isGreyFill(cell)) duties[sid][nc.name] = "N/A";
      }
    }

    for (const sid of Object.keys(duties)) {
      for (const v of Object.values(duties[sid])) {
        if (LWLAT.isUnavailable(v)) unavailableCount++; else dutyCount++;
      }
    }

    for (const sid of Object.keys(sessionTimes)) if (!sessionTimes[sid].start && sessionTimes[sid].duration == null) delete sessionTimes[sid];
    if (badTimes) warnings.push(`${badTimes} start time or duration cell${badTimes === 1 ? " was" : "s were"} not understood and left blank. Times should look like 9:00 and durations like 1:30 or 90 mins.`);
    if (extraRows.length) warnings.push(`Sessions with more than one exam row: ${extraRows.join(", ")}. Each row is kept as a separate exam.`);
    const labels = LWLAT.slotLabels({ sessions });
    const clashes = [];
    for (const s of sessions) {
      for (const [name, v] of Object.entries(duties[s.id])) {
        if (LWLAT.isClash(v)) clashes.push(`${name} (${LWLAT.formatShort(s.date)} ${labels.get(s.id)})`);
      }
    }
    if (clashes.length) warnings.push(`Cells marked CLASH: ${clashes.join(", ")}. Invigilators will be asked to speak to the Exams Officer.`);
    const anyDateless = sessions.length === 0;
    if (anyDateless) throw new Error("I found the names but no sessions with dates underneath them.");

    // Optional exam details sheet.
    const exams = {};
    const examSheet = wb.getWorksheet(EXAM_SHEET);
    if (examSheet) {
      // Older copies of this sheet had a Finish column where Duration now is.
      const oldFinish = /finish|end/i.test(String(cellText(examSheet.getRow(1).getCell(6)) || ""));
      examSheet.eachRow((row, r) => {
        if (r === 1) return;
        const dateVal = cellText(row.getCell(1));
        const [period, slot] = String(cellText(row.getCell(2)) || "").toUpperCase().trim().split(/\s+/);
        const room = String(cellText(row.getCell(3)) || "");
        const dm = parseDayMonth(dateVal);
        if (!dm || !period || !room) return;
        const date = iso(chooseYear(dm, ""), dm.month, dm.day);
        const start = LWLAT.parseTime(timeText(row.getCell(5))) || "";
        const details = {
          exam: String(cellText(row.getCell(4)) || ""),
          start,
          duration: oldFinish ? LWLAT.minutesBetween(start, timeText(row.getCell(6))) : durationMinutes(row.getCell(6)),
        };
        const sid = `${date}|${period}${slot > 1 ? `|${slot}` : ""}`;
        if (LWLAT.normalise(room) === LWLAT.normalise(ALL_ROOMS)) {
          if (details.exam) (sessionTimes[sid] ??= {}).exam = details.exam;
        } else if (details.exam || details.start || details.duration) exams[LWLAT.examKey(sid, room)] = details;
      });
    }

    return {
      sheetName: ws.name,
      invigilators: uniqueCols.map((nc) => nc.name),
      sessions,
      duties,
      sessionTimes,
      hasTimeColumns: Boolean(startCol || durationCol),
      exams,
      warnings,
      stats: { dutyCount, unavailableCount, timedSessions: Object.keys(sessionTimes).length },
    };
  }

  function timeText(cell) {
    const v = cell?.value;
    if (v instanceof Date) return `${String(v.getUTCHours()).padStart(2, "0")}:${String(v.getUTCMinutes()).padStart(2, "0")}`;
    if (typeof v === "number") { // Excel fraction of a day
      const mins = Math.round(v * 24 * 60);
      return `${String(Math.floor(mins / 60)).padStart(2, "0")}:${String(mins % 60).padStart(2, "0")}`;
    }
    const t = String(cellText(cell) || "");
    const m = t.match(/^(\d{1,2})[:.](\d{2})/);
    return m ? `${m[1].padStart(2, "0")}:${m[2]}` : t;
  }

  // ---- Export, in the same layout so it can be uploaded again ----

  async function write(data) {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet(data.title ? data.title.slice(0, 31).replace(/[\\/?*[\]:]/g, "") : "Exam invigilators");
    const names = data.invigilators;
    const periodOrder = LWLAT.periodOrderOf(data);
    const sessions = LWLAT.sortSessions(data.sessions, periodOrder);
    const teal = "FF009189", grey = "FF595959", red = "FFCE153F";

    const FIRST = 6; // column of the first invigilator
    const header = ws.addRow(["Day", "Date", "Session", "Start time", "Duration", ...names]);
    header.font = { bold: true, color: { argb: "FFFFFFFF" } };
    header.eachCell((c) => { c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: teal } }; });
    ws.views = [{ state: "frozen", xSplit: FIRST - 1, ySplit: 1 }];
    ws.getColumn(1).width = 6;
    ws.getColumn(2).width = 11;
    ws.getColumn(3).width = 9;
    ws.getColumn(4).width = 10;
    ws.getColumn(5).width = 20;
    names.forEach((_, i) => { ws.getColumn(FIRST + i).width = 12; });

    let lastMonday = null, lastDate = null;
    for (const s of sessions) {
      const monday = LWLAT.mondayOf(s.date);
      if (lastMonday && monday !== lastMonday) ws.addRow([]);
      lastMonday = monday;
      const d = LWLAT.parseISO(s.date);
      const first = s.date !== lastDate;
      lastDate = s.date;
      const dutiesHere = data.duties[LWLAT.idOf(s)] || {};
      const row = ws.addRow([
        first ? LWLAT.DAY_NAMES[d.getDay()].slice(0, 3) : "",
        first ? new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate())) : "",
        s.period,
        data.sessionTimes?.[LWLAT.idOf(s)]?.start || "",
        LWLAT.formatDuration(data.sessionTimes?.[LWLAT.idOf(s)]?.duration),
        ...names.map((n) => (dutiesHere[n] && !LWLAT.isUnavailable(dutiesHere[n]) ? dutiesHere[n] : "")),
      ]);
      row.getCell(2).numFmt = "d mmm yyyy";
      names.forEach((n, i) => {
        const v = dutiesHere[n];
        const cell = row.getCell(FIRST + i);
        if (v && LWLAT.isUnavailable(v)) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: grey } };
        else if (v && LWLAT.isClash(v)) { cell.font = { bold: true, color: { argb: "FFFFFFFF" } }; cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: red } }; }
      });
    }

    const ex = wb.addWorksheet(EXAM_SHEET);
    const exHead = ex.addRow(["Date", "Session", "Room / duty", "Exam", "Start time", "Duration"]);
    exHead.font = { bold: true, color: { argb: "FFFFFFFF" } };
    exHead.eachCell((c) => { c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: teal } }; });
    [14, 9, 16, 34, 10, 20].forEach((w, i) => { ex.getColumn(i + 1).width = w; });
    // An exam name typed for a whole exam row is written as "All rooms".
    const labels = LWLAT.slotLabels(data);
    const lines = [];
    for (const s of sessions) {
      const name = data.sessionTimes?.[LWLAT.idOf(s)]?.exam;
      if (name) lines.push({ session: s, label: labels.get(LWLAT.idOf(s)), location: ALL_ROOMS, details: { exam: name } });
      lines.push(...examRows(data).filter((r) => r.sid === LWLAT.idOf(s)));
    }
    for (const item of lines) {
      const d = LWLAT.parseISO(item.session.date);
      const r = ex.addRow([new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate())), item.label,
        item.location, item.details.exam || "", item.details.start || "", LWLAT.formatDuration(item.details.duration)]);
      r.getCell(1).numFmt = "d mmm yyyy";
    }

    const buffer = await wb.xlsx.writeBuffer();
    return new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  }

  // Every room or duty in use in each session, with any exam details entered for it.
  function examRows(data) {
    const periodOrder = LWLAT.periodOrderOf(data);
    const labels = LWLAT.slotLabels(data);
    const rows = [];
    for (const s of LWLAT.sortSessions(data.sessions, periodOrder)) {
      const sid = LWLAT.idOf(s);
      const byPlace = new Map();
      for (const name of data.invigilators) {
        const v = (data.duties[sid]?.[name] || "").trim();
        if (!v || LWLAT.isUnavailable(v) || LWLAT.isClash(v)) continue;
        const key = LWLAT.normalise(v);
        if (!byPlace.has(key)) byPlace.set(key, { location: v, names: [] });
        byPlace.get(key).names.push(name);
      }
      for (const { location, names } of [...byPlace.values()].sort((a, b) => a.location.localeCompare(b.location, "en", { numeric: true }))) {
        const key = LWLAT.examKey(sid, location);
        rows.push({ session: s, sid, label: labels.get(sid), location, names, key, details: data.exams?.[key] || {} });
      }
    }
    return rows;
  }

  return { read, write, examRows };
})();
