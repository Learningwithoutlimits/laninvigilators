// Reads the exams office's invigilation spreadsheet, and writes it back out.
//
// The expected layout (as in "Overview of invigilation - Year 11 mocks"):
//   - One header row with invigilator names across the top.
//   - A column of day names (Mon, Tue...), a column of dates ("9th Nov" or real dates)
//     and a column of sessions (AM / PM). Day and date cells are often merged.
//   - Each cell under a name holds where that person is (Hall, Gym, Ma 5, Roving...).
//   - Grey-shaded empty cells mean the person is not available.
// An optional sheet called "Exam details" holds exam names and times per room and session.

const LWLAT_EXCEL = (() => {
  const MONTHS = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
  const DAYS = ["sun", "mon", "tue", "wed", "thu", "fri", "sat"];
  const EXAM_SHEET = "Exam details";

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
      return { day: +m[1], month: +m[2] - 1, year };
    }
    return null;
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

  function findDateColumn(ws, periodCol) {
    for (let c = periodCol - 1; c >= 1; c--) {
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

    const dateCol = findDateColumn(ws, periodCol);
    const dayCol = findDayColumn(ws, periodCol);
    if (!dateCol) throw new Error("I could not find a column of dates to the left of the AM and PM column.");

    // Header row: the first row above the first AM/PM row with names to the right.
    let firstDataRow = null;
    for (let r = 1; r <= ws.rowCount && !firstDataRow; r++) {
      const t = cellText(ws.getRow(r).getCell(periodCol));
      if (typeof t === "string" && /^(am|pm)$/i.test(t)) firstDataRow = r;
    }
    let headerRow = null, nameCols = [];
    for (let r = firstDataRow - 1; r >= 1 && !headerRow; r--) {
      const cols = [];
      ws.getRow(r).eachCell((cell, c) => {
        const t = cellText(cell);
        if (c > periodCol && typeof t === "string" && t) cols.push({ col: c, name: t.replace(/\s+/g, " ") });
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
    let current = null;      // { date, period }
    let currentDate = null;
    let mergedRows = 0, unavailableCount = 0, dutyCount = 0;

    for (let r = headerRow + 1; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      const dateRaw = cellText(row.getCell(dateCol));
      const dayRaw = dayCol ? cellText(row.getCell(dayCol)) : "";
      const periodRaw = cellText(row.getCell(periodCol));
      const values = uniqueCols.map((nc) => ({ nc, cell: row.getCell(nc.col), text: cellText(row.getCell(nc.col)) }));
      const hasAnything = dateRaw || periodRaw || values.some((v) => v.text);

      if (!hasAnything) { current = null; currentDate = null; continue; }

      const dm = parseDayMonth(dateRaw);
      if (dm) {
        const year = chooseYear(dm, dayRaw);
        currentDate = iso(year, dm.month, dm.day);
      }
      if (!currentDate) continue;

      let period = typeof periodRaw === "string" ? periodRaw.toUpperCase().replace(/\s+/g, "") : "";
      if (!period) {
        if (current && current.date === currentDate) {
          period = current.period;   // an unlabelled extra row belongs to the session above
          if (values.some((v) => v.text)) mergedRows++;
        } else continue;
      }

      const sid = `${currentDate}|${period}`;
      if (!duties[sid]) {
        duties[sid] = {};
        sessions.push({ date: currentDate, period });
      }
      current = { date: currentDate, period };

      for (const { nc, cell, text } of values) {
        if (typeof text === "string" && text) {
          const tidy = text.replace(/\s+/g, " ");
          const existing = duties[sid][nc.name];
          duties[sid][nc.name] = existing && existing !== tidy && !LWLAT.isUnavailable(existing) ? `${existing} / ${tidy}` : tidy;
        } else if (!duties[sid][nc.name] && isGreyFill(cell)) {
          duties[sid][nc.name] = "N/A";
        }
      }
    }

    for (const sid of Object.keys(duties)) {
      for (const v of Object.values(duties[sid])) {
        if (LWLAT.isUnavailable(v)) unavailableCount++; else dutyCount++;
      }
    }

    if (mergedRows) warnings.push(`${mergedRows} unlabelled row${mergedRows === 1 ? " was" : "s were"} added to the session above it.`);
    const clashes = [];
    for (const s of sessions) {
      for (const [name, v] of Object.entries(duties[`${s.date}|${s.period}`])) {
        if (LWLAT.isClash(v)) clashes.push(`${name} (${LWLAT.formatShort(s.date)} ${s.period})`);
      }
    }
    if (clashes.length) warnings.push(`Cells marked CLASH: ${clashes.join(", ")}. Invigilators will be asked to check with the exams office.`);
    const anyDateless = sessions.length === 0;
    if (anyDateless) throw new Error("I found the names but no sessions with dates underneath them.");

    // Optional exam details sheet.
    const exams = {};
    const examSheet = wb.getWorksheet(EXAM_SHEET);
    if (examSheet) {
      examSheet.eachRow((row, r) => {
        if (r === 1) return;
        const dateVal = cellText(row.getCell(1));
        const period = String(cellText(row.getCell(2)) || "").toUpperCase();
        const room = String(cellText(row.getCell(3)) || "");
        const dm = parseDayMonth(dateVal);
        if (!dm || !period || !room) return;
        const date = iso(chooseYear(dm, ""), dm.month, dm.day);
        const details = {
          exam: String(cellText(row.getCell(4)) || ""),
          start: timeText(row.getCell(5)),
          end: timeText(row.getCell(6)),
        };
        if (details.exam || details.start || details.end) exams[LWLAT.examKey(`${date}|${period}`, room)] = details;
      });
    }

    return {
      sheetName: ws.name,
      invigilators: uniqueCols.map((nc) => nc.name),
      sessions,
      duties,
      exams,
      warnings,
      stats: { dutyCount, unavailableCount },
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
    const ws = wb.addWorksheet(data.title ? data.title.slice(0, 31).replace(/[\\/?*[\]:]/g, "") : "Invigilation");
    const names = data.invigilators;
    const periodOrder = LWLAT.periodOrderOf(data);
    const sessions = LWLAT.sortSessions(data.sessions, periodOrder);
    const teal = "FF009189", grey = "FF595959", red = "FFCE153F";

    const header = ws.addRow(["Day", "Date", "Session", ...names]);
    header.font = { bold: true, color: { argb: "FFFFFFFF" } };
    header.eachCell((c) => { c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: teal } }; });
    ws.views = [{ state: "frozen", xSplit: 3, ySplit: 1 }];
    ws.getColumn(1).width = 6;
    ws.getColumn(2).width = 11;
    ws.getColumn(3).width = 9;
    names.forEach((_, i) => { ws.getColumn(4 + i).width = 12; });

    let lastMonday = null, lastDate = null;
    for (const s of sessions) {
      const monday = LWLAT.mondayOf(s.date);
      if (lastMonday && monday !== lastMonday) ws.addRow([]);
      lastMonday = monday;
      const d = LWLAT.parseISO(s.date);
      const first = s.date !== lastDate;
      lastDate = s.date;
      const dutiesHere = data.duties[`${s.date}|${s.period}`] || {};
      const row = ws.addRow([
        first ? LWLAT.DAY_NAMES[d.getDay()].slice(0, 3) : "",
        first ? new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate())) : "",
        s.period,
        ...names.map((n) => (dutiesHere[n] && !LWLAT.isUnavailable(dutiesHere[n]) ? dutiesHere[n] : "")),
      ]);
      row.getCell(2).numFmt = "d mmm yyyy";
      names.forEach((n, i) => {
        const v = dutiesHere[n];
        const cell = row.getCell(4 + i);
        if (v && LWLAT.isUnavailable(v)) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: grey } };
        else if (v && LWLAT.isClash(v)) { cell.font = { bold: true, color: { argb: "FFFFFFFF" } }; cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: red } }; }
      });
    }

    const ex = wb.addWorksheet(EXAM_SHEET);
    const exHead = ex.addRow(["Date", "Session", "Room / duty", "Exam", "Start", "Finish"]);
    exHead.font = { bold: true, color: { argb: "FFFFFFFF" } };
    exHead.eachCell((c) => { c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: teal } }; });
    [14, 9, 16, 34, 9, 9].forEach((w, i) => { ex.getColumn(i + 1).width = w; });
    for (const item of examRows(data)) {
      const d = LWLAT.parseISO(item.session.date);
      const r = ex.addRow([new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate())), item.session.period,
        item.location, item.details.exam || "", item.details.start || "", item.details.end || ""]);
      r.getCell(1).numFmt = "d mmm yyyy";
    }

    const buffer = await wb.xlsx.writeBuffer();
    return new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  }

  // Every room or duty in use in each session, with any exam details entered for it.
  function examRows(data) {
    const periodOrder = LWLAT.periodOrderOf(data);
    const rows = [];
    for (const s of LWLAT.sortSessions(data.sessions, periodOrder)) {
      const sid = `${s.date}|${s.period}`;
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
        rows.push({ session: s, location, names, key, details: data.exams?.[key] || {} });
      }
    }
    return rows;
  }

  return { read, write, examRows };
})();
