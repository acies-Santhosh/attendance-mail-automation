// ============================================================
//  ATTENDANCE COMPLIANCE — Google Apps Script Web App
//  File: Code.gs  (server side)
// ============================================================

const CONFIG = {
  SPREADSHEET_ID: "1pE1DEIpytWLRzsfUxSn9MQ-PFUE8Rrg5oodKEEj4noo",
  RESULTS_SHEET: "Compliance Results",
  LOG_SHEET: "Email Log",
  MIN_OFFICE_DAYS: 3,
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Attendance Compliance — Acies Global")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setupSheets() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  function ensureSheet(name, headers, widths) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    sh.clearContents();
    sh.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setBackground("#1e2026")
      .setFontColor("#e6a817")
      .setFontWeight("bold")
      .setFontFamily("Courier New");
    sh.setFrozenRows(1);
    widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
    return sh;
  }
  ensureSheet(
    CONFIG.RESULTS_SHEET,
    [
      "Employee ID",
      "Employee Name",
      "Email",
      "Week",
      "Mon",
      "Tue",
      "Wed",
      "Thu",
      "Fri",
      "Total Days",
      "Compliant",
      "Email Status",
      "Sent At",
    ],
    [100, 180, 220, 130, 90, 90, 90, 90, 90, 100, 90, 110, 160],
  );
  ensureSheet(
    CONFIG.LOG_SHEET,
    [
      "Timestamp",
      "Week",
      "Employee ID",
      "Employee Name",
      "Email",
      "Total Days",
      "Mon",
      "Tue",
      "Wed",
      "Thu",
      "Fri",
      "Email Status",
    ],
    [160, 130, 100, 180, 220, 100, 90, 90, 90, 90, 90, 110],
  );
  return "✅ Sheets created successfully.";
}

// ── CALLED FROM CLIENT ────────────────────────────────────────
// payload now includes fileTypes: { chennai, bangalore, zoho, ullen, excuse }
function processFiles(payload) {
  try {
    const {
      chennaiB64,
      bangaloreB64,
      zohoB64,
      ullenB64,
      excuseB64,
      fileTypes,
      weekStart,
      weekEnd,
    } = payload;

    const ft = fileTypes || {};

    const chennaiSheets = decodeFile(chennaiB64, ft.chennai || "xlsx");
    const chennai = parseChennaiData(chennaiSheets);
    Logger.log("=== CHENNAI DEBUG ===");
    Logger.log("Script timezone: " + Session.getScriptTimeZone());
    const _ws = chennaiSheets[Object.keys(chennaiSheets)[0]];
    if (_ws && _ws[0]) Logger.log("Chennai header[0..8]: " + JSON.stringify(_ws[0].slice(0,9).map(h => ({ v: String(h), t: typeof h, isDate: h instanceof Date }))));
    Logger.log("Chennai parsed keys (first 3 employees): " + JSON.stringify(Object.keys(chennai).slice(0,3)));
    const _sampleId = Object.keys(chennai)[0];
    if (_sampleId) Logger.log("Chennai[" + _sampleId + "]: " + JSON.stringify(chennai[_sampleId]));
    Logger.log("Chennai['1348']: " + JSON.stringify(chennai['1348'] || "NOT FOUND"));
    const bangalore = parseBangaloreData(
      decodeFile(bangaloreB64, ft.bangalore || "csv"),
    );
    const { leaveMap: zoho, emailMap } = parseZohoData(
      decodeFile(zohoB64, ft.zoho || "xls"),
    );
    const { officeMap: ullen, employees } = parseUllenAyyahData(
      decodeFile(ullenB64, ft.ullen || "csv"),
    );
    const excuse = excuseB64
      ? parseExcuseData(decodeFile(excuseB64, ft.excuse || "xlsx"))
      : {};

    if (!employees || employees.length === 0)
      return {
        ok: false,
        error: "No employees found in Ullen Ayyah file. Check the format.",
      };

    const weekdays = getWeekdays(weekStart, weekEnd);
    if (weekdays.length === 0)
      return { ok: false, error: "No working days in selected range." };

    const enriched = employees.map((e) => ({
      ...e,
      email: e.email || emailMap[e.id] || "",
    }));
    const results = computeCompliance(
      enriched,
      weekdays,
      chennai,
      bangalore,
      zoho,
      ullen,
      excuse,
    );
    results.sort(
      (a, b) => a.compliant - b.compliant || a.name.localeCompare(b.name),
    );

    return { ok: true, results, weekdays };
  } catch (e) {
    return { ok: false, error: e.message + "\n" + e.stack };
  }
}

function sendEmails(payload) {
  try {
    const { results, weekLabel, weekdays } = payload;
    const dayNames = ["Mon", "Tue", "Wed", "Thu", "Fri"];
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const rs =
      ss.getSheetByName(CONFIG.RESULTS_SHEET) ||
      ss.insertSheet(CONFIG.RESULTS_SHEET);
    const ls =
      ss.getSheetByName(CONFIG.LOG_SHEET) || ss.insertSheet(CONFIG.LOG_SHEET);

    const now = new Date();
    const sent = [],
      failed = [],
      noEmail = [];
    const resultRows = [],
      logRows = [];

    for (const r of results) {
      const dayVals = weekdays.map((d) => r.dayBreakdown[d]?.type || "absent");
      const compliant = r.total >= CONFIG.MIN_OFFICE_DAYS;
      let emailStatus = "Pending",
        sentAt = "";

      if (!compliant) {
        if (!r.email || !r.email.includes("@")) {
          emailStatus = "No email on record";
          noEmail.push(r.name);
        } else {
          try {
            GmailApp.sendEmail(
              r.email,
              `Office Attendance Reminder — Week of ${weekLabel}`,
              buildEmailBody(r, weekLabel, weekdays, dayNames),
            );
            emailStatus = "Sent";
            sentAt = now.toLocaleString();
            sent.push(r.name);
          } catch (e) {
            emailStatus = "Failed: " + e.message;
            failed.push(r.name);
          }
          Utilities.sleep(150);
        }
      } else {
        emailStatus = "Compliant — no email";
      }

      resultRows.push([
        r.id,
        r.name,
        r.email,
        weekLabel,
        ...dayVals,
        r.total,
        compliant ? "Yes" : "No",
        emailStatus,
        sentAt,
      ]);
      if (!compliant)
        logRows.push([
          now.toLocaleString(),
          weekLabel,
          r.id,
          r.name,
          r.email,
          r.total,
          ...dayVals,
          emailStatus,
        ]);
    }

    rs.clearContents();
    const rHeaders = [
      "Employee ID",
      "Employee Name",
      "Email",
      "Week",
      "Mon",
      "Tue",
      "Wed",
      "Thu",
      "Fri",
      "Total Days",
      "Compliant",
      "Email Status",
      "Sent At",
    ];
    rs.getRange(1, 1, 1, rHeaders.length)
      .setValues([rHeaders])
      .setBackground("#1e2026")
      .setFontColor("#e6a817")
      .setFontWeight("bold");
    if (resultRows.length > 0)
      rs.getRange(2, 1, resultRows.length, resultRows[0].length).setValues(
        resultRows,
      );

    for (let i = 0; i < resultRows.length; i++) {
      rs.getRange(i + 2, 11).setFontColor(
        resultRows[i][10] === "Yes" ? "#3fb950" : "#f85149",
      );
    }
    if (logRows.length > 0) {
      const lastRow = ls.getLastRow();
      ls.getRange(lastRow + 1, 1, logRows.length, logRows[0].length).setValues(
        logRows,
      );
    }

    return {
      ok: true,
      sent: sent.length,
      failed: failed.length,
      noEmail: noEmail.length,
      failedNames: failed,
    };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ── FILE DECODER ──────────────────────────────────────────────
// FIX 1: Handle CSV, XLS, and XLSX differently
function decodeFile(b64, ext) {
  const e = (ext || "xlsx").toLowerCase().replace(".", "");

  if (e === "csv") {
    // CSV: decode bytes → UTF-8 string → return as single-sheet object
    const bytes = Utilities.base64Decode(b64);
    const text = Utilities.newBlob(bytes).getDataAsString("UTF-8");
    return { Sheet1: parseCSVText(text) };
  }

  // XLS or XLSX: upload to Drive, convert to Google Sheets, read, then delete
  const mimeMap = {
    xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    xls: "application/vnd.ms-excel",
    xlsm: "application/vnd.ms-excel.sheet.macroEnabled.12",
  };
  const mime = mimeMap[e] || mimeMap.xlsx;
  const bytes = Utilities.base64Decode(b64);
  const blob = Utilities.newBlob(bytes, mime, "_tmp_" + Date.now() + "." + e);
  const tmpFile = DriveApp.createFile(blob);
  const tmpId = tmpFile.getId();

  const resource = {
    title: "_tmp_gs_" + Date.now(),
    mimeType: "application/vnd.google-apps.spreadsheet",
  };
  const converted = Drive.Files.copy(resource, tmpId, { convert: true });
  const convId = converted.id;

  const tempSS = SpreadsheetApp.openById(convId);
  const result = {};
  for (const sh of tempSS.getSheets()) {
    result[sh.getName()] = sh.getDataRange().getValues();
  }
  try {
    DriveApp.getFileById(tmpId).setTrashed(true);
  } catch (e) {}
  try {
    DriveApp.getFileById(convId).setTrashed(true);
  } catch (e) {}
  return result;
}

// Parse CSV text into a 2D array (same shape as Google Sheets getValues())
function parseCSVText(text) {
  const rows = [];
  const lines = text.split(/\r?\n/);
  for (const line of lines) {
    if (line.trim() === "") continue;
    const fields = [];
    let field = "",
      inQuote = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') {
        inQuote = !inQuote;
      } else if (ch === "," && !inQuote) {
        fields.push(field.trim());
        field = "";
      } else {
        field += ch;
      }
    }
    fields.push(field.trim());
    rows.push(fields);
  }
  return rows;
}

// ── HELPERS ───────────────────────────────────────────────────
function getWeekdays(start, end) {
  const days = [];
  const tz = Session.getScriptTimeZone();
  const s = new Date(start + "T00:00:00"),
    e = new Date(end + "T00:00:00");
  for (let d = new Date(s); d <= e; d.setDate(d.getDate() + 1)) {
    const dow = d.getDay();
    if (dow >= 1 && dow <= 5) days.push(Utilities.formatDate(d, tz, "yyyy-MM-dd"));
  }
  return days;
}

function normId(v) {
  return v == null ? null : String(v).trim().toLowerCase().replace(/\s+/g, "");
}

function toISO(v) {
  if (!v) return null;
  if (v instanceof Date)
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  const s = String(v).trim();
  let m;
  m = s.match(/^(\d{2})-(\d{2})-(\d{4})$/);
  if (m) return `${m[3]}-${m[2]}-${m[1]}`;
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const mo = {
    jan: "01",
    feb: "02",
    mar: "03",
    apr: "04",
    may: "05",
    jun: "06",
    jul: "07",
    aug: "08",
    sep: "09",
    oct: "10",
    nov: "11",
    dec: "12",
  };
  m = s.match(/^(\d{1,2})[- ]([A-Za-z]+)[- ](\d{4})$/);
  if (m) {
    const x = mo[m[2].toLowerCase().slice(0, 3)];
    if (x) return `${m[3]}-${x}-${m[1].padStart(2, "0")}`;
  }
  return null;
}

// FIX 2: parseColDate now handles Date objects (Chennai xlsx col headers are Date values)
function parseColDate(h, yr) {
  if (!h) return null;
  // Handle actual Date objects from Google Sheets (Chennai biometric date columns)
  if (h instanceof Date)
    return Utilities.formatDate(h, Session.getScriptTimeZone(), "yyyy-MM-dd");
  const s = String(h).trim();
  const mo = {
    jan: "01",
    feb: "02",
    mar: "03",
    apr: "04",
    may: "05",
    jun: "06",
    jul: "07",
    aug: "08",
    sep: "09",
    oct: "10",
    nov: "11",
    dec: "12",
  };
  let m;
  m = s.match(/^(\d{1,2})\s*[-–]\s*([A-Za-z]+)$/);
  if (m) {
    const x = mo[m[2].toLowerCase().slice(0, 3)];
    if (x) return `${yr}-${x}-${m[1].padStart(2, "0")}`;
  }
  m = s.match(/^(\d{1,2})-([A-Za-z]+)-(\d{4})$/);
  if (m) {
    const x = mo[m[2].toLowerCase().slice(0, 3)];
    if (x) return `${m[3]}-${x}-${m[1].padStart(2, "0")}`;
  }
  m = s.match(/^(\d{2})-(\d{2})-(\d{4})$/);
  if (m) return `${m[3]}-${m[2]}-${m[1]}`;
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  return null;
}

const LEAVE_CODES = new Set([
  "EL",
  "SCL",
  "SCLK",
  "ML",
  "MSL",
  "LWP",
  "PLM",
  "H",
  "WH",
]);
const isLeave = (s) => LEAVE_CODES.has(s) || s.startsWith("0.5");
const isTime = (s) => /^\d{1,2}:\d{2}(:\d{2})?$/.test(s);
const halfDay = (s) => (s.startsWith("0.5") ? 0.5 : 0);

// ── PARSERS ───────────────────────────────────────────────────

// FIX 3: Chennai biometric is ONE sheet with date columns, not multiple sheets.
// Columns: Employee, First Name, Last Name, Seating Location, [date cols...], No. of Days, Remarks, Status
// Cell values: "09:45" = in office, "EL"/"SCL" = leave, blank/NaN = absent
function parseChennaiData(sheets) {
  const result = {};
  const ws = sheets[Object.keys(sheets)[0]];
  if (!ws || ws.length < 2) return result;

  const hdr = ws[0];

  // Find Employee ID column — header is literally "Employee" in their file
  let iId = hdr.findIndex((h) => {
    const s = String(h || "")
      .trim()
      .toLowerCase();
    return s === "employee" || s === "employee id" || s.includes("employee id");
  });
  if (iId < 0) return result; // safety

  // Find date columns using parseColDate (which now handles Date objects too)
  const dateCols = [];
  for (let c = 0; c < hdr.length; c++) {
    const d = parseColDate(hdr[c], 2026);
    if (d) dateCols.push({ col: c, date: d });
  }

  for (let i = 1; i < ws.length; i++) {
    const row = ws[i];
    const empId = normId(row[iId]);
    if (!empId) continue;
    if (!result[empId]) result[empId] = {};

    for (const { col, date } of dateCols) {
      const val = row[col];
      let entry;
      if (val === null || val === "" || val === undefined) {
        entry = { type: "absent" };
      } else if (val instanceof Date) {
        // Time stored as a Date object in Google Sheets (e.g. 1899-12-30 09:45:00)
        entry = {
          type: "office",
          value: Utilities.formatDate(val, "UTC", "HH:mm"),
        };
      } else {
        const s = String(val).trim().toUpperCase();
        if (s === "" || s === "NAN") {
          entry = { type: "absent" };
        } else if (LEAVE_CODES.has(s)) {
          entry = { type: "leave", value: 1 };
        } else if (isTime(s)) {
          entry = { type: "office", value: s };
        } else if (s.length > 0) {
          // Non-empty, non-leave string → treat as office (e.g. any time format variation)
          entry = { type: "office", value: s };
        } else {
          entry = { type: "absent" };
        }
      }
      result[empId][date] = entry;
    }
  }
  return result;
}

// ── DEBUG HELPER (run manually from Apps Script editor, paste a Chennai file ID) ──
function debugChennai() {
  // 1. Upload your Chennai XLSX to Drive manually, copy its file ID here:
  const FILE_ID = "PASTE_CHENNAI_FILE_ID_HERE";
  const file = DriveApp.getFileById(FILE_ID);
  const resource = { title: "_dbg_" + Date.now(), mimeType: "application/vnd.google-apps.spreadsheet" };
  const conv = Drive.Files.copy(resource, FILE_ID, { convert: true });
  const ss = SpreadsheetApp.openById(conv.id);
  const ws = ss.getSheets()[0].getDataRange().getValues();
  DriveApp.getFileById(conv.id).setTrashed(true);

  const hdr = ws[0];
  Logger.log("Header row: " + JSON.stringify(hdr.map(h => ({ val: h, type: typeof h, isDate: h instanceof Date }))));

  const dateCols = [];
  for (let c = 0; c < hdr.length; c++) {
    const d = parseColDate(hdr[c], 2026);
    if (d) dateCols.push({ col: c, date: d, raw: hdr[c] });
  }
  Logger.log("Date columns found: " + JSON.stringify(dateCols));

  // Log Ashween Raj's row (or first 3 data rows)
  for (let i = 1; i <= Math.min(5, ws.length - 1); i++) {
    const row = ws[i];
    Logger.log("Row " + i + " empId=" + row[0] + " dateCells=" + JSON.stringify(dateCols.map(d => ({ date: d.date, val: row[d.col], type: typeof row[d.col] }))));
  }
}

function parseBangaloreData(sheets) {
  const result = {};
  const ws = sheets[Object.keys(sheets)[0]];
  if (!ws || ws.length < 2) return result;
  const hdr = ws[0];
  const iId = hdr.findIndex((h) =>
    String(h || "")
      .toLowerCase()
      .includes("employee id"),
  );
  const dateCols = [];
  for (let c = 0; c < hdr.length; c++) {
    const d = parseColDate(hdr[c], 2026);
    if (d) dateCols.push({ col: c, date: d });
  }
  for (let i = 1; i < ws.length; i++) {
    const row = ws[i];
    const empId = normId(row[iId]);
    if (!empId) continue;
    for (const { col, date } of dateCols) {
      const s = String(row[col] || "")
        .trim()
        .toUpperCase();
      let entry;
      if (s === "WFO") entry = { type: "office" };
      else if (LEAVE_CODES.has(s)) entry = { type: "leave", value: 1 };
      else entry = { type: "absent" };
      if (!result[empId]) result[empId] = {};
      result[empId][date] = entry;
    }
  }
  return result;
}

function parseZohoData(sheets) {
  const leaveMap = {},
    emailMap = {};
  const ws = sheets[Object.keys(sheets)[0]];
  if (!ws) return { leaveMap, emailMap };
  let hdrRow = -1;
  for (let i = 0; i < ws.length; i++) {
    if (
      ws[i].some((c) =>
        String(c || "")
          .toLowerCase()
          .includes("employee id"),
      )
    ) {
      hdrRow = i;
      break;
    }
  }
  if (hdrRow < 0) return { leaveMap, emailMap };
  const hdr = ws[hdrRow];
  const iId = hdr.findIndex((h) =>
    String(h || "")
      .toLowerCase()
      .includes("employee id"),
  );
  const iEmail = hdr.findIndex((h) =>
    String(h || "")
      .toLowerCase()
      .includes("email"),
  );
  const dateCols = [];
  for (let c = 0; c < hdr.length; c++) {
    const d = parseColDate(hdr[c], 2026);
    if (d) dateCols.push({ col: c, date: d });
  }
  for (let i = hdrRow + 1; i < ws.length; i++) {
    const row = ws[i];
    const empId = normId(row[iId]);
    if (!empId) continue;
    if (iEmail >= 0 && row[iEmail])
      emailMap[empId] = String(row[iEmail]).trim();
    for (const { col, date } of dateCols) {
      const s = String(row[col] || "")
        .trim()
        .toUpperCase();
      if (s === "-" || s === "") continue;
      const h = halfDay(s);
      let entry;
      if (h) entry = { type: "half_leave", value: h };
      else if (LEAVE_CODES.has(s)) entry = { type: "leave", value: 1 };
      else continue;
      if (!leaveMap[empId]) leaveMap[empId] = {};
      if (!leaveMap[empId][date]) leaveMap[empId][date] = entry;
    }
  }
  return { leaveMap, emailMap };
}

function parseUllenAyyahData(sheets) {
  const officeMap = {},
    employees = [];
  const OFFICE_UA = new Set(["WFC", "WFB", "WFCO"]);
  const ws = sheets[Object.keys(sheets)[0]];
  if (!ws || ws.length < 2) return { officeMap, employees };
  const hdr = ws[0];
  const iId = hdr.findIndex((h) =>
    String(h || "")
      .toLowerCase()
      .includes("employee id"),
  );
  const iName = hdr.findIndex(
    (h) =>
      String(h || "")
        .toLowerCase()
        .includes("employee name") || String(h || "").toLowerCase() === "name",
  );
  const iEmail = hdr.findIndex((h) =>
    String(h || "")
      .toLowerCase()
      .includes("email"),
  );
  const dateCols = [];
  for (let c = 0; c < hdr.length; c++) {
    const d = parseColDate(hdr[c], 2026);
    if (d) dateCols.push({ col: c, date: d });
  }
  for (let i = 1; i < ws.length; i++) {
    const row = ws[i];
    const empId = normId(row[iId]);
    if (!empId) continue;
    const name = iName >= 0 ? String(row[iName] || "").trim() : "";
    const email = iEmail >= 0 ? String(row[iEmail] || "").trim() : "";
    if (name) employees.push({ id: empId, name, email });
    for (const { col, date } of dateCols) {
      const s = String(row[col] || "")
        .trim()
        .toUpperCase();
      if (s === "-" || s === "") continue;
      let entry;
      if (OFFICE_UA.has(s)) entry = { type: "office" };
      else if (s === "L") entry = { type: "leave", value: 1 };
      else if (s === "WR") entry = { type: "remote" };
      else continue;
      if (!officeMap[empId]) officeMap[empId] = {};
      if (!officeMap[empId][date]) officeMap[empId][date] = entry;
    }
  }
  return { officeMap, employees };
}

function parseExcuseData(sheets) {
  const result = {};
  const ws = sheets[Object.keys(sheets)[0]];
  if (!ws || ws.length < 2) return result;
  const hdr = ws[0].map((h) =>
    String(h || "")
      .trim()
      .toLowerCase(),
  );
  const iId = hdr.findIndex((h) => h.includes("employee id") || h === "id");
  const iDate = hdr.findIndex((h) => h.includes("date"));
  for (let i = 1; i < ws.length; i++) {
    const row = ws[i];
    const empId = normId(row[iId]);
    const date = toISO(iDate >= 0 ? row[iDate] : null);
    if (empId && date) {
      if (!result[empId]) result[empId] = {};
      result[empId][date] = true;
    }
  }
  return result;
}

// ── MERGE ENGINE ──────────────────────────────────────────────
function computeCompliance(
  employees,
  weekdays,
  chennai,
  bangalore,
  zoho,
  ullen,
  excuse,
) {
  return employees.map((emp) => {
    const id = emp.id;
    let total = 0;
    const dayBreakdown = {};
    for (const date of weekdays) {
      const cE = (chennai[id] || {})[date];
      const bE = (bangalore[id] || {})[date];
      const zE = (zoho[id] || {})[date];
      const uE = (ullen[id] || {})[date];
      const ex = (excuse[id] || {})[date];
      let score = 0,
        type = "absent";
      if (cE && cE.type === "office") {
        score = 1;
        type = "office";
      } else if (bE && bE.type === "office") {
        score = 1;
        type = "office";
      } else if (cE && cE.type === "leave") {
        score = cE.value ?? 1;
        type = "leave";
      } else if (cE && cE.type === "half_leave") {
        score = cE.value ?? 0.5;
        type = "half_leave";
      } else if (bE && bE.type === "leave") {
        score = 1;
        type = "leave";
      } else if (zE && zE.type === "leave") {
        score = 1;
        type = "leave";
      } else if (zE && zE.type === "half_leave") {
        score = zE.value ?? 0.5;
        type = "half_leave";
      } else if (uE && uE.type === "office") {
        score = 1;
        type = "office";
      } else if (uE && uE.type === "leave") {
        score = 1;
        type = "leave";
      } else if (ex) {
        score = 1;
        type = "excuse";
      }
      total += score;
      dayBreakdown[date] = { score, type };
    }
    return {
      ...emp,
      total,
      compliant: total >= CONFIG.MIN_OFFICE_DAYS,
      dayBreakdown,
    };
  });
}

// ── EMAIL BUILDER ─────────────────────────────────────────────
function buildEmailBody(emp, weekLabel, weekdays, dayNames) {
  const lbl = {
    office: "✓  Office",
    leave: "✓  Leave",
    half_leave: "✓  Half-day leave",
    excuse: "✓  Approved excuse",
    absent: "✗  Absent",
    remote: "✗  Remote (not counted)",
  };
  const lines = weekdays
    .map((d, i) => {
      const t = emp.dayBreakdown[d]?.type || "absent";
      return `  ${dayNames[i].padEnd(4)}  ${lbl[t] || "✗  Absent"}`;
    })
    .join("\n");
  const missing = Math.max(0, 3 - emp.total);
  return `Dear ${emp.name},\n\nThis is a reminder regarding the office attendance policy which requires a minimum of 3 days of in-office presence per week (Monday–Friday).\n\nYour attendance summary for the week of ${weekLabel}:\n\n${lines}\n\n  Total effective days : ${emp.total} / 5\n  Required             : 3 / 5\n  Shortfall            : ${missing} day${missing !== 1 ? "s" : ""}\n\nIf you believe this is incorrect, please reach out to your manager or HR.\n\nGoing forward, please ensure you are present in the office for at least 3 working days each week.\n\nRegards,\nHR & Administration`;
}
