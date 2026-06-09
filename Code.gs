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

    const { attendance: chennai, employees: chennaiEmps } = parseChennaiData(
      decodeFile(chennaiB64, ft.chennai || "xlsx"),
    );
    const { attendance: bangalore, employees: bangaloreEmps } = parseBangaloreData(
      decodeFile(bangaloreB64, ft.bangalore || "csv"),
    );
    const { leaveMap: zoho, emailMap, employees: zohoEmps } = parseZohoData(
      decodeFile(zohoB64, ft.zoho || "xls"),
    );
    const { officeMap: ullen, employees: ullenEmps } = parseUllenAyyahData(
      decodeFile(ullenB64, ft.ullen || "csv"),
    );
    const excuse = excuseB64
      ? parseExcuseData(decodeFile(excuseB64, ft.excuse || "xlsx"))
      : {};

    // Zoho is the employee master — only Zoho employees are included.
    // Biometric (Chennai/Bangalore) and Ullen Ayyah data are used for attendance
    // computation only; anyone not in Zoho is excluded from the results.
    const SYSTEM_NAMES = new Set(["acies global", "acies auditor", "acies", "auditor auditor"]);
    const SYSTEM_EMAILS = new Set(["license@aciesglobal.com", "auditor@aciesglobal.com"]);
    const isSystemRow = e => {
      const name  = (e.name  || "").trim().toLowerCase();
      const email = (e.email || "").trim().toLowerCase();
      return SYSTEM_EMAILS.has(email) || SYSTEM_NAMES.has(name);
    };
    const seenIds = new Set();
    const employees = [];
    for (const e of zohoEmps) {
      if (!seenIds.has(e.id) && e.name && !isSystemRow(e)) {
        seenIds.add(e.id);
        employees.push({ ...e, email: e.email || emailMap[e.id] || "" });
      }
    }

    if (!employees || employees.length === 0)
      return {
        ok: false,
        error: "No employees found across any file. Check the format.",
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

    // Auto-write all employees to Compliance Results sheet immediately
    writeComplianceSheet(results, weekdays, getWeekLabel(weekStart, weekEnd));

    return { ok: true, results, weekdays };
  } catch (e) {
    return { ok: false, error: e.message + "\n" + e.stack };
  }
}

// ── WEEK LABEL ────────────────────────────────────────────────
function getWeekLabel(start, end) {
  const fmt = d => Utilities.formatDate(new Date(d + "T00:00:00"), Session.getScriptTimeZone(), "MMM d");
  return fmt(start) + " – " + fmt(end);
}

// ── WRITE ALL EMPLOYEES TO COMPLIANCE RESULTS SHEET ──────────
function writeComplianceSheet(results, weekdays, weekLabel) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let rs = ss.getSheetByName(CONFIG.RESULTS_SHEET);
    if (!rs) rs = ss.insertSheet(CONFIG.RESULTS_SHEET);

    const dayNames = ["Mon", "Tue", "Wed", "Thu", "Fri"];

    const headers = [
      "Employee ID", "Employee Name", "Email", "Week",
      ...dayNames,
      "Total Days", "Compliant", "Flagged", "Status",
    ];

    // Build data rows for ALL employees
    const rows = results.map(r => {
      const dayVals = weekdays.map(d => r.dayBreakdown[d]?.type || "absent");
      let status;
      if (r.flagged_ua)      status = "UA Defaulter";
      else if (r.flagged_bio)status = "Bio Defaulter";
      else if (r.compliant)  status = "Compliant";
      else                   status = "Non-Compliant";

      return [
        r.id,
        r.name,
        r.email || "",
        weekLabel,
        ...dayVals,
        r.total,
        r.flagged ? "—" : (r.compliant ? "Yes" : "No"),
        r.flagged ? "Yes" : "No",
        status,
      ];
    });

    // Clear & rewrite sheet
    rs.clearContents();
    rs.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setBackground("#1e2026")
      .setFontColor("#e6a817")
      .setFontWeight("bold")
      .setFontFamily("Courier New");
    rs.setFrozenRows(1);

    if (rows.length > 0) {
      rs.getRange(2, 1, rows.length, rows[0].length).setValues(rows);

      // Colour-code the Status column (last col)
      const statusCol = headers.length;
      for (let i = 0; i < rows.length; i++) {
        const cell = rs.getRange(i + 2, statusCol);
        const s = rows[i][statusCol - 1];
        if (s === "Compliant")          cell.setFontColor("#3fb950");
        else if (s === "UA Defaulter")  cell.setFontColor("#f0883e");
        else if (s === "Bio Defaulter") cell.setFontColor("#bc8cff");
        else                             cell.setFontColor("#f85149");
      }

      // Auto-resize columns for readability
      for (let c = 1; c <= headers.length; c++) rs.autoResizeColumn(c);
    }
  } catch (e) {
    Logger.log("writeComplianceSheet error: " + e.message);
  }
}

function sendEmails(payload) {
  try {
    const { results, weekLabel, weekdays, emailTemplate, emailOverrides } = payload;
    const dayNames = ["Mon", "Tue", "Wed", "Thu", "Fri"];
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const ls =
      ss.getSheetByName(CONFIG.LOG_SHEET) || ss.insertSheet(CONFIG.LOG_SHEET);

    const now = new Date();
    const sent = [], failed = [], noEmail = [];
    const logRows = [];

    // Only email non-compliant & non-flagged employees (selected subset from UI)
    for (const r of results) {
      const dayVals = weekdays.map(d => r.dayBreakdown[d]?.type || "absent");
      if (r.flagged || r.compliant) continue; // skip — only email non-compliant

      let emailStatus = "Pending", sentAt = "";
      if (!r.email || !r.email.includes("@")) {
        emailStatus = "No email on record";
        noEmail.push(r.name);
      } else {
        try {
          const overrides = emailOverrides || {};
          const body = overrides[r.id] !== undefined
            ? overrides[r.id]
            : buildEmailBody(r, weekLabel, weekdays, dayNames, emailTemplate);
          GmailApp.sendEmail(
            r.email,
            `Office Attendance Reminder — Week of ${weekLabel}`,
            body,
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

      logRows.push([
        now.toLocaleString(),
        weekLabel,
        r.id,
        r.name,
        r.email,
        r.total,
        ...dayVals,
        emailStatus,
        sentAt,
      ]);
    }

    // Append to Email Log only
    if (logRows.length > 0) {
      const lastRow = ls.getLastRow();
      // Ensure header if sheet is empty
      if (lastRow === 0) {
        const lHeaders = ["Timestamp","Week","Employee ID","Employee Name","Email","Total Days","Mon","Tue","Wed","Thu","Fri","Email Status","Sent At"];
        ls.getRange(1, 1, 1, lHeaders.length)
          .setValues([lHeaders])
          .setBackground("#1e2026")
          .setFontColor("#e6a817")
          .setFontWeight("bold");
      }
      ls.getRange(ls.getLastRow() + 1, 1, logRows.length, logRows[0].length).setValues(logRows);
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

// Chennai biometric — auto-detects one of 3 formats:
//   LONG     : one sheet, each row = one employee + one date (transaction log)
//   MULTISHEET: tabs named "18 May", "19 May" … each tab = one day
//   WIDE     : one sheet, one row per employee, date columns across the top
function parseChennaiData(sheets) {
  const attendance = {}, employees = [];

  // ── SHARED UTILITIES ────────────────────────────────────────────

  // Find the first row that looks like a real header (≥3 non-empty cells)
  function findHdrRow(ws) {
    for (let r = 0; r < Math.min(6, ws.length); r++) {
      if (ws[r].filter(c => String(c||"").trim() !== "").length >= 3) return r;
    }
    return 0;
  }

  // Find a column index by testing header cells against a list of matcher fns (first match wins)
  function findCol(hdr, ...matchers) {
    for (const fn of matchers) {
      const i = hdr.findIndex(h => fn(String(h||"").trim()));
      if (i >= 0) return i;
    }
    return -1;
  }

  // Parse a clock-in cell to an attendance entry
  function clockEntry(ci) {
    if (ci === null || ci === "" || ci === undefined) return { type: "no_record" };
    if (ci instanceof Date) {
      const t = Utilities.formatDate(ci, "UTC", "HH:mm");
      return (t === "00:00") ? { type: "no_record" } : { type: "office", value: t };
    }
    const s = String(ci).trim().toUpperCase();
    if (!s || s === "NAN" || s === "-" || s === "0") return { type: "no_record" };
    if (LEAVE_CODES.has(s) || s.startsWith("0.5")) return { type: "leave", value: 1 };
    if (isTime(s)) return { type: "office", value: s };
    return { type: "no_record" };
  }

  // Build name from first+last or full-name columns
  function buildName(row, iFirst, iLast, iName) {
    if (iFirst >= 0 || iLast >= 0)
      return [String(row[iFirst]||"").trim(), String(row[iLast]||"").trim()].filter(Boolean).join(" ");
    if (iName >= 0) return String(row[iName]||"").trim();
    return "";
  }

  // Employee ID column matchers
  const EMP_ID_MATCHERS = [
    h => h.toLowerCase() === "employee id",
    h => h.toLowerCase() === "emp id",
    h => h.toLowerCase() === "empid",
    h => h.toLowerCase() === "employee",
    h => h.toLowerCase().includes("employee id"),
    h => h.toLowerCase().includes("emp id"),
  ];

  // Name column matchers
  const NAME_MATCHERS     = [h => /employee\s*name/i.test(h), h => /emp\s*name/i.test(h), h => /^name$/i.test(h)];
  const FIRST_MATCHERS    = [h => /first\s*name/i.test(h)];
  const LAST_MATCHERS     = [h => /last\s*name/i.test(h)];

  // Clock-In column matchers (header-based)
  const CLOCK_IN_MATCHERS = [
    h => /clock\s*in/i.test(h),
    h => /\bin\s*time\b/i.test(h),
    h => /time\s*in/i.test(h),
    h => /check\s*in/i.test(h),
    h => /entry\s*time/i.test(h),
    h => /punch\s*in/i.test(h),
    h => /in\s*punch/i.test(h),
    h => /^in$/i.test(h),
  ];

  // Date column matchers (header-based)
  const DATE_MATCHERS = [
    h => /\bdate\b/i.test(h),
    h => /atd\s*date/i.test(h),
    h => /attendance\s*date/i.test(h),
  ];

  // ── FORMAT DETECTION ─────────────────────────────────────────────
  const firstWs = sheets[Object.keys(sheets)[0]];
  const daySheetNames = Object.keys(sheets).filter(n => /^\d{1,2}\s+[A-Za-z]+$/i.test(n.trim()));

  // Try to detect long format: find a column where data rows contain date values
  function detectLongDateCol(ws) {
    if (!ws || ws.length < 3) return -1;
    const hr = findHdrRow(ws);
    const hdr = ws[hr];

    // 1. Header name match
    let idx = findCol(hdr, ...DATE_MATCHERS);
    if (idx >= 0) return idx;

    // 2. Data content scan: find a column where ≥ half of first 6 data rows look like dates
    const checkRows = Math.min(6, ws.length - hr - 1);
    if (checkRows < 1) return -1;
    for (let c = 0; c < hdr.length; c++) {
      let hits = 0;
      for (let r = hr + 1; r <= hr + checkRows; r++) {
        const v = ws[r] ? ws[r][c] : null;
        if (!v && v !== 0) continue;
        if (v instanceof Date) { hits++; continue; }
        const s = String(v).trim();
        if (/^\d{1,2}[-\/]\d{1,2}[-\/]\d{4}$/.test(s) || /^\d{4}-\d{2}-\d{2}$/.test(s)) hits++;
      }
      if (hits >= Math.ceil(checkRows / 2)) return c;
    }
    return -1;
  }

  // Try to find Clock-In column by data content (time values) — used as last resort
  function detectClockColByData(ws, hr, afterCol) {
    const checkRows = Math.min(6, ws.length - hr - 1);
    if (checkRows < 1) return -1;
    for (let c = afterCol; c < (ws[hr]||[]).length; c++) {
      let hits = 0;
      for (let r = hr + 1; r <= hr + checkRows; r++) {
        const v = ws[r] ? ws[r][c] : null;
        if (v instanceof Date) { hits++; continue; }
        if (isTime(String(v||"").trim())) hits++;
      }
      if (hits >= Math.ceil(checkRows / 2)) return c;
    }
    return -1;
  }

  // ── PROCESS ──────────────────────────────────────────────────────
  // Priority: multi-sheet first (day-named tabs are unambiguous),
  // then long format, then wide format.

  if (daySheetNames.length > 0) {
    // ════════════════════════════════════════════════
    //  MULTI-SHEET FORMAT  (one tab per day)
    // ════════════════════════════════════════════════
    for (const sheetName of daySheetNames) {
      const ws = sheets[sheetName];
      if (!ws || ws.length < 2) continue;
      const date = parseDaySheetName(sheetName.trim(), 2026);
      if (!date) continue;

      const hr  = findHdrRow(ws);
      const hdr = ws[hr];

      const iId    = findCol(hdr, ...EMP_ID_MATCHERS);
      if (iId < 0) continue;
      const iFirst = findCol(hdr, ...FIRST_MATCHERS);
      const iLast  = findCol(hdr, ...LAST_MATCHERS);
      const iName  = findCol(hdr, ...NAME_MATCHERS);
      let iClock   = findCol(hdr, ...CLOCK_IN_MATCHERS);
      if (iClock < 0) iClock = detectClockColByData(ws, hr, iId + 1);

      for (let i = hr + 1; i < ws.length; i++) {
        const row = ws[i];
        const empId = normId(row[iId]);
        if (!empId) continue;
        const name = buildName(row, iFirst, iLast, iName);
        if (name) employees.push({ id: empId, name, email: "" });
        const entry = iClock >= 0 ? clockEntry(row[iClock]) : { type: "no_record" };
        if (!attendance[empId]) attendance[empId] = {};
        attendance[empId][date] = entry;
      }
    }

  } else {
    // No day-named tabs — check for long format, then fall back to wide
    const iLongDate = detectLongDateCol(firstWs);

    if (iLongDate >= 0) {
      // ════════════════════════════════════════════════
      //  LONG FORMAT  (each row = one employee-date)
      // ════════════════════════════════════════════════
      const ws  = firstWs;
      const hr  = findHdrRow(ws);
      const hdr = ws[hr];

      const iId    = findCol(hdr, ...EMP_ID_MATCHERS);
      if (iId >= 0) {
        const iFirst = findCol(hdr, ...FIRST_MATCHERS);
        const iLast  = findCol(hdr, ...LAST_MATCHERS);
        const iName  = findCol(hdr, ...NAME_MATCHERS);
        let iClock   = findCol(hdr, ...CLOCK_IN_MATCHERS);
        if (iClock < 0) iClock = detectClockColByData(ws, hr, iLongDate + 1);

        const seenEmp = new Set();
        for (let i = hr + 1; i < ws.length; i++) {
          const row = ws[i];
          const empId = normId(row[iId]);
          if (!empId) continue;
          const name = buildName(row, iFirst, iLast, iName);
          if (name && !seenEmp.has(empId)) { seenEmp.add(empId); employees.push({ id: empId, name, email: "" }); }
          const date = toISO(row[iLongDate]);
          if (!date) continue;
          const entry = iClock >= 0 ? clockEntry(row[iClock]) : { type: "no_record" };
          if (!attendance[empId]) attendance[empId] = {};
          if (!attendance[empId][date]) attendance[empId][date] = entry;
        }
      }

    } else {
      // ════════════════════════════════════════════════
      //  WIDE FORMAT  (one row per employee, date cols)
      // ════════════════════════════════════════════════
    const ws = firstWs;
    if (!ws || ws.length < 2) return { attendance, employees };
    const hr  = findHdrRow(ws);
    const hdr = ws[hr];

    const iId    = findCol(hdr, ...EMP_ID_MATCHERS);
    if (iId < 0) return { attendance, employees };
    const iFirst = findCol(hdr, ...FIRST_MATCHERS);
    const iLast  = findCol(hdr, ...LAST_MATCHERS);
    const iName  = findCol(hdr, ...NAME_MATCHERS);

    const dateCols = [];
    for (let c = 0; c < hdr.length; c++) {
      const d = parseColDate(hdr[c], 2026);
      if (d) dateCols.push({ col: c, date: d });
    }

    for (let i = hr + 1; i < ws.length; i++) {
      const row = ws[i];
      const empId = normId(row[iId]);
      if (!empId) continue;
      const name = buildName(row, iFirst, iLast, iName);
      if (name) employees.push({ id: empId, name, email: "" });
      if (!attendance[empId]) attendance[empId] = {};
      for (const { col, date } of dateCols) {
        attendance[empId][date] = clockEntry(row[col]);
      }
    }
    } // end wide format else
  } // end outer else (no day-named tabs)

  return { attendance, employees };
}

// Parse day-sheet tab names like "19 May", "19-May", "19 May 2026"
function parseDaySheetName(s, yr) {
  const mo = { jan:"01",feb:"02",mar:"03",apr:"04",may:"05",jun:"06",jul:"07",aug:"08",sep:"09",oct:"10",nov:"11",dec:"12" };
  let m = s.match(/^(\d{1,2})\s+([A-Za-z]+)(?:\s+(\d{4}))?$/);
  if (m) { const x=mo[m[2].toLowerCase().slice(0,3)], y=m[3]||yr; if(x) return `${y}-${x}-${m[1].padStart(2,"0")}`; }
  m = s.match(/^(\d{1,2})-([A-Za-z]+)(?:-(\d{4}))?$/);
  if (m) { const x=mo[m[2].toLowerCase().slice(0,3)], y=m[3]||yr; if(x) return `${y}-${x}-${m[1].padStart(2,"0")}`; }
  return null;
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
  const attendance = {};
  const employees = [];
  const ws = sheets[Object.keys(sheets)[0]];
  if (!ws || ws.length < 2) return { attendance, employees };
  const hdr = ws[0];
  const iId   = hdr.findIndex(h => String(h||"").toLowerCase().includes("employee id"));
  const iName = hdr.findIndex(h => /employee\s*name/i.test(String(h||"")) || /^name$/i.test(String(h||"").trim()));
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
    if (name) employees.push({ id: empId, name, email: "" });
    for (const { col, date } of dateCols) {
      const s = String(row[col] || "").trim().toUpperCase();
      let entry;
      if (s === "WFO") entry = { type: "office" };
      else if (LEAVE_CODES.has(s)) entry = { type: "leave", value: 1 };
      else entry = { type: "absent" };
      if (!attendance[empId]) attendance[empId] = {};
      attendance[empId][date] = entry;
    }
  }
  return { attendance, employees };
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
  if (hdrRow < 0) return { leaveMap, emailMap, employees: [] };
  const hdr = ws[hdrRow];
  const iId       = hdr.findIndex(h => String(h||"").toLowerCase().includes("employee id"));
  const iEmail    = hdr.findIndex(h => String(h||"").toLowerCase().includes("email"));
  const iName     = hdr.findIndex(h => /employee\s*name/i.test(String(h||"")) || /^name$/i.test(String(h||"").trim()));

  // Location column — try several Zoho field name variants
  const iLocation = hdr.findIndex(h => {
    const s = String(h||"").trim().toLowerCase();
    return s === "location" || s === "branch" || s === "office location"
        || s === "work location" || s.includes("location");
  });

  // Reporting manager name column — must NOT match "Reporting Manager Email"
  const iMgrName  = hdr.findIndex(h => {
    const s = String(h||"").trim().toLowerCase();
    if (s.includes("email")) return false;          // never match email columns
    return s === "reporting manager" || s === "manager name" || s === "manager"
        || s.includes("reporting manager") || s.includes("manager name");
  });

  // Reporting manager email column
  const iMgrEmail = hdr.findIndex(h => {
    const s = String(h||"").trim().toLowerCase();
    return s === "reporting manager email" || s === "manager email"
        || s.includes("manager email") || s.includes("reporting manager email");
  });

  const dateCols = [];
  for (let c = 0; c < hdr.length; c++) {
    const d = parseColDate(hdr[c], 2026);
    if (d) dateCols.push({ col: c, date: d });
  }
  const employees = [];
  for (let i = hdrRow + 1; i < ws.length; i++) {
    const row = ws[i];
    const empId = normId(row[iId]);
    if (!empId) continue;
    if (iEmail >= 0 && row[iEmail]) {
      const em = String(row[iEmail]).trim();
      if (em && em !== "#N/A" && !em.startsWith("#") && em.includes("@")) emailMap[empId] = em;
    }
    // Helper: return cleaned string, treating #N/A / 0 / errors as empty
    const clean = v => {
      if (v === null || v === undefined || v === 0 || v === false) return "";
      const s = String(v).trim();
      return (s === "" || s === "#N/A" || s === "0" || s.startsWith("#")) ? "" : s;
    };
    const name     = iName     >= 0 ? clean(row[iName])     : "";
    const location = iLocation >= 0 ? clean(row[iLocation]) : "";
    const mgrName  = iMgrName  >= 0 ? clean(row[iMgrName])  : "";
    const mgrEmail = iMgrEmail >= 0 ? clean(row[iMgrEmail]) : "";
    if (name) employees.push({ id: empId, name, email: emailMap[empId] || "", location, managerName: mgrName, managerEmail: mgrEmail });
    for (const { col, date } of dateCols) {
      const s = String(row[col] || "").trim().toUpperCase();
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
  return { leaveMap, emailMap, employees };
}

function parseUllenAyyahData(sheets) {
  const officeMap = {},
    employees = [];
  const OFFICE_UA_BIOMETRIC = new Set(["WFC", "WFB"]); // biometric expected for these
  const OFFICE_UA_CLIENT   = new Set(["WFCO"]);         // client office — no biometric needed
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
      if (OFFICE_UA_BIOMETRIC.has(s))   entry = { type: "office", code: s };   // WFC=Chennai / WFB=Bangalore
      else if (OFFICE_UA_CLIENT.has(s)) entry = { type: "office_co", code: s }; // WFCO — client office
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
function computeCompliance(employees, weekdays, chennai, bangalore, zoho, ullen, excuse) {
  return employees.map(emp => {
    const id = emp.id;
    let total = 0;
    const dayBreakdown = {};

    for (const date of weekdays) {
      const cE = (chennai[id] || {})[date];
      const bE = (bangalore[id] || {})[date];
      const zE = (zoho[id]    || {})[date];
      const uE = (ullen[id]   || {})[date];

      // ── SCORING PRIORITY ──────────────────────────────────────────
      const eE = (excuse[id] || {})[date];

      const hasBioOffice  = (cE && cE.type === "office") || (bE && bE.type === "office");
      const hasBioLeave   = (cE && cE.type === "leave")  || (bE && bE.type === "leave");
      const hasZohoLeave  = zE && (zE.type === "leave" || zE.type === "half_leave");
      const hasUllenLeave = uE && uE.type === "leave";          // L in Ullen Ayyah → counts as leave
      const hasWfco       = uE && uE.type === "office_co";      // WFCO → score 1, no biometric needed
      const hasExcuse     = !!eE;                               // Excuse sheet → display only, score 0
      const hasRemote     = uE && uE.type === "remote";         // WR → display only, score 0
      const hasWfcWfb     = uE && uE.type === "office";         // WFC/WFB without biometric → flagged

      let score = 0, type = "no_record";

      if (hasBioOffice) {
        score = 1; type = "office";
      } else if (hasBioLeave || hasZohoLeave) {
        score = 1; type = "leave";
      } else if (hasUllenLeave) {
        score = 1; type = "leave";                              // Ullen Ayyah L = leave, score 1
      } else if (hasWfco) {
        score = 1; type = "wfco";                              // WFCO = counts, shown separately
      } else if (hasExcuse) {
        score = 0; type = "excuse";                            // Excused = display only
      } else if (hasRemote) {
        score = 0; type = "remote";
      } else if (hasWfcWfb) {
        score = 0; type = "claimed";                           // Reacted WFC/WFB but no biometric
      }
      // else: no_record — score stays 0

      total += score;
      dayBreakdown[date] = { score, type };
    }

    return {
      ...emp,
      total,
      compliant: total >= CONFIG.MIN_OFFICE_DAYS,
      location:     emp.location     || "—",
      managerName:  emp.managerName  || "—",
      managerEmail: emp.managerEmail || "—",
      dayBreakdown,
    };
  });
}

// ── EMAIL BUILDER ─────────────────────────────────────────────
function buildEmailBody(emp, weekLabel, weekdays, dayNames, customTemplate) {
  const lbl = {
    office: "✓  Office",
    leave: "✓  Leave",
    half_leave: "✓  Half-day leave",
    excuse: "✓  Approved excuse",
    absent: "✗  Absent",
    remote: "✗  Remote (not counted)",
  };
  const breakdown = weekdays
    .map((d, i) => {
      const t = emp.dayBreakdown[d]?.type || "absent";
      return `  ${dayNames[i].padEnd(4)}  ${lbl[t] || "✗  Absent"}`;
    })
    .join("\n");
  const missing = Math.max(0, 3 - emp.total);
  const missingStr = missing + " day" + (missing !== 1 ? "s" : "");
  if (customTemplate) {
    return customTemplate
      .replace(/\{\{name\}\}/g, emp.name)
      .replace(/\{\{weekLabel\}\}/g, weekLabel)
      .replace(/\{\{breakdown\}\}/g, breakdown)
      .replace(/\{\{total\}\}/g, emp.total)
      .replace(/\{\{missing\}\}/g, missingStr);
  }
  return `Dear ${emp.name},\n\nThis is a reminder regarding the office attendance policy which requires a minimum of 3 days of in-office presence per week (Monday–Friday).\n\nYour attendance summary for the week of ${weekLabel}:\n\n${breakdown}\n\n  Total effective days : ${emp.total} / 5\n  Required             : 3 / 5\n  Shortfall            : ${missingStr}\n\nIf you believe this is incorrect, please reach out to your manager or HR.\n\nGoing forward, please ensure you are present in the office for at least 3 working days each week.\n\nRegards,\nHR & Administration`;
}
