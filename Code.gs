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

// FIX 3: Chennai biometric is ONE sheet with date columns, not multiple sheets.
// Columns: Employee, First Name, Last Name, Seating Location, [date cols...], No. of Days, Remarks, Status
// Cell values: "09:45" = in office, "EL"/"SCL" = leave, blank/NaN = absent
function parseChennaiData(sheets) {
  const attendance = {};
  const employees = [];
  const ws = sheets[Object.keys(sheets)[0]];
  if (!ws || ws.length < 2) return { attendance, employees };

  const hdr = ws[0];

  let iId = hdr.findIndex((h) => {
    const s = String(h || "").trim().toLowerCase();
    return s === "employee" || s === "employee id" || s.includes("employee id");
  });
  if (iId < 0) return { attendance, employees };

  // Find First Name / Last Name columns for building full name
  const iFirst = hdr.findIndex(h => /first\s*name/i.test(String(h||"")));
  const iLast  = hdr.findIndex(h => /last\s*name/i.test(String(h||"")));
  const iName  = hdr.findIndex(h => /^name$/i.test(String(h||"").trim()));

  const dateCols = [];
  for (let c = 0; c < hdr.length; c++) {
    const d = parseColDate(hdr[c], 2026);
    if (d) dateCols.push({ col: c, date: d });
  }

  for (let i = 1; i < ws.length; i++) {
    const row = ws[i];
    const empId = normId(row[iId]);
    if (!empId) continue;

    // Build name from First+Last or Name column
    let name = "";
    if (iFirst >= 0 || iLast >= 0) {
      const f = String(row[iFirst] || "").trim();
      const l = String(row[iLast]  || "").trim();
      name = [f, l].filter(Boolean).join(" ");
    } else if (iName >= 0) {
      name = String(row[iName] || "").trim();
    }
    if (name) employees.push({ id: empId, name, email: "" });

    if (!attendance[empId]) attendance[empId] = {};
    for (const { col, date } of dateCols) {
      const val = row[col];
      let entry;
      if (val === null || val === "" || val === undefined) {
        entry = { type: "absent" };
      } else if (val instanceof Date) {
        entry = { type: "office", value: Utilities.formatDate(val, "UTC", "HH:mm") };
      } else {
        const s = String(val).trim().toUpperCase();
        if (s === "" || s === "NAN") {
          entry = { type: "absent" };
        } else if (LEAVE_CODES.has(s)) {
          entry = { type: "leave", value: 1 };
        } else if (isTime(s)) {
          entry = { type: "office", value: s };
        } else if (s.length > 0) {
          entry = { type: "office", value: s };
        } else {
          entry = { type: "absent" };
        }
      }
      attendance[empId][date] = entry;
    }
  }
  return { attendance, employees };
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
  const iId    = hdr.findIndex(h => String(h||"").toLowerCase().includes("employee id"));
  const iEmail = hdr.findIndex(h => String(h||"").toLowerCase().includes("email"));
  const iName  = hdr.findIndex(h => /employee\s*name/i.test(String(h||"")) || /^name$/i.test(String(h||"").trim()));
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
    if (iEmail >= 0 && row[iEmail]) emailMap[empId] = String(row[iEmail]).trim();
    const name = iName >= 0 ? String(row[iName] || "").trim() : "";
    if (name) employees.push({ id: empId, name, email: emailMap[empId] || "" });
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
      if (OFFICE_UA_BIOMETRIC.has(s))   entry = { type: "office" };      // WFC/WFB — biometric cross-check applies
      else if (OFFICE_UA_CLIENT.has(s)) entry = { type: "office_co" };   // WFCO — client office, no biometric needed
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
      const hasBiometricOffice = (cE && cE.type === "office") || (bE && bE.type === "office");
      const hasWfcWfb  = uE && uE.type === "office";    // WFC or WFB — biometric cross-check applies
      const hasWfco    = uE && uE.type === "office_co"; // WFCO — client office, no biometric needed

      if (hasWfco) {
        // WFCO: always counts as office, never flagged
        score = 1; type = "office";
      } else if (hasBiometricOffice && hasWfcWfb) {
        // Both biometric AND Ullen Ayyah (WFC/WFB) agree → clean office day
        score = 1; type = "office";
      } else if (hasBiometricOffice && !uE) {
        // Biometric says office but person never marked in Ullen Ayyah → bio defaulter
        score = 1; type = "flagged_bio";
      } else if (hasWfcWfb && !hasBiometricOffice) {
        // WFC/WFB in Ullen Ayyah but no biometric record → UA defaulter
        score = 1; type = "flagged_ua";
      } else if (hasBiometricOffice) {
        // Biometric office + Ullen Ayyah has something (leave/remote) — biometric wins
        score = 1; type = "office";
      } else if (cE && cE.type === "leave") {
        score = cE.value ?? 1; type = "leave";
      } else if (cE && cE.type === "half_leave") {
        score = cE.value ?? 0.5; type = "half_leave";
      } else if (bE && bE.type === "leave") {
        score = 1; type = "leave";
      } else if (zE && zE.type === "leave") {
        score = 1; type = "leave";
      } else if (zE && zE.type === "half_leave") {
        score = zE.value ?? 0.5; type = "half_leave";
      } else if (uE && uE.type === "leave") {
        score = 1; type = "leave";
      } else if (ex) {
        score = 1; type = "excuse";
      }
      total += score;
      dayBreakdown[date] = { score, type };
    }
    const roundedTotal = Math.ceil(total);
    const hasUaFlag  = Object.values(dayBreakdown).some(d => d.type === "flagged_ua");
    const hasBioFlag = Object.values(dayBreakdown).some(d => d.type === "flagged_bio");
    const hasFlagged = hasUaFlag || hasBioFlag;
    return {
      ...emp,
      total: roundedTotal,
      flagged: hasFlagged,
      flagged_ua:  hasUaFlag,
      flagged_bio: hasBioFlag,
      compliant: !hasFlagged && roundedTotal >= CONFIG.MIN_OFFICE_DAYS,
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
