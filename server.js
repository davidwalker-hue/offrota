const http = require("http");
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const zlib = require("zlib");

const PORT = Number(process.env.PORT || 4173);
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || "ChangeMe2026!";
const ROOT = __dirname;
const PUBLIC_DIR = path.join(ROOT, "public");
const DATA_DIR = process.env.DATA_DIR || path.join(ROOT, "data");
const LEGACY_RESPONSES_JSON = path.join(DATA_DIR, "responses.json");
const SETTINGS_JSON = path.join(DATA_DIR, "settings.json");
const LEGACY_WORKBOOK = path.join(DATA_DIR, "off-rota-responses.xlsx");
const SERVER_LOG = path.join(DATA_DIR, "server.log");

const HEADERS = [
  "Timestamp",
  "Email address",
  "Staff Initials",
  "Dates to avoid from:",
  "Dates to avoid to:",
  "Are these Dates part of a continuous period of annual leave?"
];

const MONTHS = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December"
];

function buildTitle(month, year) {
  return `Off Rota Requests for ${MONTHS[month - 1]} ${year}`;
}

function periodKey(settings) {
  return `${settings.requestYear}-${String(settings.requestMonth).padStart(2, "0")}`;
}

function periodLabel(settings) {
  return `${MONTHS[settings.requestMonth - 1]} ${settings.requestYear}`;
}

function settingsFromPeriodKey(key) {
  const match = /^(\d{4})-(\d{2})$/.exec(key);
  if (!match) return null;
  const requestYear = Number(match[1]);
  const requestMonth = Number(match[2]);
  if (requestMonth < 1 || requestMonth > 12 || requestYear < 2000 || requestYear > 2100) return null;
  return { requestMonth, requestYear };
}

function workbookFileName(settings) {
  return `off-rota-responses-${periodLabel(settings).replace(/\s+/g, "-")}.xlsx`;
}

function responsesPath(settings) {
  return path.join(DATA_DIR, `responses-${periodKey(settings)}.json`);
}

function workbookPath(settings) {
  return path.join(DATA_DIR, workbookFileName(settings));
}

const DEFAULT_SETTINGS = {
  requestMonth: 6,
  requestYear: 2026,
  title: buildTitle(6, 2026),
  deadline: "2026-05-08T09:00",
  timezoneLabel: "UK time",
  source: "Replicated from Google Drive response sheet: Off Rota Requests for June 2026 - Open until 9 am 8th May"
};

fs.mkdirSync(DATA_DIR, { recursive: true });

function logLine(message) {
  fs.appendFileSync(SERVER_LOG, `${new Date().toISOString()} ${message}\n`);
}

process.on("uncaughtException", error => {
  logLine(`uncaughtException: ${error.stack || error.message}`);
  process.exit(1);
});

process.on("unhandledRejection", error => {
  logLine(`unhandledRejection: ${error.stack || error}`);
  process.exit(1);
});

function readJson(file, fallback) {
  try {
    return JSON.parse(fs.readFileSync(file, "utf8"));
  } catch {
    return fallback;
  }
}

function writeJson(file, value) {
  fs.writeFileSync(file, JSON.stringify(value, null, 2));
}

function getSettings() {
  const settings = { ...DEFAULT_SETTINGS, ...readJson(SETTINGS_JSON, {}) };
  if (!Number.isInteger(settings.requestMonth) || settings.requestMonth < 1 || settings.requestMonth > 12) {
    settings.requestMonth = DEFAULT_SETTINGS.requestMonth;
  }
  if (!Number.isInteger(settings.requestYear) || settings.requestYear < 2000 || settings.requestYear > 2100) {
    settings.requestYear = DEFAULT_SETTINGS.requestYear;
  }
  settings.title = buildTitle(settings.requestMonth, settings.requestYear);
  writeJson(SETTINGS_JSON, settings);
  return settings;
}

function getResponses(settings) {
  const file = responsesPath(settings);
  const responses = readJson(file, []);
  writeJson(file, responses);
  return responses;
}

function writeResponses(settings, responses) {
  writeJson(responsesPath(settings), responses);
}

function deadlineHasPassed(settings) {
  return new Date(settings.deadline).getTime() <= Date.now();
}

function send(res, status, body, type = "application/json") {
  const payload = type === "application/json" ? JSON.stringify(body) : body;
  res.writeHead(status, {
    "Content-Type": type,
    "Cache-Control": "no-store"
  });
  res.end(payload);
}

function parseBody(req) {
  return new Promise((resolve, reject) => {
    let body = "";
    req.on("data", chunk => {
      body += chunk;
      if (body.length > 1_000_000) req.destroy();
    });
    req.on("end", () => {
      try {
        resolve(body ? JSON.parse(body) : {});
      } catch (error) {
        reject(error);
      }
    });
  });
}

function isValidEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value);
}

function requireAdmin(req) {
  const supplied = req.headers["x-admin-password"] || "";
  const a = Buffer.from(String(supplied));
  const b = Buffer.from(ADMIN_PASSWORD);
  return a.length === b.length && crypto.timingSafeEqual(a, b);
}

function escapeXml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function columnName(index) {
  let name = "";
  let n = index + 1;
  while (n > 0) {
    const remainder = (n - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
}

function sheetXml(rows) {
  const rowXml = rows.map((row, rowIndex) => {
    const cells = row.map((value, columnIndex) => {
      const ref = `${columnName(columnIndex)}${rowIndex + 1}`;
      return `<c r="${ref}" t="inlineStr"><is><t>${escapeXml(value)}</t></is></c>`;
    }).join("");
    return `<row r="${rowIndex + 1}">${cells}</row>`;
  }).join("");

  const lastColumn = columnName(HEADERS.length - 1);
  const lastRow = Math.max(rows.length, 1);

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:${lastColumn}${lastRow}"/>
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols>
    <col min="1" max="1" width="24" customWidth="1"/>
    <col min="2" max="2" width="34" customWidth="1"/>
    <col min="3" max="3" width="18" customWidth="1"/>
    <col min="4" max="5" width="18" customWidth="1"/>
    <col min="6" max="6" width="56" customWidth="1"/>
  </cols>
  <sheetData>${rowXml}</sheetData>
  <autoFilter ref="A1:${lastColumn}${lastRow}"/>
</worksheet>`;
}

function crc32(buffer) {
  if (!crc32.table) {
    crc32.table = Array.from({ length: 256 }, (_, n) => {
      let c = n;
      for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      return c >>> 0;
    });
  }
  let crc = 0xffffffff;
  for (const byte of buffer) crc = crc32.table[(crc ^ byte) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

function dosDateTime(date = new Date()) {
  const time = ((date.getHours() & 31) << 11) | ((date.getMinutes() & 63) << 5) | Math.floor(date.getSeconds() / 2);
  const dosDate = (((date.getFullYear() - 1980) & 127) << 9) | (((date.getMonth() + 1) & 15) << 5) | (date.getDate() & 31);
  return { time, date: dosDate };
}

function u16(value) {
  const buffer = Buffer.alloc(2);
  buffer.writeUInt16LE(value);
  return buffer;
}

function u32(value) {
  const buffer = Buffer.alloc(4);
  buffer.writeUInt32LE(value >>> 0);
  return buffer;
}

function createZip(files) {
  const localParts = [];
  const centralParts = [];
  let offset = 0;
  const stamp = dosDateTime();

  for (const file of files) {
    const name = Buffer.from(file.name);
    const content = Buffer.from(file.content);
    const compressed = zlib.deflateRawSync(content);
    const crc = crc32(content);

    const local = Buffer.concat([
      u32(0x04034b50), u16(20), u16(0), u16(8), u16(stamp.time), u16(stamp.date),
      u32(crc), u32(compressed.length), u32(content.length), u16(name.length), u16(0), name, compressed
    ]);

    const central = Buffer.concat([
      u32(0x02014b50), u16(20), u16(20), u16(0), u16(8), u16(stamp.time), u16(stamp.date),
      u32(crc), u32(compressed.length), u32(content.length), u16(name.length), u16(0), u16(0),
      u16(0), u16(0), u32(0), u32(offset), name
    ]);

    localParts.push(local);
    centralParts.push(central);
    offset += local.length;
  }

  const central = Buffer.concat(centralParts);
  const end = Buffer.concat([
    u32(0x06054b50), u16(0), u16(0), u16(files.length), u16(files.length),
    u32(central.length), u32(offset), u16(0)
  ]);

  return Buffer.concat([...localParts, central, end]);
}

function buildWorkbook(settings, responses) {
  const rows = [
    HEADERS,
    ...responses.map(response => [
      response.timestamp,
      response.email,
      response.staffInitials,
      response.dateFrom,
      response.dateTo,
      response.continuousLeave
    ])
  ];

  const files = [
    {
      name: "[Content_Types].xml",
      content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`
    },
    {
      name: "_rels/.rels",
      content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`
    },
    {
      name: "xl/workbook.xml",
      content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="${escapeXml(periodLabel(settings))}" sheetId="1" r:id="rId1"/></sheets>
</workbook>`
    },
    {
      name: "xl/_rels/workbook.xml.rels",
      content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`
    },
    {
      name: "xl/styles.xml",
      content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
</styleSheet>`
    },
    { name: "xl/worksheets/sheet1.xml", content: sheetXml(rows) }
  ];

  fs.writeFileSync(workbookPath(settings), createZip(files));
}

function ensureWorkbook(settings, force = false) {
  const responses = getResponses(settings);
  const file = workbookPath(settings);
  if (force || !fs.existsSync(file)) buildWorkbook(settings, responses);
  return file;
}

function workbookOptions(settings) {
  ensureWorkbook(settings);
  const currentKey = periodKey(settings);
  const options = new Map();

  for (const file of fs.readdirSync(DATA_DIR)) {
    const jsonMatch = /^responses-(\d{4}-\d{2})\.json$/.exec(file);
    if (jsonMatch) {
      const parsed = settingsFromPeriodKey(jsonMatch[1]);
      if (parsed) {
        const optionSettings = { ...settings, ...parsed };
        ensureWorkbook(optionSettings);
      }
    }
  }

  for (const file of fs.readdirSync(DATA_DIR)) {
    const match = /^off-rota-responses-([A-Za-z]+)-(\d{4})\.xlsx$/.exec(file);
    if (!match) continue;
    const requestMonth = MONTHS.indexOf(match[1]) + 1;
    const requestYear = Number(match[2]);
    if (!requestMonth || !requestYear) continue;
    const optionSettings = { ...settings, requestMonth, requestYear };
    const key = periodKey(optionSettings);
    options.set(key, {
      period: key,
      label: periodLabel(optionSettings),
      fileName: file,
      isCurrent: key === currentKey
    });
  }

  if (fs.existsSync(LEGACY_WORKBOOK)) {
    options.set("legacy", {
      period: "legacy",
      label: "Legacy responses",
      fileName: path.basename(LEGACY_WORKBOOK),
      isCurrent: false
    });
  }

  return [...options.values()].sort((a, b) => {
    if (a.isCurrent) return -1;
    if (b.isCurrent) return 1;
    return b.period.localeCompare(a.period);
  });
}

function validateSubmission(body) {
  const email = String(body.email || "").trim();
  const staffInitials = String(body.staffInitials || "").trim();
  const dateFrom = String(body.dateFrom || "").trim();
  const dateTo = String(body.dateTo || "").trim();
  const continuousLeave = String(body.continuousLeave || "").trim();
  const errors = [];

  if (!isValidEmail(email)) errors.push("Enter a valid email address.");
  if (!staffInitials) errors.push("Enter your staff initials.");
  if (!dateFrom) errors.push("Choose the first date to avoid.");
  if (!dateTo) errors.push("Choose the final date to avoid.");
  if (dateFrom && dateTo && dateFrom > dateTo) errors.push("The final date must be the same as or after the first date.");
  if (!["Yes", "No"].includes(continuousLeave)) errors.push("Choose whether this is part of continuous annual leave.");

  return { errors, value: { email, staffInitials, dateFrom, dateTo, continuousLeave } };
}

async function handleApi(req, res) {
  const url = new URL(req.url, `http://${req.headers.host}`);
  const settings = getSettings();

  if (req.method === "GET" && url.pathname === "/api/status") {
    return send(res, 200, {
      settings,
      isClosed: deadlineHasPassed(settings),
      responseCount: getResponses(settings).length
    });
  }

  if (req.method === "POST" && url.pathname === "/api/submit") {
    if (deadlineHasPassed(settings)) {
      return send(res, 403, { error: "This request form is now closed." });
    }

    const body = await parseBody(req);
    const { errors, value } = validateSubmission(body);
    if (errors.length) return send(res, 400, { errors });

    const responses = getResponses(settings);
    responses.push({
      timestamp: new Date().toISOString(),
      ...value
    });
    writeResponses(settings, responses);
    buildWorkbook(settings, responses);
    return send(res, 201, { ok: true, responseCount: responses.length });
  }

  if (req.method === "GET" && url.pathname === "/api/admin") {
    if (!requireAdmin(req)) return send(res, 401, { error: "Incorrect admin password." });
    return send(res, 200, {
      settings,
      responseCount: getResponses(settings).length,
      workbookPath: workbookPath(settings),
      workbooks: workbookOptions(settings)
    });
  }

  if (req.method === "POST" && url.pathname === "/api/admin/deadline") {
    if (!requireAdmin(req)) return send(res, 401, { error: "Incorrect admin password." });
    const body = await parseBody(req);
    const deadline = String(body.deadline || "").trim();
    if (!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$/.test(deadline) || Number.isNaN(new Date(deadline).getTime())) {
      return send(res, 400, { error: "Enter a valid deadline date and time." });
    }
    const next = { ...settings, deadline };
    writeJson(SETTINGS_JSON, next);
    return send(res, 200, { settings: next, isClosed: deadlineHasPassed(next) });
  }

  if (req.method === "POST" && url.pathname === "/api/admin/title") {
    if (!requireAdmin(req)) return send(res, 401, { error: "Incorrect admin password." });
    const body = await parseBody(req);
    const requestMonth = Number(body.requestMonth);
    const requestYear = Number(body.requestYear);
    if (!Number.isInteger(requestMonth) || requestMonth < 1 || requestMonth > 12) {
      return send(res, 400, { error: "Choose a valid month." });
    }
    if (!Number.isInteger(requestYear) || requestYear < 2000 || requestYear > 2100) {
      return send(res, 400, { error: "Choose a valid year." });
    }
    const next = {
      ...settings,
      requestMonth,
      requestYear,
      title: buildTitle(requestMonth, requestYear)
    };
    writeJson(SETTINGS_JSON, next);
    ensureWorkbook(next);
    return send(res, 200, {
      settings: next,
      isClosed: deadlineHasPassed(next),
      responseCount: getResponses(next).length,
      workbooks: workbookOptions(next)
    });
  }

  if (req.method === "GET" && url.pathname === "/api/admin/responses.xlsx") {
    if (!requireAdmin(req)) return send(res, 401, { error: "Incorrect admin password." });
    const selectedPeriod = url.searchParams.get("period") || periodKey(settings);
    let selectedSettings = settings;
    let selectedWorkbook = null;
    let selectedFileName = null;

    if (selectedPeriod === "legacy") {
      selectedWorkbook = LEGACY_WORKBOOK;
      selectedFileName = "off-rota-responses-legacy.xlsx";
      if (!fs.existsSync(selectedWorkbook)) return send(res, 404, { error: "That workbook is not available." });
    } else {
      const parsed = settingsFromPeriodKey(selectedPeriod);
      if (!parsed) return send(res, 400, { error: "Choose a valid workbook." });
      selectedSettings = { ...settings, ...parsed };
      selectedWorkbook = ensureWorkbook(selectedSettings);
      selectedFileName = workbookFileName(selectedSettings);
    }

    res.writeHead(200, {
      "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "Content-Disposition": `attachment; filename="${selectedFileName}"`,
      "Cache-Control": "no-store"
    });
    return fs.createReadStream(selectedWorkbook).pipe(res);
  }

  if (req.method === "POST" && url.pathname === "/api/admin/workbook/clear") {
    if (!requireAdmin(req)) return send(res, 401, { error: "Incorrect admin password." });
    const body = await parseBody(req);
    const selectedPeriod = String(body.period || "").trim();
    const parsed = settingsFromPeriodKey(selectedPeriod);
    if (!parsed) return send(res, 400, { error: "Choose a valid workbook to clear." });
    const selectedSettings = { ...settings, ...parsed };
    writeResponses(selectedSettings, []);
    buildWorkbook(selectedSettings, []);
    return send(res, 200, {
      clearedPeriod: selectedPeriod,
      settings,
      responseCount: getResponses(settings).length,
      workbooks: workbookOptions(settings)
    });
  }

  return false;
}

function serveStatic(req, res) {
  const url = new URL(req.url, `http://${req.headers.host}`);
  const safePath = url.pathname === "/" ? "/index.html" : url.pathname;
  const filePath = path.normalize(path.join(PUBLIC_DIR, safePath));
  if (!filePath.startsWith(PUBLIC_DIR)) return send(res, 403, "Forbidden", "text/plain");

  fs.readFile(filePath, (error, content) => {
    if (error) return send(res, 404, "Not found", "text/plain");
    const ext = path.extname(filePath);
    const type = ext === ".css" ? "text/css" : ext === ".js" ? "text/javascript" : "text/html";
    send(res, 200, content, type);
  });
}

ensureWorkbook(getSettings());

const server = http.createServer(async (req, res) => {
  try {
    if (req.url.startsWith("/api/")) {
      const handled = await handleApi(req, res);
      if (handled === false) send(res, 404, { error: "Not found" });
      return;
    }
    serveStatic(req, res);
  } catch (error) {
    send(res, 500, { error: "Something went wrong.", detail: error.message });
  }
});

server.listen(PORT, () => {
  logLine(`Off Rota Requests site running at http://localhost:${PORT}`);
  logLine(`Admin password: ${ADMIN_PASSWORD}`);
});
