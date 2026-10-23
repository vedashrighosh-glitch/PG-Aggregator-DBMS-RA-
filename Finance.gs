// ================================================================
//  RoomAdda — Finance.gs  (v4 — Final merged)
//  - Role-based access: Founder + Finance for all writes
//  - getAccessRole() reads from Interns_Details by OAuth email
//  - deleteFinanceRow(): ONLY Founder + Finance
//  - updateFinanceEntry(): edit existing row by _rowIndex
//  - addFinanceEntry(): validates _callerRole before saving
//  - isFinanceWriteRole(): central write-permission check
//  - getTopPGs(): leaderboard helper
//  NOTE: SS_ID, doGet(), and all URL helpers live in Code.gs
// ================================================================

const SHEET_FINANCE         = "Finance";
const SHEET_FINANCE_SUMMARY = "Finance_Summary";

const F_COL = {
  month:            0, year:             1, sr_no:            2,
  invoice_no:       3, date:             4, name:             5,
  pg_name:          6, sharing_type:     7, mobile:           8,
  monthly_rent:     9, security_deposit: 10, total_amc:       11,
  commission:       12, advance:         13, paid_to_owner:   14
};
const F_TOTAL_COLS = 15;

// ─── ROLE HELPERS ────────────────────────────────────────────────
function isFinanceWriteRole(role) {
  var r = (role || '').toString().trim().toLowerCase();
  return r === 'founder' || r === 'finance';
}

function _ok(data)  { return { status: "success", data: data }; }
function _err(msg)  { return { status: "error",   message: String(msg) }; }

function _getSheet(name) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const s  = ss.getSheetByName(name);
  if (!s) throw new Error('Sheet "' + name + '" not found.');
  return s;
}

function _parseCurrency(val) {
  if (val === null || val === undefined || val === "") return 0;
  if (typeof val === "number") return val;
  const n = parseFloat(String(val).replace(/[₹,\s]/g, "").trim());
  return isNaN(n) ? 0 : n;
}

function _parseDate(raw) {
  if (!raw) return "";
  if (raw instanceof Date) return Utilities.formatDate(raw, Session.getScriptTimeZone(), "dd/MM/yyyy");
  if (typeof raw === "number") {
    return Utilities.formatDate(new Date(Math.round((raw - 25569) * 86400 * 1000)), Session.getScriptTimeZone(), "dd/MM/yyyy");
  }
  const s = String(raw).trim();
  const dmy = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if (dmy) return dmy[1].padStart(2, "0") + "/" + dmy[2].padStart(2, "0") + "/" + dmy[3];
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return iso[3] + "/" + iso[2] + "/" + iso[1];
  return s;
}

// Invoice No is NEVER a date. Numbers stay as numbers, text stays as text.
// _parseDate must never touch this column — it converts 1 → 28/06/1900 etc.
function _parseInvoiceNo(raw) {
  if (raw === null || raw === undefined || raw === "") return "";
  if (raw instanceof Date) return ""; // mis-formatted cell — ignore
  if (typeof raw === "number") return String(Math.round(raw));
  return String(raw).trim();
}

function _parseMobile(raw) {
  if (!raw) return "";
  if (typeof raw === "number") return String(Math.round(raw));
  return String(raw).trim();
}

function _isSeparatorRow(row) {
  const month = String(row[F_COL.month] || "").trim();
  const name  = String(row[F_COL.name]  || "").trim();
  if (month.startsWith("─") || month.startsWith("—") || month.includes("──")) return true;
  if (!month && !name && !String(row[F_COL.sr_no] || "").trim() && !String(row[F_COL.commission] || "").trim()) return true;
  return false;
}

// ─── READ ALL FINANCE ROWS ───────────────────────────────────────
function _readAllFinanceRows() {
  const sheet = _getSheet(SHEET_FINANCE);
  const last  = sheet.getLastRow();
  if (last < 2) return [];

  const data = sheet.getRange(2, 1, last - 1, F_TOTAL_COLS).getValues();
  const rows = [];
  data.forEach(function(row, idx) {
    if (_isSeparatorRow(row)) return;
    const name = String(row[F_COL.name] || "").trim();
    if (!name) return;
    rows.push({
      _rowIndex:        idx + 2,
      month:            String(row[F_COL.month]            || "").trim(),
      year:             String(row[F_COL.year]             || "").trim().replace(".0", ""),
      sr_no:            String(row[F_COL.sr_no]            || "").trim().replace(".0", ""),
      invoice_no:       _parseInvoiceNo(row[F_COL.invoice_no]),
      date:             _parseDate(row[F_COL.date]),
      name:             name,
      pg_name:          String(row[F_COL.pg_name]          || "").trim(),
      sharing_type:     String(row[F_COL.sharing_type]     || "").trim(),
      mobile:           _parseMobile(row[F_COL.mobile]),
      monthly_rent:     _parseCurrency(row[F_COL.monthly_rent]),
      security_deposit: _parseCurrency(row[F_COL.security_deposit]),
      total_amc:        _parseCurrency(row[F_COL.total_amc]),
      commission:       _parseCurrency(row[F_COL.commission]),
      advance:          _parseCurrency(row[F_COL.advance]),
      paid_to_owner:    _parseCurrency(row[F_COL.paid_to_owner])
    });
  });
  return rows;
}

// ─── GET ACCESS ROLE ─────────────────────────────────────────────
// Reads access_role from Interns_Details by Google OAuth email.
// Changing a cell in the sheet immediately affects what the user
// can do on next page load — no code deployment needed.
function getAccessRole() {
  try {
    const email = Session.getActiveUser().getEmail();
    if (!email) return { role: '' };

    const ss    = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName('Interns_Details');
    if (!sheet) return { role: '' };

    const data    = sheet.getDataRange().getValues();
    const headers = data[0].map(function(h) { return String(h).trim().toLowerCase(); });
    const emailCol      = headers.indexOf('email');
    const accessRoleCol = headers.indexOf('access_role');
    const deptCol       = headers.indexOf('department');

    if (emailCol === -1 || accessRoleCol === -1) return { role: '' };

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][emailCol] || '').trim().toLowerCase() === email.toLowerCase()) {
        return {
          role:       String(data[i][accessRoleCol] || '').trim(),
          department: deptCol >= 0 ? String(data[i][deptCol] || '').trim() : ''
        };
      }
    }
    return { role: '' };
  } catch(e) { return { role: '' }; }
}

// ─── GET FINANCE DATA ────────────────────────────────────────────
function getFinanceData(filters) {
  try {
    const all = _readAllFinanceRows();
    filters   = filters || {};
    const name   = (filters.name    || "").trim().toLowerCase();
    const pgName = (filters.pg_name || "").trim().toLowerCase();
    const month  = (filters.month   || "").trim().toLowerCase();
    const year   = (filters.year    || "").trim();

    const filtered = all.filter(function(r) {
      if (name   && !r.name.toLowerCase().includes(name))       return false;
      if (pgName && !r.pg_name.toLowerCase().includes(pgName))  return false;
      if (month  && r.month.toLowerCase() !== month)            return false;
      if (year   && String(r.year) !== String(year))            return false;
      return true;
    });

    let tc = 0, ta = 0, tp = 0, ts = 0, tam = 0;
    filtered.forEach(function(r) { tc += r.commission; ta += r.advance; tp += r.paid_to_owner; ts += r.security_deposit; tam += r.total_amc; });

    return _ok({
      rows:   filtered,
      count:  filtered.length,
      totals: { commission: tc, advance: ta, paid_to_owner: tp, security_deposit: ts, total_amc: tam }
    });
  } catch(e) { return _err(e.message); }
}

// ─── GET FINANCE SUMMARY ─────────────────────────────────────────
function getFinanceSummary() {
  try {
    const all = _readAllFinanceRows();
    const MO  = { "january":1,"february":2,"march":3,"april":4,"may":5,"june":6,"july":7,"august":8,"september":9,"october":10,"november":11,"december":12 };
    const groups = {};
    all.forEach(function(r) {
      const key = (r.month || "Unknown") + "||" + (r.year || "0");
      if (!groups[key]) groups[key] = { month: r.month, year: r.year, label: r.month + " " + r.year, entries: 0, commission: 0, advance: 0, paid_to_owner: 0, security_deposit: 0, total_amc: 0 };
      groups[key].entries++;
      groups[key].commission       += r.commission;
      groups[key].advance          += r.advance;
      groups[key].paid_to_owner    += r.paid_to_owner;
      groups[key].security_deposit += r.security_deposit;
      groups[key].total_amc        += r.total_amc;
    });
    const arr = Object.values(groups).sort(function(a, b) {
      const yd = parseInt(a.year || 0) - parseInt(b.year || 0);
      if (yd !== 0) return yd;
      return (MO[(a.month || "").toLowerCase()] || 13) - (MO[(b.month || "").toLowerCase()] || 13);
    });
    let gc = 0, ga = 0, gp = 0, ge = 0, gs = 0, gam = 0;
    arr.forEach(function(s) { gc += s.commission; ga += s.advance; gp += s.paid_to_owner; ge += s.entries; gs += s.security_deposit; gam += s.total_amc; });
    return _ok({ rows: arr, totals: { entries: ge, commission: gc, advance: ga, paid_to_owner: gp, security_deposit: gs, total_amc: gam } });
  } catch(e) { return _err(e.message); }
}

// ─── GET FILTER OPTIONS ──────────────────────────────────────────
function getFinanceFilterOptions() {
  try {
    const all = _readAllFinanceRows();
    const MO  = { "january":1,"february":2,"march":3,"april":4,"may":5,"june":6,"july":7,"august":8,"september":9,"october":10,"november":11,"december":12 };
    const pg = {}, mo = {}, yr = {};
    all.forEach(function(r) {
      if (r.pg_name) pg[r.pg_name] = 1;
      if (r.month)   mo[r.month]   = MO[r.month.toLowerCase()] || 99;
      if (r.year)    yr[r.year]    = 1;
    });
    return _ok({
      pgNames: Object.keys(pg).sort(),
      months:  Object.keys(mo).sort(function(a, b) { return (mo[a] || 99) - (mo[b] || 99); }),
      years:   Object.keys(yr).sort()
    });
  } catch(e) { return _err(e.message); }
}

// ─── GET DASHBOARD KPIs ──────────────────────────────────────────
function getFinanceKPIs() {
  try {
    const all = _readAllFinanceRows();
    if (!all.length) return _ok({ totalCommission: 0, totalEntries: 0, totalAdvance: 0, totalSecurityDeposit: 0, totalAMC: 0, thisMonthCommission: 0, thisMonthEntries: 0, topPG: "—", topPGCommission: 0, avgCommission: 0 });

    const now      = new Date();
    const MONTHS   = ["January","February","March","April","May","June","July","August","September","October","November","December"];
    const curMonth = MONTHS[now.getMonth()];
    const curYear  = String(now.getFullYear());

    let tc = 0, ta = 0, ts = 0, tam = 0, tmc = 0, tme = 0;
    const pgC = {};
    all.forEach(function(r) {
      tc += r.commission; ta += r.advance; ts += r.security_deposit; tam += r.total_amc;
      if (r.pg_name) pgC[r.pg_name] = (pgC[r.pg_name] || 0) + r.commission;
      if (r.month.toLowerCase() === curMonth.toLowerCase() && String(r.year) === curYear) { tmc += r.commission; tme++; }
    });

    let topPG = "—", topPGC = 0;
    Object.keys(pgC).forEach(function(pg) { if (pgC[pg] > topPGC) { topPGC = pgC[pg]; topPG = pg; } });

    return _ok({
      totalCommission:    tc,
      totalEntries:       all.length,
      totalAdvance:       ta,
      totalSecurityDeposit: ts,
      totalAMC:           tam,
      thisMonthCommission: tmc,
      thisMonthEntries:   tme,
      topPG:              topPG,
      topPGCommission:    topPGC,
      avgCommission:      all.length > 0 ? Math.round(tc / all.length) : 0
    });
  } catch(e) { return _err(e.message); }
}

// ─── ADD FINANCE ENTRY ───────────────────────────────────────────
function addFinanceEntry(payload) {
  try {
    if (!isFinanceWriteRole(payload._callerRole))
      return _err("Only Founder or Finance role can add entries.");
    if (!payload.name || !payload.name.trim()) return _err("Customer name is required.");
    if (!payload.month || !payload.month.trim()) return _err("Month is required.");
    if (!payload.year  || !payload.year.trim())  return _err("Year is required.");

    const sheet = _getSheet(SHEET_FINANCE);
    let srNo = payload.sr_no;
    if (!srNo) {
      const maxSr = _readAllFinanceRows().reduce(function(mx, r) { const n = parseInt(r.sr_no) || 0; return n > mx ? n : mx; }, 0);
      srNo = maxSr + 1;
    }

    const newRow = new Array(F_TOTAL_COLS).fill("");
    newRow[F_COL.month]            = payload.month           || "";
    newRow[F_COL.year]             = payload.year            || "";
    newRow[F_COL.sr_no]            = srNo;
    newRow[F_COL.invoice_no]       = payload.invoice_no      || "";
    newRow[F_COL.date]             = payload.date            || "";
    newRow[F_COL.name]             = payload.name.trim();
    newRow[F_COL.pg_name]          = payload.pg_name         || "";
    newRow[F_COL.sharing_type]     = payload.sharing_type    || "";
    newRow[F_COL.mobile]           = payload.mobile          || "";
    newRow[F_COL.monthly_rent]     = payload.monthly_rent     ? "₹" + Number(payload.monthly_rent).toLocaleString("en-IN")     : "";
    newRow[F_COL.security_deposit] = payload.security_deposit ? "₹" + Number(payload.security_deposit).toLocaleString("en-IN") : "";
    newRow[F_COL.total_amc]        = payload.total_amc        ? "₹" + Number(payload.total_amc).toLocaleString("en-IN")        : "";
    newRow[F_COL.commission]       = payload.commission       ? "₹" + Number(payload.commission).toLocaleString("en-IN")       : "";
    newRow[F_COL.advance]          = payload.advance          ? "₹" + Number(payload.advance).toLocaleString("en-IN")          : "";
    newRow[F_COL.paid_to_owner]    = payload.paid_to_owner    ? "₹" + Number(payload.paid_to_owner).toLocaleString("en-IN")    : "";

    sheet.appendRow(newRow);
    SpreadsheetApp.flush();
    return _ok({ message: "Entry added successfully.", sr_no: srNo });
  } catch(e) { return _err(e.message); }
}

// ─── UPDATE FINANCE ENTRY ────────────────────────────────────────
function updateFinanceEntry(payload) {
  try {
    if (!isFinanceWriteRole(payload._callerRole))
      return _err("Only Founder or Finance role can edit entries.");

    const rowIndex = parseInt(payload._rowIndex);
    if (isNaN(rowIndex) || rowIndex < 2) return _err("Invalid row index.");

    const sheet = _getSheet(SHEET_FINANCE);
    if (rowIndex > sheet.getLastRow()) return _err("Row " + rowIndex + " not found.");

    function sc(col, val, isCurr) {
      if (val === undefined || val === null) return;
      var out = isCurr && val !== "" && val !== "0"
        ? "₹" + Number(val).toLocaleString("en-IN")
        : (isCurr && (val === "" || val === "0") ? "" : val);
      sheet.getRange(rowIndex, col + 1).setValue(out);
    }

    sc(F_COL.month,            payload.month,            false);
    sc(F_COL.year,             payload.year,             false);
    sc(F_COL.invoice_no,       payload.invoice_no,       false);
    sc(F_COL.date,             payload.date,             false);
    sc(F_COL.name,             payload.name,             false);
    sc(F_COL.pg_name,          payload.pg_name,          false);
    sc(F_COL.sharing_type,     payload.sharing_type,     false);
    sc(F_COL.mobile,           payload.mobile,           false);
    sc(F_COL.monthly_rent,     payload.monthly_rent,     true);
    sc(F_COL.security_deposit, payload.security_deposit, true);
    sc(F_COL.total_amc,        payload.total_amc,        true);
    sc(F_COL.commission,       payload.commission,       true);
    sc(F_COL.advance,          payload.advance,          true);
    sc(F_COL.paid_to_owner,    payload.paid_to_owner,    true);

    SpreadsheetApp.flush();
    return _ok({ message: "Entry updated successfully." });
  } catch(e) { return _err(e.message); }
}

// ─── DELETE FINANCE ROW ──────────────────────────────────────────
// ONLY Founder + Finance. Admin intentionally cannot delete Finance data.
function deleteFinanceRow(callerAccessRole, rowIndex) {
  try {
    if (!isFinanceWriteRole(callerAccessRole))
      return _err("Only Founder or Finance role can delete entries. Got: '" + callerAccessRole + "'");

    rowIndex = parseInt(rowIndex);
    if (isNaN(rowIndex) || rowIndex < 2) return _err("Invalid row index: " + rowIndex);

    const sheet = _getSheet(SHEET_FINANCE);
    if (rowIndex > sheet.getLastRow()) return _err("Row " + rowIndex + " not found.");

    sheet.deleteRow(rowIndex);
    SpreadsheetApp.flush();
    return _ok({ message: "Entry deleted successfully.", deletedRow: rowIndex });
  } catch(e) { return _err(e.message); }
}

// ─── GET TOP PGs LEADERBOARD ─────────────────────────────────────
function getTopPGs(limit) {
  try {
    limit = limit || 10;
    const all = _readAllFinanceRows();
    const pgMap = {};
    all.forEach(function(r) {
      if (!r.pg_name) return;
      if (!pgMap[r.pg_name]) pgMap[r.pg_name] = { pg_name: r.pg_name, commission: 0, entries: 0 };
      pgMap[r.pg_name].commission += r.commission;
      pgMap[r.pg_name].entries++;
    });
    return _ok(Object.values(pgMap).sort(function(a, b) { return b.commission - a.commission; }).slice(0, limit));
  } catch(e) { return _err(e.message); }
}