// ================================================================
//  RoomAdda — Leads.gs  (v12 — full lead data for emails + date sync)
//
//  FIXES:
//  1. allotLead fetches full lead row from sheet before sending
//     notification — so budget, area, PG, room etc. are populated
//  2. _writeVisitToSharedCal uses full lead data for structured note
//  3. _updateVisitInSharedCal updates date when visit is rescheduled
//  4. updateVisitDateTime exposed as callable server function
// ================================================================

var LEADS_SS_ID  = DATA_SS_ID;
var LEADS_TAB    = 'leads';
var VISITLOG_TAB = 'VisitLog';

var LEADS_HEADER_ROW = 4;
var LEADS_DATA_START = 5;

// ── SHEET ACCESSORS ──────────────────────────────────────────────
function leadsSheet() {
  var ss = SpreadsheetApp.openById(LEADS_SS_ID);
  var sh = ss.getSheetByName(LEADS_TAB);
  if (!sh) throw new Error('Sheet "' + LEADS_TAB + '" not found.');
  return sh;
}

function visitLogSheet() {
  var ss = SpreadsheetApp.openById(LEADS_SS_ID);
  var sh = ss.getSheetByName(VISITLOG_TAB);
  if (!sh) {
    sh = ss.insertSheet(VISITLOG_TAB);
    sh.appendRow([
      'visit_id','lead_name','lead_phone','allotted_to','allotted_by',
      'allotment_time','visit_date_time','office_note',
      'visit_outcome','outcome_note','outcome_time','calendar_event_id'
    ]);
    sh.setFrozenRows(1);
  }
  return sh;
}

// ── HEADER HELPERS ───────────────────────────────────────────────
function _getRawHeaders(sh) {
  var lastCol = sh.getLastColumn();
  return sh.getRange(LEADS_HEADER_ROW, 1, 1, lastCol).getValues()[0]
    .map(function(h) { return String(h || '').trim(); });
}

function _h2k(h) {
  var map = {
    '#':                    'serial',
    'date':                 'date',
    'name':                 'name',
    'contact no.':          'phone',
    'contact no':           'phone',
    'source':               'source',
    'location preference':  'preferred_area',
    'location\npreference': 'preferred_area',
    'budget (₹/mo)':        'budget_max',
    'budget (rs/mo)':       'budget_max',
    'budget':               'budget_max',
    'handled by':           'attended_by',
    'status':               'lead_status',
    'column 1':             'extra_1',
  };
  var lc = h.toLowerCase();
  return map[lc] || lc.replace(/[^a-z0-9]+/g, '_').replace(/^_|_$/g, '');
}

function _fmt(val) {
  if (val instanceof Date) {
    try { return Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd/MM/yyyy'); }
    catch(e) { return String(val); }
  }
  if (val === null || val === undefined) return '';
  var sv = String(val);
  if (sv.match(/^\d{4}-\d{2}-\d{2}T/)) return sv.split('T')[0];
  return val;
}

function _cleanPhone(v) {
  var sv = String(v || '');
  if (sv.indexOf('E') !== -1 || sv.indexOf('e') !== -1) {
    return String(Math.round(parseFloat(sv)));
  }
  if (sv.indexOf('.') !== -1 && parseFloat(sv) > 1000000000) {
    return String(Math.round(parseFloat(sv)));
  }
  return sv.replace(/\.0$/, '');
}

function generateLeadId() {
  return 'LD-' + new Date().getTime() + '-' + Math.floor(Math.random() * 1000);
}

// ── GET ALL LEADS ────────────────────────────────────────────────
function getAllLeads() {
  try {
    var sh      = leadsSheet();
    var lastRow = sh.getLastRow();
    if (lastRow < LEADS_DATA_START) return { status: 'success', data: [] };

    var rawH    = _getRawHeaders(sh);
    var numCols = rawH.length;
    var numRows = lastRow - LEADS_DATA_START + 1;
    var rows    = sh.getRange(LEADS_DATA_START, 1, numRows, numCols).getValues();

    var data = [];
    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      if (!row.some(function(c) { return c !== '' && c !== null && c !== undefined; })) continue;
      var obj = {};
      rawH.forEach(function(h, idx) { obj[_h2k(h)] = _fmt(row[idx]); });
      obj.phone           = _cleanPhone(obj.phone);
      obj.name            = obj.name            || '';
      obj.lead_status     = obj.lead_status     || 'New';
      obj.budget_min      = obj.budget_min      || '';
      obj.budget_max      = obj.budget_max      || '';
      obj.preferred_pg    = obj.preferred_pg    || '';
      obj.preferred_area  = obj.preferred_area  || '';
      obj.room_preference = obj.room_preference || '';
      obj.food_preference = obj.food_preference || '';
      obj.attended_by     = obj.attended_by     || '';
      obj.comment_1       = obj.comment_1       || '';
      obj.allotted_to     = obj.allotted_to     || '';
      obj.source          = obj.source          || '';
      obj.gender          = obj.gender          || '';
      obj.email           = obj.email           || '';
      data.push(obj);
    }
    return { status: 'success', data: data };
  } catch(e) {
    Logger.log('getAllLeads: ' + e.message);
    return { status: 'error', message: e.message };
  }
}

// ── FIND LEAD ROW (returns row index + full object) ──────────────
function _findLeadRow(sh, rawH, name, phone) {
  var ni = -1, pi = -1;
  rawH.forEach(function(h, i) {
    if (_h2k(h) === 'name')  ni = i;
    if (_h2k(h) === 'phone') pi = i;
  });
  if (ni === -1 || pi === -1) return null;

  var lastRow = sh.getLastRow();
  if (lastRow < LEADS_DATA_START) return null;

  var numRows = lastRow - LEADS_DATA_START + 1;
  var rows    = sh.getRange(LEADS_DATA_START, 1, numRows, rawH.length).getValues();
  var cleanPh = String(phone).replace(/[^0-9]/g, '');

  for (var i = 0; i < rows.length; i++) {
    var nm = String(rows[i][ni] || '').trim().toLowerCase();
    var ph = _cleanPhone(rows[i][pi]).replace(/[^0-9]/g, '');
    if (nm === String(name).trim().toLowerCase() && ph === cleanPh) {
      var obj = {};
      rawH.forEach(function(h, idx) { obj[_h2k(h)] = _fmt(rows[i][idx]); });
      obj.phone = _cleanPhone(obj.phone);
      return { sheetRow: LEADS_DATA_START + i, data: obj, rawRow: rows[i] };
    }
  }
  return null;
}

// ── SEARCH BY PHONE ──────────────────────────────────────────────
function searchLeadByPhone(phone) {
  try {
    var sh      = leadsSheet();
    var lastRow = sh.getLastRow();
    if (lastRow < LEADS_DATA_START) return { status: 'success', data: null };
    var rawH     = _getRawHeaders(sh);
    var phoneIdx = -1;
    rawH.forEach(function(h, i) { if (_h2k(h) === 'phone') phoneIdx = i; });
    if (phoneIdx === -1) return { status: 'error', message: 'Phone column not found' };
    var numRows = lastRow - LEADS_DATA_START + 1;
    var rows    = sh.getRange(LEADS_DATA_START, 1, numRows, rawH.length).getValues();
    var clean   = String(phone).replace(/[^0-9]/g, '');
    for (var i = 0; i < rows.length; i++) {
      var rp = _cleanPhone(rows[i][phoneIdx]).replace(/[^0-9]/g, '');
      if (rp && (rp === clean || rp.endsWith(clean) || clean.endsWith(rp))) {
        var obj = {};
        rawH.forEach(function(h, j) { obj[_h2k(h)] = _fmt(rows[i][j]); });
        obj.phone     = _cleanPhone(obj.phone);
        obj._rowIndex = LEADS_DATA_START + i;
        return { status: 'success', data: obj };
      }
    }
    return { status: 'success', data: null, message: 'No lead found.' };
  } catch(e) { return { status: 'error', message: e.message }; }
}

// ── ADD LEAD ─────────────────────────────────────────────────────
function addLeadRow(payload) {
  try {
    if (!payload.name || !payload.phone)
      return { status: 'error', message: 'Name and Phone are required.' };
    var sh      = leadsSheet();
    var lastRow = sh.getLastRow();
    var rawH    = _getRawHeaders(sh);
    if (lastRow >= LEADS_DATA_START) {
      var ni = -1, pi = -1;
      rawH.forEach(function(h, i) {
        if (_h2k(h) === 'name')  ni = i;
        if (_h2k(h) === 'phone') pi = i;
      });
      if (ni !== -1 && pi !== -1) {
        var existing = sh.getRange(LEADS_DATA_START, 1, lastRow - LEADS_DATA_START + 1, rawH.length).getValues();
        for (var i = 0; i < existing.length; i++) {
          var sn = String(existing[i][ni] || '').trim().toLowerCase();
          var sp = _cleanPhone(existing[i][pi]).replace(/[^0-9]/g, '');
          if (sn === payload.name.trim().toLowerCase() &&
              sp === String(payload.phone).replace(/[^0-9]/g, ''))
            return { status: 'error', message: 'Lead already exists.' };
        }
      }
    }
    var nextSerial = 1;
    if (lastRow >= LEADS_DATA_START) {
      var si = -1;
      rawH.forEach(function(h, i) { if (_h2k(h) === 'serial') si = i; });
      if (si !== -1) {
        var svals = sh.getRange(LEADS_DATA_START, si + 1, lastRow - LEADS_DATA_START + 1, 1).getValues();
        svals.forEach(function(r) {
          var n = parseFloat(r[0]);
          if (!isNaN(n) && n >= nextSerial) nextSerial = n + 1;
        });
      }
    }
    var newRow = rawH.map(function(h) {
      switch(_h2k(h)) {
        case 'serial':         return nextSerial;
        case 'date':           return new Date();
        case 'name':           return payload.name           || '';
        case 'phone':          return payload.phone          || '';
        case 'source':         return payload.source         || '';
        case 'preferred_area': return payload.preferred_area || '';
        case 'budget_max':     return payload.budget_max || payload.budget_min || '';
        case 'attended_by':    return payload.attended_by    || '';
        case 'lead_status':    return payload.lead_status    || 'New';
        default:               return '';
      }
    });
    sh.appendRow(newRow);
    SpreadsheetApp.flush();
    return { status: 'success', message: 'Lead added.' };
  } catch(e) { return { status: 'error', message: e.message }; }
}

// ── UPDATE LEAD ──────────────────────────────────────────────────
function updateLeadRow(payload) {
  try {
    var sh   = leadsSheet();
    var rawH = _getRawHeaders(sh);
    var found = _findLeadRow(sh, rawH, payload.name, payload.phone);
    if (!found) return { status: 'error', message: 'Lead not found.' };

    rawH.forEach(function(h, idx) {
      var k = _h2k(h);
      if (k === 'name' || k === 'serial' || k === 'date') return;
      var val = payload[k];
      if (val !== undefined && val !== null)
        sh.getRange(found.sheetRow, idx + 1).setValue(val);
    });
    SpreadsheetApp.flush();
    return { status: 'success', message: 'Lead updated.' };
  } catch(e) { return { status: 'error', message: e.message }; }
}

// ── DELETE LEAD ──────────────────────────────────────────────────
function deleteLeadRow(payload) {
  try {
    var sh   = leadsSheet();
    var rawH = _getRawHeaders(sh);
    var found = _findLeadRow(sh, rawH, payload.name, payload.phone);
    if (!found) return { status: 'error', message: 'Lead not found.' };
    sh.deleteRow(found.sheetRow);
    SpreadsheetApp.flush();
    return { status: 'success', message: 'Lead deleted.' };
  } catch(e) { return { status: 'error', message: e.message }; }
}

// ── ALLOT ────────────────────────────────────────────────────────
// FIX: reads full lead row from sheet so all fields are populated
// in the notification email (budget, area, PG, room, food pref etc.)
function allotLead(payload) {
  try {
    var sh   = leadsSheet();
    var rawH = _getRawHeaders(sh);

    // Find lead and get FULL row data
    var found = _findLeadRow(sh, rawH, payload.name, payload.phone);
    if (!found) return { status: 'error', message: 'Lead not found.' };

    // Full lead object from sheet — has all columns
    var fullLead = found.data;

    // Update lead status to Visit Scheduled
    var si = -1;
    rawH.forEach(function(h, i) { if (_h2k(h) === 'lead_status') si = i; });
    if (si > -1) sh.getRange(found.sheetRow, si + 1).setValue('Visit Scheduled');
    SpreadsheetApp.flush();

    // Merge allotment-specific fields into fullLead
    fullLead.visit_date_time = payload.visit_date_time || '';
    fullLead.office_note     = payload.office_note     || '';
    // food_preference: prefer what's in sheet, fall back to payload
    if (!fullLead.food_preference && payload.food_preference) {
      fullLead.food_preference = payload.food_preference;
    }

    // Generate visit ID — shared across log, calendar, and all 4 emails
    var visitId = 'V-' + new Date().getTime();

    // Write to SharedCalendar tagged as 'Sales'
    var calEventId = null;
    try {
      calEventId = _writeVisitToSharedCal(fullLead, visitId, payload);
    } catch(ex) { Logger.log('SharedCalendar: ' + ex.message); }

    // Write to VisitLog
    try {
      visitLogSheet().appendRow([
        visitId,
        fullLead.name              || '',
        fullLead.phone             || '',
        payload.allotted_to        || '',
        payload.allotted_by        || '',
        new Date(),
        payload.visit_date_time    || '',
        payload.office_note        || '',
        '',   // visit_outcome
        '',   // outcome_note
        '',   // outcome_time
        calEventId || ''
      ]);
      SpreadsheetApp.flush();
    } catch(ex) { Logger.log('VisitLog: ' + ex.message); }

    // Send all 4 emails with FULL lead data from sheet
    try {
      sendAllotmentNotification(
        fullLead,
        payload.allotted_to  || '',
        payload.agent_phone  || '',
        payload.agent_email  || '',
        payload.office_note  || '',
        visitId
      );
    } catch(ex) { Logger.log('Notification: ' + ex.message); }

    return {
      status:          'success',
      message:         fullLead.name + ' allotted to ' + payload.allotted_to,
      calendarCreated: !!calEventId,
      visitId:         visitId
    };
  } catch(e) {
    return { status: 'error', message: e.message };
  }
}

// ── UPDATE VISIT STATUS ──────────────────────────────────────────
function updateVisitStatus(payload) {
  try {
    var sh   = leadsSheet();
    var rawH = _getRawHeaders(sh);
    var found = _findLeadRow(sh, rawH, payload.name, payload.phone);
    if (!found) return { status: 'error', message: 'Lead not found.' };

    var si = -1;
    rawH.forEach(function(h, i) { if (_h2k(h) === 'lead_status') si = i; });

    var statusMap = {
      done:           'Visited',
      postponed:      'Visit Scheduled',
      not_interested: 'New',
      no_show:        'Contacted'
    };
    if (si > -1) sh.getRange(found.sheetRow, si + 1).setValue(statusMap[payload.visit_status] || 'Contacted');
    SpreadsheetApp.flush();

    // Update VisitLog outcome
    try {
      _updateVisitLogOutcome(payload.name, payload.phone, payload.visit_status, payload.outcome_note || '');
    } catch(ex) { Logger.log('VisitLog update: ' + ex.message); }

    return { status: 'success', message: 'Visit status updated.' };
  } catch(e) { return { status: 'error', message: e.message }; }
}

// ── UPDATE VISIT DATE TIME ───────────────────────────────────────
// Called when a visit is rescheduled — updates SharedCalendar date
function updateVisitDateTime(payload) {
  // payload: { name, phone, visit_date_time }
  try {
    if (!payload.name || !payload.visit_date_time) {
      return { status: 'error', message: 'Name and new date/time required.' };
    }
    _updateVisitInSharedCal(payload.name, payload.phone, payload.visit_date_time);

    // Also update VisitLog
    try { _updateVisitLogDateTime(payload.name, payload.phone, payload.visit_date_time); }
    catch(ex) { Logger.log('VisitLog date update: ' + ex.message); }

    return { status: 'success', message: 'Visit date updated in calendar.' };
  } catch(e) {
    return { status: 'error', message: e.message };
  }
}

// ── UPDATE VISIT LOG OUTCOME ─────────────────────────────────────
function _updateVisitLogOutcome(name, phone, status, outcomeNote) {
  var ss      = SpreadsheetApp.openById(LEADS_SS_ID);
  var sh      = ss.getSheetByName(VISITLOG_TAB);
  if (!sh) return;
  var data    = sh.getDataRange().getValues();
  var headers = data[0].map(function(x){ return String(x).trim().toLowerCase().replace(/\s+/g,'_'); });
  var nameCol    = headers.indexOf('lead_name');
  var phoneCol   = headers.indexOf('lead_phone');
  var outcomeCol = headers.indexOf('visit_outcome');
  var noteCol    = headers.indexOf('outcome_note');
  var timeCol    = headers.indexOf('outcome_time');
  var cleanPh    = String(phone).replace(/[^0-9]/g,'');

  for (var i = 1; i < data.length; i++) {
    var rn = String(data[i][nameCol]  || '').trim().toLowerCase();
    var rp = String(data[i][phoneCol] || '').replace(/[^0-9]/g,'');
    if (rn === String(name).trim().toLowerCase() && rp === cleanPh) {
      if (outcomeCol !== -1) sh.getRange(i+1, outcomeCol+1).setValue(status);
      if (noteCol    !== -1) sh.getRange(i+1, noteCol+1).setValue(outcomeNote || '');
      if (timeCol    !== -1) sh.getRange(i+1, timeCol+1).setValue(new Date().toISOString());
      SpreadsheetApp.flush();
      return;
    }
  }
}

// ── UPDATE VISIT LOG DATE TIME ───────────────────────────────────
function _updateVisitLogDateTime(name, phone, newDateTime) {
  var ss      = SpreadsheetApp.openById(LEADS_SS_ID);
  var sh      = ss.getSheetByName(VISITLOG_TAB);
  if (!sh) return;
  var data    = sh.getDataRange().getValues();
  var headers = data[0].map(function(x){ return String(x).trim().toLowerCase().replace(/\s+/g,'_'); });
  var nameCol  = headers.indexOf('lead_name');
  var phoneCol = headers.indexOf('lead_phone');
  var dtCol    = headers.indexOf('visit_date_time');
  var cleanPh  = String(phone).replace(/[^0-9]/g,'');

  for (var i = 1; i < data.length; i++) {
    var rn = String(data[i][nameCol]  || '').trim().toLowerCase();
    var rp = String(data[i][phoneCol] || '').replace(/[^0-9]/g,'');
    if (rn === String(name).trim().toLowerCase() && rp === cleanPh) {
      if (dtCol !== -1) sh.getRange(i+1, dtCol+1).setValue(newDateTime);
      SpreadsheetApp.flush();
      return;
    }
  }
}

/// ── WRITE VISIT TO SHARED CALENDAR ──────────────────────────────
// Uses SC_HEADERS order: date|title|note|addedBy|addedAt|id|team|deleted
// If a visit already exists for this lead, UPDATE it instead of adding new row
function _writeVisitToSharedCal(fullLead, visitId, payload) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sh = ss.getSheetByName('SharedCalendar');
    if (!sh) {
      sh = ss.insertSheet('SharedCalendar');
      sh.appendRow(['date','title','note','addedBy','addedAt','id','team','deleted']);
      sh.getRange(1,1,1,8).setFontWeight('bold');
      sh.setFrozenRows(1);
    }

    var visitDT = payload.visit_date_time || '';
    var dateStr = '';
    if (visitDT) {
      var dt = new Date(visitDT);
      if (!isNaN(dt.getTime())) {
        dateStr = dt.getFullYear() + '-'
          + (dt.getMonth()+1 < 10 ? '0' : '') + (dt.getMonth()+1) + '-'
          + (dt.getDate()    < 10 ? '0' : '') + dt.getDate();
      }
    }
    if (!dateStr) {
      var today = new Date();
      dateStr = today.getFullYear() + '-'
        + (today.getMonth()+1 < 10 ? '0' : '') + (today.getMonth()+1) + '-'
        + (today.getDate()    < 10 ? '0' : '') + today.getDate();
    }

    var timeLabel = '';
    if (visitDT) {
      try {
        timeLabel = Utilities.formatDate(new Date(visitDT), Session.getScriptTimeZone(), 'hh:mm a');
      } catch(ex) {}
    }

    var title = '🏠 Visit: ' + (fullLead.name || 'Customer')
      + ' → ' + (payload.allotted_to || 'Agent')
      + (timeLabel ? ' @ ' + timeLabel : '');

    // Build structured note from full lead data
    var noteParts = [];
    if (fullLead.name            && String(fullLead.name).trim())            noteParts.push('Customer: '    + String(fullLead.name).trim());
    if (fullLead.phone           && String(fullLead.phone).trim())           noteParts.push('Phone: '       + String(fullLead.phone).trim());
    if (fullLead.preferred_area  && String(fullLead.preferred_area).trim())  noteParts.push('Area: '        + String(fullLead.preferred_area).trim());
    if (fullLead.preferred_pg    && String(fullLead.preferred_pg).trim())    noteParts.push('PG: '          + String(fullLead.preferred_pg).trim());
    if (fullLead.room_preference && String(fullLead.room_preference).trim()) noteParts.push('Room: '        + String(fullLead.room_preference).trim());
    if (fullLead.budget_max      && String(fullLead.budget_max).trim())      noteParts.push('Budget: ₹'     + String(fullLead.budget_max).trim());
    if (fullLead.food_preference && String(fullLead.food_preference).trim()) noteParts.push('Food: '        + String(fullLead.food_preference).trim());
    if (payload.office_note      && String(payload.office_note).trim())      noteParts.push('Note: '        + String(payload.office_note).trim());
    if (payload.allotted_by      && String(payload.allotted_by).trim())      noteParts.push('Allotted by: ' + String(payload.allotted_by).trim());
    var note = noteParts.join(' | ');

    var addedAt = '';
    try { addedAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm'); }
    catch(ex) {}

    // Check if a visit for this lead already exists — UPDATE instead of INSERT
    var data      = sh.getDataRange().getValues();
    var headers   = data[0].map(function(h){ return String(h).trim().toLowerCase(); });
    var ii        = headers.indexOf('id');
    var dii       = headers.indexOf('date');
    var tii       = headers.indexOf('title');
    var nii       = headers.indexOf('note');
    var biIdx     = headers.indexOf('addedby');
    var atiIdx    = headers.indexOf('addedat');
    var dliIdx    = headers.indexOf('deleted');
    var cleanName = String(fullLead.name || '').trim().toLowerCase();

    for (var i = 1; i < data.length; i++) {
      var rowId    = String(data[i][ii]     || '').trim();
      var rowTitle = String(data[i][tii]    || '').trim();
      var rowDel   = String(dliIdx !== -1 ? data[i][dliIdx] : '').toUpperCase();
      if (rowDel === 'TRUE' || rowDel === 'YES') continue;

      // Match existing visit for same lead
      if ((rowId.indexOf('VISIT-') === 0 || rowId.indexOf('V-') === 0) &&
          rowTitle.toLowerCase().indexOf(cleanName) !== -1) {
        // UPDATE existing row
        if (dii  !== -1) sh.getRange(i+1, dii+1).setValue(dateStr);
        if (tii  !== -1) sh.getRange(i+1, tii+1).setValue(title);
        if (nii  !== -1) sh.getRange(i+1, nii+1).setValue(note);
        if (biIdx !== -1) sh.getRange(i+1, biIdx+1).setValue('System (Re-allotment)');
        if (atiIdx !== -1) sh.getRange(i+1, atiIdx+1).setValue(addedAt);
        SpreadsheetApp.flush();
        Logger.log('SharedCalendar visit UPDATED for: ' + fullLead.name + ' → ' + dateStr);
        return rowId; // return existing ID
      }
    }

    // No existing visit found — INSERT new row
    // Order: date|title|note|addedBy|addedAt|id|team|deleted
    var id = visitId || ('V-' + new Date().getTime());
    sh.appendRow([
      dateStr,
      title,
      note,
      'System (Allotment)',
      addedAt,
      id,
      'Sales',
      'FALSE'
    ]);
    SpreadsheetApp.flush();
    Logger.log('SharedCalendar visit INSERTED: ' + id + ' on ' + dateStr + ' for ' + fullLead.name);
    return id;
  } catch(e) {
    Logger.log('_writeVisitToSharedCal: ' + e.message);
    return null;
  }
}

// ── UPDATE VISIT DATE IN SHARED CALENDAR ────────────────────────
function _updateVisitInSharedCal(leadName, leadPhone, newVisitDateTime) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sh = ss.getSheetByName('SharedCalendar');
    if (!sh) return;

    var data    = sh.getDataRange().getValues();
    var headers = data[0].map(function(h){ return String(h).trim().toLowerCase(); });
    var ii      = headers.indexOf('id');
    var di      = headers.indexOf('date');
    var ti      = headers.indexOf('title');
    var dli     = headers.indexOf('deleted');

    var dt = new Date(newVisitDateTime);
    var newDateStr = '';
    if (!isNaN(dt.getTime())) {
      newDateStr = dt.getFullYear() + '-'
        + (dt.getMonth()+1 < 10 ? '0' : '') + (dt.getMonth()+1) + '-'
        + (dt.getDate()    < 10 ? '0' : '') + dt.getDate();
    }

    var timeLabel = '';
    try { timeLabel = Utilities.formatDate(dt, Session.getScriptTimeZone(), 'hh:mm a'); }
    catch(ex) {}

    var cleanName = String(leadName || '').trim().toLowerCase();

    for (var i = 1; i < data.length; i++) {
      var rowId    = String(data[i][ii]  || '').trim();
      var rowTitle = String(data[i][ti]  || '').trim();
      var deleted  = dli !== -1 ? String(data[i][dli] || '').toUpperCase() : '';
      if (deleted === 'TRUE' || deleted === 'YES') continue;

      if ((rowId.indexOf('VISIT-') === 0 || rowId.indexOf('V-') === 0) &&
           rowTitle.toLowerCase().indexOf(cleanName) !== -1) {
        if (newDateStr && di !== -1) sh.getRange(i+1, di+1).setValue(newDateStr);
        if (timeLabel  && ti !== -1) {
          var newTitle = rowTitle.replace(/ @ \d{1,2}:\d{2} [AP]M/i, ' @ ' + timeLabel);
          if (newTitle === rowTitle) newTitle = rowTitle.replace(/\s*$/, '') + ' @ ' + timeLabel;
          sh.getRange(i+1, ti+1).setValue(newTitle);
        }
        SpreadsheetApp.flush();
        Logger.log('SharedCalendar date updated for: ' + leadName + ' → ' + newDateStr);
        return;
      }
    }
  } catch(e) {
    Logger.log('_updateVisitInSharedCal: ' + e.message);
  }
}
// ── TEST ─────────────────────────────────────────────────────────
function testLeadsFinal() {
  var r = getAllLeads();
  Logger.log('Status: ' + r.status);
  if (r.status === 'error') { Logger.log('ERROR: ' + r.message); return; }
  Logger.log('Total leads: ' + r.data.length);
  if (r.data.length > 0) {
    var f = r.data[0];
    Logger.log('First: name=' + f.name + ', phone=' + f.phone + ', status=' + f.lead_status);
    Logger.log('  area=' + f.preferred_area + ', budget=' + f.budget_max + ', room=' + f.room_preference);
  }
}