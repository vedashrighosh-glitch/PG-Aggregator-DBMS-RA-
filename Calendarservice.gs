// ================================================================
//  RoomAdda — CalendarService.gs  (v4 — team-scoped notes)
//
//  CHANGES from v3:
//  1. CalendarNotes sheet now has a 'team' column
//  2. getCalendarNote(date, team) filters by date + team
//  3. saveCalendarNote(date, note, addedBy, team) stores team
//  4. _ensureCalNotesSheet() helper creates sheet with correct headers
//
//  CalendarNotes sheet column order:
//  date | team | note | addedBy | updatedAt
//
//  SharedCalendar sheet column order (unchanged):
//  date | title | note | addedBy | addedAt | id | team | deleted
// ================================================================

var SHEET_SHARED_CAL = 'SharedCalendar';
var SHEET_CAL_NOTES  = 'CalendarNotes';

// Exact column order for the SharedCalendar sheet
var SC_HEADERS = ['date','title','note','addedBy','addedAt','id','team','deleted'];

// Exact column order for CalendarNotes sheet
var CN_HEADERS = ['date','team','note','addedBy','updatedAt'];

// ── ENSURE SharedCalendar SHEET EXISTS ──────────────────────────
function _ensureSharedCalSheet() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sh = ss.getSheetByName(SHEET_SHARED_CAL);
  if (!sh) {
    sh = ss.insertSheet(SHEET_SHARED_CAL);
    sh.appendRow(SC_HEADERS);
    sh.getRange(1, 1, 1, SC_HEADERS.length).setFontWeight('bold');
    sh.setFrozenRows(1);
    Logger.log('SharedCalendar sheet created with headers: ' + SC_HEADERS.join(', '));
  }
  return sh;
}

// ── ENSURE CalendarNotes SHEET EXISTS ───────────────────────────
function _ensureCalNotesSheet() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sh = ss.getSheetByName(SHEET_CAL_NOTES);
  if (!sh) {
    sh = ss.insertSheet(SHEET_CAL_NOTES);
    sh.appendRow(CN_HEADERS);
    sh.getRange(1, 1, 1, CN_HEADERS.length).setFontWeight('bold');
    sh.setFrozenRows(1);
    Logger.log('CalendarNotes sheet created with headers: ' + CN_HEADERS.join(', '));
  }
  return sh;
}

// ── GET HEADER INDICES (SharedCalendar) ─────────────────────────
function _getSharedCalHeaders(sh) {
  var data = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var h    = data.map(function(x){ return String(x).trim().toLowerCase(); });
  return {
    di:  h.indexOf('date'),
    ti:  h.indexOf('title'),
    ni:  h.indexOf('note'),
    bi:  h.indexOf('addedby'),
    ati: h.indexOf('addedat'),
    ii:  h.indexOf('id'),
    tmi: h.indexOf('team'),
    dli: h.indexOf('deleted')
  };
}

// ── GET HEADER INDICES (CalendarNotes) ──────────────────────────
function _getCalNotesHeaders(sh) {
  var data = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var h    = data.map(function(x){ return String(x).trim().toLowerCase(); });
  return {
    di:  h.indexOf('date'),
    tmi: h.indexOf('team'),
    ni:  h.indexOf('note'),
    bi:  h.indexOf('addedby'),
    ui:  h.indexOf('updatedat')
  };
}

// ── GET SHARED CALENDAR EVENTS ──────────────────────────────────
// team = 'overall' → only rows tagged 'overall'
// team = 'Sales'   → rows tagged 'Sales' + rows tagged 'overall'
function getSharedCalendarEvents(team) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sh = ss.getSheetByName(SHEET_SHARED_CAL);
    if (!sh || sh.getLastRow() < 2) return { success: true, events: [] };

    var data = sh.getDataRange().getValues();
    var idx  = _getSharedCalHeaders(sh);

    var events = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];

      // Skip deleted
      var deleted = idx.dli !== -1 ? String(row[idx.dli] || '').trim().toUpperCase() : '';
      if (deleted === 'TRUE' || deleted === 'YES') continue;

      // Skip empty rows
      var hasDate  = idx.di  !== -1 && row[idx.di];
      var hasTitle = idx.ti  !== -1 && row[idx.ti];
      if (!hasDate && !hasTitle) continue;

      var rowTeam = idx.tmi !== -1 ? String(row[idx.tmi] || '').trim() : 'overall';
      if (!rowTeam) rowTeam = 'overall';

      // Overall tab → only 'overall' events
      // Team tab    → that team's events + overall events
      if (team === 'overall') {
        if (rowTeam !== 'overall') continue;
      } else {
        if (rowTeam !== team && rowTeam !== 'overall') continue;
      }

      // Parse date
      var rawDate = idx.di !== -1 ? row[idx.di] : '';
      var dateStr = '';
      if (rawDate instanceof Date) {
        try {
          dateStr = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        } catch(ex) {
          dateStr = String(rawDate).trim().slice(0, 10);
        }
      } else {
        dateStr = String(rawDate || '').trim().slice(0, 10);
      }
      if (!dateStr || dateStr.length < 10) continue;

      // Parse addedAt
      var addedAtRaw = idx.ati !== -1 ? row[idx.ati] : '';
      var addedAt    = '';
      if (addedAtRaw instanceof Date) {
        try { addedAt = Utilities.formatDate(addedAtRaw, Session.getScriptTimeZone(), 'HH:mm'); }
        catch(ex) { addedAt = String(addedAtRaw); }
      } else {
        addedAt = String(addedAtRaw || '').trim();
      }

      events.push({
        id:      idx.ii  !== -1 ? String(row[idx.ii]  || '').trim() : '',
        date:    dateStr,
        title:   idx.ti  !== -1 ? String(row[idx.ti]  || '').trim() : '',
        note:    idx.ni  !== -1 ? String(row[idx.ni]  || '').trim() : '',
        team:    rowTeam,
        addedBy: idx.bi  !== -1 ? String(row[idx.bi]  || '').trim() : '',
        addedAt: addedAt
      });
    }

    Logger.log('getSharedCalendarEvents(' + team + '): ' + events.length + ' events');
    return { success: true, events: events };
  } catch(e) {
    Logger.log('getSharedCalendarEvents error: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ── ADD EVENT ───────────────────────────────────────────────────
function addSharedCalendarEvent(payload) {
  try {
    var sh  = _ensureSharedCalSheet();
    var id  = 'E-' + new Date().getTime();
    var now = new Date();
    var addedAt = '';
    try { addedAt = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm'); }
    catch(ex) { addedAt = ''; }

    sh.appendRow([
      payload.date    || '',          // date
      payload.title   || '',          // title
      payload.note    || '',          // note
      payload.addedBy || '',          // addedBy
      addedAt,                        // addedAt
      id,                             // id
      payload.team    || 'overall',   // team
      'FALSE'                         // deleted
    ]);
    SpreadsheetApp.flush();
    Logger.log('addSharedCalendarEvent: ' + id + ' team=' + (payload.team || 'overall'));
    return { success: true, id: id };
  } catch(e) {
    Logger.log('addSharedCalendarEvent error: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ── UPDATE EVENT ─────────────────────────────────────────────────
function updateSharedCalendarEvent(payload) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sh = ss.getSheetByName(SHEET_SHARED_CAL);
    if (!sh) return { success: false, error: 'SharedCalendar sheet not found' };

    var data = sh.getDataRange().getValues();
    var idx  = _getSharedCalHeaders(sh);

    for (var i = 1; i < data.length; i++) {
      if (idx.ii === -1) continue;
      if (String(data[i][idx.ii] || '').trim() !== String(payload.id || '').trim()) continue;

      if (payload.date  !== undefined && idx.di  !== -1) sh.getRange(i+1, idx.di+1).setValue(payload.date);
      if (payload.title !== undefined && idx.ti  !== -1) sh.getRange(i+1, idx.ti+1).setValue(payload.title);
      if (payload.note  !== undefined && idx.ni  !== -1) sh.getRange(i+1, idx.ni+1).setValue(payload.note);
      if (payload.team  !== undefined && idx.tmi !== -1) sh.getRange(i+1, idx.tmi+1).setValue(payload.team);
      SpreadsheetApp.flush();
      return { success: true };
    }
    return { success: false, error: 'Event not found' };
  } catch(e) {
    Logger.log('updateSharedCalendarEvent error: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ── DELETE EVENT (soft delete) ───────────────────────────────────
function deleteSharedCalendarEvent(id) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sh = ss.getSheetByName(SHEET_SHARED_CAL);
    if (!sh) return { success: false, error: 'SharedCalendar sheet not found' };

    var data = sh.getDataRange().getValues();
    var idx  = _getSharedCalHeaders(sh);

    for (var i = 1; i < data.length; i++) {
      if (idx.ii === -1) continue;
      if (String(data[i][idx.ii] || '').trim() !== String(id || '').trim()) continue;
      if (idx.dli !== -1) sh.getRange(i+1, idx.dli+1).setValue('TRUE');
      SpreadsheetApp.flush();
      return { success: true };
    }
    return { success: false, error: 'Event not found' };
  } catch(e) {
    Logger.log('deleteSharedCalendarEvent error: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ================================================================
//  CALENDAR NOTES  (team-scoped)
//
//  CalendarNotes sheet columns: date | team | note | addedBy | updatedAt
//
//  Scoping rules (mirrors event scoping):
//  • Saved in Overall tab  → team = 'overall'  → visible to ALL users
//  • Saved in Sales tab    → team = 'Sales'    → visible only to Sales
//    (founders/admins see all teams' notes when on that tab)
//
//  getCalendarNote(date, team):
//    team = 'overall' → fetch note where team = 'overall'
//    team = 'Sales'   → fetch note where team = 'Sales'
//    (caller passes the active tab's team value)
// ================================================================

// ── GET CALENDAR NOTE ────────────────────────────────────────────
function getCalendarNote(date, team) {
  try {
    var targetTeam = String(team || 'overall').trim();

    var ss = SpreadsheetApp.openById(SS_ID);
    var sh = ss.getSheetByName(SHEET_CAL_NOTES);
    if (!sh || sh.getLastRow() < 2) return { success: true, note: '' };

    var data = sh.getDataRange().getValues();
    var idx  = _getCalNotesHeaders(sh);

    for (var i = 1; i < data.length; i++) {
      var rowDate = data[i][idx.di] instanceof Date
        ? Utilities.formatDate(data[i][idx.di], Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : String(data[i][idx.di] || '').trim().slice(0, 10);

      var rowTeam = idx.tmi !== -1 ? String(data[i][idx.tmi] || '').trim() : 'overall';
      if (!rowTeam) rowTeam = 'overall';

      if (rowDate === date && rowTeam === targetTeam) {
        return {
          success:   true,
          note:      idx.ni !== -1 ? String(data[i][idx.ni] || '').trim() : '',
          addedBy:   idx.bi !== -1 ? String(data[i][idx.bi] || '').trim() : '',
          updatedAt: idx.ui !== -1 ? String(data[i][idx.ui] || '').trim() : ''
        };
      }
    }
    return { success: true, note: '' };
  } catch(e) {
    Logger.log('getCalendarNote error: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ── SAVE CALENDAR NOTE ───────────────────────────────────────────
// team = 'overall' → stored as 'overall' (company-wide)
// team = 'Sales'   → stored as 'Sales'   (Sales team only)
function saveCalendarNote(date, note, addedBy, team) {
  try {
    var targetTeam = String(team || 'overall').trim();
    var sh         = _ensureCalNotesSheet();
    var data       = sh.getDataRange().getValues();
    var idx        = _getCalNotesHeaders(sh);
    var now        = new Date().toISOString();

    // Look for existing row matching date + team
    for (var i = 1; i < data.length; i++) {
      var rowDate = data[i][idx.di] instanceof Date
        ? Utilities.formatDate(data[i][idx.di], Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : String(data[i][idx.di] || '').trim().slice(0, 10);

      var rowTeam = idx.tmi !== -1 ? String(data[i][idx.tmi] || '').trim() : 'overall';
      if (!rowTeam) rowTeam = 'overall';

      if (rowDate === date && rowTeam === targetTeam) {
        // Update existing row
        if (idx.ni  !== -1) sh.getRange(i+1, idx.ni+1).setValue(note || '');
        if (idx.bi  !== -1) sh.getRange(i+1, idx.bi+1).setValue(addedBy || '');
        if (idx.ui  !== -1) sh.getRange(i+1, idx.ui+1).setValue(now);
        SpreadsheetApp.flush();
        Logger.log('saveCalendarNote: updated date=' + date + ' team=' + targetTeam);
        return { success: true };
      }
    }

    // Not found — append new row in CN_HEADERS order: date|team|note|addedBy|updatedAt
    sh.appendRow([date, targetTeam, note || '', addedBy || '', now]);
    SpreadsheetApp.flush();
    Logger.log('saveCalendarNote: added date=' + date + ' team=' + targetTeam);
    return { success: true };
  } catch(e) {
    Logger.log('saveCalendarNote error: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ── CALENDAR ACCESS LIST ─────────────────────────────────────────
function getCalendarAccessList(accessRole, department) {
  var r         = String(accessRole  || '').trim().toLowerCase();
  var d         = String(department  || '').trim().toLowerCase();
  var effective = r || d;

  var tabs    = [{ id: 'overall', label: 'Overall' }];
  var teamMap = {
    sales:   { id: 'Sales',   label: 'Sales'   },
    hr:      { id: 'HR',      label: 'HR'       },
    tech:    { id: 'Tech',    label: 'Tech'     },
    finance: { id: 'Finance', label: 'Finance'  }
  };

  if (effective === 'founder' || effective === 'admin') {
    Object.keys(teamMap).forEach(function(k){ tabs.push(teamMap[k]); });
  } else if (teamMap[effective]) {
    tabs.push(teamMap[effective]);
  }

  return { success: true, tabs: tabs };
}

// ── DEBUG HELPER ─────────────────────────────────────────────────
function debugSharedCal() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sh = ss.getSheetByName(SHEET_SHARED_CAL);
  if (!sh) { Logger.log('SharedCalendar sheet not found'); return; }
  var data = sh.getDataRange().getValues();
  Logger.log('Headers: ' + JSON.stringify(data[0]));
  Logger.log('Row count: ' + (data.length - 1));
  for (var i = 1; i < Math.min(data.length, 6); i++) {
    Logger.log('Row ' + i + ': ' + JSON.stringify(data[i]));
  }
}

function debugCalNotes() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sh = ss.getSheetByName(SHEET_CAL_NOTES);
  if (!sh) { Logger.log('CalendarNotes sheet not found'); return; }
  var data = sh.getDataRange().getValues();
  Logger.log('Headers: ' + JSON.stringify(data[0]));
  Logger.log('Row count: ' + (data.length - 1));
  for (var i = 1; i < Math.min(data.length, 6); i++) {
    Logger.log('Row ' + i + ': ' + JSON.stringify(data[i]));
  }
}