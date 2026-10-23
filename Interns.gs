// ================================================================
//  RoomAdda — Interns.gs  (v6 — Team Calendars)
//  CHANGES FROM v5:
//  - SharedCalendar sheet now has a 'team' column (col 7)
//    '' / null  = Overall (visible to everyone)
//    'Sales','HR','Tech','Finance','Marketing','Admin','Founder'
//               = team-scoped (visible to that team + Founder)
//  - getSharedCalendarEvents(team)  — team-filtered fetch
//  - addSharedCalendarEvent(payload) — payload.team field
//  - updateSharedCalendarEvent(payload) — preserves team field
//  - deleteSharedCalendarEvent(eventId) — unchanged
//  - getCalendarAccessList(accessRole, department)
//    → returns which calendar tabs the user may see
//  All previous logic (login, attendance, directory, etc.) unchanged
// ================================================================

var SHEET_INTERNS    = 'Interns_Details';
var SHEET_ATTEND     = 'Attendance';
var SHEET_SHARED_CAL = 'SharedCalendar';

var ATTEND_SKIP_VALUES    = ['Sunday','Holiday','NA',''];
var ATTEND_PRESENT_VALUES = ['Present','WFH'];
var ATTEND_ABSENT_VALUES  = ['Absent'];
var ATTEND_WFH_VALUES     = ['WFH'];
var EDIT_ROLES            = ['HR','Admin','Founder','Technical Director'];
var ADMIN_VIEW_ROLES      = ['hr','admin','founder','technical director'];

// Teams that get their own calendar tab
var ALL_TEAMS = ['Sales','HR','Tech','Finance','Marketing','Admin'];

function _internsSS() { return SpreadsheetApp.openById(SS_ID); }

// ── ROLE RESOLUTION ─────────────────────────────────────────────
function getAccessRole() {
  try {
    var email = Session.getActiveUser().getEmail();
    if (!email) return { role: '', department: '' };
    var sheet = _internsSS().getSheetByName(SHEET_INTERNS);
    if (!sheet) return { role: '', department: '' };
    var data    = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h){ return String(h).trim().toLowerCase(); });
    var emailCol      = headers.indexOf('email');
    var accessRoleCol = headers.indexOf('access_role');
    var deptCol       = headers.indexOf('department');
    if (emailCol === -1 || accessRoleCol === -1) return { role: '', department: '' };
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][emailCol]||'').trim().toLowerCase() === email.toLowerCase()) {
        return {
          role:       String(data[i][accessRoleCol]||'').trim(),
          department: deptCol >= 0 ? String(data[i][deptCol]||'').trim() : ''
        };
      }
    }
    return { role: '', department: '' };
  } catch(e) { return { role: '', department: '' }; }
}

// ── CALENDAR ACCESS LIST ─────────────────────────────────────────
// Returns the ordered list of calendar tabs this user may see.
// Format: [{ id: 'overall', label: 'Overall' }, { id: 'Sales', label: 'Sales' }, ...]
// Rules:
//   Founder / Admin → Overall + ALL teams
//   HR              → Overall + HR team
//   Technical Dir   → Overall + Tech team
//   Sales dept      → Overall + Sales team
//   Finance dept    → Overall + Finance team
//   Marketing dept  → Overall + Marketing team
//   Tech dept       → Overall + Tech team   (Tech users now have a calendar!)
//   Everyone else   → Overall only
function getCalendarAccessList(accessRole, department) {
  var role = (accessRole  || '').trim().toLowerCase();
  var dept = (department  || '').trim().toLowerCase();

  var tabs = [{ id: 'overall', label: 'Overall' }];

  var isFounder = role === 'founder' || role === 'admin';

  if (isFounder) {
    // Founders/Admins see every team calendar
    ALL_TEAMS.forEach(function(t) {
      tabs.push({ id: t, label: t });
    });
  } else {
    // Map role/dept → team calendar they belong to
    var teamMap = {
      'hr':                 'HR',
      'technical director': 'Tech',
      'finance':            'Finance',
      'sales':              'Sales',
      'marketing':          'Marketing',
      'admin':              'Admin',
      'tech':               'Tech'
    };
    // Check access_role first, then department
    var team = teamMap[role] || teamMap[dept] || null;
    if (team) tabs.push({ id: team, label: team });
  }

  return { success: true, tabs: tabs };
}

// ── SHARED CALENDAR: ENSURE SHEET ───────────────────────────────
// Columns: date | title | note | added_by | added_at | id | team
function _ensureSharedCalSheet() {
  var ss = _internsSS();
  var sh = ss.getSheetByName(SHEET_SHARED_CAL);
  if (!sh) {
    sh = ss.insertSheet(SHEET_SHARED_CAL);
    sh.appendRow(['date','title','note','added_by','added_at','id','team']);
    sh.setFrozenRows(1);
  } else {
    // Add 'team' column if it doesn't exist yet (migration for existing sheets)
    var headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    var hasTeam = headers.some(function(h){ return String(h).trim().toLowerCase() === 'team'; });
    if (!hasTeam) {
      var nextCol = sh.getLastColumn() + 1;
      sh.getRange(1, nextCol).setValue('team');
    }
  }
  return sh;
}

// ── GET EVENTS ───────────────────────────────────────────────────
// teamFilter:
//   'overall'          → only rows where team is blank / 'overall'
//   'Sales' etc        → only rows where team === that value
//   null / undefined   → all rows (admin use)
function getSharedCalendarEvents(teamFilter) {
  try {
    var sh   = _ensureSharedCalSheet();
    var last = sh.getLastRow();
    if (last < 2) return { success: true, events: [] };

    var rows = sh.getRange(2, 1, last - 1, 7).getValues();
    var tz   = Session.getScriptTimeZone();

    var events = rows
      .filter(function(r) { return r[0] || r[1]; })
      .filter(function(r) {
        if (!teamFilter) return true; // no filter → all
        var rowTeam = String(r[6] || '').trim().toLowerCase();
        var filter  = teamFilter.toLowerCase();
        if (filter === 'overall') {
          return rowTeam === '' || rowTeam === 'overall';
        }
        return rowTeam === filter;
      })
      .map(function(r) {
        var ds = r[0] instanceof Date
          ? Utilities.formatDate(r[0], tz, 'yyyy-MM-dd')
          : String(r[0] || '').split('T')[0].trim();
        var as = r[4] instanceof Date
          ? Utilities.formatDate(r[4], tz, 'dd MMM yyyy, hh:mm a')
          : String(r[4] || '');
        return {
          date:    ds,
          title:   String(r[1] || '').trim(),
          note:    String(r[2] || '').trim(),
          addedBy: String(r[3] || '').trim(),
          addedAt: as,
          id:      String(r[5] || '').trim(),
          team:    String(r[6] || '').trim()
        };
      });

    return { success: true, events: events };
  } catch(e) { return { success: false, error: e.message }; }
}

// ── ADD EVENT ────────────────────────────────────────────────────
// payload.team: '' = Overall, 'Sales' = Sales team only, etc.
function addSharedCalendarEvent(payload) {
  try {
    if (!payload.date || !payload.title)
      return { success: false, error: 'Date and title required.' };
    var sh   = _ensureSharedCalSheet();
    var id   = 'CAL-' + new Date().getTime();
    var team = (payload.team || '').trim(); // '' means Overall
    sh.appendRow([payload.date, payload.title, payload.note || '', payload.addedBy || '', new Date(), id, team]);
    SpreadsheetApp.flush();
    return { success: true, id: id };
  } catch(e) { return { success: false, error: e.message }; }
}

// ── UPDATE EVENT ─────────────────────────────────────────────────
function updateSharedCalendarEvent(payload) {
  try {
    var sh   = _ensureSharedCalSheet();
    var last = sh.getLastRow();
    if (last < 2) return { success: false, error: 'No events.' };
    var data = sh.getRange(2, 1, last - 1, 7).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][5]).trim() === payload.id) {
        var row = i + 2;
        sh.getRange(row, 1).setValue(payload.date);
        sh.getRange(row, 2).setValue(payload.title);
        sh.getRange(row, 3).setValue(payload.note || '');
        sh.getRange(row, 4).setValue(payload.addedBy || data[i][3]);
        sh.getRange(row, 5).setValue(new Date());
        // Preserve team — don't change it on edit unless explicitly passed
        if (payload.team !== undefined) sh.getRange(row, 7).setValue(payload.team || '');
        SpreadsheetApp.flush();
        return { success: true };
      }
    }
    return { success: false, error: 'Event not found.' };
  } catch(e) { return { success: false, error: e.message }; }
}

// ── DELETE EVENT ─────────────────────────────────────────────────
function deleteSharedCalendarEvent(eventId) {
  try {
    var sh   = _ensureSharedCalSheet();
    var last = sh.getLastRow();
    if (last < 2) return { success: false, error: 'No events.' };
    var data = sh.getRange(2, 1, last - 1, 7).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][5]).trim() === eventId) {
        sh.deleteRow(i + 2);
        SpreadsheetApp.flush();
        return { success: true };
      }
    }
    return { success: false, error: 'Event not found.' };
  } catch(e) { return { success: false, error: e.message }; }
}

// ════════════════════════════════════════════════════════════════
//  Everything below is UNCHANGED from v5
// ════════════════════════════════════════════════════════════════

// ── AUTH: LOGIN ─────────────────────────────────────────────────
function loginUser(nameInput, passwordInput) {
  try {
    var ss    = _internsSS();
    var sheet = ss.getSheetByName(SHEET_INTERNS);
    var data  = sheet.getDataRange().getValues();
    var h     = data[0].map(function(x){ return String(x).trim().toLowerCase(); });
 
    var nameCol=h.indexOf('name'), passCol=h.indexOf('password');
    var roleCol=h.indexOf('role'), deptCol=h.indexOf('department');
    var emailCol=h.indexOf('email'), phoneCol=h.indexOf('phone');
    var typeCol=h.indexOf('type'), accessCol=h.indexOf('access_role');
    var activeCol=h.indexOf('is_active'), joinCol=h.indexOf('join_date');
    var diaryCol=h.indexOf('work diary url');
 
    for (var i = 1; i < data.length; i++) {
      var row     = data[i];
      var rowName = String(row[nameCol]||'').trim();
      var rowPass = String(row[passCol]||'').trim();
 
      if (rowName.toLowerCase() === nameInput.trim().toLowerCase() &&
          rowPass.replace(/\s+/g,'').toLowerCase() === passwordInput.trim().replace(/\s+/g,'').toLowerCase()) {
 
        if (String(row[activeCol]||'').trim().toUpperCase() === 'NO')
          return { success:false, error:'Account is inactive. Contact HR.' };
 
        var joinDate = '';
        var jr = row[joinCol];
        var MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
        if (jr instanceof Date) {
          joinDate = jr.getDate()+' '+MONTHS[jr.getMonth()]+' '+jr.getFullYear();
        } else if (jr && typeof jr === 'number') {
          var d = new Date(Math.round((jr-25569)*86400*1000));
          joinDate = d.getDate()+' '+MONTHS[d.getMonth()]+' '+d.getFullYear();
        } else if (jr) {
          joinDate = String(jr).split('T')[0].split(' ')[0].trim();
        }
 
        var accessRole = String(row[accessCol]||'').trim();
        var department = String(row[deptCol]||'').trim();
 
        // ── NEW: compute nav permissions ──────────────────────────
        var navPerms = getNavPermissions(accessRole, department);
 
        return { success:true, user:{
          name:           rowName,
          role:           String(row[roleCol]  ||'').trim(),
          department:     department,
          email:          String(row[emailCol] ||'').trim(),
          phone:          String(row[phoneCol] ||'').trim(),
          type:           String(row[typeCol]  ||'').trim(),
          accessRole:     accessRole,
          joinDate:       joinDate,
          diaryUrl:       diaryCol >= 0 ? String(row[diaryCol]||'').trim() : '',
          rowIndex:       i + 1,
          // ── permission payload sent to client ──
          navPermissions: navPerms
        }};
      }
    }
    return { success:false, error:'Invalid name or password.' };
  } catch(e) { return { success:false, error:'Server error: '+e.message }; }
}

// ── ATTENDANCE: MY OWN ──────────────────────────────────────────
function getMyAttendance(name) {
  try {
    var ss    = _internsSS();
    var sheet = ss.getSheetByName(SHEET_ATTEND);
    var data  = sheet.getDataRange().getValues();
    var hRow  = data[0];
    var tz    = Session.getScriptTimeZone();

    function findRow(nm) {
      var nFirst = nm.trim().toLowerCase().split(/\s+/)[0];
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim().toLowerCase() === nm.trim().toLowerCase()) return data[i];
      }
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim().toLowerCase().split(/\s+/)[0] === nFirst) return data[i];
      }
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim().toLowerCase().includes(nFirst)) return data[i];
      }
      return null;
    }

    var personRow = findRow(name);
    if (!personRow) return { error: 'Attendance record not found for ' + name };

    var presentCount = 0, absentCount = 0, wfhCount = 0;
    var records = [];

    for (var c = 2; c < hRow.length; c++) {
      var rawDate = hRow[c];
      var status  = String(personRow[c] || '').trim();
      if (!rawDate) continue;
      if (ATTEND_SKIP_VALUES.map(function(v){ return v.toLowerCase(); }).indexOf(status.toLowerCase()) !== -1) continue;

      var dateLabel = '';
      if (rawDate instanceof Date && !isNaN(rawDate.getTime())) {
        var realDay   = rawDate.getMonth() + 1;
        var realMonth = rawDate.getDate();
        var realYear  = rawDate.getFullYear();
        dateLabel = realYear + '-' + (realMonth < 10 ? '0' : '') + realMonth + '-' + (realDay < 10 ? '0' : '') + realDay;
      } else if (typeof rawDate === 'number') {
        var d = new Date(Math.round((rawDate - 25569) * 86400 * 1000));
        dateLabel = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
      } else {
        var sv = String(rawDate).trim();
        if (/^\d{4}-\d{2}-\d{2}$/.test(sv)) {
          dateLabel = sv;
        } else if (/^\d{1,2}\/\d{1,2}\/\d{2}$/.test(sv)) {
          var parts = sv.split('/');
          dateLabel = (2000 + parseInt(parts[2])) + '-'
            + (parts[1].length === 1 ? '0' : '') + parts[1] + '-'
            + (parts[0].length === 1 ? '0' : '') + parts[0];
        } else if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(sv)) {
          var parts2 = sv.split('/');
          dateLabel = parts2[2] + '-'
            + (parts2[1].length === 1 ? '0' : '') + parts2[1] + '-'
            + (parts2[0].length === 1 ? '0' : '') + parts2[0];
        } else {
          try {
            var parsed = new Date(sv);
            if (!isNaN(parsed.getTime())) {
              dateLabel = Utilities.formatDate(parsed, tz, 'yyyy-MM-dd');
            } else { dateLabel = sv; }
          } catch(ex) { dateLabel = sv; }
        }
      }

      var category = 'other';
      var sLow = status.toLowerCase();
      if (ATTEND_WFH_VALUES.map(function(v){ return v.toLowerCase(); }).indexOf(sLow) !== -1) {
        wfhCount++; presentCount++; category = 'wfh';
      } else if (ATTEND_PRESENT_VALUES.map(function(v){ return v.toLowerCase(); }).indexOf(sLow) !== -1) {
        presentCount++; category = 'present';
      } else if (ATTEND_ABSENT_VALUES.map(function(v){ return v.toLowerCase(); }).indexOf(sLow) !== -1) {
        absentCount++; category = 'absent';
      }
      records.push({ date: dateLabel, status: status, category: category });
    }

    var total = presentCount + absentCount;
    var pct   = total > 0 ? Math.round((presentCount / total) * 100) : 0;
    return { records: records, presentCount: presentCount, absentCount: absentCount, wfhCount: wfhCount, total: total, percentage: pct };
  } catch(e) { return { error: 'Attendance error: ' + e.message }; }
}

// ── ALL EMPLOYEES (role-filtered) ───────────────────────────────
function getAllEmployees(callerAccessRole, callerDepartment) {
  try {
    var ss    = _internsSS();
    var sheet = ss.getSheetByName(SHEET_INTERNS);
    var data  = sheet.getDataRange().getValues();
    var h     = data[0].map(function(x){ return String(x).trim().toLowerCase(); });
    var MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

    var ni=h.indexOf('name'), ri=h.indexOf('role'), di=h.indexOf('department');
    var ei=h.indexOf('email'), pi=h.indexOf('phone'), ti=h.indexOf('type');
    var ai=h.indexOf('access_role'), aci=h.indexOf('is_active'), ji=h.indexOf('join_date');
    var diari=h.indexOf('work diary url');

    var callerRole = (callerAccessRole || '').trim().toLowerCase();
    var callerDept = (callerDepartment  || '').trim().toLowerCase();
    var canSeeAll  = ADMIN_VIEW_ROLES.indexOf(callerRole) !== -1;

    var employees = [];
    for (var i=1;i<data.length;i++) {
      var row  = data[i];
      var name = String(row[ni]||'').trim();
      if (!name) continue;
      if (!canSeeAll && callerDept) {
        var empDept = String(row[di]||'').trim().toLowerCase();
        if (empDept !== callerDept) continue;
      }
      var jr = row[ji], joinDate = '';
      if (jr instanceof Date) { joinDate = jr.getDate()+' '+MONTHS[jr.getMonth()]+' '+jr.getFullYear(); }
      else if (jr && typeof jr === 'number') { var d=new Date(Math.round((jr-25569)*86400*1000));joinDate=d.getDate()+' '+MONTHS[d.getMonth()]+' '+d.getFullYear(); }
      else if (jr) { joinDate = String(jr).split('T')[0].split(' ')[0].trim(); }

      employees.push({
        rowIndex:   i+1, name: name,
        role:       String(row[ri]  ||'').trim(),
        department: String(row[di]  ||'').trim(),
        email:      String(row[ei]  ||'').trim(),
        phone:      String(row[pi]  ||'').trim(),
        type:       String(row[ti]  ||'').trim(),
        accessRole: String(row[ai]  ||'').trim(),
        isActive:   String(row[aci] ||'').trim().toUpperCase(),
        joinDate:   joinDate,
        diaryUrl:   diari>=0 ? String(row[diari]||'').trim() : ''
      });
    }
    return { success:true, employees:employees };
  } catch(e) { return { success:false, error:e.message }; }
}

// ── EDIT EMPLOYEE ────────────────────────────────────────────────
function updateEmployee(callerAccessRole, rowIndex, fields) {
  try {
    if (EDIT_ROLES.map(function(r){return r.toLowerCase();}).indexOf((callerAccessRole||'').toLowerCase()) === -1)
      return { success:false, error:'Permission denied.' };
    var ss    = _internsSS();
    var sheet = ss.getSheetByName(SHEET_INTERNS);
    var h     = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(function(x){return String(x).trim().toLowerCase();});
    var colMap = {
      role:       h.indexOf('role')+1,
      department: h.indexOf('department')+1,
      email:      h.indexOf('email')+1,
      phone:      h.indexOf('phone')+1,
      type:       h.indexOf('type')+1,
      isActive:   h.indexOf('is_active')+1
    };
    for (var key in colMap) {
      if (colMap[key] > 0 && fields[key] !== undefined && fields[key] !== null)
        sheet.getRange(rowIndex, colMap[key]).setValue(fields[key]);
    }
    SpreadsheetApp.flush();
    return { success:true, message:'Employee updated.' };
  } catch(e) { return { success:false, error:e.message }; }
}

// ── ATTENDANCE GRID ─────────────────────────────────────────────
function getAttendanceGrid(personName) {
  try {
    var ss    = _internsSS();
    var sheet = ss.getSheetByName(SHEET_ATTEND);
    var data  = sheet.getDataRange().getValues();
    var hRow  = data[0];

    function findRowIdx(nm) {
      var nFirst = nm.trim().toLowerCase().split(/\s+/)[0];
      for (var i=1;i<data.length;i++) {
        if (String(data[i][0]).trim().toLowerCase() === nm.trim().toLowerCase()) return i;
      }
      for (var i=1;i<data.length;i++) {
        if (String(data[i][0]).trim().toLowerCase().split(/\s+/)[0] === nFirst) return i;
      }
      for (var i=1;i<data.length;i++) {
        if (String(data[i][0]).trim().toLowerCase().includes(nFirst)) return i;
      }
      return -1;
    }

    var ri = findRowIdx(personName);
    if (ri === -1) return { success:false, error:'Person not found.' };
    var personRow = data[ri];
    var grid = [];

    for (var c=2;c<hRow.length;c++) {
      var rawDate = hRow[c];
      if (!rawDate) continue;
      var dateLabel = '';
      if (rawDate instanceof Date) {
        dateLabel = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), 'dd/MM/yy');
      } else if (typeof rawDate === 'number') {
        dateLabel = Utilities.formatDate(new Date(Math.round((rawDate-25569)*86400*1000)), Session.getScriptTimeZone(), 'dd/MM/yy');
      } else {
        dateLabel = String(rawDate).split('T')[0].split(' ')[0].trim();
      }
      grid.push({ dateHeader:dateLabel, status:String(personRow[c]||'').trim() });
    }

    var sheetUrl = _internsSS().getUrl() + '#gid=' + sheet.getSheetId();
    return { success:true, grid:grid, sheetUrl:sheetUrl };
  } catch(e) { return { success:false, error:e.message }; }
}

function updateAttendanceCell(callerAccessRole, personName, dateHeader, newValue) {
  try {
    if (EDIT_ROLES.map(function(r){return r.toLowerCase();}).indexOf((callerAccessRole||'').toLowerCase()) === -1)
      return { success:false, error:'Permission denied.' };
    var validValues = ['Present','Absent','WFH','Holiday',''];
    if (validValues.indexOf(newValue) === -1) return { success:false, error:'Invalid status.' };

    var ss    = _internsSS();
    var sheet = ss.getSheetByName(SHEET_ATTEND);
    var data  = sheet.getDataRange().getValues();
    var hRow  = data[0];

    var dateCol = -1;
    for (var c=2;c<hRow.length;c++) {
      var hl = String(hRow[c]).trim();
      if (hl === dateHeader.trim()) { dateCol=c; break; }
      if (hRow[c] instanceof Date) {
        var fmt = Utilities.formatDate(hRow[c], Session.getScriptTimeZone(), 'dd/MM/yy');
        if (fmt === dateHeader.trim()) { dateCol=c; break; }
      }
    }
    if (dateCol === -1) return { success:false, error:'Date column not found: '+dateHeader };

    var pRow = -1;
    for (var r=1;r<data.length;r++) {
      if (String(data[r][0]).trim().toLowerCase() === personName.trim().toLowerCase()) { pRow=r; break; }
    }
    if (pRow === -1) return { success:false, error:'Person not found: '+personName };

    sheet.getRange(pRow+1, dateCol+1).setValue(newValue);
    SpreadsheetApp.flush();
    return { success:true, message:'Updated.' };
  } catch(e) { return { success:false, error:e.message }; }
}

function getAttendanceSheetInfo() {
  try {
    var ss = _internsSS();
    var sh = ss.getSheetByName(SHEET_ATTEND);
    return { success:true, url:ss.getUrl()+'#gid='+sh.getSheetId() };
  } catch(e) { return { success:false, error:e.message }; }
}

// ── WORK DIARY ───────────────────────────────────────────────────
function updateDiaryUrl(callerName, newUrl) {
  try {
    var ss    = _internsSS();
    var sheet = ss.getSheetByName(SHEET_INTERNS);
    var data  = sheet.getDataRange().getValues();
    var rawH  = data[0];
    var h     = rawH.map(function(x){ return String(x).trim().toLowerCase(); });

    var nameCol  = h.indexOf('name');
    var diaryCol = h.indexOf('work diary url');

    if (nameCol === -1) return { success:false, error:'Name column not found.' };
    if (diaryCol === -1) {
      for (var hi=0;hi<h.length;hi++) {
        if (h[hi].indexOf('diary')!==-1||h[hi].indexOf('work')!==-1){ diaryCol=hi; break; }
      }
      if (diaryCol === -1) return { success:false, error:'Work diary URL column not found.' };
    }

    var callerLower = callerName.trim().toLowerCase();
    for (var i=1;i<data.length;i++) {
      if (String(data[i][nameCol]||'').trim().toLowerCase() === callerLower) {
        sheet.getRange(i+1, diaryCol+1).setValue(newUrl ? newUrl.trim() : '');
        SpreadsheetApp.flush();
        return { success:true, message:'Work diary URL updated successfully.' };
      }
    }
    return { success:false, error:'User "'+callerName+'" not found in sheet.' };
  } catch(e) { return { success:false, error:'Server error: '+e.message }; }
}

function colLetter(idx) { return String.fromCharCode(65+idx); }