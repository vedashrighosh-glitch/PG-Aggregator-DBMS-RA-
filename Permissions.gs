// ================================================================
//  RoomAdda — Permissions.gs
// ================================================================

var PAGE_KEYS = {
  PORTAL:    'portal',
  LEADS:     'leads',
  CUSTOMERS: 'customers',
  FINANCE:   'finance',
  PG:        'pg'
};

function _normRole(r) {
  return String(r || '').trim().toLowerCase();
}

function getCanonicalRole(accessRole, department) {
  var r = _normRole(accessRole);
  var d = _normRole(department);
  if (r) return r;
  return d;   // fall back to department if no explicit access_role
}

// ── PAGE ACCESS CHECK ──────────────────────────────────────────
function canAccessPage(accessRole, department, page) {
  var r = _normRole(accessRole);
  var d = _normRole(department);

  // Resolve effective role: explicit access_role wins, else use department
  var effective = r || d;

  switch (page) {

    case PAGE_KEYS.PORTAL:
      return true; // everyone, scope filtered server-side

    case PAGE_KEYS.LEADS:
    case PAGE_KEYS.CUSTOMERS:
    case PAGE_KEYS.PG:
      // Founder, Admin, Sales (by access_role OR department)
      return effective === 'founder' ||
             effective === 'admin'   ||
             effective === 'sales';

    case PAGE_KEYS.FINANCE:
      // Founder, Finance (by access_role OR department)
      return effective === 'founder' ||
             effective === 'finance';

    default:
      return false;
  }
}

// ── NAV PERMISSIONS ────────────────────────────────────────────
function getNavPermissions(accessRole, department) {
  var r = _normRole(accessRole);
  var d = _normRole(department);

  // Effective role: explicit access_role wins, else department
  var effective = r || d;

  var isFounder    = (effective === 'founder');
  var isAdmin      = (effective === 'admin');
  var isSalesUser  = (effective === 'sales');
  var isFinance    = (effective === 'finance');
  var isHR         = (effective === 'hr');

  return {
    showLeads:     isFounder || isAdmin || isSalesUser,
    showCustomers: isFounder || isAdmin || isSalesUser,
    showFinance:   isFounder || isFinance,
    showPG:        isFounder || isAdmin || isSalesUser,

    // Team Directory scope
    dirScope: (isFounder || isAdmin || isHR) ? 'all' : (d || 'own'),

    canSetPasswords: isFounder,

    // Raw flags
    isFounder: isFounder,
    isAdmin:   isAdmin,
    isSales:   isSalesUser,
    isFinance: isFinance,
    isHR:      isHR,

    // Useful for display — never shows "unknown"
    displayRole: effective || 'member'
  };
}

// ── SERVER-SIDE PAGE GUARD ─────────────────────────────────────
function isPageDenied(user, page) {
  if (!user) return true;
  return !canAccessPage(user.accessRole, user.department, page);
}

// ── PASSWORD GUARDS ────────────────────────────────────────────
function canSetPassword(callerAccessRole) {
  return _normRole(callerAccessRole) === 'founder';
}

function setUserPassword(callerAccessRole, targetName, newPassword) {
  if (!canSetPassword(callerAccessRole)) {
    return { success: false, error: 'Only Founders can set passwords.' };
  }
  if (!targetName || !newPassword) {
    return { success: false, error: 'Target name and new password are required.' };
  }
  try {
    var ss    = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName(SHEET_INTERNS);
    if (!sheet) return { success: false, error: 'Interns sheet not found.' };

    var data    = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h).trim().toLowerCase(); });
    var nameCol = headers.indexOf('name');
    var passCol = headers.indexOf('password');

    if (nameCol === -1 || passCol === -1) {
      return { success: false, error: 'Name or password column not found.' };
    }

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][nameCol] || '').trim().toLowerCase() ===
          targetName.trim().toLowerCase()) {
        sheet.getRange(i + 1, passCol + 1).setValue(newPassword.trim());
        SpreadsheetApp.flush();
        return { success: true, message: 'Password updated for ' + targetName + '.' };
      }
    }
    return { success: false, error: 'User "' + targetName + '" not found.' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function getAllUsersForAdmin(callerAccessRole) {
  if (!canSetPassword(callerAccessRole)) {
    return { success: false, error: 'Only Founders can view all user accounts.' };
  }
  try {
    var ss    = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName(SHEET_INTERNS);
    var data  = sheet.getDataRange().getValues();
    var h     = data[0].map(function(x) { return String(x).trim().toLowerCase(); });
    var ni  = h.indexOf('name'),
        di  = h.indexOf('department'),
        ri  = h.indexOf('role'),
        ai  = h.indexOf('access_role'),
        aci = h.indexOf('is_active');

    var users = [];
    for (var i = 1; i < data.length; i++) {
      var name = String(data[i][ni] || '').trim();
      if (!name) continue;
      users.push({
        name:       name,
        department: String(data[i][di]  || '').trim(),
        role:       String(data[i][ri]  || '').trim(),
        accessRole: String(data[i][ai]  || '').trim(),
        isActive:   String(data[i][aci] || '').trim().toUpperCase()
      });
    }
    return { success: true, users: users };
  } catch (e) {
    return { success: false, error: e.message };
  }
}