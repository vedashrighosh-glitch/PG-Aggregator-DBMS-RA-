// ================================================================
//  RoomAdda — Code.gs  (v9 — server-side role injection fix)
//
//  ROOT CAUSE:
//  window.location.search is ALWAYS empty in Google Apps Script
//  web apps — the client HTML never sees URL params. So _ROLE and
//  _DEPT were always '' in every module page, making getNavUrls
//  build links with no role, causing Access Denied on navigation.
//
//  FIX:
//  _serve() now uses createTemplateFromFile() and injects
//  roleParam + deptParam directly into the HTML as JS variables
//  using <?= ?> template tags. The HTML pages read these injected
//  values instead of trying to read window.location.search.
// ================================================================

var SS_ID         = '1XHDlNzwFUNy5Ekl92DePS_KoKUDpec601ghwOwhPqrI';
var DATA_SS_ID    = '1XHDlNzwFUNy5Ekl92DePS_KoKUDpec601ghwOwhPqrI';
var SHEET_INTERNS = 'Interns';

// ── doGet ROUTER ────────────────────────────────────────────────
function doGet(e) {
  var page      = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'portal';
  var roleParam = (e && e.parameter && e.parameter.role) ? e.parameter.role : '';
  var deptParam = (e && e.parameter && e.parameter.dept) ? e.parameter.dept : '';

  if (page === 'portal' || page === '') {
    return _serve('Interns1', 'RoomAdda — Team Portal', '', '');
  }

  if (!canAccessPage(roleParam, deptParam, page)) {
    var displayRole = roleParam || deptParam || 'unknown';
    return _serveAccessDenied(page, displayRole);
  }

  switch (page) {
    case 'leads':     return _serve('Leads1',     'RoomAdda — Leads',      roleParam, deptParam);
    case 'customers': return _serve('Customers1', 'RoomAdda — Customers',  roleParam, deptParam);
    case 'finance':   return _serve('Finance1',   'RoomAdda — Finance',    roleParam, deptParam);
    case 'pg':        return _serve('Pg',         'RoomAdda — PG Details', roleParam, deptParam);
    default:          return _serve('Interns1',   'RoomAdda — Team Portal', '', '');
  }
}

// ── SERVE HELPER (uses template to inject role+dept into HTML) ───
function _serve(file, title, roleParam, deptParam) {
  var tpl = HtmlService.createTemplateFromFile(file);
  tpl.roleParam = roleParam || '';
  tpl.deptParam = deptParam || '';
  return tpl.evaluate()
    .setTitle(title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── LOGIN ────────────────────────────────────────────────────────
function loginUser(name, password) {
  try {
    var ss    = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName(SHEET_INTERNS);
    if (!sheet) return { success: false, error: 'User sheet not found.' };

    var data    = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h).trim().toLowerCase(); });

    var nameCol   = headers.indexOf('name');
    var passCol   = headers.findIndex(function(h) {
      return h.replace(/\s+/g,'').toLowerCase() === 'password';
    });
    var deptCol   = headers.indexOf('department');
    var roleCol   = headers.indexOf('role');
    var accessCol = headers.indexOf('access_role');
    var activeCol = headers.indexOf('is_active');

    for (var i = 1; i < data.length; i++) {
      var rowName = String(data[i][nameCol] || '').trim().toLowerCase();
      var rowPass = String(data[i][passCol] || '').trim();

      if (rowName === name.trim().toLowerCase() && rowPass === password.trim()) {

        var isActive = String(data[i][activeCol] || '').trim().toUpperCase();
        if (isActive === 'FALSE' || isActive === 'NO' || isActive === '0') {
          return { success: false, error: 'Account is inactive. Contact a Founder.' };
        }

        var userObj = {
          name:       String(data[i][nameCol]   || '').trim(),
          department: String(data[i][deptCol]   || '').trim(),
          role:       String(data[i][roleCol]   || '').trim(),
          accessRole: String(data[i][accessCol] || '').trim(),
          isActive:   isActive
          
        };

        var nav = getNavPermissions(userObj.accessRole, userObj.department);
        return { success: true, user: userObj, nav: nav };
      }
    }
    return { success: false, error: 'Invalid name or password.' };

  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ── getNavUrls ───────────────────────────────────────────────────
function getNavUrls(accessRole, department) {
  try {
    var base = ScriptApp.getService().getUrl();
    var r    = encodeURIComponent(accessRole || '');
    var d    = encodeURIComponent(department || '');
    var nav  = getNavPermissions(accessRole, department);

    return {
      success: true,
      nav: nav,
      urls: {
        portal:    base,
        leads:     base + '?page=leads&role='     + r + '&dept=' + d,
        customers: base + '?page=customers&role=' + r + '&dept=' + d,
        finance:   base + '?page=finance&role='   + r + '&dept=' + d,
        pg:        base + '?page=pg&role='        + r + '&dept=' + d
      }
    };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ── getUserContext ───────────────────────────────────────────────
function getUserContext(accessRole, department) {
  try {
    var nav  = getNavPermissions(accessRole, department);
    var urls = getNavUrls(accessRole, department);
    return {
      success:    true,
      accessRole: accessRole || '',
      department: department || '',
      nav:        nav,
      urls:       urls.urls
    };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ── ORIGINAL URL HELPERS (kept for backward compat) ──────────────
function getPortalUrl() {
  return { success: true, url: ScriptApp.getService().getUrl() };
}
function getLeadsUrl(accessRole, department) {
  return { success: true, url: ScriptApp.getService().getUrl()
    + '?page=leads&role=' + encodeURIComponent(accessRole || '')
    + '&dept=' + encodeURIComponent(department || '') };
}
function getCustomersUrl(accessRole, department) {
  return { success: true, url: ScriptApp.getService().getUrl()
    + '?page=customers&role=' + encodeURIComponent(accessRole || '')
    + '&dept=' + encodeURIComponent(department || '') };
}
function getFinanceUrl(accessRole, department) {
  return { success: true, url: ScriptApp.getService().getUrl()
    + '?page=finance&role=' + encodeURIComponent(accessRole || '')
    + '&dept=' + encodeURIComponent(department || '') };
}
function getPGDetailsUrl(accessRole, department) {
  return { success: true, url: ScriptApp.getService().getUrl()
    + '?page=pg&role=' + encodeURIComponent(accessRole || '')
    + '&dept=' + encodeURIComponent(department || '') };
}

// ── ACCESS DENIED PAGE ───────────────────────────────────────────
function _serveAccessDenied(page, role) {
  var body = '<!DOCTYPE html><html><head><meta charset="UTF-8"/>'
    + '<title>RoomAdda — Access Denied</title>'
    + '<style>*{box-sizing:border-box;margin:0;padding:0}'
    + 'body{font-family:sans-serif;background:#06070F;color:#F0F0F0;'
    + 'display:flex;align-items:center;justify-content:center;min-height:100vh}'
    + '.box{text-align:center;padding:48px;max-width:420px}'
    + '.icon{font-size:56px;margin-bottom:20px}'
    + 'h1{font-size:24px;font-weight:700;margin-bottom:10px;color:#FF5B1F}'
    + 'p{font-size:14px;color:#8890B0;line-height:1.6;margin-bottom:24px}'
    + 'a{display:inline-block;padding:10px 24px;background:#FF5B1F;color:#fff;'
    + 'border-radius:8px;text-decoration:none;font-weight:600;font-size:14px}'
    + '</style></head><body>'
    + '<div class="box"><div class="icon">🔒</div>'
    + '<h1>Access Denied</h1>'
    + '<p>Your role (<strong>' + role + '</strong>) does not have permission to access '
    + 'the <strong>' + page + '</strong> module.<br/><br/>'
    + 'Contact a Founder to request access or update your role.</p>'
    + '<a href="' + ScriptApp.getService().getUrl() + '">← Back to Portal</a>'
    + '</div></body></html>';
  return HtmlService.createHtmlOutput(body)
    .setTitle('RoomAdda — Access Denied')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── DEBUG HELPERS ────────────────────────────────────────────────
function debugSheets() {
  var ss = SpreadsheetApp.openById('1f2A6RP4Ux_wIozT4DibwIohEoRd29yyzOV5PfOIATuM');
  Logger.log('All sheets: ' + ss.getSheets().map(function(s){return s.getName()}).join(', '));
  var leads = ss.getSheetByName('leads');
  Logger.log('Leads last row: ' + leads.getLastRow());
  Logger.log('Row 1: ' + JSON.stringify(leads.getRange(1,1,1,10).getValues()[0]));
}
function debugFull() {
  var ss    = SpreadsheetApp.openById('1f2A6RP4Ux_wIozT4DibwIohEoRd29yyzOV5PfOIATuM');
  var leads = ss.getSheetByName('leads');
  Logger.log('LEADS HEADERS: ' + JSON.stringify(leads.getRange(1,1,1,leads.getLastColumn()).getValues()[0]));
  var cust  = ss.getSheetByName('customers');
  Logger.log('CUSTOMERS HEADERS: ' + JSON.stringify(cust.getRange(1,1,1,cust.getLastColumn()).getValues()[0]));
}
function testLeadsNow() {
  var r = getAllLeads();
  Logger.log('Status: ' + r.status);
  if (r.status === 'error') { Logger.log('ERROR: ' + r.message); return; }
  Logger.log('Count: ' + r.data.length);
}
function debugBoth() {
  try { var r1 = getAllLeads();   Logger.log('LEADS: '     + r1.status + ' count=' + (r1.data||[]).length); }
  catch(e) { Logger.log('LEADS exception: ' + e.message); }
  try { var r2 = getCustomers(); Logger.log('CUSTOMERS: ' + r2.status + ' count=' + (r2.data||[]).length); }
  catch(e) { Logger.log('CUSTOMERS exception: ' + e.message); }
}