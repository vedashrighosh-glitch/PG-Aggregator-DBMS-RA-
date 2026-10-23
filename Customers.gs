// ================================================================
//  RoomAdda — Customers.gs  (v8 — all 19 columns mapped)
//
//  ACTUAL CUSTOMERS SHEET HEADERS (19 cols):
//  Inc No | Date | Guest Name | Room NO | Location | Electricity |
//  Mobile | Sharing Type | Monthly Rent | SD | Duration | TOTAL AMC |
//  RoomAdda Collection | Paid | Commission | check-in date |
//  Check-in status | Payment Mode | Bank Statement
// ================================================================

var CUSTOMER_SHEET = 'customers';

var CUST_COLS = [
  'name','phone','email','emergency_contact_name','emergency_contact_phone',
  'pg_name','room_type','room_number','move_in_date','stay_type',
  'customer_status','notes','created_at','updated_at',
  // Extended sheet columns
  'inc_no','electricity','monthly_rent','sd','duration',
  'total_amc','roomadda_collection','paid','commission','payment_mode','bank_statement'
];

function customerSheet() {
  var ss = SpreadsheetApp.openById(DATA_SS_ID);
  var sh = ss.getSheetByName(CUSTOMER_SHEET);
  if (!sh) throw new Error('Sheet "' + CUSTOMER_SHEET + '" not found.');
  return sh;
}

function _getCustHeaders(sh) {
  var last = sh.getLastColumn();
  return sh.getRange(1, 1, 1, last).getValues()[0].map(function(h) {
    return String(h).trim();
  });
}

// Maps real sheet header → internal UI key
function _custH2k(h) {
  var map = {
    'guest name':                'name',
    'guest nam':                 'name',
    'name':                      'name',
    'mobile':                    'phone',
    'mobile number':             'phone',
    'phone':                     'phone',
    'contact no.':               'phone',
    'contact no':                'phone',
    'location':                  'pg_name',
    'pg name':                   'pg_name',
    'pg_name':                   'pg_name',
    'room no':                   'room_number',
    'room no.':                  'room_number',
    'room number':               'room_number',
    'room_number':               'room_number',
    'sharing type':              'room_type',
    'room type':                 'room_type',
    'room_type':                 'room_type',
    'check-in date':             'move_in_date',
    'move_in_date':              'move_in_date',
    'move in date':              'move_in_date',
    'check-in status':           'customer_status',
    'customer_status':           'customer_status',
    'status':                    'customer_status',
    'comments':                  'notes',
    'notes':                     'notes',
    'email':                     'email',
    'date':                      'created_at',
    'created_at':                'created_at',
    'updated_at':                'updated_at',
    'inc no':                    'inc_no',
    'inc no.':                   'inc_no',
    'inc_no':                    'inc_no',
    'monthly rent':              'monthly_rent',
    'monthly_rent':              'monthly_rent',
    'sd':                        'sd',
    'duration':                  'duration',
    'total amc':                 'total_amc',
    'total_amc':                 'total_amc',
    'electricity':               'electricity',
    'roomadda collection':       'roomadda_collection',
    'roomadda collection ':      'roomadda_collection',
    'roomadda_collection':       'roomadda_collection',
    'paid':                      'paid',
    'commission':                'commission',
    'payment mode':              'payment_mode',
    'payment_mode':              'payment_mode',
    'bank statement (if any)':   'bank_statement',
    'bank statement':            'bank_statement',
    'bank_statement':            'bank_statement',
    'stay_type':                 'stay_type',
    'stay type':                 'stay_type',
    'emergency_contact_name':    'emergency_contact_name',
    'emergency_contact_phone':   'emergency_contact_phone',
  };
  return map[h.toLowerCase().trim()] || h.toLowerCase().trim().replace(/[^a-z0-9]+/g, '_').replace(/^_|_$/g, '');
}

// ── DATE FORMATTER ───────────────────────────────────────────────
function fmtCVal(val) {
  if (val instanceof Date) {
    if (!isNaN(val.getTime())) {
      try {
        return Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd/MM/yyyy');
      } catch(e) {
        var d = val.getDate(), m = val.getMonth() + 1, y = val.getFullYear();
        return (d < 10 ? '0' + d : d) + '/' + (m < 10 ? '0' + m : m) + '/' + y;
      }
    }
    return '';
  }
  if (val === null || val === undefined) return '';
  if (typeof val === 'boolean') return val ? 'true' : 'false';

  var sv = String(val).trim();
  if (sv === '') return '';

  if (sv.match(/^\d{4}-\d{2}-\d{2}T/)) {
    var parts = sv.split('T')[0].split('-');
    if (parts.length === 3) return parts[2] + '/' + parts[1] + '/' + parts[0];
    return sv.split('T')[0];
  }
  if (sv.match(/^\d{4}-\d{2}-\d{2}$/)) {
    var p2 = sv.split('-');
    return p2[2] + '/' + p2[1] + '/' + p2[0];
  }
  if (sv.match(/^\d{2}\/\d{2}\/\d{4}\s+\d{1,2}:\d{2}/)) return sv.split(/\s+/)[0];
  if (sv.match(/^\d{2}-\d{2}-\d{4}\s+\d{1,2}:\d{2}/)) return sv.split(/\s+/)[0].replace(/-/g, '/');
  if (sv.match(/^\d{2}\/\d{2}\/\d{4}$/)) return sv;

  return sv;
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

// ── GET ALL CUSTOMERS ────────────────────────────────────────────
function getCustomers() {
  try {
    var sh      = customerSheet();
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { status: 'success', data: [] };

    var rawHeaders = _getCustHeaders(sh);
    var numCols    = rawHeaders.length;
    var rows       = sh.getRange(2, 1, lastRow - 1, numCols).getValues();

    var data = rows
      .filter(function(r) {
        return r.some(function(c) { return c !== '' && c !== null && c !== undefined; });
      })
      .map(function(r, rowIdx) {
        var obj = {};
        rawHeaders.forEach(function(h, i) {
          var key = _custH2k(h);
          // Only format as date for date-type columns; keep numbers as-is
          var dateKeys = ['move_in_date','created_at','updated_at'];
          if (dateKeys.indexOf(key) !== -1) {
            obj[key] = fmtCVal(r[i]);
          } else {
            // For numeric columns, preserve the value cleanly
            if (r[i] instanceof Date) {
              obj[key] = fmtCVal(r[i]);
            } else if (r[i] === null || r[i] === undefined) {
              obj[key] = '';
            } else {
              obj[key] = String(r[i]).replace(/\.0$/, '');
            }
          }
        });

        // Guarantee all UI-required keys exist
        CUST_COLS.forEach(function(k) {
          if (obj[k] === undefined || obj[k] === null) obj[k] = '';
        });

        // Clean phone
        obj.phone = _cleanPhone(obj.phone);

        // Default status
        if (!obj.customer_status) obj.customer_status = 'Active';

        // Store actual sheet row number
        obj._sheetRow = rowIdx + 2;

        return obj;
      });

    return { status: 'success', data: data };
  } catch(e) {
    return { status: 'error', message: e.message };
  }
}

// ── ADD CUSTOMER ─────────────────────────────────────────────────
function addCustomer(payload) {
  try {
    if (!payload.name || !payload.phone)
      return { status: 'error', message: 'Name and Phone are required.' };

    var sh         = customerSheet();
    var lastRow    = sh.getLastRow();
    var rawHeaders = _getCustHeaders(sh);

    // Duplicate check
    if (lastRow >= 2) {
      var ni = -1, pi = -1;
      rawHeaders.forEach(function(h, i) {
        var k = _custH2k(h);
        if (k === 'name')  ni = i;
        if (k === 'phone') pi = i;
      });
      if (ni !== -1 && pi !== -1) {
        var existing = sh.getRange(2, 1, lastRow - 1, rawHeaders.length).getValues();
        for (var i = 0; i < existing.length; i++) {
          var sn = String(existing[i][ni] || '').trim().toLowerCase();
          var sp = _cleanPhone(existing[i][pi]).replace(/[^0-9]/g, '');
          if (sn === payload.name.trim().toLowerCase() &&
              sp === String(payload.phone).replace(/[^0-9]/g, ''))
            return { status: 'error', message: 'Customer already exists.' };
        }
      }
    }

    var now    = new Date();
    var newRow = rawHeaders.map(function(h) {
      var k = _custH2k(h);
      switch(k) {
        case 'name':                    return payload.name                    || '';
        case 'phone':                   return payload.phone                   || '';
        case 'email':                   return payload.email                   || '';
        case 'emergency_contact_name':  return payload.emergency_contact_name  || '';
        case 'emergency_contact_phone': return payload.emergency_contact_phone || '';
        case 'pg_name':                 return payload.pg_name                 || '';
        case 'room_type':               return payload.room_type               || '';
        case 'room_number':             return payload.room_number             || '';
        case 'move_in_date':
          if (payload.move_in_date) {
            var d = new Date(payload.move_in_date);
            return isNaN(d.getTime()) ? payload.move_in_date : d;
          }
          return '';
        case 'stay_type':               return payload.stay_type               || '';
        case 'customer_status':         return payload.customer_status         || 'Active';
        case 'notes':                   return payload.notes                   || '';
        case 'monthly_rent':            return payload.monthly_rent            || '';
        case 'sd':                      return payload.sd                      || '';
        case 'duration':                return payload.duration                || '';
        case 'total_amc':               return payload.total_amc               || '';
        case 'electricity':             return payload.electricity             || '';
        case 'roomadda_collection':     return payload.roomadda_collection     || '';
        case 'paid':                    return payload.paid                    || '';
        case 'commission':              return payload.commission              || '';
        case 'payment_mode':            return payload.payment_mode            || '';
        case 'bank_statement':          return payload.bank_statement          || '';
        case 'created_at':              return now;
        case 'updated_at':              return now;
        default:                        return '';
      }
    });

    sh.appendRow(newRow);
    SpreadsheetApp.flush();
    return { status: 'success', message: 'Customer added.' };
  } catch(e) {
    return { status: 'error', message: e.message };
  }
}

// ── UPDATE CUSTOMER ──────────────────────────────────────────────
function updateCustomer(payload) {
  try {
    var sh         = customerSheet();
    var lastRow    = sh.getLastRow();
    if (lastRow < 2) return { status: 'error', message: 'No customer data.' };

    var rawHeaders = _getCustHeaders(sh);
    var ni = -1, pi = -1;
    rawHeaders.forEach(function(h, i) {
      var k = _custH2k(h);
      if (k === 'name')  ni = i;
      if (k === 'phone') pi = i;
    });

    var rows = sh.getRange(2, 1, lastRow - 1, rawHeaders.length).getValues();
    var sheetRow = -1;

    for (var i = 0; i < rows.length; i++) {
      var nm = String(rows[i][ni] || '').trim().toLowerCase();
      var ph = _cleanPhone(rows[i][pi]).replace(/[^0-9]/g, '');
      if (nm === String(payload.name).trim().toLowerCase() &&
          ph === String(payload.phone).replace(/[^0-9]/g, '')) {
        sheetRow = i + 2;
        break;
      }
    }

    if (sheetRow === -1) return { status: 'error', message: 'Customer not found.' };

    rawHeaders.forEach(function(h, idx) {
      var k = _custH2k(h);
      if (k === 'name' || k === 'created_at') return;
      if (k === 'updated_at') {
        sh.getRange(sheetRow, idx + 1).setValue(new Date());
        return;
      }
      if (payload[k] !== undefined && payload[k] !== null) {
        sh.getRange(sheetRow, idx + 1).setValue(payload[k]);
      }
    });

    SpreadsheetApp.flush();
    return { status: 'success', message: 'Customer updated.' };
  } catch(e) {
    return { status: 'error', message: e.message };
  }
}

// ── DELETE CUSTOMER ──────────────────────────────────────────────
function deleteCustomer(payload) {
  try {
    var sh         = customerSheet();
    var lastRow    = sh.getLastRow();
    if (lastRow < 2) return { status: 'error', message: 'No customer data.' };

    var rawHeaders = _getCustHeaders(sh);
    var ni = -1, pi = -1;
    rawHeaders.forEach(function(h, i) {
      var k = _custH2k(h);
      if (k === 'name')  ni = i;
      if (k === 'phone') pi = i;
    });

    // If _sheetRow provided, verify it still matches before deleting
    if (payload._sheetRow && payload._sheetRow >= 2 && payload._sheetRow <= lastRow) {
      var verifyRow = sh.getRange(payload._sheetRow, 1, 1, rawHeaders.length).getValues()[0];
      var vName  = String(verifyRow[ni] || '').trim().toLowerCase();
      var vPhone = _cleanPhone(verifyRow[pi]).replace(/[^0-9]/g, '');
      if (vName === String(payload.name).trim().toLowerCase() &&
          vPhone === String(payload.phone).replace(/[^0-9]/g, '')) {
        sh.deleteRow(payload._sheetRow);
        SpreadsheetApp.flush();
        return { status: 'success', message: 'Customer deleted.' };
      }
    }

    // Fallback: scan by name + phone
    var rows = sh.getRange(2, 1, lastRow - 1, rawHeaders.length).getValues();
    var sheetRow = -1;
    for (var i = 0; i < rows.length; i++) {
      var nm = String(rows[i][ni] || '').trim().toLowerCase();
      var ph = _cleanPhone(rows[i][pi]).replace(/[^0-9]/g, '');
      if (nm === String(payload.name).trim().toLowerCase() &&
          ph === String(payload.phone).replace(/[^0-9]/g, '')) {
        sheetRow = i + 2;
        break;
      }
    }

    if (sheetRow === -1) return { status: 'error', message: 'Customer not found.' };

    sh.deleteRow(sheetRow);
    SpreadsheetApp.flush();
    return { status: 'success', message: 'Customer deleted.' };
  } catch(e) {
    return { status: 'error', message: e.message };
  }
}

// ── TEST ─────────────────────────────────────────────────────────
function testCustomers() {
  var r = getCustomers();
  Logger.log('Status: ' + r.status);
  if (r.status === 'error') { Logger.log('ERROR: ' + r.message); return; }
  Logger.log('Total customers: ' + r.data.length);
  if (r.data.length > 0) {
    var first = r.data[0];
    Logger.log('Inc No: '          + first.inc_no);
    Logger.log('Name: '            + first.name);
    Logger.log('Phone: '           + first.phone);
    Logger.log('PG: '              + first.pg_name);
    Logger.log('Room: '            + first.room_number);
    Logger.log('Sharing Type: '    + first.room_type);
    Logger.log('Move-in: '         + first.move_in_date);
    Logger.log('Status: '          + first.customer_status);
    Logger.log('Monthly Rent: '    + first.monthly_rent);
    Logger.log('SD: '              + first.sd);
    Logger.log('Duration: '        + first.duration);
    Logger.log('Total AMC: '       + first.total_amc);
    Logger.log('Electricity: '     + first.electricity);
    Logger.log('RA Collection: '   + first.roomadda_collection);
    Logger.log('Paid: '            + first.paid);
    Logger.log('Commission: '      + first.commission);
    Logger.log('Payment Mode: '    + first.payment_mode);
    Logger.log('Bank Statement: '  + first.bank_statement);
    Logger.log('Sheet Row: '       + first._sheetRow);
  }
}