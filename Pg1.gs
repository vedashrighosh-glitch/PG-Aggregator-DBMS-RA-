// ================================================================
//  RoomAdda — Pg1.gs  (June 2026 rewrite)
//
//  Three sheets (all in DATA_SS_ID workbook):
//    • Sg Palya Details  — one row per property (list view)
//    • Final Rent        — one row per room type (detail table)
//    • Final Property    — one row per property  (detail info)
//
//  DATA_SS_ID is declared in Code.gs — do not redeclare here.
// ================================================================

var SG_PALYA_SHEET   = 'Sg Palya Details';
var FINAL_RENT_SHEET = 'Final Rent';
var FINAL_PROP_SHEET = 'Final Property';

// ── SHEET ACCESSORS ──────────────────────────────────────────────
function _sgSheet() {
  var ss = SpreadsheetApp.openById(DATA_SS_ID);
  var sh = ss.getSheetByName(SG_PALYA_SHEET);
  if (!sh) throw new Error('Sheet "' + SG_PALYA_SHEET + '" not found.');
  return sh;
}

function _rentSheet() {
  var ss = SpreadsheetApp.openById(DATA_SS_ID);
  var sh = ss.getSheetByName(FINAL_RENT_SHEET);
  if (!sh) throw new Error('Sheet "' + FINAL_RENT_SHEET + '" not found.');
  return sh;
}

function _propSheet() {
  var ss = SpreadsheetApp.openById(DATA_SS_ID);
  var sh = ss.getSheetByName(FINAL_PROP_SHEET);
  if (!sh) throw new Error('Sheet "' + FINAL_PROP_SHEET + '" not found.');
  return sh;
}

// ── GENERIC HELPERS ──────────────────────────────────────────────
function _getHeaders(sh) {
  var last = sh.getLastColumn();
  if (last < 1) return [];
  return sh.getRange(1, 1, 1, last).getValues()[0].map(function(h) {
    return String(h).trim();
  });
}

function _fmtVal(val) {
  if (val instanceof Date) {
    try {
      return Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    } catch(e) { return String(val); }
  }
  if (val === null || val === undefined) return '';
  return val;
}

// Header → key maps for each sheet
function _sgKey(h) {
  var map = {
    'timestamp':       'timestamp',
    'property':        'property',
    'house name':      'house_name',
    'actual location': 'actual_location',
    'message':         'message',
    'rent details':    'rent_details',
    'locality':        'locality',
    'food':            'food',
    'group name':      'group_name',
    'manager number':  'manager_number',
    'owner number':    'owner_number',
    'other details':   'other_details',
    'canned messages': 'canned_messages',
    'canned messages ':'canned_messages'
  };
  return map[h.toLowerCase().trim()] ||
    h.toLowerCase().trim().replace(/[^a-z0-9]+/g, '_').replace(/^_|_$/g, '');
}

function _rentKey(h) {
  var map = {
    'property code':             'property_code',
    'property name':             'property_name',
    'house name':                'house_name',
    'gender':                    'gender',
    'room type':                 'room_type',
    'sharing':                   'sharing',
    'monthly rent (₹)':          'monthly_rent',
    'monthly rent':              'monthly_rent',
    'deposit (₹)':               'deposit',
    'deposit':                   'deposit',
    'caution deposit (₹)':       'caution_deposit',
    'caution deposit':           'caution_deposit',
    'notice period':             'notice_period',
    'lock in period':            'lock_in_period',
    'maintenance amount':        'maintenance_amount',
    'maintenance post discount': 'maintenance_disc'
  };
  return map[h.toLowerCase().trim()] ||
    h.toLowerCase().trim().replace(/[^a-z0-9]+/g, '_').replace(/^_|_$/g, '');
}

function _propKey(h) {
  var map = {
    'property id':              'property_id',
    'property name':            'property_name',
    'house name':               'house_name',
    'manager contact number':   'manager_contact',
    'owner contact no':         'owner_contact',
    'location':                 'location',
    'nearby location':          'nearby_location',
    'gender':                   'gender',
    'furnishing':               'furnishing',
    'food':                     'food',
    'status':                   'status',
    'veg or non veg':           'veg_nonveg'
  };
  return map[h.toLowerCase().trim()] ||
    h.toLowerCase().trim().replace(/[^a-z0-9]+/g, '_').replace(/^_|_$/g, '');
}

// ================================================================
//  SG PALYA DETAILS — READ
// ================================================================
function getSgPalaData() {
  try {
    var sh      = _sgSheet();
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { status: 'success', data: [] };

    var rawH    = _getHeaders(sh);
    var numCols = rawH.length;
    var rows    = sh.getRange(2, 1, lastRow - 1, numCols).getValues();

    var data = [];
    rows.forEach(function(r, ri) {
      if (!r.some(function(c) { return c !== '' && c !== null && c !== undefined; })) return;
      var obj = { _row: ri + 2 }; // store actual sheet row for updates
      rawH.forEach(function(h, i) {
        obj[_sgKey(h)] = _fmtVal(r[i]);
      });
      data.push(obj);
    });

    return { status: 'success', data: data };
  } catch(e) {
    Logger.log('getSgPalaData error: ' + e.message);
    return { status: 'error', message: e.message };
  }
}

// ================================================================
//  FINAL PROPERTY — READ
// ================================================================
function getFinalProperty() {
  try {
    var sh      = _propSheet();
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { status: 'success', data: [] };

    var rawH    = _getHeaders(sh);
    var numCols = rawH.length;
    var rows    = sh.getRange(2, 1, lastRow - 1, numCols).getValues();

    var data = [];
    rows.forEach(function(r, ri) {
      if (!r.some(function(c) { return c !== '' && c !== null && c !== undefined; })) return;
      var obj = { _row: ri + 2 };
      rawH.forEach(function(h, i) {
        obj[_propKey(h)] = _fmtVal(r[i]);
      });
      data.push(obj);
    });

    return { status: 'success', data: data };
  } catch(e) {
    Logger.log('getFinalProperty error: ' + e.message);
    return { status: 'error', message: e.message };
  }
}

// ── GET SINGLE PROPERTY DETAIL (by property name) ────────────────
function getFinalPropertyByName(propertyName) {
  try {
    var res = getFinalProperty();
    if (res.status !== 'success') return res;
    var pn = String(propertyName || '').trim().toLowerCase();
    // exact match first
    var match = res.data.filter(function(r) {
      return String(r.property_name || '').trim().toLowerCase() === pn;
    });
    // fuzzy fallback — one contains the other
    if (!match.length) {
      match = res.data.filter(function(r) {
        var rn = String(r.property_name || '').trim().toLowerCase();
        return rn.includes(pn) || pn.includes(rn);
      });
    }
    return { status: 'success', data: match };
  } catch(e) {
    return { status: 'error', message: e.message };
  }
}

// ================================================================
//  FINAL RENT — READ
// ================================================================
function getFinalRent() {
  try {
    var sh      = _rentSheet();
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { status: 'success', data: [] };

    var rawH    = _getHeaders(sh);
    var numCols = rawH.length;
    var rows    = sh.getRange(2, 1, lastRow - 1, numCols).getValues();

    // Fill down property_name and house_name (merged-cell style blanks)
    var lastPropName = '', lastHouseName = '', lastPropCode = '';
    var propNameIdx  = -1, houseNameIdx = -1, propCodeIdx = -1;
    rawH.forEach(function(h, i) {
      var k = _rentKey(h);
      if (k === 'property_name') propNameIdx  = i;
      if (k === 'house_name')    houseNameIdx = i;
      if (k === 'property_code') propCodeIdx  = i;
    });

    var data = [];
    rows.forEach(function(r, ri) {
      if (!r.some(function(c) { return c !== '' && c !== null && c !== undefined; })) return;

      var obj = { _row: ri + 2 };
      rawH.forEach(function(h, i) {
        obj[_rentKey(h)] = _fmtVal(r[i]);
      });

      // Fill down
      if (propNameIdx  >= 0) {
        var pn = String(r[propNameIdx]  || '').trim();
        if (pn)  lastPropName  = pn; else obj.property_name  = lastPropName;
      }
      if (houseNameIdx >= 0) {
        var hn = String(r[houseNameIdx] || '').trim();
        if (hn) lastHouseName = hn; else obj.house_name = lastHouseName;
      }
      if (propCodeIdx  >= 0) {
        var pc = String(r[propCodeIdx]  || '').trim();
        if (pc)  lastPropCode  = pc; else obj.property_code  = lastPropCode;
      }

      data.push(obj);
    });

    return { status: 'success', data: data };
  } catch(e) {
    Logger.log('getFinalRent error: ' + e.message);
    return { status: 'error', message: e.message };
  }
}

// ── GET RENT ROWS FOR ONE PROPERTY ───────────────────────────────
function getFinalRentByProperty(propertyName) {
  try {
    var res = getFinalRent();
    if (res.status !== 'success') return res;
    var match = res.data.filter(function(r) {
      return String(r.property_name || '').trim().toLowerCase() ===
             String(propertyName || '').trim().toLowerCase();
    });
    return { status: 'success', data: match };
  } catch(e) {
    return { status: 'error', message: e.message };
  }
}

// ================================================================
//  SG PALYA DETAILS — ADD / UPDATE / DELETE
// ================================================================

// ADD — appends new row to Sg Palya Details
function addSgPalaRow(payload) {
  try {
    if (!payload.property || !payload.house_name)
      return { status: 'error', message: 'Property and House Name are required.' };

    var sh   = _sgSheet();
    var rawH = _getHeaders(sh);
    var newRow = rawH.map(function(h) {
      var k = _sgKey(h);
      return payload[k] !== undefined ? payload[k] : '';
    });
    sh.appendRow(newRow);
    SpreadsheetApp.flush();
    return { status: 'success', message: 'Row added to Sg Palya Details.' };
  } catch(e) {
    return { status: 'error', message: e.message };
  }
}

// UPDATE — matched by property + house_name
function updateSgPalaRow(payload) {
  try {
    var sh      = _sgSheet();
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { status: 'error', message: 'No data in Sg Palya Details.' };

    var rawH = _getHeaders(sh);
    var propIdx  = -1, houseIdx = -1;
    rawH.forEach(function(h, i) {
      var k = _sgKey(h);
      if (k === 'property')   propIdx  = i;
      if (k === 'house_name') houseIdx = i;
    });

    var rows     = sh.getRange(2, 1, lastRow - 1, rawH.length).getValues();
    var sheetRow = -1;

    for (var i = 0; i < rows.length; i++) {
      var rProp  = String(rows[i][propIdx]  || '').trim().toLowerCase();
      var rHouse = String(rows[i][houseIdx] || '').trim().toLowerCase();
      if (rProp  === String(payload.property   || '').trim().toLowerCase() &&
          rHouse === String(payload.house_name || '').trim().toLowerCase()) {
        sheetRow = i + 2; break;
      }
    }

    if (sheetRow === -1) return { status: 'error', message: 'Row not found in Sg Palya Details.' };

    rawH.forEach(function(h, idx) {
      var k = _sgKey(h);
      if (k === 'timestamp') return; // never overwrite timestamp
      if (payload[k] !== undefined && payload[k] !== null) {
        sh.getRange(sheetRow, idx + 1).setValue(payload[k]);
      }
    });

    SpreadsheetApp.flush();
    return { status: 'success', message: 'Sg Palya Details updated.' };
  } catch(e) {
    return { status: 'error', message: e.message };
  }
}

// DELETE — matched by property + house_name
function deleteSgPalaRow(payload) {
  try {
    var sh      = _sgSheet();
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { status: 'error', message: 'No data in Sg Palya Details.' };

    var rawH = _getHeaders(sh);
    var propIdx = -1, houseIdx = -1;
    rawH.forEach(function(h, i) {
      var k = _sgKey(h);
      if (k === 'property')   propIdx  = i;
      if (k === 'house_name') houseIdx = i;
    });

    var rows     = sh.getRange(2, 1, lastRow - 1, rawH.length).getValues();
    var sheetRow = -1;

    for (var i = 0; i < rows.length; i++) {
      var rProp  = String(rows[i][propIdx]  || '').trim().toLowerCase();
      var rHouse = String(rows[i][houseIdx] || '').trim().toLowerCase();
      if (rProp  === String(payload.property   || '').trim().toLowerCase() &&
          rHouse === String(payload.house_name || '').trim().toLowerCase()) {
        sheetRow = i + 2; break;
      }
    }

    if (sheetRow === -1) return { status: 'error', message: 'Row not found in Sg Palya Details.' };

    sh.deleteRow(sheetRow);
    SpreadsheetApp.flush();
    return { status: 'success', message: 'Deleted from Sg Palya Details.' };
  } catch(e) {
    return { status: 'error', message: e.message };
  }
}

// ================================================================
//  FINAL RENT — ADD / UPDATE / DELETE
// ================================================================

// ADD — appends new room-type row
function addFinalRentRow(payload) {
  try {
    if (!payload.property_name || !payload.room_type || !payload.monthly_rent)
      return { status: 'error', message: 'Property Name, Room Type and Monthly Rent are required.' };

    var sh   = _rentSheet();
    var rawH = _getHeaders(sh);
    var newRow = rawH.map(function(h) {
      var k = _rentKey(h);
      return payload[k] !== undefined ? payload[k] : '';
    });
    sh.appendRow(newRow);
    SpreadsheetApp.flush();
    return { status: 'success', message: 'Room type row added to Final Rent.' };
  } catch(e) {
    return { status: 'error', message: e.message };
  }
}

// UPDATE — matched by property_name + room_type + sharing
function updateFinalRentRow(payload) {
  try {
    var sh      = _rentSheet();
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { status: 'error', message: 'No data in Final Rent.' };

    var rawH = _getHeaders(sh);
    var propIdx = -1, rtIdx = -1, shIdx = -1;
    var lastPropName = '';
    rawH.forEach(function(h, i) {
      var k = _rentKey(h);
      if (k === 'property_name') propIdx = i;
      if (k === 'room_type')     rtIdx   = i;
      if (k === 'sharing')       shIdx   = i;
    });

    var rows     = sh.getRange(2, 1, lastRow - 1, rawH.length).getValues();
    var sheetRow = -1;

    for (var i = 0; i < rows.length; i++) {
      var cellProp = String(rows[i][propIdx] || '').trim();
      if (cellProp !== '') lastPropName = cellProp;
      var effectiveProp = cellProp !== '' ? cellProp : lastPropName;

      var rRoom    = String(rows[i][rtIdx] || '').trim().toLowerCase();
      var rSharing = String(rows[i][shIdx] || '').trim().toLowerCase();

      if (effectiveProp.toLowerCase() === String(payload.property_name || '').trim().toLowerCase() &&
          rRoom    === String(payload.room_type || '').trim().toLowerCase() &&
          rSharing === String(payload.sharing   || '').trim().toLowerCase()) {
        sheetRow = i + 2; break;
      }
    }

    if (sheetRow === -1) return { status: 'error', message: 'Row not found in Final Rent.' };

    rawH.forEach(function(h, idx) {
      var k = _rentKey(h);
      if (k === 'property_name' || k === 'property_code') return;
      if (payload[k] !== undefined && payload[k] !== null) {
        sh.getRange(sheetRow, idx + 1).setValue(payload[k]);
      }
    });

    SpreadsheetApp.flush();
    return { status: 'success', message: 'Final Rent updated.' };
  } catch(e) {
    return { status: 'error', message: e.message };
  }
}

// DELETE — matched by property_name + room_type + sharing
function deleteFinalRentRow(payload) {
  try {
    var sh      = _rentSheet();
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { status: 'error', message: 'No data in Final Rent.' };

    var rawH = _getHeaders(sh);
    var propIdx = -1, rtIdx = -1, shIdx = -1;
    var lastPropName = '';
    rawH.forEach(function(h, i) {
      var k = _rentKey(h);
      if (k === 'property_name') propIdx = i;
      if (k === 'room_type')     rtIdx   = i;
      if (k === 'sharing')       shIdx   = i;
    });

    var rows     = sh.getRange(2, 1, lastRow - 1, rawH.length).getValues();
    var sheetRow = -1;

    for (var i = 0; i < rows.length; i++) {
      var cellProp = String(rows[i][propIdx] || '').trim();
      if (cellProp !== '') lastPropName = cellProp;
      var effectiveProp = cellProp !== '' ? cellProp : lastPropName;

      var rRoom    = String(rows[i][rtIdx] || '').trim().toLowerCase();
      var rSharing = String(rows[i][shIdx] || '').trim().toLowerCase();

      if (effectiveProp.toLowerCase() === String(payload.property_name || '').trim().toLowerCase() &&
          rRoom    === String(payload.room_type || '').trim().toLowerCase() &&
          rSharing === String(payload.sharing   || '').trim().toLowerCase()) {
        sheetRow = i + 2; break;
      }
    }

    if (sheetRow === -1) return { status: 'error', message: 'Row not found in Final Rent.' };

    sh.deleteRow(sheetRow);
    SpreadsheetApp.flush();
    return { status: 'success', message: 'Deleted from Final Rent.' };
  } catch(e) {
    return { status: 'error', message: e.message };
  }
}

// ================================================================
//  FINAL PROPERTY — ADD / UPDATE / DELETE
// ================================================================

// ADD — appends new property row
function addFinalPropertyRow(payload) {
  try {
    if (!payload.property_name)
      return { status: 'error', message: 'Property Name is required.' };

    var sh   = _propSheet();
    var rawH = _getHeaders(sh);
    var newRow = rawH.map(function(h) {
      var k = _propKey(h);
      return payload[k] !== undefined ? payload[k] : '';
    });
    sh.appendRow(newRow);
    SpreadsheetApp.flush();
    return { status: 'success', message: 'Property added to Final Property.' };
  } catch(e) {
    return { status: 'error', message: e.message };
  }
}

// UPDATE — matched by property_id OR property_name
function updateFinalPropertyRow(payload) {
  try {
    var sh      = _propSheet();
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { status: 'error', message: 'No data in Final Property.' };

    var rawH = _getHeaders(sh);
    var idIdx = -1, nameIdx = -1;
    rawH.forEach(function(h, i) {
      var k = _propKey(h);
      if (k === 'property_id')   idIdx   = i;
      if (k === 'property_name') nameIdx = i;
    });

    var rows     = sh.getRange(2, 1, lastRow - 1, rawH.length).getValues();
    var sheetRow = -1;

    for (var i = 0; i < rows.length; i++) {
      var rId   = String(rows[i][idIdx]   || '').trim().toUpperCase();
      var rName = String(rows[i][nameIdx] || '').trim().toLowerCase();

      var matchById   = payload.property_id   && rId   === String(payload.property_id   || '').trim().toUpperCase();
      var matchByName = payload.property_name && rName === String(payload.property_name || '').trim().toLowerCase();

      if (matchById || matchByName) {
        sheetRow = i + 2; break;
      }
    }

    if (sheetRow === -1) return { status: 'error', message: 'Row not found in Final Property.' };

    rawH.forEach(function(h, idx) {
      var k = _propKey(h);
      if (k === 'property_id') return; // never overwrite the ID
      if (payload[k] !== undefined && payload[k] !== null) {
        sh.getRange(sheetRow, idx + 1).setValue(payload[k]);
      }
    });

    SpreadsheetApp.flush();
    return { status: 'success', message: 'Final Property updated.' };
  } catch(e) {
    return { status: 'error', message: e.message };
  }
}

// DELETE — matched by property_id
function deleteFinalPropertyRow(payload) {
  try {
    var sh      = _propSheet();
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { status: 'error', message: 'No data in Final Property.' };

    var rawH = _getHeaders(sh);
    var idIdx = -1;
    rawH.forEach(function(h, i) {
      if (_propKey(h) === 'property_id') idIdx = i;
    });

    var rows     = sh.getRange(2, 1, lastRow - 1, rawH.length).getValues();
    var sheetRow = -1;

    for (var i = 0; i < rows.length; i++) {
      var rId = String(rows[i][idIdx] || '').trim().toUpperCase();
      if (rId === String(payload.property_id || '').trim().toUpperCase()) {
        sheetRow = i + 2; break;
      }
    }

    if (sheetRow === -1) return { status: 'error', message: 'Row not found in Final Property.' };

    sh.deleteRow(sheetRow);
    SpreadsheetApp.flush();
    return { status: 'success', message: 'Deleted from Final Property.' };
  } catch(e) {
    return { status: 'error', message: e.message };
  }
}

// ================================================================
//  DEBUG
// ================================================================
function debugPGSheets() {
  var ss = SpreadsheetApp.openById(DATA_SS_ID);
  Logger.log('All sheets: ' + ss.getSheets().map(function(s) { return s.getName(); }).join(', '));

  [SG_PALYA_SHEET, FINAL_RENT_SHEET, FINAL_PROP_SHEET].forEach(function(name) {
    try {
      var sh = ss.getSheetByName(name);
      if (!sh) { Logger.log(name + ': NOT FOUND'); return; }
      var h = _getHeaders(sh);
      Logger.log(name + ' headers (' + h.length + '): ' + JSON.stringify(h));
      Logger.log(name + ' last row: ' + sh.getLastRow());
    } catch(e) {
      Logger.log(name + ' error: ' + e.message);
    }
  });
}

function debugSgPalaData() {
  var r = getSgPalaData();
  Logger.log('Status: ' + r.status);
  if (r.status === 'error') { Logger.log('ERROR: ' + r.message); return; }
  Logger.log('Count: ' + r.data.length);
  if (r.data.length > 0) Logger.log('First row: ' + JSON.stringify(r.data[0]));
}

function debugFinalRent() {
  var r = getFinalRent();
  Logger.log('Status: ' + r.status);
  if (r.status === 'error') { Logger.log('ERROR: ' + r.message); return; }
  Logger.log('Count: ' + r.data.length);
  if (r.data.length > 0) Logger.log('First row: ' + JSON.stringify(r.data[0]));
}

function debugFinalProperty() {
  var r = getFinalProperty();
  Logger.log('Status: ' + r.status);
  if (r.status === 'error') { Logger.log('ERROR: ' + r.message); return; }
  Logger.log('Count: ' + r.data.length);
  if (r.data.length > 0) Logger.log('First row: ' + JSON.stringify(r.data[0]));
}