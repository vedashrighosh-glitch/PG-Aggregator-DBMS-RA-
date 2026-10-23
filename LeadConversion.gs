// ================================================================
//  RoomAdda — LeadConversion.gs  (v4 FIXED)
//  FIX 1: _getLeadsLayout() now DEFINED here (was missing — runtime error)
//  FIX 2: CUST_COL_MAP defined here
//  FIX 3: Uses SS_ID global from Code.gs
// ================================================================

var CUST_COL_MAP = {
  'name'               : 'Guest Name',
  'phone'              : 'Mobile',
  'inc_no'             : 'Inc No',
  'date'               : 'Date',
  'room_num'           : 'Room NO',
  'pg_name'            : 'Location',
  'electricity'        : 'Electricity',
  'room_type'          : 'Sharing Type',
  'monthly_rent'       : 'Monthly Rent',
  'sd'                 : 'SD',
  'duration'           : 'Duration',
  'total_amc'          : 'TOTAL AMC',
  'roomadda_collection': 'RoomAdda Collection ',
  'paid'               : 'Paid',
  'commission'         : 'Commission',
  'move_in_date'       : 'check-in date',
  'customer_status'    : 'Check-in status',
  'payment_mode'       : 'Payment Mode',
  'bank_statement'     : 'Bank Statement (if any)',
  'notes'              : 'Comments'
};

// ── LAYOUT HELPER ───────────────────────────────────────────────
// Returns the headers and data-start-row for the Leads sheet.
// The Leads sheet has:
//   Row 1: Title
//   Row 2: Subtitle
//   Row 3: Empty
//   Row 4: HEADERS
//   Row 5+: DATA
function _getLeadsLayout(leadSh) {
  var HEADER_ROW  = 4;
  var DATA_START  = 5;
  var lastCol     = leadSh.getLastColumn();
  var rawHeaders  = leadSh.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0]
                          .map(function(h){ return String(h||'').trim(); });
  return { headers: rawHeaders, dataStartRow: DATA_START };
}

// ── CONVERT LEAD TO CUSTOMER ────────────────────────────────────
function convertLeadToCustomer(payload) {
  try {
    var ss     = SpreadsheetApp.openById(SS_ID);
    var leadSh = ss.getSheetByName(LEADS_TAB);
    var custSh = ss.getSheetByName(CUSTOMER_SHEET);
    if (!leadSh) return { status:'error', message:'Leads sheet "'+LEADS_TAB+'" not found.' };
    if (!custSh) return { status:'error', message:'Customers sheet "'+CUSTOMER_SHEET+'" not found.' };

    var layout  = _getLeadsLayout(leadSh);
    var headers = layout.headers;
    var normH   = headers.map(function(h){
      return h.toLowerCase().replace(/[^a-z0-9]+/g,'_').replace(/^_|_$/g,'');
    });
    var ni = normH.indexOf('name');  if (ni === -1) ni = 0;
    var pi = normH.indexOf('phone'); if (pi === -1) pi = 1;
    var lastRow = leadSh.getLastRow();
    if (lastRow < layout.dataStartRow) return { status:'error', message:'No leads found.' };

    var numRows     = lastRow - layout.dataStartRow + 1;
    var allLeadRows = leadSh.getRange(layout.dataStartRow, 1, numRows, headers.length).getValues();
    var leadRowIdx  = -1;
    var lead        = {};

    for (var i = 0; i < allLeadRows.length; i++) {
      var nm = String(allLeadRows[i][ni]).trim().toLowerCase();
      var ph = String(allLeadRows[i][pi]).replace(/[^0-9]/g,'');
      if (nm === String(payload.name).trim().toLowerCase() &&
          ph === String(payload.phone).replace(/[^0-9]/g,'')) {
        leadRowIdx = layout.dataStartRow + i;
        normH.forEach(function(h,j){ lead[h] = allLeadRows[i][j]; });
        break;
      }
    }
    if (leadRowIdx === -1) return { status:'error', message:'Lead "'+payload.name+'" not found.' };

    // Get customer sheet headers
    var custHeaders = custSh.getRange(1,1,1,custSh.getLastColumn())
                            .getValues()[0]
                            .map(function(h){ return String(h).trim(); });

    // Build reverse map: customer sheet header → CUST_COL_MAP key
    var reverseMap = {};
    for (var dbKey in CUST_COL_MAP) {
      reverseMap[CUST_COL_MAP[dbKey].toLowerCase().trim()] = dbKey;
    }

    // Duplicate check by phone
    var custPhoneColIdx = -1;
    custHeaders.forEach(function(h,i){
      var k = reverseMap[h.toLowerCase().trim()];
      if (k === 'phone') custPhoneColIdx = i;
    });

    var custLast = custSh.getLastRow();
    if (custLast >= 2 && custPhoneColIdx >= 0) {
      var existingPhones = custSh.getRange(2, custPhoneColIdx+1, custLast-1, 1)
                                 .getValues()
                                 .map(function(r){ return String(r[0]).replace(/[^0-9]/g,''); });
      var lp = String(lead.phone || payload.phone || '').replace(/[^0-9]/g,'');
      if (existingPhones.indexOf(lp) !== -1)
        return { status:'error', message:'This lead is already a customer.' };
    }

    // Build new customer row
    var today  = new Date();
    var newRow = custHeaders.map(function(h) {
      var k = reverseMap[h.toLowerCase().trim()];
      if (!k) return '';
      switch(k) {
        case 'name':            return lead.name            || payload.name   || '';
        case 'phone':           return lead.phone           || payload.phone  || '';
        case 'pg_name':         return lead.preferred_pg    || '';
        case 'room_type':       return lead.room_preference || '';
        case 'move_in_date':    return lead.expected_move_in|| '';
        case 'customer_status': return 'Active';
        case 'date':            return today;
        case 'created_at':      return today;
        default:                return '';
      }
    });
    custSh.appendRow(newRow);

    // Update lead status to Converted
    var lsIdx = normH.indexOf('lead_status');
    if (lsIdx === -1) lsIdx = normH.indexOf('status');
    var ctIdx = normH.indexOf('converted_to');
    if (lsIdx > -1)  leadSh.getRange(leadRowIdx, lsIdx+1).setValue('Converted');
    if (ctIdx > -1)  leadSh.getRange(leadRowIdx, ctIdx+1).setValue(true);
    SpreadsheetApp.flush();

    return { status:'success', message:(lead.name||payload.name)+' added to Customers.' };
  } catch(e) {
    return { status:'error', message: e.message };
  }
}