// ================================================================
//  RoomAdda — NotificationService.gs  (v4)
//  - 4-email sequence for visit allotment
//  - VisitLog sheet integration
// ================================================================

var TWILIO_ENABLED       = false;
var TWILIO_ACCOUNT_SID   = 'YOUR_TWILIO_SID';
var TWILIO_AUTH_TOKEN    = 'YOUR_TWILIO_AUTH_TOKEN';
var TWILIO_WHATSAPP_FROM = 'whatsapp:+14155238886';
var FOUNDER_EMAIL        = 'vanshbansal@roomadda.co.in';
var SHEET_VISIT_LOG      = 'VisitLog';

// ── LOG VISIT TO SHEET ──────────────────────────────────────────
function logVisitToSheet(lead, agentName, visitDateTime, allottedBy, officeNote) {
  try {
    var ss    = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName(SHEET_VISIT_LOG);

    // Create sheet with headers if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_VISIT_LOG);
      sheet.appendRow([
        'Visit ID','Lead Name','Phone','Agent Name','Allotted By',
        'Visit Date Time','Preferred PG','Preferred Area','Room Type',
        'Budget Min','Budget Max','Food Preference','Office Note',
        'Status','Outcome Note','Created At','Updated At'
      ]);
      sheet.getRange(1,1,1,17).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
    }

    var visitId = 'V-' + new Date().getTime();
    var now     = new Date();

    sheet.appendRow([
      visitId,
      lead.name        || '',
      lead.phone       || '',
      agentName        || '',
      allottedBy       || '',
      visitDateTime    || '',
      lead.preferred_pg   || '',
      lead.preferred_area || '',
      lead.room_preference || '',
      lead.budget_min  || '',
      lead.budget_max  || '',
      lead.food_preference || '',
      officeNote       || '',
      'Scheduled',
      '',
      now.toISOString(),
      now.toISOString()
    ]);

    return visitId;
  } catch(e) {
    Logger.log('logVisitToSheet: ' + e.message);
    return null;
  }
}

// ── UPDATE VISIT LOG STATUS ─────────────────────────────────────
function updateVisitLogStatus(visitId, status, outcomeNote) {
  try {
    var ss    = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName(SHEET_VISIT_LOG);
    if (!sheet) return;

    var data    = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h){ return String(h).trim().toLowerCase().replace(/\s/g,'_'); });
    var vidCol  = headers.indexOf('visit_id');
    var stCol   = headers.indexOf('status');
    var onCol   = headers.indexOf('outcome_note');
    var upCol   = headers.indexOf('updated_at');

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][vidCol]).trim() === visitId) {
        if (stCol  !== -1) sheet.getRange(i+1, stCol+1).setValue(status);
        if (onCol  !== -1) sheet.getRange(i+1, onCol+1).setValue(outcomeNote || '');
        if (upCol  !== -1) sheet.getRange(i+1, upCol+1).setValue(new Date().toISOString());
        SpreadsheetApp.flush();
        return;
      }
    }
  } catch(e) {
    Logger.log('updateVisitLogStatus: ' + e.message);
  }
}

// ── GET SALES TEAM ──────────────────────────────────────────────
function getSalesTeam() {
  try {
    var ss    = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName(SHEET_INTERNS);
    if (!sheet) return { status:'error', message:'Interns sheet not found' };

    var data    = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h){ return String(h).trim().toLowerCase(); });
    var ni=headers.indexOf('name'), ri=headers.indexOf('role');
    var di=headers.indexOf('department'), pi=headers.indexOf('phone');
    var ei=headers.indexOf('email'), aci=headers.indexOf('is_active');
    var ai=headers.indexOf('access_role');

    var team = [];
    for (var i = 1; i < data.length; i++) {
      var dept   = String(data[i][di]  || '').trim().toLowerCase();
      var access = String(data[i][ai]  || '').trim().toLowerCase();
      var active = String(data[i][aci] || '').trim().toUpperCase();
      var name   = String(data[i][ni]  || '').trim();
      var role   = String(data[i][ri]  || '').trim();
      if (active === 'YES' && name && (dept === 'sales' || access === 'sales')) {
        team.push({
          name:  name,
          phone: String(data[i][pi] || '').trim(),
          email: String(data[i][ei] || '').trim(),
          role:  role
        });
      }
    }
    return { status:'success', data:team };
  } catch(e) {
    return { status:'error', message:e.message };
  }
}

// ── FETCH AGENT LIVE ────────────────────────────────────────────
function _getLiveAgentDetails(agentName) {
  try {
    var ss    = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName(SHEET_INTERNS);
    if (!sheet) return null;
    var data    = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h){ return String(h).trim().toLowerCase(); });
    var ni=headers.indexOf('name'), ei=headers.indexOf('email');
    var pi=headers.indexOf('phone'), ri=headers.indexOf('role');
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][ni] || '').trim().toLowerCase() === agentName.trim().toLowerCase()) {
        return {
          name:  String(data[i][ni] || '').trim(),
          email: String(data[i][ei] || '').trim(),
          phone: String(data[i][pi] || '').trim(),
          role:  String(data[i][ri] || '').trim()
        };
      }
    }
    return null;
  } catch(e) {
    Logger.log('_getLiveAgentDetails: ' + e.message);
    return null;
  }
}

// ── MAIN: SEND ALLOTMENT + SCHEDULE ALL 4 EMAILS ───────────────
function sendAllotmentNotification(lead, agentName, agentPhone, agentEmail, officeNote, visitId) {
  try {
    var live       = _getLiveAgentDetails(agentName);
    var finalEmail = (live && live.email) ? live.email : agentEmail;
    var finalPhone = (live && live.phone) ? live.phone : agentPhone;

    if (!finalEmail) {
      Logger.log('No email for ' + agentName);
      return { success: false, error: 'No email found for agent' };
    }

    // EMAIL 1: Immediate allotment notification
    var msg = buildAllotmentMessage(lead, agentName, officeNote);
    MailApp.sendEmail({
      to:       finalEmail,
      bcc:      FOUNDER_EMAIL,
      subject:  '🏠 New Visit Allotted — ' + (lead.name || 'Customer') + ' | RoomAdda',
      body:     msg.plain,
      htmlBody: msg.html
    });
    Logger.log('Email 1 sent to ' + finalEmail);

    // Schedule emails 2, 3, 4
    if (lead.visit_date_time) {
      scheduleAllVisitReminders(lead, agentName, finalPhone, finalEmail, visitId);
    }

    if (TWILIO_ENABLED && finalPhone) {
      sendTwilioWhatsApp(finalPhone, msg.whatsapp);
    }

    return { success: true };
  } catch(e) {
    Logger.log('sendAllotmentNotification: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ── SCHEDULE ALL 3 REMAINING EMAILS ────────────────────────────
function scheduleAllVisitReminders(lead, agentName, agentPhone, agentEmail, visitId) {
  try {
    var visitTime = new Date(lead.visit_date_time);
    if (isNaN(visitTime.getTime())) {
      Logger.log('Invalid visit time: ' + lead.visit_date_time);
      return;
    }

    var props   = PropertiesService.getScriptProperties();
    var dataKey = 'visit_' + (visitId || new Date().getTime());

    props.setProperty(dataKey, JSON.stringify({
      lead:       lead,
      agentName:  agentName,
      agentPhone: agentPhone,
      agentEmail: agentEmail,
      visitId:    visitId,
      _key:       dataKey
    }));

    var now = new Date();

    function makeTrigger(fnName, fireAt) {
      if (fireAt <= now) {
        Logger.log('Skipping trigger ' + fnName + ' — time already passed');
        return;
      }
      var t = ScriptApp.newTrigger(fnName).timeBased().at(fireAt).create();
      props.setProperty('trig_' + t.getUniqueId(), dataKey);
      Logger.log('Trigger set: ' + fnName + ' at ' + fireAt);
    }

    // EMAIL 2: 24hrs before
    makeTrigger('send24hrReminder',   new Date(visitTime.getTime() - 24*60*60*1000));
    // EMAIL 3: 30mins before
    makeTrigger('send30minReminder',  new Date(visitTime.getTime() - 30*60*1000));
    // EMAIL 4: 1hr after (feedback)
    makeTrigger('sendPostVisitCheck', new Date(visitTime.getTime() + 60*60*1000));

  } catch(e) {
    Logger.log('scheduleAllVisitReminders: ' + e.message);
  }
}

// ── HELPER: GET LEAD DATA FROM TRIGGER ─────────────────────────
function _getLeadDataFromTrigger(e) {
  try {
    var props   = PropertiesService.getScriptProperties();
    var dataKey = props.getProperty('trig_' + e.triggerUid);
    if (!dataKey) return null;
    var raw = props.getProperty(dataKey);
    return raw ? JSON.parse(raw) : null;
  } catch(ex) {
    Logger.log('_getLeadDataFromTrigger: ' + ex.message);
    return null;
  }
}

// ── EMAIL 2: 24HR REMINDER ──────────────────────────────────────
function send24hrReminder(e) {
  try {
    var data = _getLeadDataFromTrigger(e);
    if (!data) { Logger.log('send24hrReminder: no data'); return; }

    var vs = _fmtDate(data.lead.visit_date_time);
    var budgetStr = _budgetStr(data.lead);

    var html = _baseEmailWrapper(
      '⏰ Visit Tomorrow — ' + data.lead.name,
      'Hi <strong>' + data.agentName + '</strong>, your visit is <strong>tomorrow</strong>!',
      data.lead, vs, budgetStr,
      '<div style="background:#1C3A1C;border-left:3px solid #00E5A0;padding:12px 16px;border-radius:0 8px 8px 0;margin-top:14px">' +
      '⏰ <strong style="color:#00E5A0">Reminder:</strong> Please confirm with the customer today and be ready by ' + vs + '</div>'
    );

    MailApp.sendEmail({
      to:       data.agentEmail,
      bcc:      FOUNDER_EMAIL,
      subject:  '⏰ Visit Tomorrow — ' + data.lead.name + ' | RoomAdda',
      body:     '24hr Reminder\n\nCustomer: ' + data.lead.name + '\nPhone: ' + data.lead.phone + '\nVisit: ' + vs,
      htmlBody: html
    });
    Logger.log('Email 2 (24hr) sent for ' + data.lead.name);

    if (TWILIO_ENABLED && data.agentPhone) {
      sendTwilioWhatsApp(data.agentPhone,
        '⏰ *Visit Tomorrow!*\nCustomer: ' + data.lead.name + '\nPhone: ' + data.lead.phone + '\nTime: ' + vs);
    }
  } catch(e) {
    Logger.log('send24hrReminder error: ' + e.message);
  }
}

// ── EMAIL 3: 30MIN REMINDER ─────────────────────────────────────
function send30minReminder(e) {
  try {
    var data = _getLeadDataFromTrigger(e);
    if (!data) { Logger.log('send30minReminder: no data'); return; }

    var vs = _fmtDate(data.lead.visit_date_time);
    var budgetStr = _budgetStr(data.lead);

    var html = _baseEmailWrapper(
      '🚨 Visit in 30 Minutes — ' + data.lead.name,
      'Hi <strong>' + data.agentName + '</strong>, your visit starts in <strong style="color:#FF5B1F">30 minutes!</strong>',
      data.lead, vs, budgetStr,
      '<div style="background:#3A1C1C;border-left:3px solid #FF4466;padding:12px 16px;border-radius:0 8px 8px 0;margin-top:14px">' +
      '🚨 <strong style="color:#FF4466">Action needed:</strong> Call the customer now to confirm you\'re on your way!</div>'
    );

    MailApp.sendEmail({
      to:       data.agentEmail,
      bcc:      FOUNDER_EMAIL,
      subject:  '🚨 Visit in 30 mins — ' + data.lead.name + ' | RoomAdda',
      body:     'Visit in 30 minutes!\nCustomer: ' + data.lead.name + '\nPhone: ' + data.lead.phone + '\nTime: ' + vs,
      htmlBody: html
    });
    Logger.log('Email 3 (30min) sent for ' + data.lead.name);

    if (TWILIO_ENABLED && data.agentPhone) {
      sendTwilioWhatsApp(data.agentPhone,
        '🚨 *Visit in 30 Minutes!*\nCustomer: ' + data.lead.name + '\n📞 Call now to confirm!');
    }
  } catch(e) {
    Logger.log('send30minReminder error: ' + e.message);
  }
}

// ── EMAIL 4: POST-VISIT FEEDBACK ────────────────────────────────
function sendPostVisitCheck(e) {
  try {
    var data = _getLeadDataFromTrigger(e);
    if (!data) { Logger.log('sendPostVisitCheck: no data'); return; }

    var html = '<div style="font-family:sans-serif;max-width:520px;padding:28px;background:#0A0A12;color:#F0F0F0;border-radius:12px">'
      + '<h2 style="color:#00E5A0;margin:0 0 6px">✅ How Did the Visit Go?</h2>'
      + '<p style="color:#888;margin:0 0 18px;font-size:13px">Post-Visit Feedback — RoomAdda</p>'
      + '<hr style="border:none;border-top:1px solid #2A2A3A;margin-bottom:18px"/>'
      + '<p>Hi <strong>' + data.agentName + '</strong>,</p>'
      + '<p style="color:#aaa;margin:12px 0">Your visit with <strong style="color:#fff">' + data.lead.name + '</strong> was scheduled an hour ago.</p>'
      + '<p style="margin:0 0 16px">Please update the outcome in the RoomAdda portal:</p>'
      + '<table style="width:100%;font-size:13px;border-collapse:collapse;background:#13151A;border-radius:8px;overflow:hidden;margin-bottom:18px">'
      + _r2('Customer', data.lead.name)
      + _r2('Phone', data.lead.phone)
      + _r2('PG', data.lead.preferred_pg || '—')
      + _r2('Area', data.lead.preferred_area || '—')
      + '</table>'
      + '<div style="background:#1C2A3A;border:1px solid #2A3A5A;border-radius:8px;padding:16px;margin-bottom:16px">'
      + '<p style="margin:0 0 10px;font-weight:700;color:#4A9EFF">📋 Please update in portal:</p>'
      + '<ul style="margin:0;padding-left:20px;color:#aaa;font-size:13px;line-height:2">'
      + '<li>✅ Visit Done — customer interested</li>'
      + '<li>📅 Postponed — new date agreed</li>'
      + '<li>❌ Not Interested — mark as lost</li>'
      + '<li>🚫 No Show — customer didn\'t turn up</li>'
      + '</ul></div>'
      + '<p style="font-size:11px;color:#555;margin-top:12px">Timely updates help the team track pipeline accurately. — RoomAdda</p>'
      + '</div>';

    MailApp.sendEmail({
      to:       data.agentEmail,
      bcc:      FOUNDER_EMAIL,
      subject:  '✅ How did the visit go? — ' + data.lead.name + ' | RoomAdda',
      body:     'Please update the visit outcome for ' + data.lead.name + ' in the RoomAdda portal.',
      htmlBody: html
    });
    Logger.log('Email 4 (post-visit) sent for ' + data.lead.name);

    if (TWILIO_ENABLED && data.agentPhone) {
      sendTwilioWhatsApp(data.agentPhone,
        '✅ *How did the visit go?*\nCustomer: ' + data.lead.name + '\nPlease update the outcome in the portal!');
    }

    // Cleanup trigger data
    var props = PropertiesService.getScriptProperties();
    if (data._key) props.deleteProperty(data._key);
    props.deleteProperty('trig_' + e.triggerUid);

  } catch(e) {
    Logger.log('sendPostVisitCheck error: ' + e.message);
  }
}

// ── EMAIL HTML BUILDER ──────────────────────────────────────────
function _baseEmailWrapper(heading, subheading, lead, vs, budgetStr, extraHtml) {
  return '<div style="font-family:sans-serif;max-width:520px;padding:28px;background:#0A0A12;color:#F0F0F0;border-radius:12px">'
    + '<h2 style="color:#FF6B2B;margin:0 0 6px">' + heading + '</h2>'
    + '<p style="color:#888;margin:0 0 18px;font-size:13px">RoomAdda Field Operations</p>'
    + '<hr style="border:none;border-top:1px solid #2A2A3A;margin-bottom:18px"/>'
    + '<p style="margin-bottom:16px">' + subheading + '</p>'
    + '<table style="width:100%;font-size:14px;border-collapse:collapse;background:#13151A;border-radius:8px;overflow:hidden">'
    + _r2('Name',   lead.name)
    + _r2('Phone',  '<a href="tel:' + lead.phone + '" style="color:#FF6B2B">' + lead.phone + '</a>')
    + _r2('Budget', budgetStr)
    + _r2('Room',   lead.room_preference)
    + _r2('Area',   lead.preferred_area)
    + _r2('PG',     lead.preferred_pg)
    + _r2('Visit',  '<strong style="color:#FF6B2B">' + vs + '</strong>')
    + '</table>'
    + (extraHtml || '')
    + '<p style="font-size:11px;color:#555;margin-top:16px">RoomAdda DBMS — automated notification</p>'
    + '</div>';
}

function buildAllotmentMessage(lead, agentName, officeNote) {
  var vs        = _fmtDate(lead.visit_date_time);
  var budgetStr = _budgetStr(lead);

  var plain = [
    '🏠 NEW VISIT ALLOTTED — RoomAdda', '─────────────────────────',
    'Hi ' + agentName + ',', '', 'A visit has been allotted to you:', '',
    'Name:       ' + (lead.name || '—'),
    'Phone:      ' + (lead.phone || '—'),
    'Budget:     ' + (budgetStr || '—'),
    'Room Type:  ' + (lead.room_preference || '—'),
    'Food Pref:  ' + (lead.food_preference || '—'),
    'Area:       ' + (lead.preferred_area || '—'),
    'PG:         ' + (lead.preferred_pg || '—'),
    'Visit Time: ' + vs, '',
    '📝 Note: ' + (officeNote || 'None'), '',
    'Please call the customer before the visit.', 'RoomAdda'
  ].join('\n');

  var extraHtml = officeNote
    ? '<div style="margin-top:14px;padding:12px 16px;background:#1C1F27;border-left:3px solid #FF6B2B;border-radius:0 8px 8px 0">'
      + '<strong>📝 Office Note:</strong><br/>' + officeNote + '</div>'
    : '';

  var html = _baseEmailWrapper(
    '🏠 New Visit Allotted — ' + (lead.name || 'Customer'),
    'Hi <strong>' + agentName + '</strong>, a visit has been allotted to you.',
    lead, vs, budgetStr, extraHtml
  );

  var whatsapp = '🏠 *New Visit Allotted — RoomAdda*\n\nHi ' + agentName + '!\n\n'
    + '*Name:* '   + (lead.name || '—') + '\n'
    + '*Phone:* '  + (lead.phone || '—') + '\n'
    + '*Budget:* ' + (budgetStr || '—') + '\n'
    + '*Room:* '   + (lead.room_preference || '—') + '\n'
    + '*Area:* '   + (lead.preferred_area || '—') + '\n'
    + '*PG:* '     + (lead.preferred_pg || '—') + '\n'
    + '*Visit:* '  + vs + '\n\n'
    + '📝 *Note:* ' + (officeNote || 'None') + '\n\n'
    + '_Call the customer before the visit._';

  return { plain: plain, html: html, whatsapp: whatsapp };
}

// ── UTILS ───────────────────────────────────────────────────────
function _fmtDate(dt) {
  if (!dt) return 'TBD';
  try {
    return Utilities.formatDate(new Date(dt), Session.getScriptTimeZone(), 'dd MMM yyyy, hh:mm a');
  } catch(ex) { return String(dt); }
}

function _budgetStr(lead) {
  if (lead.budget_min && lead.budget_max) return '₹' + lead.budget_min + ' – ₹' + lead.budget_max;
  if (lead.budget_max) return 'Up to ₹' + lead.budget_max;
  return '—';
}

function _r2(label, value) {
  return '<tr>'
    + '<td style="padding:9px 14px;color:#888;width:90px;border-bottom:1px solid #2A2A3A">' + label + '</td>'
    + '<td style="padding:9px 14px;color:#F0F0F0;border-bottom:1px solid #2A2A3A"><strong>' + (value || '—') + '</strong></td>'
    + '</tr>';
}

function sendTwilioWhatsApp(toPhone, message) {
  try {
    var clean = String(toPhone).replace(/[^0-9]/g, '');
    if (clean.length === 10) clean = '91' + clean;
    UrlFetchApp.fetch(
      'https://api.twilio.com/2010-04-01/Accounts/' + TWILIO_ACCOUNT_SID + '/Messages.json',
      {
        method:  'post',
        payload: { From: TWILIO_WHATSAPP_FROM, To: 'whatsapp:+' + clean, Body: message },
        headers: { Authorization: 'Basic ' + Utilities.base64Encode(TWILIO_ACCOUNT_SID + ':' + TWILIO_AUTH_TOKEN) },
        muteHttpExceptions: true
      }
    );
  } catch(e) { Logger.log('sendTwilioWhatsApp: ' + e.message); }
}