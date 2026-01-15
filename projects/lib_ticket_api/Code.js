//const SHEET_ID = "12XqoWMiko6czUDFHWK411qLSVfBKAgBf5XppDCtNwtU";   // 第1处：改这里
const SHEET_ID = PropertiesService.getScriptProperties().getProperty('TICKET_SHEET_ID');
const SHEET_TICKETS = "Tickets";
const API_TOKEN = PropertiesService.getScriptProperties().getProperty('API_TOKEN');   // 第2处：改这里
const SHEET_4_LOG = PropertiesService.getScriptProperties().getProperty('SHEET_4_LOG');
const ALLOWED_EMAIL_REGEX = /^[A-Za-z0-9._%+-]+@dukekunshan\.edu\.cn$/i;
const RATE_LIMIT_STATE_KEY = 'RATE_LIMIT_STATE';

function doGet() {
  return json_({ ok: true, service: "ticket-api" });
}

function doPost(e) {
  let email = '';
  let ticketId = '';
  let requestId = '';
  try {

    const raw = e?.postData?.contents || "";
    const body = raw ? JSON.parse(raw) : {};

    if (body.token !== API_TOKEN) {
      return json_({ ok: false, error: "Unauthorized" }, 401);
    }

    email = (body.email || body.requesterEmail || "").trim();
    const title = (body.title || "").trim();
    const description = (body.description || "").trim();
    const urgency = (body.urgency || "Medium").trim();
    const impact = (body.impact || "Medium").trim();
    requestId = String(body.requestId || '').trim();

    if (!email) return json_({ ok:false, error:"Missing: email" }, 400);
    if (!title) return json_({ ok:false, error:"Missing: title" }, 400);
    if (!description) return json_({ ok:false, error:"Missing: description" }, 400);
    if (!requestId) return json_({ ok:false, error:"Missing: requestId" }, 400);
    if (!ALLOWED_EMAIL_REGEX.test(email)) {
      return json_({ ok: false, error: "Email must be @dukekunshan.edu.cn" }, 400);
    }

    const abuseCheck = enforceAbuseControls_(email);
    if (!abuseCheck.ok) {
      return json_({ ok: false, error: abuseCheck.error }, abuseCheck.status || 403);
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName(SHEET_TICKETS);
    if (!sh) return json_({ ok:false, error:`Sheet not found: ${SHEET_TICKETS}` }, 500);

    const headers = getHeaderMap_(sh);
    requireHeaders_(headers, ['TicketID', 'RequestId', 'CreatedAt', 'Last updated']);

    const existingTicketId = findTicketIdByRequestId_(
      sh,
      headers['RequestId'],
      headers['TicketID'],
      requestId
    );
    if (existingTicketId) {
      logEvent_(
        'INFO',
        'REQUEST_ID_DEDUPED',
        email,
        `Existing Ticket ID: ${existingTicketId}`
      );
      return json_({ ok: true, ticketId: existingTicketId, deduped: true }, 200);
    }

    const createdAt = new Date();
    const year = createdAt.getFullYear();
    ticketId = `LIB-${year}-${Utilities.getUuid().slice(0,8).toUpperCase()}`;

    const priority = DKUTicketCore.calcPriority(impact, urgency);
    const dueAt = DKUTicketCore.calcDueAt(createdAt, priority);
    logEvent_(
      'INFO',
      'Start create ticket',
       email,
      `Try to create Ticket ID: ${ticketId || ''}`.trim()
      );

    const lastCol = sh.getLastColumn();
    const row = new Array(lastCol).fill('');
    const valuesByHeader = {
      'TicketID': ticketId,
      'CreatedAt': createdAt,
      'Email': email,
      'Title': title,
      'Description': description,
      'Urgency': urgency,
      'Impact': impact,
      'Priority': priority,
      'Status': 'New',
      'Assignee': 'Systems Librarian',
      'DueAt': dueAt,
      'FirstResponseAt': '',
      'ResolvedAt': '',
      'Notes': '',
      'Resolution': '',
      'Last updated': new Date(),
      'IsOverdue': 'FALSE',
      'ClosedAt': '',
      'RequestId': requestId
    };

    Object.keys(valuesByHeader).forEach((key) => {
      const col = headers[key];
      if (col) row[col - 1] = valuesByHeader[key];
    });

    sh.getRange(sh.getLastRow() + 1, 1, 1, lastCol).setValues([row]);
    // Send confirmation email AFTER ticket is successfully created
    let emailSent = false;
    let emailError = "";

    logEvent_(
      'INFO',
      'ADD_NEW_TICKET',
      email,
      ticketId ? `Ticket ID: ${ticketId}` : null
    );

    try {
      sendTicketConfirmationEmail_(email, {
        ticketId,
        createdAt,
        email,
        title,
        description,
        urgency,
        impact,
        priority
      });
      emailSent = true;
      logEvent_(
            'INFO',
            'SEND_CONFIRMATION_EMAIL',
             email,
            `Backend sent confirmation email. Ticket ID: ${ticketId || ''}`.trim()
          );
    } catch (e2) {
      emailError = String(e2);
      // do not fail ticket creation if email sending fails
      logEvent_(
           'ERROR',
            'SEND_CONFIRMATION_EMAIL_FAIL',
            email,
            emailError ? String(emailError) : 'Backend did not confirm email sending.'
          );
    }

    return json_({ ok: true, ticketId, emailSent, emailError }, 200);
  } catch (err) {
      logEvent_(
      'ERROR',
      'FAILE_TO_ADD_NEW_TICKET',
      email,
      ticketId ? `Ticket ID: ${ticketId}` : null
    );
    return json_({ ok: false, error: String(err) }, 500);
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getHeaderMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => {
    const key = String(h || '').trim();
    if (key) map[key] = i + 1;
  });
  return map;
}

function requireHeaders_(map, required) {
  const missing = required.filter(k => !map[k]);
  if (missing.length) throw new Error('Missing required headers: ' + missing.join(', '));
}

function findTicketIdByRequestId_(sheet, requestIdCol, ticketIdCol, requestId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return '';
  const requestValues = sheet.getRange(2, requestIdCol, lastRow - 1, 1).getValues();
  const ticketValues = sheet.getRange(2, ticketIdCol, lastRow - 1, 1).getValues();
  for (let i = 0; i < requestValues.length; i++) {
    if (String(requestValues[i][0] || '').trim() === requestId) {
      return String(ticketValues[i][0] || '').trim();
    }
  }
  return '';
}

function enforceAbuseControls_(email) {
  const props = PropertiesService.getScriptProperties();
  const denyEmails = String(props.getProperty('DENYLIST_EMAILS') || '')
    .split(',')
    .map(s => s.trim().toLowerCase())
    .filter(Boolean);
  const denyPatterns = String(props.getProperty('DENYLIST_PATTERNS') || '')
    .split(',')
    .map(s => s.trim())
    .filter(Boolean);

  if (denyEmails.includes(String(email || '').toLowerCase())) {
    return { ok: false, status: 403, error: 'Email is not allowed.' };
  }

  for (const pattern of denyPatterns) {
    try {
      const re = new RegExp(pattern, 'i');
      if (re.test(email)) {
        return { ok: false, status: 403, error: 'Email is not allowed.' };
      }
    } catch (err) {
      console.warn(`Invalid denylist regex: ${pattern}`, err);
    }
  }

  const max = parseInt(props.getProperty('RATE_LIMIT_MAX') || '0', 10);
  const windowSec = parseInt(props.getProperty('RATE_LIMIT_WINDOW_SEC') || '0', 10);
  if (max > 0 && windowSec > 0) {
    const now = Date.now();
    const windowMs = windowSec * 1000;
    const stateRaw = props.getProperty(RATE_LIMIT_STATE_KEY);
    let state = {};
    if (stateRaw) {
      try { state = JSON.parse(stateRaw); } catch (e) { state = {}; }
    }
    const key = String(email || '').toLowerCase();
    const history = Array.isArray(state[key]) ? state[key] : [];
    const pruned = history.filter(ts => now - ts < windowMs);
    if (pruned.length >= max) {
      state[key] = pruned;
      state = cleanupRateLimitState_(state, now, windowMs);
      props.setProperty(RATE_LIMIT_STATE_KEY, JSON.stringify(state));
      return { ok: false, status: 429, error: 'Too many requests. Please try again later.' };
    }
    pruned.push(now);
    state[key] = pruned;
    state = cleanupRateLimitState_(state, now, windowMs);
    props.setProperty(RATE_LIMIT_STATE_KEY, JSON.stringify(state));
  }

  return { ok: true };
}

function cleanupRateLimitState_(state, now, windowMs) {
  const MAX_KEYS = 100;
  const TRIM_TO = 50;
  const clean = {};
  Object.keys(state || {}).forEach((emailKey) => {
    const history = Array.isArray(state[emailKey]) ? state[emailKey] : [];
    const pruned = history.filter(ts => typeof ts === 'number' && now - ts < windowMs);
    if (pruned.length) clean[emailKey] = pruned;
  });

  const keys = Object.keys(clean);
  if (keys.length > MAX_KEYS) {
    // 100/50 hysteresis to prevent boundary thrash and accidental over-trimming.
    keys.sort((a, b) => {
      const aLast = clean[a][clean[a].length - 1] || 0;
      const bLast = clean[b][clean[b].length - 1] || 0;
      return bLast - aLast;
    });
    const keep = new Set(keys.slice(0, TRIM_TO));
    keys.forEach((k) => {
      if (!keep.has(k)) delete clean[k];
    });
  }

  return clean;
}

function testRequestIdDeduplication() {
  const requestIds = [['abc'], ['def'], ['ghi']];
  const ticketIds = [['LIB-2024-AAAA'], ['LIB-2024-BBBB'], ['LIB-2024-CCCC']];
  const found = findTicketIdByRequestIdFromValues_(requestIds, ticketIds, 'def');
  if (found !== 'LIB-2024-BBBB') {
    throw new Error('RequestId deduplication failed.');
  }
  const missing = findTicketIdByRequestIdFromValues_(requestIds, ticketIds, 'zzz');
  if (missing) {
    throw new Error('Expected no ticket for unknown requestId.');
  }
  return { ok: true };
}

function findTicketIdByRequestIdFromValues_(requestValues, ticketValues, requestId) {
  for (let i = 0; i < requestValues.length; i++) {
    if (String(requestValues[i][0] || '').trim() === requestId) {
      return String(ticketValues[i][0] || '').trim();
    }
  }
  return '';
}

function sendTicketConfirmationEmail_(toEmail, t) {
  const subject = `[DKU Library Systems Ticket (#${t.ticketId})]`;

  const body =
`Hello,

This is a confirmation that your ticket request has been submitted to DKU Library Systems.

Ticket summary
- Email: ${t.email}
- Title: ${t.title}
- Urgency: ${t.urgency}
- Impact: ${t.impact}
- Ticket ID: ${t.ticketId}

How to add more details
You can reply to this email with additional information. Attachments are welcome (screenshots, error messages, files, etc.). Please keep the ticket ID (#${t.ticketId}) in the subject line.

We will follow up if additional details are needed.

Regards,
DKU Library Systems`;

  // Use GmailApp if you want it to send from the Gmail account that owns the script
  GmailApp.sendEmail(toEmail, subject, body);
}

//通用日志函数（Code.gs）
function logEvent_(type, stage, email, message) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_4_LOG);
    const sheet = ss.getSheetByName('Logs') || ss.insertSheet('Logs');

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Timestamp', 'Type', 'Stage', 'Email', 'Message']);
    }

    sheet.appendRow([
      new Date(),
      type,          // INFO / WARN / ERROR
      stage,         // e.g. SEND_EMAIL
      email || '',
      message || ''
    ]);
  } catch (e) {
    // Last-resort: do not break main flow
    console.error('Logging failed', e);  //记录失败 
  }
}
