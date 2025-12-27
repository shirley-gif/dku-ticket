const TICKET_API_URL = "https://script.google.com/macros/s/AKfycbxtaevS9EFN6-7llqAMBaT2BNo18JFCByInzCI5RkyfqV3w4ERXyEctCYcu1Q_m40s/exec";

const SHEET_4_LOG = "1R9C5AaVvN6yJsFvK_zqlpZVZ1mbIwLg8g9nuKVidb8M"

const ALLOWED_SYSTEMS = ['Alma', 'Summon', 'AtoM', 'Archivematica', 'RFID', 'Other'];

// const ALLOWED_ENVS = ['Production', 'Sandbox', 'N/A'];

const ALLOWED_LEVELS = ['Low', 'Medium', 'High'];

// ---- Security: store token in Script Properties, not in code ----
// Set once in Apps Script: Project Settings -> Script properties
// Key: TICKET_API_TOKEN   Value: (your token)

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('DKU Library Ticket Chat');
}

function getToken_() {
  const token = PropertiesService.getScriptProperties().getProperty('TICKET_API_TOKEN');
  if (!token) throw new Error('Missing Script Property: TICKET_API_TOKEN');
  return token;
}

function cache_() {
  return CacheService.getUserCache();
}

function getSession_() {
  const raw = cache_().get('ticket_chat_session');
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}

function setSession_(obj) {
  cache_().put('ticket_chat_session', JSON.stringify(obj), 60 * 30); // 30 min
}

function clearSession_() {
  cache_().remove('ticket_chat_session');
}

function startChat(email) {
  const em = (email || '').trim();
  if (!isValidEmail_(em)) {
    return { ok: false, message: 'Email is required and must be a DKU address (name@dukekunshan.edu.cn).' };
  }

  const s = {
    step: 'ASK_TITLE',
    payload: {
      email: em,
      title: '',
      system: '',
      environment: '',
      description: '',
      urgency: 'Medium',
      impact: 'Medium'
    }
  };
  setSession_(s);

  return {
    ok: true,
    message: 'Please enter a one-sentence summary (Ticket Title, 5–120 characters). Example: “Summon record page should not display MARC 035.”'
  };
}

function isValidEmail_(email) {
  return /^[A-Za-z0-9._%+-]+@dukekunshan\.edu\.cn$/i.test(String(email || '').trim());
}

function countChars_(s) {
  // Unicode-safe-ish: count code points
  return Array.from(String(s || '')).length;
}

function pickLevel_(input) {
  const v = String(input || '').trim();
  const found = ALLOWED_LEVELS.find(x => x.toLowerCase() === v.toLowerCase());
  return found || null;
}

function pickOne_(input, allowedList) {
  const v = String(input || '').trim();
  const found = allowedList.find(x => x.toLowerCase() === v.toLowerCase());
  return found || null;
}

function validatePayload_(p) {
  const errors = [];

  if (!p.email || !isValidEmail_(p.email)) errors.push('Email is required and must be @dukekunshan.edu.cn');
  const titleLen = countChars_(p.title);
  if (!p.title || titleLen < 5 || titleLen > 120) errors.push('Title is required and must be 5–120 characters');

  if (!pickOne_(p.system, ALLOWED_SYSTEMS)) errors.push(`System must be one of: ${ALLOWED_SYSTEMS.join(' / ')}`);

  const descLen = countChars_(p.description);
  if (!p.description || descLen < 15) errors.push('Description is required and must be at least 15 characters');

  if (!pickLevel_(p.urgency)) errors.push(`Urgency must be one of: ${ALLOWED_LEVELS.join(' / ')}`);
  if (!pickLevel_(p.impact)) errors.push(`Impact must be one of: ${ALLOWED_LEVELS.join(' / ')}`);

  return errors;
}

//用于激发授权流程
function authorize() {
  // Force OAuth scopes prompt
  console.log(MailApp.getRemainingDailyQuota()); // triggers mail scope
  // SpreadsheetApp.getActiveSpreadsheet(); // triggers sheets scope (or open a sheet if needed)
  ss = SpreadsheetApp.openById(SHEET_4_LOG);
  const sheet = ss.getSheetByName('Logs') || ss.insertSheet('Logs');

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Timestamp', 'Type', 'Stage', 'Email', 'Message']);
  }

  sheet.appendRow([
    new Date(),
    'INFOR',          // INFO / WARN / ERROR
    'SEND_EMAIL',         // e.g. SEND_EMAIL
    'xc' || '',
    'test' || ''
  ]);

  return 'ok';
}


function chatTurn(input) {
  const text = (input || '').trim();
  if (!text) return { ok: false, message: 'Please enter a message.' };

  const cmd = text.toLowerCase();
  if (cmd === 'restart') { clearSession_(); return { ok: true, message: 'Restarted. Please enter your email again (or click Start).' }; }
  if (cmd === 'cancel')  { clearSession_(); return { ok: true, message: 'Cancelled. No ticket was created.' }; }

  let s = getSession_();
  if (!s) {
    return { ok: false, message: 'Session not found or expired. Please click Start to begin.' };
  }

  const p = s.payload;

  switch (s.step) {
    case 'ASK_TITLE': {
      const len = countChars_(text);
      if (len < 5 || len > 120) return { ok: false, message: 'Title must be 5–120 characters. Please re-enter.' };
      p.title = text;
      s.step = 'ASK_SYSTEM';
      setSession_(s);
      return { ok: true, message: `Which system is this about? Please enter one of: ${ALLOWED_SYSTEMS.join(' / ')}` };
    }

    case 'ASK_SYSTEM': {
      const v = pickOne_(text, ALLOWED_SYSTEMS);
      if (!v) return { ok: false, message: `System not recognized. It must be one of: ${ALLOWED_SYSTEMS.join(' / ')}` };
      p.system = v;
      s.step = 'ASK_DESC';
      setSession_(s);
      return { ok: true, message: 'Please provide details (Description, at least 15 characters): steps taken + expected result + actual result + any error message (if applicable).' };
    }

    case 'ASK_DESC': {
      const len = countChars_(text);
      if (len < 15) return { ok: false, message: `Description must be at least 15 characters (currently ${len}). Please add more details: steps + expected/actual + errors.` };
      p.description = text;
      s.step = 'ASK_URGENCY';
      setSession_(s);
      return { ok: true, message: `What is the urgency? Enter one of: ${ALLOWED_LEVELS.join(' / ')} (default: Medium).` };
    }

    case 'ASK_URGENCY': {
      const v = pickLevel_(text);
      if (!v) return { ok: false, message: `Urgency must be one of: ${ALLOWED_LEVELS.join(' / ')}` };
      p.urgency = v;
      s.step = 'ASK_IMPACT';
      setSession_(s);
      return { ok: true, message: `What is the impact? Enter one of: ${ALLOWED_LEVELS.join(' / ')}` };
    }

    case 'ASK_IMPACT': {
      const v = pickLevel_(text);
      if (!v) return { ok: false, message: `Impact must be one of: ${ALLOWED_LEVELS.join(' / ')}` };
      p.impact = v;
      s.step = 'CONFIRM';
      setSession_(s);
      return {
        ok: true,
        message: 'Please confirm the ticket below (type "confirm" to create; type "restart" to start over):\n\n' + renderSummary_(p)
      };
    }

    case 'CONFIRM':
      if (cmd !== 'confirm') {
        return { ok: false, message: 'Not created. Type "confirm" to create, "restart" to start over, or "cancel" to cancel.' };
      }

      // Validate again before submitting
      const errs = validatePayload_(p);
      if (errs.length) {
        return { ok: false, message: 'Validation failed:\n- ' + errs.join('\n- ') + '\n\nType "restart" to start over, or fix the inputs and then type "confirm" again.' };
      }

      const finalPayload = {
        email: p.email,
        title: `[${p.system}] ${p.title}`,
        description: `System: ${p.system}\nUrgency: ${p.urgency}\nImpact: ${p.impact}\n\n` + p.description,
        urgency: p.urgency,
        impact: p.impact
      };
      const res = createTicketFromWeb(finalPayload);

      // Send confirmation email only when API call succeeded
      try {
        sendTicketConfirmationEmail_(p.email, { ...finalPayload, system: p.system }, res);
      } catch (e) {
        const msg = e && e.message ? e.message : String(e);

        console.error('Failed to send confirmation email', msg);

        logEvent_(
          'ERROR',
          'SEND_CONFIRMATION_EMAIL',
          p.email,
          msg
        );
      }


      clearSession_();
      return { ok: true, message: 'Ticket creation request submitted. A confirmation email will be sent shortly.', ticket_result: res };


    default:
      clearSession_();
      return { ok: false, message: 'Session state error. Resetting. Please click Start to begin again.' };
  }
}

function ping() {
  return { ok: true, now: new Date().toISOString() };
}

function createTicketFromWeb(payload) {
  const body = {
    token: getToken_(),
    email: payload.email,
    title: payload.title,
    description: payload.description,
    urgency: payload.urgency,
    impact: payload.impact
  };

  const resp = UrlFetchApp.fetch(TICKET_API_URL, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  const text = resp.getContentText();
  let json;
  try { json = JSON.parse(text); } catch (e) { json = { ok: false, error: text }; }

  return { http: resp.getResponseCode(), ...json };
}

function normalizePick_(input, allowed) {
  const x = (input || '').trim().toLowerCase();
  if (!x) return null;
  // accept exact or common variants
  const map = { 'a2m': 'atom', 'at0m': 'atom' };
  const y = map[x] || x;
  if (allowed.includes(y)) return y;
  // accept case-insensitive matches like "Alma"
  const z = y.replace(/\s+/g,'');
  if (allowed.includes(z)) return z;
  return null;
}

function capitalize_(s) {
  s = String(s || '').toLowerCase();
  return s.charAt(0).toUpperCase() + s.slice(1);
}

function renderSummary_(p) {
  return [
    `Email: ${p.email || '(missing)'}`,
    `System: ${p.system || '(missing)'}`,
    // `Environment: ${p.environment || '(missing)'}`,
    `Title: ${p.title || '(missing)'}`,
    `Urgency: ${p.urgency || 'Medium'}`,
    `Impact: ${p.impact || 'Medium'}`,
    '',
    `Description:\n${p.description || ''}`
  ].join('\n');
}

//新增一个邮件函数
function sendTicketConfirmationEmail_(toEmail, finalPayload, ticketResult) {
  // Optional: if you want to avoid sending emails when API failed
  const ok = ticketResult && (ticketResult.ok === true || ticketResult.success === true);
  if (!ok) return;

  // Try to extract ticket id/url from API response (adjust keys based on your API)
  const ticketId = ticketResult.ticket_id || ticketResult.id || ticketResult.ticketId || '';
  const ticketUrl = ticketResult.url || ticketResult.ticket_url || ticketResult.link || '';

  const subject = `DKU Library Systems Ticket Received${ticketId ? ` (#${ticketId})` : ''}`;

  const lines = [
    `Hello,`,
    ``,
    `This is a confirmation that your ticket request has been submitted to DKU Library Systems.`,
    ``,
    `Summary`,
    `- Email: ${toEmail}`,
    `- System: ${finalPayload.system || ''}`,
    `- Title: ${finalPayload.title || ''}`,
    `- Urgency: ${finalPayload.urgency || ''}`,
    `- Impact: ${finalPayload.impact || ''}`,
    ticketId ? `- Ticket ID: ${ticketId}` : null,
    ticketUrl ? `- Ticket Link: ${ticketUrl}` : null,
    ``,
    `We will follow up if additional details are needed.`,
    ``,
    `Regards,`,
    `DKU Library Systems`
  ].filter(Boolean);

  MailApp.sendEmail({
    to: toEmail,
    subject: subject,
    body: lines.join('\n')
  });
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
    console.error('Logging failed', e);
  }
}

