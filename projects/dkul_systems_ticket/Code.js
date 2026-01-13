/**
 * DKU Library Ticket System
 * 
 * STATUS: STABLE / FROZEN
 * Last verified working: 2025-12-13
 * Owner: Xueying Cheng (Systems Librarian)
 * 
 * WARNING:
 * - Do NOT modify unless ticket workflow breaks
 * - All triggers are installable and time-driven
 * - TicketID uses ScriptProperties (TICKET_SEQ)
 */


/***********************
 * CONFIG
 ***********************/
// ---- Core config ----
const TICKET_API_URL = PropertiesService.getScriptProperties().getProperty('TICKET_API_URL'); 
// e.g. https://script.google.com/macros/s/xxxxx/exec

const TICKET_API_TOKEN = PropertiesService.getScriptProperties().getProperty('TICKET_API_TOKEN');
// same token as lib_ticket_api expects
const SHEET_TICKETS = 'Tickets';
const SHEET_FORM = 'Form Responses 1';
const YOUR_DIGEST_EMAIL = 'xueying.cheng@dukekunshan.edu.cn'; // ← 改成你自己


/***********************
 * UTILS
 ***********************/
function col_(sheet, headerName) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idx = headers.indexOf(headerName);
  if (idx === -1) throw new Error(`Missing column header: ${headerName}`);
  return idx + 1;
}

function nextTicketSeq_() {
  const props = PropertiesService.getScriptProperties();
  const k = 'TICKET_SEQ';

  const cur = parseInt(props.getProperty(k) || '0', 10);
  const next = cur + 1;

  props.setProperty(k, String(next));
  return next;
}

/***********************
 * FORM SUBMIT → CREATE TICKET
 * Trigger: From spreadsheet → On form submit
 ***********************/
function onFormSubmit(e) {
  const nv = e && e.namedValues ? e.namedValues : null;
  if (!nv) throw new Error('onFormSubmit: missing event.namedValues');

  // 根据你的表单问题标题改 key（示例：Email/Title/Description/Urgency/Impact）
  const email = String((nv['Email'] || nv['email'] || [''])[0]).trim();
  const title = String((nv['Title'] || nv['title'] || [''])[0]).trim();
  const system = String((nv['Category'] || nv['Category'] || [''])[0]).trim();
  const description = String((nv['Description'] || nv['description'] || [''])[0]).trim();
  const urgency = String((nv['Urgency'] || ['Medium'])[0]).trim();
  const impact = String((nv['Impact'] || ['Medium'])[0]).trim();

  // 最小校验
  if (!email) throw new Error('Missing form field: Email');
  if (!title) throw new Error('Missing form field: Title');
  if (!description) throw new Error('Missing form field: Description');

  const finalTitle = `[${system}] ${title}`;
  const finalDescription = `System: ${system}\n${description}`;

  const res = createTicketViaApi_({
    email,
    title: finalTitle,
    description: finalDescription,
    urgency,
    impact
  });

  if (!(res && res.ok === true && res.ticketId)) {
    // 这里不要“悄悄失败”，一定要让 execution log 里看得出来
    throw new Error(`Ticket API failed: ${JSON.stringify(res)}`);
  }

  // 可选：把 ticketId 写回 Form Responses（若你有一个列叫 TicketID）
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_FORM);
    if (sh) {
      const lastRow = sh.getLastRow();
      const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      const idx = headers.indexOf('TicketID');
      if (idx !== -1) sh.getRange(lastRow, idx + 1).setValue(res.ticketId);
    }
  } catch (err) {
    // 写回失败不影响主流程
    console.warn('Failed to write TicketID back to form sheet:', String(err));
  }
}


/***********************
 * EDIT TICKET (STATUS / SLA / EMAIL)
 * Trigger: From spreadsheet → On edit (INSTALLABLE)
 ***********************/
function onTicketsEdit(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_TICKETS) return;
  if (e.range.getRow() === 1) return;

  const row = e.range.getRow();

  const cStatus = col_(sheet, 'Status');
  const cLast = col_(sheet, 'Last updated');
  const cFirst = col_(sheet, 'FirstResponseAt');
  const cRes = col_(sheet, 'ResolvedAt');
  const cClosed = col_(sheet, 'ClosedAt');

  // Only act on Status change
  if (e.range.getColumn() !== cStatus) return;

  const newStatus = e.value;
  const oldStatus = e.oldValue;
  if (!newStatus || newStatus === oldStatus) return;
  if (typeof newStatus !== 'string') return;

  if (!DKUTicketCore.isValidStatus(newStatus)) {
    sheet.getRange(row, cStatus).setValue(oldStatus || '');
    console.warn(`Invalid status: ${newStatus}`);
    return;
  }
  if (oldStatus && !DKUTicketCore.isValidTransition(oldStatus, newStatus)) {
    sheet.getRange(row, cStatus).setValue(oldStatus);
    console.warn(`Invalid status transition: ${oldStatus} -> ${newStatus}`);
    return;
  }

  // Always update Last updated on valid status change
  sheet.getRange(row, cLast).setValue(new Date());

  // First response time
  const firstCell = sheet.getRange(row, cFirst);
  if (newStatus === 'In Progress' && !firstCell.getValue()) {
    firstCell.setValue(new Date());
  }

  // Resolved time
  const resCell = sheet.getRange(row, cRes);
  if (newStatus === 'Resolved') {
    resCell.setValue(new Date());
  }

  // Closed time
  const closedCell = sheet.getRange(row, cClosed);
  if (newStatus === 'Closed') {
    closedCell.setValue(new Date());
  }

  // Update overdue flag
  updateOverdueFlag_(sheet, row);

  // Email user
  const ticketId = sheet.getRange(row, col_(sheet, 'TicketID')).getValue();
  const email = sheet.getRange(row, col_(sheet, 'Email')).getValue();
  const title = sheet.getRange(row, col_(sheet, 'Title')).getValue();
  const description = sheet.getRange(row, col_(sheet, 'Description')).getValue();

  if (!ticketId || !email) return;

  sendStatusUpdateEmail_(email, ticketId, title, description, oldStatus, newStatus);
}

/***********************
 * OVERDUE FLAG
 ***********************/
function updateOverdueFlag_(sheet, row) {
  const dueAt = sheet.getRange(row, col_(sheet, 'DueAt')).getValue();
  const status = sheet.getRange(row, col_(sheet, 'Status')).getValue();
  const now = new Date();

  const done = (status === 'Resolved' || status === 'Closed');
  let overdue = false;

  if (!done && dueAt instanceof Date && !isNaN(dueAt.getTime())) {
    overdue = dueAt.getTime() < now.getTime();
  }

  sheet.getRange(row, col_(sheet, 'IsOverdue')).setValue(overdue ? 'TRUE' : 'FALSE');
}

/***********************
 * STATUS UPDATE EMAIL
 ***********************/
function sendStatusUpdateEmail_(to, ticketId, title, description, oldStatus, newStatus) {
  const subject = `[DKU Library Systems Ticket (#${ticketId})] Status updated: ${newStatus}`;
  const body = `
Your library support request has been updated.

Ticket ID: ${ticketId}
Title: ${title}
Description: ${description}

Previous status: ${oldStatus || 'N/A'}
Current status: ${newStatus}

${DKUTicketCore.statusExplanation(newStatus)}

You may reply to this email if you have additional information.

— DKU Library Systems
`;

  MailApp.sendEmail(to, subject, body);
}

/***********************
 * DAILY DIGEST (E1)
 * Trigger: Time-driven → Daily
 ***********************/
function refreshAllOverdueFlags_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TICKETS);
  const lastRow = sheet.getLastRow();
  for (let r = 2; r <= lastRow; r++) {
    updateOverdueFlag_(sheet, r);
  }
}

function dailyTicketDigest() {
  refreshAllOverdueFlags_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TICKETS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const now = new Date();
  const dueSoonDays = 2;
  const waitingDays = 3;

  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  const cId = col_(sheet, 'TicketID') - 1;
  const cTitle = col_(sheet, 'Title') - 1;
  const cStatus = col_(sheet, 'Status') - 1;
  const cPri = col_(sheet, 'Priority') - 1;
  const cDue = col_(sheet, 'DueAt') - 1;
  const cLast = col_(sheet, 'Last updated') - 1;
  const cOver = col_(sheet, 'IsOverdue') - 1;

  const overdue = [];
  const dueSoon = [];
  const waiting = [];

  for (const r of data) {
    const ticketId = r[cId];
    const title = r[cTitle];
    const status = r[cStatus];
    const pri = r[cPri];
    const dueAt = r[cDue];
    const lastUpd = r[cLast];
    const isOver = String(r[cOver]).toUpperCase() === 'TRUE';

    if (!ticketId) continue;
    if (status === 'Resolved' || status === 'Closed') continue;

    if (isOver) {
      overdue.push(`${ticketId} [${pri}] ${title} (${status})`);
      continue;
    }

    if (dueAt instanceof Date && !isNaN(dueAt.getTime())) {
      const diff = Math.ceil((dueAt - now) / 86400000);
      if (diff >= 0 && diff <= dueSoonDays) {
        dueSoon.push(`${ticketId} [${pri}] ${title} (due in ${diff}d)`);
      }
    }

    if (status.startsWith('Waiting') && lastUpd instanceof Date) {
      const idle = Math.floor((now - lastUpd) / 86400000);
      if (idle >= waitingDays) {
        waiting.push(`${ticketId} [${pri}] ${title} (idle ${idle}d)`);
      }
    }
  }

  if (!overdue.length && !dueSoon.length && !waiting.length) return;

  const subject = `[Tickets] Daily digest — ${Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd')}`;
  const body = `
Overdue (${overdue.length})
${overdue.join('\n') || 'None'}

Due soon (${dueSoon.length})
${dueSoon.join('\n') || 'None'}

Waiting too long (${waiting.length})
${waiting.join('\n') || 'None'}
`;

  MailApp.sendEmail(YOUR_DIGEST_EMAIL, subject, body);
}


// ===== Auto-close config =====
const AUTO_CLOSE_DAYS = 5; // ← N 天，按需改

function autoCloseResolvedTickets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TICKETS);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const cStatus = col_(sheet, 'Status');
  const cResAt  = col_(sheet, 'ResolvedAt');
  const cLast   = col_(sheet, 'Last updated');
  const cClosed = col_(sheet, 'ClosedAt');
  const cEmail  = col_(sheet, 'Email');
  const cId     = col_(sheet, 'TicketID');
  const cTitle  = col_(sheet, 'Title');
  const cDesc   = col_(sheet, 'Description');
  const cOver   = col_(sheet, 'IsOverdue');
  const cNotes  = col_(sheet, 'Notes');

  const now = new Date();
  const cutoffMs = AUTO_CLOSE_DAYS * 24 * 3600 * 1000;

  // 读整行数据
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  for (let i = 0; i < data.length; i++) {
    const row = i + 2;

    const status = data[i][cStatus - 1];
    if (status !== 'Resolved') continue;

    const resolvedAt = data[i][cResAt - 1];
    const lastUpdated = data[i][cLast - 1];

    // 需要 ResolvedAt 存在；且从 ResolvedAt 起超过 N 天
    if (!(resolvedAt instanceof Date) || isNaN(resolvedAt.getTime())) continue;
    if (now.getTime() - resolvedAt.getTime() < cutoffMs) continue;

    // “无回复/无动作”的代理：Last updated 也超过 N 天
    //（防止你最近还在改 Notes，但忘了改状态）
    if (lastUpdated instanceof Date && !isNaN(lastUpdated.getTime())) {
      if (now.getTime() - lastUpdated.getTime() < cutoffMs) continue;
    }

    // 执行 Close
    // 1) 先写 Notes
    const oldNotes = sheet.getRange(row, cNotes).getValue();
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const add = `[${stamp}] Auto-closed after ${AUTO_CLOSE_DAYS} days with no further updates.`;
    sheet.getRange(row, cNotes).setValue(oldNotes ? `${oldNotes}\n${add}` : add);

    // 2) 再改状态与时间戳
    sheet.getRange(row, cStatus).setValue('Closed');
    sheet.getRange(row, cClosed).setValue(new Date());
    sheet.getRange(row, cLast).setValue(new Date());
    sheet.getRange(row, cOver).setValue('FALSE');

    const ticketId = data[i][cId - 1];
    const email = data[i][cEmail - 1];
    const title = data[i][cTitle - 1];
    const desc = data[i][cDesc - 1];

    if (email && ticketId) {
      sendAutoCloseEmail_(email, ticketId, title, desc, AUTO_CLOSE_DAYS);
    }
  }
}

function sendAutoCloseEmail_(to, ticketId, title, description, days) {
  const subject = `[DKU Library Systems Ticket (#${ticketId})] Ticket closed after ${days} days`;
  const body = `
We’re closing this ticket because it has been marked as Resolved and we did not receive further updates for ${days} days.

Ticket ID: ${ticketId}
Title: ${title}
Description: ${description}

If the issue persists or you need further help, please submit a new ticket (and reference this Ticket ID if helpful).

— DKU Library Systems
`;
  MailApp.sendEmail(to, subject, body);
}


function createTicketViaApi_(payload) {
  if (!TICKET_API_URL) throw new Error('Missing Script Property: TICKET_API_URL');
  if (!TICKET_API_TOKEN) throw new Error('Missing Script Property: TICKET_API_TOKEN');

  const body = {
    token: TICKET_API_TOKEN,
    requestId: payload.requestId || Utilities.getUuid(),
    email: payload.email,
    title: payload.title,
    description: payload.description,
    urgency: payload.urgency || 'Medium',
    impact: payload.impact || 'Medium'
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

  // attach http code for debugging
  return { http: resp.getResponseCode(), ...json };
}
