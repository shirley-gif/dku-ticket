/***********************
 * CONFIG
 ***********************/
const TICKET_TAB = 'Tickets';
const LOG_TAB = 'Ticket_Conversation_Log';

const LABEL_REPLIES = 'DKU-Tickets/Replies';
const LABEL_PROCESSED = 'DKU-Tickets/Processed';
const LABEL_UNMATCHED = 'DKU-Tickets/Unmatched';

const NOTES_MAX_CHARS = 100;
// 从 subject 中提取 TicketID：支持 "(#LIB-...)" 或 "#LIB-..."
const TICKET_ID_REGEX = /#(LIB-[A-Za-z0-9-]+)/;

/***********************
 * ENTRYPOINT
 * Create a time-driven trigger to run this every 10 minutes.
 ***********************/
function ingestTicketReplies() {
  ensureLabels_();
  const ss = openTicketSpreadsheet_();
  const ticketSheet = ss.getSheetByName(TICKET_TAB);
  if (!ticketSheet) throw new Error(`Missing sheet tab: ${TICKET_TAB}`);

  const logSheet = ensureLogSheet_(ss);

  const cols = getHeaderMap_(ticketSheet);
  // 必须存在 TicketID / Notes / Last updated
  requireHeaders_(cols, ['TicketID', 'Notes', 'Last updated']);

  const processedLabel = GmailApp.getUserLabelByName(LABEL_PROCESSED);
  const unmatchedLabel = GmailApp.getUserLabelByName(LABEL_UNMATCHED);

  // 只扫描 Replies 标签下的 thread（Gmail 搜索不支持层级 label 的 regex，但 label: 精确）
  const threads = GmailApp.search(`label:"${LABEL_REPLIES}" -label:"${LABEL_PROCESSED}"`, 0, 50);
  //const threads = GmailApp.search( 'label:Replies -label:Processed', 0, 50);

  if (!threads.length) return { ok: true, scannedThreads: 0, ingested: 0 };

  // 预加载 Tickets 数据（用于 TicketID -> rowIndex 映射）
  const ticketIndex = buildTicketIdIndex_(ticketSheet, cols['TicketID']);

  // 预加载 Log 的 messageId 集合（去重）
  const seenMessageIds = loadSeenMessageIds_(logSheet);

  let ingested = 0;

  for (const thread of threads) {
    const messages = thread.getMessages();

    for (const msg of messages) {
      const messageId = safeGetMessageId_(msg);
      if (!messageId) continue;
      if (seenMessageIds.has(messageId)) continue; // 去重

      const subject = msg.getSubject() || '';
      const from = msg.getFrom() || '';
      const date = msg.getDate(); // Date object
      const snippet = (msg.getPlainBody ? msg.getPlainBody() : msg.getBody()).slice(0, 1000); // 防止超大
  
      const ticketId = extractTicketId_(subject);

      if (!ticketId) {
        // 无法提取 TicketID：标记 Unmatched
        try { thread.addLabel(unmatchedLabel); } catch (e) {}
        appendLogRow_(logSheet, {
          TicketID: '',
          MessageTime: date,
          Direction: 'inbound',
          From: from,
          Subject: subject,
          Snippet: truncate_(snippet, 100),
          MessageId: messageId,
          ThreadId: thread.getId()
        });
        seenMessageIds.add(messageId);
        ingested++;
        continue;
      }

      const rowIndex = ticketIndex.get(ticketId); // 1-based row index
      if (!rowIndex) {
        // 有 TicketID 但 Tickets 表找不到：标记 Unmatched
        try { thread.addLabel(unmatchedLabel); } catch (e) {}
        appendLogRow_(logSheet, {
          TicketID: ticketId,
          MessageTime: date,
          Direction: 'inbound',
          From: from,
          Subject: subject,
          Snippet: truncate_(snippet, 100),
          MessageId: messageId,
          ThreadId: thread.getId()
        });
        seenMessageIds.add(messageId);
        ingested++;
        continue;
      }

      // 1) 写入 Log（每封邮件一行）
      appendLogRow_(logSheet, {
        TicketID: ticketId,
        MessageTime: date,
        Direction: 'inbound',
        From: from,
        Subject: subject,
        Snippet: truncate_(snippet, 100),
        MessageId: messageId,
        ThreadId: thread.getId()
      });

      // 2) 更新 Tickets 主表：Notes + Last updated（只写摘要）
      const summary = buildLatestNote_(from, date, snippet);
      ticketSheet.getRange(rowIndex, cols['Notes']).setValue(truncate_(summary, NOTES_MAX_CHARS));
      ticketSheet.getRange(rowIndex, cols['Last updated']).setValue(new Date());

      seenMessageIds.add(messageId);
      ingested++;
    }

    // 该 thread 处理完，打 Processed 标签（避免重复扫描）
    try { thread.addLabel(processedLabel); } catch (e) {}
  }

  return { ok: true, scannedThreads: threads.length, ingested };
}

/***********************
 * HELPERS: Gmail
 ***********************/
function ensureLabels_() {
  // 若标签不存在则创建（可选，但建议）
  const need = [LABEL_REPLIES, LABEL_PROCESSED, LABEL_UNMATCHED];
  for (const name of need) {
    let lbl = GmailApp.getUserLabelByName(name);
    if (!lbl) GmailApp.createLabel(name);
  }
}

function extractTicketId_(subject) {
  const m = String(subject || '').match(TICKET_ID_REGEX);
  return m ? m[1] : null;
}

function safeGetMessageId_(msg) {
  // GmailMessage.getId() 存在且稳定，可用作去重
  try { return msg.getId(); } catch (e) { return null; }
}

/***********************
 * HELPERS: Sheets
 ***********************/
function openTicketSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('TICKET_SHEET_ID');
  if (!id) throw new Error('Missing Script Property: TICKET_SHEET_ID');
  return SpreadsheetApp.openById(id);
}
function ensureLogSheet_(ss) {
  let sh = ss.getSheetByName(LOG_TAB);
  if (!sh) {
    sh = ss.insertSheet(LOG_TAB);
    sh.getRange(1, 1, 1, 8).setValues([[
      'TicketID', 'MessageTime', 'Direction', 'From', 'Subject', 'Snippet', 'MessageId', 'ThreadId'
    ]]);
  } else {
    // 若已存在但没表头，尽量补齐（不强制）
    const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].filter(String);
    if (!header.length) {
      sh.getRange(1, 1, 1, 8).setValues([[
        'TicketID', 'MessageTime', 'Direction', 'From', 'Subject', 'Snippet', 'MessageId', 'ThreadId'
      ]]);
    }
  }
  return sh;
}

function getHeaderMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => {
    const key = String(h || '').trim();
    if (key) map[key] = i + 1; // 1-based
  });
  return map;
}

function requireHeaders_(map, required) {
  const missing = required.filter(k => !map[k]);
  if (missing.length) throw new Error('Missing required headers: ' + missing.join(', '));
}

function buildTicketIdIndex_(ticketSheet, ticketIdCol) {
  const lastRow = ticketSheet.getLastRow();
  const idx = new Map();
  if (lastRow < 2) return idx;

  const values = ticketSheet.getRange(2, ticketIdCol, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    const ticketId = String(values[i][0] || '').trim();
    if (!ticketId) continue;
    idx.set(ticketId, i + 2); // row number in sheet (since start at row 2)
  }
  return idx;
}

function loadSeenMessageIds_(logSheet) {
  const set = new Set();
  const lastRow = logSheet.getLastRow();
  if (lastRow < 2) return set;

  // MessageId 在第 7 列（按我们写的表头）
  const values = logSheet.getRange(2, 7, lastRow - 1, 1).getValues();
  for (const [id] of values) {
    if (id) set.add(String(id));
  }
  return set;
}

function appendLogRow_(logSheet, row) {
  logSheet.appendRow([
    row.TicketID || '',
    row.MessageTime || new Date(),
    row.Direction || 'inbound',
    row.From || '',
    row.Subject || '',
    row.Snippet || '',
    row.MessageId || '',
    row.ThreadId || ''
  ]);
}

/***********************
 * HELPERS: Text
 ***********************/
function truncate_(text, maxChars) {
  const s = String(text || '');
  if (s.length <= maxChars) return s;
  return s.slice(0, maxChars - 1) + '…';
}
function buildLatestNote_(from, date, body) {
  const ts = date instanceof Date ? date.toISOString().slice(0, 16).replace('T',' ') : '';
  const oneLine = String(body || '')
    .replace(/\r/g, '')
    .split('\n')
    .map(s => s.trim())
    .filter(Boolean)[0] || '';
  return `Email reply ${ts} from ${from}: ${oneLine}`;
}
