//const SHEET_ID = "12XqoWMiko6czUDFHWK411qLSVfBKAgBf5XppDCtNwtU";   // 第1处：改这里
const SHEET_ID = PropertiesService.getScriptProperties().getProperty('TICKET_SHEET_ID');
const SHEET_TICKETS = "Tickets";
const API_TOKEN = PropertiesService.getScriptProperties().getProperty('API_TOKEN');   // 第2处：改这里
const SHEET_4_LOG = PropertiesService.getScriptProperties().getProperty('SHEET_4_LOG');

function doGet() {
  return json_({ ok: true, service: "ticket-api" });
}

function doPost(e) {
  try {

    const raw = e?.postData?.contents || "";
    const body = raw ? JSON.parse(raw) : {};

    if (body.token !== API_TOKEN) {
      return json_({ ok: false, error: "Unauthorized" }, 401);
    }

    const email = (body.email || body.requesterEmail || "").trim();
    const title = (body.title || "").trim();
    const description = (body.description || "").trim();
    const urgency = (body.urgency || "Medium").trim();
    const impact = (body.impact || "Medium").trim();

    if (!email) return json_({ ok:false, error:"Missing: email" }, 400);
    if (!title) return json_({ ok:false, error:"Missing: title" }, 400);
    if (!description) return json_({ ok:false, error:"Missing: description" }, 400);


    const createdAt = new Date();
    const year = createdAt.getFullYear();
    const ticketId = `LIB-${year}-${Utilities.getUuid().slice(0,8).toUpperCase()}`;

    const priority = DKUTicketCore.calcPriority(impact, urgency);
    const dueAt = DKUTicketCore.calcDueAt(createdAt, priority);
    logEvent_(
      'INFO',
      'Start create ticket',
       email,
      `Try to create Ticket ID: ${ticketId || ''}`.trim()
      );

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName(SHEET_TICKETS);
    if (!sh) return json_({ ok:false, error:`Sheet not found: ${SHEET_TICKETS}` }, 500);

    // 18列，严格对齐你的表头（含 ClosedAt）
    // Build the row (keep alignment with your header)
    const row = [
      ticketId,           // TicketID
      createdAt,          // CreatedAt
      email,              // Email
      title,              // Title
      description,        // Description
      urgency,            // Urgency
      impact,             // Impact
      priority,           // Priority
      "New",              // Status
      "Systems Librarian",// Assignee
      dueAt,              // DueAt
      "",                 // FirstResponseAt
      "",                 // ResolvedAt
      "",                 // Notes
      "",                 // Resolution
      new Date(),         // Last updated
      "FALSE",            // IsOverdue
      ""                  // ClosedAt
    ];

    sh.appendRow(row);
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
