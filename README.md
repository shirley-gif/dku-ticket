# DKU Ticket Apps Script Monorepo

This repository contains four Google Apps Script projects:

- `projects/dku_ticket_web_chat`
- `projects/lib_ticket_api`
- `projects/dkul_systems_ticket`
- `projects/dku_ticket_core`

## Required Script Properties

Configure these in **Project Settings → Script Properties** for each Apps Script project.

### `lib_ticket_api`
- `TICKET_SHEET_ID` — Spreadsheet ID for the Tickets sheet.
- `API_TOKEN` — Shared API token for incoming requests.
- `SHEET_4_LOG` — Spreadsheet ID for the Logs sheet.
- `DENYLIST_EMAILS` — Comma-separated emails to block (can be empty string).
- `DENYLIST_PATTERNS` — Comma-separated regex patterns to block (can be empty string).
- `RATE_LIMIT_MAX` — Maximum requests per window (e.g., `5`).
- `RATE_LIMIT_WINDOW_SEC` — Rate limit window in seconds (e.g., `600`).

> Note: the API stores rate-limit state in Script Properties under `RATE_LIMIT_STATE`.

### `dku_ticket_web_chat`
- `TICKET_API_URL` — Web app URL for `lib_ticket_api`.
- `TICKET_API_TOKEN` — Token that matches `API_TOKEN` in `lib_ticket_api`.

### `dkul_systems_ticket`
- `TICKET_API_URL` — Web app URL for `lib_ticket_api`.
- `TICKET_API_TOKEN` — Token that matches `API_TOKEN` in `lib_ticket_api`.
- `TICKET_SEQ` — Optional: sequence counter for legacy workflows.

### `dku_ticket_core`
- No Script Properties required.

## Required Triggers (Manual Configuration)

### `lib_ticket_api`
- Time-driven trigger for `ingestTicketReplies` (recommended every 10 minutes).

### `dkul_systems_ticket`
- From spreadsheet → **On form submit**: `onFormSubmit`.
- From spreadsheet → **On edit (installable)**: `onTicketsEdit`.
- Time-driven trigger (daily): `dailyTicketDigest`.
- Time-driven trigger (daily): `autoCloseResolvedTickets`.

## Lightweight Regression Tests (Apps Script)

Run these manually from the Apps Script editor when needed:

- `DKUTicketCore.testStatusTransitionRules` (status transition validation).
- `testRequestIdDeduplication` in `lib_ticket_api` (requestId deduplication helper).
