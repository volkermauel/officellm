# Feature Specification: Phase 9 — Outlook Graph Integration

**Feature Branch**: `009-outlook-graph`

**Created**: 2026-05-16

**Status**: Draft

**Depends on**: Phase 4 (Outlook MVP) — basic Outlook tools must exist

**Input**: User description: "Implement deep Outlook integration via Microsoft Graph API for folder listing, email search, calendar access, shared mailboxes, and compose capabilities. All Graph calls proxied through the add-in via nested app authentication (NAA)."

## Architecture

The Office.js API only gives access to the currently selected item. To list folders, search emails, read calendar, and access shared mailboxes, we need Microsoft Graph API. The add-in obtains an access token via **nested app authentication (NAA)**, makes Graph calls client-side, and returns the data to the MCP server through the existing command/result pipeline.

```
Open WebUI                MCP Server (port 3000)              Outlook Add-in
    │                            │                                │
    │── tools/call ─────────────►│                                │
    │   { tool, params }         │── command ────────────────────►│
    │                            │   { graphApi, endpoint, ... }  │
    │                            │                                │── NAA token request
    │                            │                                │── Graph API call
    │                            │◄── result ─────────────────────│
    │◄── result ─────────────────│   { data, status }            │
    │                            │                                │
    │                            │◄── /instances/register ────────│  (on load)
    │                            │◄── /instances/:id/heartbeat ───│  (every 10s)
    │                            │►── /instances/:id/commands ────│  (poll every 2s)
    │                            │◄── /instances/:id/result ──────│  (after execution)
```

**Why proxy through add-in (not server-side)**:

- Auth stays in the user's context — no refresh token storage on the MCP server.
- NAA tokens are scoped to the signed-in user and automatically refreshed by the Office host.
- No additional Azure AD app registration or admin consent required beyond what the add-in already has.
- Works seamlessly with shared mailboxes the user already has delegate access to.

## User Scenarios & Testing

### User Story 1 — List mail folders with unread counts (Priority: P0)

A user asks "What folders do I have?" The LLM calls `outlook_list_folders` and receives all mail folders with unread counts and child folder hierarchy, giving them a complete overview of their mailbox structure without switching to Outlook.

**Why this priority**: Folder listing is the gateway to all mailbox operations. The LLM needs to understand the folder structure before it can list emails, search, or move items. It's also the simplest Graph call to validate the NAA proxy pipeline works end-to-end.

**Independent Test**: Call `outlook_list_folders` and verify the response includes all folders with correct unread counts and hierarchical nesting.

**Acceptance Scenarios**:

1. **Given** a mailbox with Inbox (3 unread), Sent Items (0 unread), and custom folder "Projects" (5 unread), **When** developer calls `outlook_list_folders`, **Then** the response includes all three folders with correct unread counts and display names
2. **Given** a folder "Inbox" with child folders "Filing/A", "Filing/B", **When** developer calls `outlook_list_folders`, **Then** child folders appear nested under their parent with correct hierarchy
3. **Given** a mailbox with no custom folders, **When** developer calls `outlook_list_folders`, **Then** the response includes at minimum the default folders: Inbox, Sent Items, Drafts, Trash, Junk Email

---

### User Story 2 — List recent emails in a folder (Priority: P0)

A user asks "Show me recent emails in my Inbox." The LLM calls `outlook_list_emails` with a folder path and receives a paginated list with subject, sender, date, and a preview snippet for each message.

**Why this priority**: Email listing is the most common Outlook operation. Users need to see what's in a folder before deciding what to read, move, or act on. Pagination keeps responses bounded.

**Independent Test**: Call `outlook_list_emails` with folder "Inbox", verify the response contains message summaries with correct fields and respects the page size limit.

**Acceptance Scenarios**:

1. **Given** an Inbox with 50 messages, **When** developer calls `outlook_list_emails` with folder "Inbox" and `pageSize: 25`, **Then** the response contains exactly 25 messages with subject, sender, date, and preview snippet, plus a `nextPageToken`
2. **Given** the first page was retrieved with a `nextPageToken`, **When** developer calls `outlook_list_emails` with that token, **Then** the response contains the next 25 messages in chronological order
3. **Given** an empty folder "Archive", **When** developer calls `outlook_list_emails`, **Then** the response contains an empty list with `totalItems: 0`

---

### User Story 3 — Search emails across all folders (Priority: P0)

A user asks "Find emails from Alice about Q4 budget." The LLM calls `outlook_search_emails` with a KQL query and receives matching emails from across all folders, ranked by relevance.

**Why this priority**: Search is the primary discovery mechanism when users don't know which folder an email is in. Full-text search across all folders is more powerful than any native Outlook quick filter and is essential for the LLM to find context.

**Independent Test**: Call `outlook_search_emails` with query "from:alice@contoso.com Q4 budget", verify results contain matching emails from multiple folders.

**Acceptance Scenarios**:

1. **Given** emails from alice@contoso.com about "Q4 budget" exist in Inbox and "Projects" folder, **When** developer calls `outlook_search_emails` with query "from:alice@contoso.com Q4 budget", **Then** results include emails from both folders ranked by relevance
2. **Given** a search returning 200 results, **When** developer calls `outlook_search_emails` with `pageSize: 25`, **Then** only 25 results are returned with a `nextPageToken` for the next page
3. **Given** a query matching no emails, **When** developer calls `outlook_search_emails`, **Then** the response contains an empty list with `totalItems: 0`

---

### User Story 4 — Read a specific email in full (Priority: P1)

A user asks "Show me the full email." The LLM calls `outlook_get_email` with a message ID and receives the complete body (text and HTML) plus attachment metadata (name, size, type) for each attachment.

**Why this priority**: After listing or searching, users need to read the full content. This is a natural second step after discovery. P1 because it depends on having a message ID from a prior list/search call.

**Independent Test**: Call `outlook_get_email` with a known message ID, verify the response contains the full body and attachment metadata.

**Acceptance Scenarios**:

1. **Given** a message with HTML body and 3 attachments (PDF, XLSX, DOCX), **When** developer calls `outlook_get_email`, **Then** the response includes both text and HTML body representations, plus attachment metadata (name, size, contentType) for all 3 attachments
2. **Given** a message with no attachments, **When** developer calls `outlook_get_email`, **Then** the response includes the body and an empty attachments array
3. **Given** a message with 50+ attachments, **When** developer calls `outlook_get_email`, **Then** the response includes metadata only (no content) for each attachment and a note: "Use outlook_download_attachment to retrieve individual attachment content"

---

### User Story 5 — List upcoming calendar events (Priority: P0)

A user asks "What's on my calendar today?" The LLM calls `outlook_list_calendar_events` and receives upcoming appointments with time, subject, location, and attendee summary for a configurable date range.

**Why this priority**: Calendar awareness is fundamental to productivity workflows. Knowing what's coming up is the most common calendar question. Default range of today to +7 days covers the most useful window.

**Independent Test**: Call `outlook_list_calendar_events` for today's range, verify events with correct times, subjects, and locations are returned.

**Acceptance Scenarios**:

1. **Given** 5 events today and 3 events in the next 3 days, **When** developer calls `outlook_list_calendar_events` with default range (today to +7 days), **Then** all 8 events are returned sorted by start time
2. **Given** a date range with no events, **When** developer calls `outlook_list_calendar_events`, **Then** the response contains an empty list
3. **Given** an event with 100+ attendees, **When** displayed in list view, **Then** the response includes only the organizer and `attendeeCount: 105` — full attendee list is available via `outlook_get_calendar_event`

---

### User Story 6 — Read calendar event details (Priority: P1)

A user asks "What are the details of my 2pm meeting?" The LLM calls `outlook_get_calendar_event` with an event ID and receives the full event body, complete attendee list with response status, location details, and any online meeting link.

**Why this priority**: Detail view supplements the list view with full information. P1 because it requires an event ID from a prior list call.

**Independent Test**: Call `outlook_get_calendar_event` with a known event ID, verify the response includes full details.

**Acceptance Scenarios**:

1. **Given** an event with 50 attendees, **When** developer calls `outlook_get_calendar_event`, **Then** the response includes all 50 attendees with their email, display name, and response status (accepted/declined/tentative/none)
2. **Given** a Teams meeting event, **When** developer calls `outlook_get_calendar_event`, **Then** the response includes the `onlineMeetingUrl` and `onlineMeetingProvider`
3. **Given** a recurring event, **When** developer calls `outlook_get_calendar_event`, **Then** the response includes `recurrence` pattern and the specific occurrence's start/end times

---

### User Story 7 — Compose a new email draft (Priority: P0)

A user asks "Draft an email to Bob about the project." The LLM calls `outlook_compose_email` which creates a draft in the Drafts folder. The draft is NEVER auto-sent — the user must explicitly review and approve sending.

**Why this priority**: Email composition is a core LLM productivity use case. The safety gate (draft-only, never auto-send) is critical and must be in place from the start.

**Independent Test**: Call `outlook_compose_email` with recipients, subject, and body, verify a draft is created in Drafts folder and the response confirms draft status.

**Acceptance Scenarios**:

1. **Given** recipient "bob@contoso.com", subject "Project Update", and body text, **When** developer calls `outlook_compose_email`, **Then** a draft is created in the Drafts folder with `isDraft: true` and the response includes the draft's message ID
2. **Given** HTML body content, **When** developer calls `outlook_compose_email`, **Then** the draft preserves HTML formatting
3. **Given** the draft was created, **When** the user opens their Drafts folder in Outlook, **Then** the draft is visible and editable

---

### User Story 8 — Move email to another folder (Priority: P1)

A user asks "Move this to Archive." The LLM calls `outlook_move_email` with an email ID and target folder name. The tool validates the target folder exists before moving.

**Why this priority**: Moving emails is a key organization action. P1 because it requires the user to have identified an email first (via list or search).

**Independent Test**: Call `outlook_move_email` with a message ID and folder name, verify the email moves to the target folder.

**Acceptance Scenarios**:

1. **Given** a message in Inbox and a target folder "Archive" that exists, **When** developer calls `outlook_move_email`, **Then** the message moves to Archive and the response confirms the new folder location
2. **Given** a target folder "NonExistent" that does not exist, **When** developer calls `outlook_move_email`, **Then** the tool returns an error: `errorCode: "FOLDER_NOT_FOUND"` with a suggestion to list folders first
3. **Given** a shared mailbox email where user has read-only access, **When** developer calls `outlook_move_email`, **Then** the tool returns an error: `errorCode: "INSUFFICIENT_PERMISSIONS"` with the permission level available

---

### User Story 9 — Toggle read/unread status (Priority: P2)

A user asks "Mark these as unread." The LLM calls `outlook_mark_read` with email IDs and the desired read/unread state.

**Why this priority**: Read/unread toggling is a minor but useful follow-up action. P2 because it's less critical than discovery, reading, or composing.

**Independent Test**: Call `outlook_mark_read` with message IDs and `isRead: false`, verify the messages are marked unread.

**Acceptance Scenarios**:

1. **Given** 3 read messages, **When** developer calls `outlook_mark_read` with `isRead: false`, **Then** all 3 messages are marked unread in Outlook
2. **Given** 3 unread messages, **When** developer calls `outlook_mark_read` with `isRead: true`, **Then** all 3 messages are marked read and unread counts update accordingly
3. **Given** a mix of read and unread messages, **When** developer calls `outlook_mark_read` with `isRead: true`, **Then** all specified messages are set to read regardless of prior state

---

### User Story 10 — Get master categories list (Priority: P2)

A user asks "What categories are available?" The LLM calls `outlook_get_categories` and receives all color-coded categories defined in the user's mailbox.

**Why this priority**: Categories support is useful for organization workflows but is secondary to core email/calendar operations. P2 as a nice-to-have.

**Independent Test**: Call `outlook_get_categories`, verify the response includes category names and color assignments.

**Acceptance Scenarios**:

1. **Given** a mailbox with 5 custom categories, **When** developer calls `outlook_get_categories`, **Then** the response lists all 5 categories with displayName and color
2. **Given** a mailbox with only default categories, **When** developer calls `outlook_get_categories`, **Then** the response includes the default Outlook categories (Red Category, Blue Category, etc.)

---

### User Story 11 — List shared mailboxes and delegates (Priority: P1)

A user asks "Do I have access to any shared mailboxes?" The LLM calls `outlook_get_shared_mailboxes` and receives a list of shared mailboxes with the user's permission level for each.

**Why this priority**: Shared mailbox support is essential for enterprise users who manage team or department mailboxes. P1 because it enables multi-mailbox workflows.

**Independent Test**: Call `outlook_get_shared_mailboxes`, verify the response lists accessible shared mailboxes with permissions.

**Acceptance Scenarios**:

1. **Given** the user has delegate access to "support@contoso.com" with Full Access, **When** developer calls `outlook_get_shared_mailboxes`, **Then** the response includes "support@contoso.com" with `permissionLevel: "FullAccess"`
2. **Given** the user has no shared mailbox access, **When** developer calls `outlook_get_shared_mailboxes`, **Then** the response contains an empty list
3. **Given** a shared mailbox with read-only access, **When** the response includes the mailbox, **Then** it notes `canSend: false, canMove: false`

---

### Edge Cases

- **Shared mailbox with read-only access**: Send/move operations fail gracefully with `errorCode: "INSUFFICIENT_PERMISSIONS"` indicating the available permission level.
- **Calendar event with large attendee list (100+)**: List view truncates to organizer + `attendeeCount`. Full detail view returns all attendees.
- **Email with 50+ attachments**: Return metadata only (name, size, type) — not content. Provide `outlook_download_attachment` tool for individual retrieval.
- **Offline scenario**: Graph API calls fail; return clear error `errorCode: "NETWORK_ERROR"` with message "Outlook is offline. Graph API requires an active connection."
- **Multi-account scenario**: NAA ensures the correct mailbox context is used based on the add-in's active account.
- **Rate limiting (429)**: Graph API returns 429; tool retries with exponential backoff up to 3 attempts, then returns `errorCode: "RATE_LIMITED"` with retry-after hint.
- **Token expiration during long operation**: NAA silently refreshes the token; if refresh fails, return `errorCode: "AUTH_EXPIRED"` asking the user to restart the add-in.
- **Folder not found during move**: Return `errorCode: "FOLDER_NOT_FOUND"` with suggestion to call `outlook_list_folders` first.

## Email Safety Gate (CRITICAL)

Email sending is a high-risk operation. The following safety gate MUST be enforced at all times:

```
┌─────────────────────────────────────────────────────────┐
│                    EMAIL SAFETY GATE                     │
├─────────────────────────────────────────────────────────┤
│                                                          │
│  1. outlook_compose_email                                │
│     └─ Creates draft in Drafts folder                    │
│     └─ Returns { isDraft: true, messageId }              │
│     └─ User can see/edit draft in Outlook                │
│                                                          │
│  2. outlook_send_message (separate tool)                 │
│     └─ Requires: confirmationToken from task pane        │
│     └─ Returns error if no valid token                   │
│                                                          │
│  3. Task Pane Confirmation Flow:                         │
│     └─ Shows: recipients, subject, attachment count      │
│     └─ Shows: total size, external recipient warning     │
│     └─ User clicks "Approve Send"                        │
│     └─ Token generated (single-use, 60s TTL)             │
│                                                          │
│  4. NEVER auto-send without human approval               │
│     └─ No tool parameter can bypass the gate             │
│     └─ No "auto-confirm" option exists                   │
│                                                          │
└─────────────────────────────────────────────────────────┘
```

**Policy filters (optional, configurable)**:

- Recipient domain allowlist/blocklist
- Attachment type restrictions (e.g., block .exe, .bat)
- External recipient warning (recipients outside the user's domain)
- Maximum recipient count
- Maximum attachment size

## Requirements

### Functional Requirements

- **FR-001**: The add-in MUST implement Microsoft Graph API integration via nested app authentication (NAA) to obtain access tokens scoped for Mail.Read, Mail.ReadWrite, Calendars.Read, and Mail.Send.
- **FR-002**: `outlook_list_folders` MUST return the complete folder hierarchy with unread counts (`unreadItemCount`), child folder nesting, and display names.
- **FR-003**: `outlook_list_emails` MUST support pagination with `pageSize` (default: 25, max: 100), `pageToken` for continuation, and folder filtering by folder ID or well-known name.
- **FR-004**: `outlook_search_emails` MUST support KQL (Keyword Query Language) syntax for full-text search across all folders, with pagination.
- **FR-005**: `outlook_get_email` MUST return the full body in both text and HTML representations, plus attachment metadata (name, size, contentType) for each attachment.
- **FR-006**: `outlook_list_calendar_events` MUST support date range filtering with `startDateTime` and `endDateTime` (default: today to +7 days), returning events sorted by start time.
- **FR-007**: `outlook_compose_email` MUST create a draft in the Drafts folder with `isDraft: true` — the tool MUST NEVER auto-send.
- **FR-008**: A separate `outlook_send_message` tool MUST require a single-use confirmation token obtained from the task pane UI before sending any email.
- **FR-009**: `outlook_move_email` MUST validate that the target folder exists before attempting the move, returning `errorCode: "FOLDER_NOT_FOUND"` if invalid.
- **FR-010**: ALL Graph API calls MUST be proxied through the add-in (not made server-side) to use the user's NAA auth context.
- **FR-011**: Graph API rate limiting (HTTP 429) MUST be handled with exponential backoff — up to 3 retries with jitter, then return `errorCode: "RATE_LIMITED"`.
- **FR-012**: The system MUST handle delegated vs. app-only permission differences gracefully, returning `errorCode: "INSUFFICIENT_PERMISSIONS"` when the user lacks the required scope.
- **FR-013**: `outlook_mark_read` MUST accept an array of message IDs and a boolean `isRead` flag, applying the change to all specified messages.
- **FR-014**: `outlook_get_categories` MUST return all user-defined categories with displayName and color.
- **FR-015**: `outlook_get_shared_mailboxes` MUST list mailboxes the user has delegate access to, with permission level (FullAccess, Read, SendAs).

### Key Entities

- **Mail Folder**: An Outlook mail folder identified by ID and display name. Has `unreadItemCount`, `totalItemCount`, and optional `parentFolderId` for hierarchy. Well-known names: Inbox, SentItems, Drafts, DeletedItems, JunkEmail.
- **Message Summary**: A lightweight email representation for list/search views containing: `id`, `subject`, `sender` (name + email), `receivedDateTime`, `bodyPreview` (first 255 chars), `isRead`, `hasAttachments`, `categories`.
- **Message Detail**: Full email content containing: all Message Summary fields plus `body` (text + HTML), `toRecipients`, `ccRecipients`, `bccRecipients`, `attachments[]` (metadata only), `importance`, `flag`.
- **Calendar Event Summary**: Lightweight event for list view containing: `id`, `subject`, `start`, `end`, `location` (displayName), `organizer` (name + email), `attendeeCount`, `isAllDay`, `onlineMeetingUrl`.
- **Calendar Event Detail**: Full event containing: all summary fields plus `body` (text + HTML), `attendees[]` (email, name, status), `recurrence`, `sensitivity`, `responseStatus`.
- **Confirmation Token**: A single-use, time-limited (60s TTL) token generated by the task pane when the user approves an email send. Required by `outlook_send_message`.
- **Shared Mailbox**: An external mailbox the user has delegate access to, with `emailAddress`, `displayName`, and `permissionLevel`.

## Success Criteria

### Measurable Outcomes

- **SC-001**: User can list all mail folders with unread counts within 3 seconds of calling `outlook_list_folders`.
- **SC-002**: Email search returns results from across all folders within 5 seconds of calling `outlook_search_emails`.
- **SC-003**: Calendar events for the next 7 days are returned within 3 seconds of calling `outlook_list_calendar_events`.
- **SC-004**: No email is ever sent without explicit human confirmation via the task pane safety gate — this is verified by the absence of any code path in `outlook_send_message` that bypasses the confirmation token check.
- **SC-005**: Shared mailbox access works with appropriate permission checks, returning `INSUFFICIENT_PERMISSIONS` when the user lacks write access.
- **SC-006**: All 11 new tools are independently testable via unit tests against mock Graph API responses.

## Assumptions

- The Outlook add-in is loaded in Outlook on the web or Outlook desktop (Microsoft 365).
- Nested app authentication (NAA) is available in the user's Outlook version (requires Microsoft 365 build 16.0.15111 or later).
- The add-in's manifest declares the required Graph scopes: `Mail.Read`, `Mail.ReadWrite`, `Mail.Send`, `Calendars.Read`.
- Graph API v1.0 endpoints are used (not beta).
- The user has an active internet connection — Graph API calls require network access.
- Shared mailbox discovery relies on the `/users/{id}/mailFolders` endpoint with the shared mailbox's SMTP address.
- The MCP server does not store any tokens, credentials, or email content — all state lives in the add-in and Outlook.
