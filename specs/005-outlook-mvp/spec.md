# Feature Specification: Phase 4 — Outlook MVP

**Feature Branch**: `005-outlook-mvp`

**Created**: 2026-05-14

**Status**: Draft

**Input**: User description: "Implement current item read, thread summary and draft reply tools. Keep sending disabled or gated behind manual Outlook confirmation. Add optional policy filter for recipient domains and attachments."

## User Scenarios & Testing

### User Story 1 — Professional reads an email thread context (Priority: P1)

A professional has a long email thread open in Outlook and wants to understand the key points without reading every message. They ask Open WebUI to summarize the conversation, receiving a structured summary with timeline, key decisions, action items, and unresolved questions.

**Why this priority**: Email threads can be dozens of messages long. A summarization tool is the most immediately valuable Outlook feature — it saves time and reduces cognitive load. Reading the current item is the prerequisite for all other Outlook operations.

**Independent Test**: Open an email thread, call `outlook_get_current_item`, verify metadata and bounded body text are returned. Then call `outlook_summarize_thread` and verify the summary captures key points accurately.

**Acceptance Scenarios**:

1. **Given** a selected email in the reading pane, **When** developer calls `outlook_get_current_item`, **Then** the response includes sender, recipients, subject, received date, and bounded body text (truncated at the output limit)
2. **Given** a multi-message thread, **When** developer calls `outlook_summarize_thread`, **Then** the response includes a structured summary with timeline, key decisions, action items, and unresolved questions
3. **Given** an email with attachments, **When** `outlook_get_current_item` is called, **Then** attachment metadata (name, size, type) is included but file contents are NOT extracted by default

---

### User Story 2 — Professional drafts a reply with controlled tone (Priority: P1)

A professional needs to reply to an email thread. They ask Open WebUI to draft a reply in a specific tone (concise, formal, friendly, or technical), review the draft in the task pane, and either approve it (which creates a draft in Outlook) or request revisions.

**Why this priority**: Drafting replies is a high-frequency email task. Controlling tone and requiring manual approval before creating a draft aligns with the safety-first principle. The draft is created in Outlook's Drafts folder, never sent automatically.

**Independent Test**: Call `outlook_draft_reply` with different tones, verify the draft appears in Outlook's Drafts folder with the correct content.

**Acceptance Scenarios**:

1. **Given** a selected email thread, **When** developer calls `outlook_draft_reply` with `tone: "formal"` and `keyPoints: ["acknowledge receipt", "request clarification on section 3"]`, **Then** a draft reply appears in Outlook's Drafts folder with the specified tone and key points
2. **Given** a draft is created, **When** developer opens the draft in Outlook, **Then** the reply is addressed to the correct recipients (To/CC based on the original thread)
3. **Given** `includeThreadSummary: true`, **When** developer calls `outlook_draft_reply`, **Then** the draft includes a brief thread summary at the top before the new content
4. **Given** the user requests revisions via Open WebUI, **When** a new draft is created, **Then** the previous draft remains in Drafts (both drafts are visible)

---

### User Story 3 — Organizer applies categories to emails (Priority: P2)

A project organizer wants to apply a specific Outlook category (color-coded label) to selected emails for organizational purposes. They ask Open WebUI to apply a category, and the change is applied immediately (non-destructive, reversible via Outlook's UI).

**Why this priority**: Category management is a low-risk, high-value organizational task. It doesn't modify email content and can be reversed entirely within Outlook. It validates the non-mutation tool pattern for Outlook.

**Independent Test**: Call `outlook_apply_category`, verify the category badge appears on the selected email(s) in Outlook.

**Acceptance Scenarios**:

1. **Given** one or more selected emails, **When** developer calls `outlook_apply_category` with category name "Project Alpha", **Then** the selected emails display the "Project Alpha" category badge
2. **Given** a category does not exist in the user's category list, **When** developer calls `outlook_apply_category`, **Then** the tool returns an error listing available categories
3. **Given** multiple emails are selected, **When** developer calls `outlook_apply_category`, **Then** the category is applied to all selected items

---

### User Story 4 — Security policy filters outbound messages (Priority: P2)

An IT administrator configures a policy that prevents drafts from being sent to external domains or containing certain attachment types. When a user attempts to send via Open WebUI, the policy filter evaluates the draft and blocks or warns about violations.

**Why this priority**: Sending emails has external side effects (Principle I: Safety-First). A policy filter adds an additional safety layer beyond the manual confirmation gate, protecting against accidental data leakage to unauthorized recipients.

**Independent Test**: Configure a policy blocking external domains, attempt to send a draft to an external address, verify the send is blocked with a clear explanation.

**Acceptance Scenarios**:

1. **Given** a policy that blocks external domains, **When** developer attempts to send a draft addressed to an external recipient, **Then** the send operation is blocked and the task pane shows which recipients violated the policy
2. **Given** a policy that blocks certain attachment types (e.g., .exe, .bat), **When** developer attempts to send a draft with a blocked attachment, **Then** the send operation is blocked and the task pane identifies the blocked attachments
3. **Given** all policy checks pass, **When** developer confirms the send, **Then** the email is sent through Outlook's normal send mechanism

---

### User Story 5 — Professional sends a drafted message with explicit confirmation (Priority: P2)

A professional has reviewed a draft reply in Outlook and decides to send it. They request the send action via Open WebUI, see a final confirmation dialog showing recipient, subject, and a truncated body preview, and explicitly approve the send. The send operation requires this explicit approval — it cannot be triggered by an LLM tool call alone.

**Why this priority**: Sending email is a high-risk external side effect (Principle I: Safety-First). The constitutional requirement is that `outlook_send_message` must NEVER execute from an LLM tool call alone — it requires explicit user approval in the Outlook UI.

**Independent Test**: Request a send via Open WebUI, verify the confirmation dialog shows recipient/subject/body preview, approve, and verify the email is sent.

**Acceptance Scenarios**:

1. **Given** a draft exists in Outlook's Drafts folder, **When** developer calls `outlook_send_message`, **Then** the task pane shows a final confirmation with recipient list, subject line, and truncated body preview
2. **Given** user approves the confirmation, **When** the send executes, **Then** the email is sent through Outlook's normal send mechanism (not bypassing Outlook)
3. **Given** user rejects the confirmation, **When** no action is taken, **Then** the draft remains in Drafts unchanged
4. **Given** an LLM tool call attempts to send without a valid confirmation token, **When** the MCP server processes the request, **Then** the send is rejected with `requiresConfirmation: true` and a confirmation request is returned

---

### Edge Cases

- What happens when the selected item is not an email (e.g., a calendar appointment or contact)? `outlook_get_current_item` should return an error indicating the item type is unsupported.
- What happens when an email thread has hundreds of messages? The summary tool should process only a bounded number of messages (configurable, default 50) and indicate if truncation occurred.
- What happens when the user is offline? The send operation should fail gracefully with a clear error message and leave the draft intact.
- What happens when multiple items are selected and `outlook_draft_reply` is called? The tool should operate on the first selected item or return an error asking the user to select a single email.
- What happens when the draft exceeds Outlook's maximum message size (typically 25 MB)? The tool should detect this during preview and warn the user before attempting to create the draft.

## Requirements

### Functional Requirements

- **FR-001**: The MCP server MUST expose `outlook_get_current_item` returning metadata (sender, recipients, subject, received date) and bounded body text for the selected Outlook item.
- **FR-002**: The MCP server MUST expose `outlook_summarize_thread` returning a structured summary with timeline, key decisions, action items, and unresolved questions for the selected email thread.
- **FR-003**: The MCP server MUST expose `outlook_draft_reply` creating a draft reply in Outlook's Drafts folder for the currently selected email thread.
- **FR-004**: `outlook_draft_reply` MUST accept parameters: `tone` (enum: concise, formal, friendly, technical), `keyPoints` (array of strings), and `includeThreadSummary` (default: false).
- **FR-005**: Draft replies MUST be created in Outlook's Drafts folder — they are NEVER sent automatically.
- **FR-006**: The MCP server MUST expose `outlook_apply_category` applying an Outlook category to selected items.
- **FR-007**: The MCP server MUST expose `outlook_send_message` that sends a drafted message only after explicit user confirmation via the Outlook UI.
- **FR-008**: `outlook_send_message` MUST NEVER execute from an LLM tool call alone — it MUST require a valid confirmation token and explicit user approval in the Outlook task pane.
- **FR-009**: The system MUST support an optional policy filter that can block sends based on recipient domains, attachment types, or other configurable criteria.
- **FR-010**: All Outlook tools MUST return the standard MCP response envelope with `ok`, `app`, `documentId`, `result`, `warnings`, `requiresConfirmation`, and `auditId` fields.

### Key Entities

- **Outlook Item**: Any item in Outlook (MailItem, AppointmentItem, ContactItem). Identified by its position in the Explorer/Inspector selection.
- **Email Thread**: A group of related messages connected by subject and reply/forward chains. Identified by a thread identifier or conversation ID.
- **Draft Reply**: An unsent MailItem created in Outlook's Drafts folder, addressed to the appropriate recipients based on the original thread.
- **Category**: An Outlook color-coded label applied to items for organization. Defined in Outlook's category list.
- **Policy Filter**: A configurable rule set evaluating drafts before sending, checking recipient domains, attachment types, and other criteria.
- **Confirmation Token**: A session-scoped token required for `outlook_send_message` that proves explicit user approval was obtained in the Outlook UI.

## Success Criteria

### Measurable Outcomes

- **SC-001**: A developer can retrieve metadata and bounded body text from the selected email within 3 seconds.
- **SC-002**: A thread summary of 20 messages is generated and displayed in the task pane within 10 seconds.
- **SC-003**: A draft reply appears in Outlook's Drafts folder within 5 seconds of approval.
- **SC-004**: An `outlook_send_message` call without a valid confirmation token is always rejected.
- **SC-005**: A policy filter can be configured to block sends to specific domains or attachment types.

## Assumptions

- Outlook 2019 or later (or Microsoft 365) is available on the developer machine.
- The VSTO add-in runs in the same process space as Outlook.
- Drafts are created using Outlook's native API, ensuring they appear in the correct folder and respect Outlook's own validation rules.
- The send operation uses Outlook's standard send mechanism (not bypassing it), preserving all Outlook features (send receipts, encryption, etc.).
- Categories used by the policy filter are pre-configured in the user's Outlook category list.
