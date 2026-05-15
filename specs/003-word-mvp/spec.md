# Feature Specification: Phase 2 — Word MVP

**Feature Branch**: `003-word-mvp`

**Created**: 2026-05-14

**Status**: Draft

**Input**: User description: "Implement Word outline, selection read, rewrite selection and add comment tools. Use Word comments or tracked changes for review-oriented workflows. Add shared document context abstraction across Word and PowerPoint."

## Architecture

The MCP server acts as a central hub. The **unified Office JS Add-in** auto-detects the host via `Office.onReady(info.host)`. For Word instances, the add-in registers with host type `Word` and instance IDs like `word_1`, `word_2`. Tools accept optional `instanceId` parameter to target specific instances. Shared document context abstraction provides unified response structure across hosts.

```
Open WebUI                    MCP Server (port 3000)              Office Add-ins
    │                                │                               │
    │── tools/call ─────────────────►│                               │
    │   { tool, { instanceId?, ... } │                               │
    │                                │── route to instance ─────────►│
    │◄── result ─────────────────────│                               │
    │                                │◄── result ────────────────────│
    │                                │◄── /instances/register ──────│  (on load)
    │                                │◄── /instances/:id/heartbeat ─│  (every 10s)
    │                                │►── /instances/:id/commands ──│  (poll every 2s)
    │                                │◄── /instances/:id/result ────│  (after execution)
```

## User Scenarios & Testing

### User Story 1 — Reviewer obtains document structure (Priority: P1)

A legal reviewer opens a lengthy Word document and asks Open WebUI to extract the document outline (headings, styles, paragraph ranges). This gives them a navigable map of the document without manually scrolling through hundreds of pages.

**Why this priority**: Word documents can be extremely long. An outline tool is the essential first step for any LLM-assisted Word workflow — it provides context without sending the entire document to the model.

**Independent Test**: Call `word_get_outline` and verify the response contains all headings with their hierarchy levels, styles, and paragraph ranges.

**Acceptance Scenarios**:

1. **Given** a 50-page Word document with headings at levels 1–4, **When** developer calls `word_get_outline`, **Then** the response includes all headings with level, text, style name, and paragraph range (start/end)
2. **Given** `maxDepth: 2`, **When** developer calls `word_get_outline`, **Then** only headings at levels 1 and 2 are returned
3. **Given** a document with no styled headings, **When** developer calls `word_get_outline`, **Then** the response is empty or contains body text paragraphs if no headings exist

---

### User Story 2 — Editor rewrites selected text with tracked preview (Priority: P1)

A writer selects a paragraph in their Word document, asks Open WebUI to rewrite it in a different tone, sees the before/after diff in the task pane, and either approves or rejects the change. The approved change is applied as a tracked change (not directly into the document).

**Why this priority**: This is the core Word mutation workflow. Using tracked changes preserves document integrity and gives the user full undo capability through Word's native review tools.

**Independent Test**: Select text, call `word_rewrite_selection`, approve the diff, verify a tracked change appears in Word.

**Acceptance Scenarios**:

1. **Given** selected text "The project was completed on time", **When** developer calls `word_rewrite_selection` with tone "formal", **Then** the task pane shows a diff preview and the change is NOT applied until approved
2. **Given** user approves the diff, **When** the change is applied, **Then** it appears as a Word tracked change (inserted text marked with the add-in's author name)
3. **Given** user rejects the diff, **When** no action is taken, **Then** the document remains unchanged
4. **Given** the selection spans multiple paragraphs, **When** developer calls `word_rewrite_selection`, **Then** the diff covers all selected content

---

### User Story 3 — Contributor inserts content under a heading (Priority: P2)

A contributor asks Open WebUI to generate additional content (e.g., an executive summary, a section expansion) and insert it under a specific heading in the document. The insertion point is confirmed via diff preview.

**Why this priority**: Content generation into existing documents is a common workflow. Inserting at a known structural point (under a heading) is more reliable than free-form insertion.

**Independent Test**: Call `word_insert_after_heading`, approve the preview, verify new content appears at the correct location as a tracked change.

**Acceptance Scenarios**:

1. **Given** a document with a heading "Methodology", **When** developer calls `word_insert_after_heading` with generated text, **Then** the task pane shows where the content will be inserted
2. **Given** user approves, **When** the change is applied, **Then** the new paragraphs appear after the heading as tracked changes
3. **Given** multiple headings match the query, **When** developer specifies the exact heading text, **Then** only the matching heading's insertion point is affected

---

### User Story 4 — Reviewer adds comments instead of editing (Priority: P2)

A reviewer wants to leave feedback without modifying the document content. They ask Open WebUI to generate review comments for specific sections, review the comment text, and approve it. Comments are added as Word comments, not tracked changes.

**Why this priority**: Not all LLM interactions should modify content. Comments provide a non-destructive review workflow that preserves the original document while capturing model-generated feedback.

**Independent Test**: Call `word_add_review_comments`, verify Word comments appear at the correct locations.

**Acceptance Scenarios**:

1. **Given** selected text in a document, **When** developer calls `word_add_review_comments` with generated feedback, **Then** a Word comment appears attached to the selected range
2. **Given** no selection exists but a cursor position is active, **When** developer calls `word_add_review_comments`, **Then** the comment is placed at the cursor position
3. **Given** multiple comments are generated, **When** user approves, **Then** all comments are added as separate Word comments

---

### User Story 5 — Shared document context abstraction (Priority: P2)

The system provides a unified document context model that works across both Word and PowerPoint. Both hosts report their context using the same envelope structure (app, document, selection, content, limits), enabling Open WebUI to handle either host with the same tool-calling logic.

**Why this priority**: This is an architectural requirement that reduces duplication and enables future hosts (Excel, Outlook) to plug into the same model. It's foundational for the multi-host roadmap.

**Independent Test**: Call context tools on both Word and PowerPoint; verify responses use the same envelope structure with host-appropriate field values.

**Acceptance Scenarios**:

1. **Given** Word is active, **When** developer calls `office_get_active_app`, **Then** the response uses the shared envelope with Word-specific selection metadata
2. **Given** PowerPoint is active, **When** developer calls `office_get_active_app`, **Then** the response uses the same envelope structure with PowerPoint-specific selection metadata
3. **Given** both hosts can be queried, **When** Open WebUI processes the response, **Then** it handles both formats identically (no host-specific parsing logic)

---

### Edge Cases

- What happens when the selected text spans across a table or section break? The tool should handle cross-boundary selections and report which boundaries were crossed.
- What happens when the document is in Read Mode (not Edit Mode)? Word's object model behaves differently in Read Mode — the tool should detect this and return an error suggesting the user switch to Edit Mode.
- What happens when tracked changes are already enabled by another user? The add-in should respect existing author names or use a distinct add-in author to differentiate LLM-generated changes.
- What happens when the document has content controls (structured document tags)? The tool should avoid modifying content within content controls unless explicitly requested.

## Requirements

### Functional Requirements

- **FR-001**: The MCP server MUST expose `word_get_outline` returning headings with level, text, style name, and paragraph range for the active Word document.
- **FR-002**: `word_get_outline` MUST accept an optional `maxDepth` parameter (default: 3).
- **FR-003**: The MCP server MUST expose `word_get_paragraphs` returning paragraphs by range, heading, or selection neighborhood.
- **FR-004**: The MCP server MUST expose `word_rewrite_selection` that rewrites selected text and returns a diff preview requiring user confirmation.
- **FR-005**: Approved rewrites MUST be applied as Word tracked changes (not direct edits) to preserve undo capability.
- **FR-006**: The MCP server MUST expose `word_insert_after_heading` that inserts generated text after a specified heading.
- **FR-007**: The MCP server MUST expose `word_add_review_comments` that adds Word comments to the current selection or cursor position.
- **FR-008**: All Word mutation tools MUST use the shared document context envelope for reporting.
- **FR-009**: The task pane MUST display before/after diffs for Word rewrites, showing the original and proposed text side by side.
- **FR-010**: The shared document context abstraction MUST provide a unified response structure across Word and PowerPoint with host-specific selection fields.

### Key Entities

- **Heading**: A paragraph styled with a heading style (Heading 1, Heading 2, etc.). Contains level, text, and paragraph range.
- **Paragraph**: A block of text in the document. Identified by its position in the document flow.
- **Tracked Change**: A Word-native revision marker indicating inserted or deleted text, attributed to the add-in's author.
- **Word Comment**: A non-destructive review annotation attached to a document range.
- **Document Context Envelope**: Shared structure `{app, document, selection, content, limits}` used by all Office hosts.

## Success Criteria

### Measurable Outcomes

- **SC-001**: A developer can extract the outline of a 200-page Word document from Open WebUI within 5 seconds.
- **SC-002**: A rewrite operation requires explicit user approval and appears as a tracked change (not a direct edit).
- **SC-003**: Comments added via Open WebUI are correctly attached to the intended document ranges in Word.
- **SC-004**: The shared document context envelope is structurally identical for Word and PowerPoint responses.
- **SC-005**: Insert operations after headings place content at the correct paragraph boundary (no offset errors).

## Assumptions

- Word 2019 or later (or Microsoft 365) is available on the developer machine.
- Tracked changes and comments are supported in the target document (not protected/locked).
- Heading styles follow standard Word conventions (Heading 1, Heading 2, etc.).
- The add-in author name for tracked changes can be configured or defaults to the add-in product name.
