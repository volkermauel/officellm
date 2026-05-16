# Feature Specification: Phase 2 — Word MVP

**Feature Branch**: `003-word-mvp`

**Created**: 2026-05-14

**Status**: Draft

**Input**: User description: "Implement Word outline, selection read, replace text (tracked changes), add comment, and tracked change management tools. Use Word tracked changes as the default mutation mode for Word (no confirmation gate needed). PowerPoint has no tracked changes API and uses undo-grouped direct writes. Add shared document context abstraction across Word and PowerPoint."

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

### User Story 2 — Editor replaces text via tracked changes (Priority: P1)

A writer selects a paragraph in their Word document, asks Open WebUI to rewrite it in a different tone. The change is applied as a Word tracked change (insertion + deletion) using `changeTrackingMode: "TrackMineOnly"`. The user can then accept or reject the change using Word's native Review ribbon or via the `word_accept_all_changes` / `word_reject_all_changes` tools.

**Why this priority**: This is the core Word mutation workflow. Using Word's native tracked changes preserves document integrity, gives the user full accept/reject capability through Word's built-in review tools, and avoids the need for a custom confirmation gate in the add-in.

**Independent Test**: Select text, call `word_replace_text`, verify a tracked change (deletion of old text + insertion of new text) appears in Word.

**Acceptance Scenarios**:

1. **Given** selected text "The project was completed on time", **When** developer calls `word_replace_text` with `replacement` "The project concluded within the scheduled timeframe", **Then** a tracked deletion of the original and tracked insertion of the replacement appear in Word
2. **Given** a tracked change from the add-in, **When** user calls `word_get_tracked_changes`, **Then** the response lists the change with type (insert/delete), author, range, and text
3. **Given** tracked changes from the add-in, **When** user calls `word_accept_all_changes`, **Then** all tracked changes are accepted and the document reflects the new text
4. **Given** tracked changes from the add-in, **When** user calls `word_reject_all_changes`, **Then** all tracked changes are rejected and the original text is restored
5. **Given** the selection spans multiple paragraphs, **When** developer calls `word_replace_text`, **Then** all selected paragraphs are tracked as deleted and the replacement inserted

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
- **FR-004**: The MCP server MUST expose `word_replace_text` that replaces the current selection (or specified range) with new text using Word tracked changes (`changeTrackingMode: "TrackMineOnly"`). The tool saves the current tracking mode, enables TrackMineOnly, performs the replacement, and restores the original mode. Returns `{ tracked: true }`.
- **FR-005**: Word mutations MUST use tracked changes by default (not direct edits). This satisfies the confirmation requirement — users accept/reject through Word's native Review ribbon or the tracked change management tools.
- **FR-005a**: The MCP server MUST expose `word_get_tracked_changes` returning a list of tracked changes in the document with type (insert/delete), author, range, and text content.
- **FR-005b**: The MCP server MUST expose `word_accept_all_changes` that accepts all tracked changes in the document.
- **FR-005c**: The MCP server MUST expose `word_reject_all_changes` that rejects all tracked changes in the document, restoring original content.
- **FR-005d**: PowerPoint has NO tracked changes API. PowerPoint mutations are direct-write with undo grouped per `PowerPoint.run()` batch. This is a fundamental platform difference.
- **FR-006**: The MCP server MUST expose `word_insert_after_heading` that inserts generated text after a specified heading.
- **FR-007**: The MCP server MUST expose `word_add_review_comments` that adds Word comments to the current selection or cursor position.
- **FR-008**: All Word mutation tools MUST use the shared document context envelope for reporting.
- **FR-009**: Word mutations are tracked changes visible in Word's native Review pane. No separate diff display in the task pane is required — Word's built-in tracked changes UI serves this purpose.
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
- **SC-002**: A replace operation applies the change as a tracked change. User approval is handled through Word's native Review tools (Accept/Reject) or the `word_accept_all_changes`/`word_reject_all_changes` tools.
- **SC-003**: Comments added via Open WebUI are correctly attached to the intended document ranges in Word.
- **SC-004**: The shared document context envelope is structurally identical for Word and PowerPoint responses.
- **SC-005**: Insert operations after headings place content at the correct paragraph boundary (no offset errors).

## Assumptions

- Word 2019 or later (or Microsoft 365) is available on the developer machine.
- Tracked changes and comments are supported in the target document (not protected/locked).
- Heading styles follow standard Word conventions (Heading 1, Heading 2, etc.).
- The add-in author name for tracked changes can be configured or defaults to the add-in product name.
