# Feature Specification: Phase 1 — PowerPoint MVP

**Feature Branch**: `002-powerpoint-mvp`

**Created**: 2026-05-14

**Status**: Draft

**Input**: User description: "Implement powerpoint_get_deck_outline, powerpoint_get_slide, powerpoint_update_shape_text and powerpoint_update_speaker_notes. Add diff preview for text changes. Add local audit log. Add basic task pane with context preview."

## User Scenarios & Testing

### User Story 1 — Presenter reviews deck structure before a meeting (Priority: P1)

A presenter opens their PowerPoint deck, asks Open WebUI to summarize the presentation structure, and receives a hierarchical outline of all slides including titles, text placeholders, speaker notes, and slide order. This helps them quickly orient themselves and prepare talking points.

**Why this priority**: Reading deck structure is the most common starting point for any LLM-assisted PowerPoint workflow. Without it, the model has no context to work with.

**Independent Test**: Call `powerpoint_get_deck_outline` from Open WebUI and verify the response contains all slides with titles, text content summaries, and speaker notes (when enabled).

**Acceptance Scenarios**:

1. **Given** a 20-slide presentation with varied slide types, **When** developer calls `powerpoint_get_deck_outline`, **Then** the response includes slide index, title, text placeholder summaries, and optional speaker notes for each slide
2. **Given** `includeSpeakerNotes: false`, **When** developer calls `powerpoint_get_deck_outline`, **Then** speaker notes are omitted from the response
3. **Given** hidden slides exist in the deck, **When** `includeHiddenSlides: false` (default), **Then** hidden slides are excluded from the outline

---

### User Story 2 — Editor corrects a slide's content with preview (Priority: P1)

A document editor notices a typo or outdated information on a specific slide. They ask Open WebUI to update the text, see a before/after diff in the task pane, approve the change, and watch it applied to the shape.

**Why this priority**: This is the core mutation workflow — propose → preview → confirm → apply. It validates the safety gate pattern that all subsequent phases must follow.

**Independent Test**: Call `powerpoint_update_shape_text` with a new text value, verify the diff preview appears in the task pane, approve, and verify the shape text changed in PowerPoint.

**Acceptance Scenarios**:

1. **Given** a slide with a title shape containing "Quarterly results", **When** developer calls `powerpoint_update_shape_text` with `{"slideIndex": 3, "shapeId": "Title 1", "text": "Q3 performance summary"}`, **Then** the task pane shows a diff preview and the change is NOT applied until user approves
2. **Given** a diff preview is shown, **When** user clicks "Approve", **Then** the shape text is updated in PowerPoint
3. **Given** a diff preview is shown, **When** user clicks "Reject", **Then** no changes are made to the slide
4. **Given** `requiresConfirmation: true` response, **When** developer calls the tool without a valid confirmation token, **Then** the change is NOT applied and a confirmation request is returned

---

### User Story 3 — Presenter prepares speaker notes (Priority: P2)

A presenter wants to add or refine speaker notes for specific slides. They ask Open WebUI to draft notes based on slide content, review the notes in the task pane, and apply them.

**Why this priority**: Speaker notes are a key PowerPoint feature that LLMs can significantly improve. This workflow is non-destructive (notes don't affect the slide itself) and has low risk.

**Independent Test**: Call `powerpoint_update_speaker_notes`, verify notes appear in PowerPoint's notes pane after approval.

**Acceptance Scenarios**:

1. **Given** a slide with no speaker notes, **When** developer calls `powerpoint_update_speaker_notes` with draft text, **Then** the notes appear in PowerPoint after approval
2. **Given** an existing speaker note, **When** developer calls `powerpoint_update_speaker_notes` with new text, **Then** the old notes are replaced after approval (with diff preview showing old vs. new)
3. **Given** multiple slides, **When** developer calls the tool for slide 5 only, **Then** only slide 5's notes are affected

---

### User Story 4 — Developer audits all PowerPoint tool calls (Priority: P2)

A security-conscious user or IT admin reviews the local audit log to understand what LLM tools were called, what changes were proposed, which were approved, and when.

**Why this priority**: Audit logging is a constitutional requirement (Principle III: User Control & Transparency). Without it, the system cannot be trusted for production use.

**Independent Test**: After running several tool calls, open the audit log file and verify each entry contains timestamp, tool name, inputs, confirmation status, and outcome.

**Acceptance Scenarios**:

1. **Given** three tool calls were made (one read, one mutation approved, one mutation rejected), **When** developer opens the audit log, **Then** all three entries are present with correct fields
2. **Given** an audit entry for a mutation tool, **When** developer inspects the entry, **Then** it contains `toolName`, `timestamp`, `documentHandle`, `inputs`, `requiresConfirmation`, `confirmationStatus`, and `outcome`
3. **Given** the audit log file, **When** the MCP helper process restarts, **Then** new entries are appended without truncating existing entries

---

### Edge Cases

- What happens when a shape ID referenced in `powerpoint_update_shape_text` does not exist on the slide? The tool should return `{"ok": false, "errorCode": "SHAPE_NOT_FOUND"}` with a recoverable error.
- What happens when the user closes the presentation while a confirmation dialog is pending? The pending confirmation should be cancelled and logged.
- What happens when a slide contains grouped shapes with nested text? The tool should handle shape hierarchies and report which level of the hierarchy was modified.
- What happens when the deck is in protected view or read-only mode? The tool should return a clear error indicating the mode.

## Requirements

### Functional Requirements

- **FR-001**: The MCP server MUST expose `powerpoint_get_deck_outline` returning slide titles, text placeholders, speaker notes metadata, and slide order.
- **FR-002**: `powerpoint_get_deck_outline` MUST accept optional parameters `includeSpeakerNotes` (default: false) and `includeHiddenSlides` (default: false).
- **FR-003**: The MCP server MUST expose `powerpoint_get_slide` returning text, notes, and shape metadata for one slide by index.
- **FR-004**: The MCP server MUST expose `powerpoint_update_shape_text` that updates a specific shape's text on a specified slide.
- **FR-005**: `powerpoint_update_shape_text` MUST return `requiresConfirmation: true` with a diff preview before applying any change.
- **FR-006**: The MCP server MUST expose `powerpoint_update_speaker_notes` that creates or updates speaker notes for specified slides.
- **FR-007**: All mutation tools MUST generate a before/after diff preview viewable in the task pane.
- **FR-008**: The VSTO add-in MUST display a task pane showing current context (active presentation, selected slide) and pending change previews.
- **FR-009**: The MCP server MUST write every tool call to a local audit log file (JSONL format) with timestamp, tool name, inputs, confirmation status, and outcome.
- **FR-010**: The system MUST return standard error envelopes with `ok: false`, `errorCode`, `message`, and `recoverable` fields for all failure modes.

### Key Entities

- **Slide**: Identified by zero-based index. Contains text placeholders, shapes, speaker notes, and visibility state.
- **Shape**: A text-containing element on a slide (title, content placeholder, textbox). Identified by name/ID within the slide.
- **Speaker Notes**: Text notes associated with a slide, displayed in presenter view.
- **Diff Preview**: A structured representation of before/after text changes for a specific shape or notes block.
- **Audit Log Entry**: A JSONL record containing tool invocation details, confirmation state, and result.

## Success Criteria

### Measurable Outcomes

- **SC-001**: A developer can retrieve the full deck outline (100 slides) from Open WebUI within 5 seconds.
- **SC-002**: A shape text change requires explicit user approval via the task pane before being applied.
- **SC-003**: The diff preview accurately shows the exact text that will be replaced (character-level diff).
- **SC-004**: Every tool call is recorded in the audit log within 1 second of execution.
- **SC-005**: The task pane displays current context and pending changes with less than 200ms visual latency.

## Assumptions

- PowerPoint 2019 or later (or Microsoft 365) is available on the developer machine.
- The VSTO add-in runs in the same process space as PowerPoint (standard VSTO host-add-in model).
- Shape IDs are stable within a presentation session (they may change if shapes are recreated).
- Speaker notes are stored within the `.pptx` file and do not require external storage.
- The MCP helper process runs on a dynamic port; the VSTO add-in discovers it via a local manifest or well-known default port.
