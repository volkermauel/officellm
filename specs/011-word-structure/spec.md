# Feature Specification: Phase 11 — Word Document Structure Tools

**Feature Branch**: `011-word-structure`

**Created**: 2026-05-16

**Status**: Draft

**Depends on**: Phase 2 (Word MVP) — basic read/write tools must exist

**Input**: User description: "Implement Word document structure tools: tables, headers/footers, replace-selection, images, styles, and lists. All write operations use tracked changes pattern."

## Architecture

Same as Word MVP — uses Word.run() context batching with tracked changes support. The **unified Office JS Add-in** auto-detects the host via `Office.onReady(info.host)`. For Word instances, the add-in registers with host type `Word` and instance IDs like `word_1`, `word_2`. Tools accept optional `instanceId` parameter to target specific instances.

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

All write operations use the tracked changes pattern:

```
save mode → set TrackMineOnly → mutate → restore mode → return { tracked: true }
```

## User Scenarios & Testing

### User Story 1 — Read tables in document (Priority: P0)

A user asks "What tables are in this document?" The LLM calls `word_get_tables` and receives table count, dimensions, and cell content for each table. This provides a structural overview of tabular data in the document.

**Why this priority**: Tables are the most common structured data container in Word documents. Understanding table layout is essential before any read or modification of tabular content.

**Independent Test**: Call `word_get_tables` on a document with known tables, verify the response lists each table with correct row/column counts and cell text.

**Acceptance Scenarios**:

1. **Given** a document with 3 tables of varying sizes, **When** developer calls `word_get_tables`, **Then** the response includes table index, row count, column count, and cell text for each table
2. **Given** a document with no tables, **When** developer calls `word_get_tables`, **Then** the response returns an empty array with `tableCount: 0`
3. **Given** a table with merged cells, **When** developer calls `word_get_tables`, **Then** the response includes cell content as-is with a `mergedRegions` note for affected cells

---

### User Story 2 — Insert table at location (Priority: P0)

A user asks "Insert a 3x4 table after the 'Results' heading." The LLM calls `word_insert_table` with row/column count, location (after paragraph index), and optional header row data. The table appears in the document with tracked changes.

**Why this priority**: Table insertion is a fundamental document editing operation. Combined with `word_get_tables`, it enables full table management.

**Independent Test**: Call `word_insert_table` with known dimensions and location, verify the table appears at the correct position with correct structure.

**Acceptance Scenarios**:

1. **Given** a document with 10 paragraphs, **When** developer calls `word_insert_table` with `rows: 3, columns: 4, afterParagraphIndex: 5`, **Then** a 3x4 table is inserted between paragraph 5 and paragraph 6
2. **Given** a call with `headerRow: ["Name", "Value", "Status", "Date"]`, **When** developer calls `word_insert_table`, **Then** the first row of the table contains the provided header values
3. **Given** `afterParagraphIndex: "end"`, **When** developer calls `word_insert_table`, **Then** the table is appended at the end of the document body
4. **Given** the table is inserted, **When** the user opens the Review pane, **Then** the insertion appears as a tracked change

---

### User Story 3 — Update table cell (Priority: P1)

A user asks "Change the value in row 2, column 3 of the first table." The LLM calls `word_update_table_cell` with table index, row, column, and new text. The change is tracked.

**Why this priority**: Cell-level editing is the most granular table operation needed for targeted corrections and data updates.

**Independent Test**: Call `word_update_table_cell` on a known cell, verify the value changes and a tracked change is created.

**Acceptance Scenarios**:

1. **Given** table 0 has a cell at row 1, column 2 with value "old", **When** developer calls `word_update_table_cell` with `tableIndex: 0, row: 1, column: 2, text: "new"`, **Then** the cell value becomes "new" and a tracked change is visible
2. **Given** coordinates that are out of bounds, **When** developer calls `word_update_table_cell`, **Then** the tool returns an error with `errorCode: "CELL_OUT_OF_BOUNDS"` and the table's actual dimensions
3. **Given** a table with merged cells, **When** developer targets a merged region, **Then** the text is set on the merge origin cell

---

### User Story 4 — Read headers and footers (Priority: P1)

A user asks "What's in the header?" The LLM calls `word_get_headers_footers` and receives header/footer content for each section, including default, first page, and odd/even page variants.

**Why this priority**: Headers and footers contain critical document metadata (page numbers, confidentiality notices, revision dates). Reading them is essential for document review workflows.

**Independent Test**: Call `word_get_headers_footers` on a document with known header/footer content, verify the response matches.

**Acceptance Scenarios**:

1. **Given** a document with 2 sections, each with a default header, **When** developer calls `word_get_headers_footers`, **Then** the response includes header content for each section
2. **Given** a section with "Different first page" enabled, **When** developer calls `word_get_headers_footers`, **Then** the response includes separate `default` and `firstPage` header content
3. **Given** a header containing page number fields, **When** developer calls `word_get_headers_footers`, **Then** field codes are preserved in the response as `{{PAGE}}` or similar placeholders

---

### User Story 5 — Write headers and footers (Priority: P1)

A user asks "Add 'Confidential' to the footer." The LLM calls `word_set_header_footer` with section index, type (header/footer), and content. The change is tracked.

**Why this priority**: Setting headers/footers programmatically is needed for document preparation workflows (adding confidentiality notices, dates, document IDs).

**Independent Test**: Call `word_set_header_footer` with known content, verify it appears in the specified header/footer location.

**Acceptance Scenarios**:

1. **Given** section 0 has an empty default footer, **When** developer calls `word_set_header_footer` with `sectionIndex: 0, type: "footer", variant: "default", text: "Confidential"`, **Then** the footer displays "Confidential"
2. **Given** a header with existing content, **When** developer calls `word_set_header_footer`, **Then** the existing content is replaced with the new text
3. **Given** a write operation, **When** the change is applied, **Then** a tracked change appears in the Review pane

---

### User Story 6 — Replace current selection (Priority: P0)

A user asks "Rewrite the selected text to be more formal." The LLM calls `word_replace_selection` with new text. The replacement uses tracked changes so the user can accept or reject each modification.

**Why this priority**: Selection replacement is the most natural LLM interaction pattern — the user selects text, asks for improvement, and reviews the change. This is the primary write workflow for Word.

**Independent Test**: Select text in Word, call `word_replace_selection` with replacement text, verify the tracked change appears.

**Acceptance Scenarios**:

1. **Given** the user has selected "Hello world" in the document, **When** developer calls `word_replace_selection` with `text: "Greetings, everyone"`, **Then** the selection is replaced and a tracked change is visible in the Review pane
2. **Given** no text is selected, **When** developer calls `word_replace_selection`, **Then** the tool returns an error with `errorCode: "EMPTY_SELECTION"`
3. **Given** a tracked change is created, **When** the user opens the Review ribbon, **Then** they can accept or reject the change using native Word controls

---

### User Story 7 — Insert image (Priority: P2)

A user asks "Insert the chart image after paragraph 5." The LLM calls `word_insert_image` with base64-encoded image data and location. The image is inserted as an inline picture.

**Why this priority**: Image insertion enriches documents with visual content but is less critical than text manipulation operations.

**Independent Test**: Call `word_insert_image` with a small base64 image, verify it appears at the specified location.

**Acceptance Scenarios**:

1. **Given** a valid base64-encoded PNG image under 10MB, **When** developer calls `word_insert_image` with `afterParagraphIndex: 4`, **Then** the image appears as an inline picture after paragraph 4
2. **Given** an image exceeding 10MB, **When** developer calls `word_insert_image`, **Then** the tool returns an error with `errorCode: "IMAGE_TOO_LARGE"` and the actual size
3. **Given** invalid base64 data, **When** developer calls `word_insert_image`, **Then** the tool returns an error with `errorCode: "INVALID_IMAGE_DATA"`

---

### User Story 8 — Apply style to paragraph (Priority: P1)

A user asks "Make paragraph 3 a Heading 2." The LLM calls `word_apply_style` with paragraph index and style name. The style is validated before application.

**Why this priority**: Style management is essential for document structure consistency, especially in long documents with many headings.

**Independent Test**: Call `word_apply_style` on a paragraph, verify the style changes and the formatting updates.

**Acceptance Scenarios**:

1. **Given** paragraph 2 is "Normal" style, **When** developer calls `word_apply_style` with `paragraphIndex: 2, styleName: "Heading 2"`, **Then** paragraph 2 becomes Heading 2 style
2. **Given** a style name that does not exist in the document, **When** developer calls `word_apply_style`, **Then** the tool returns an error with `errorCode: "STYLE_NOT_FOUND"` and a list of available styles
3. **Given** the style is applied, **When** the user views the document, **Then** the paragraph formatting updates immediately

---

### User Story 9 — Get document sections (Priority: P2)

A user asks "How is this document structured?" The LLM calls `word_get_sections` and receives section breaks, page layout info, and headers/footers per section. This provides a structural map of the document.

**Why this priority**: Section awareness is important for complex documents (reports, contracts) but less critical for basic editing workflows.

**Independent Test**: Call `word_get_sections` on a document with known section breaks, verify the response lists sections correctly.

**Acceptance Scenarios**:

1. **Given** a document with 3 sections (each with different page orientation), **When** developer calls `word_get_sections`, **Then** the response lists 3 sections with their page layout properties
2. **Given** a single-section document, **When** developer calls `word_get_sections`, **Then** the response contains one section with default layout values
3. **Given** sections with different header/footer configurations, **When** developer calls `word_get_sections`, **Then** each section includes its header/footer setup (different first page, odd/even pages)

---

### User Story 10 — Insert list (Priority: P2)

A user asks "Add a bulleted list with these items." The LLM calls `word_insert_list` with list type (bulleted/numbered) and an items array. The list is inserted at the specified location with tracked changes.

**Why this priority**: List creation is a common formatting operation but can be achieved through multiple paragraph insertions as a fallback.

**Independent Test**: Call `word_insert_list` with known items, verify a properly formatted list appears at the specified location.

**Acceptance Scenarios**:

1. **Given** a list of 5 items and `type: "bulleted"`, **When** developer calls `word_insert_list` with `afterParagraphIndex: 3`, **Then** a bulleted list with 5 items appears after paragraph 3
2. **Given** `type: "numbered"`, **When** developer calls `word_insert_list`, **Then** the list uses numbered formatting (1., 2., 3., etc.)
3. **Given** an empty items array, **When** developer calls `word_insert_list`, **Then** the tool returns an error with `errorCode: "EMPTY_ITEMS"`

---

### Edge Cases

- What happens when a table has merged cells? The read tool returns content as-is and notes merged regions. The write tool targets the merge origin cell.
- What happens when nested tables exist? The read tool returns the outer table only and notes nesting. Nested table access is not supported in this phase.
- What happens when `word_replace_selection` is called with an empty selection? The tool returns an error with `errorCode: "EMPTY_SELECTION"`.
- What happens when `word_apply_style` is called with a non-existent style name? The tool returns an error with `errorCode: "STYLE_NOT_FOUND"` and a list of available styles in the document.
- What happens when `word_insert_image` receives an image exceeding 10MB? The tool rejects the operation with `errorCode: "IMAGE_TOO_LARGE"`.
- What happens when headers/footers contain page number fields or other field codes? The tools preserve field codes and represent them as placeholders (e.g., `{{PAGE}}`, `{{DATE}}`).

## Requirements

### Functional Requirements

- **FR-001**: The MCP server MUST expose `word_get_tables` returning table index, row count, column count, and cell text for each table in the document.
- **FR-002**: The MCP server MUST expose `word_insert_table` supporting location by paragraph index or `"end"`.
- **FR-003**: `word_insert_table` MUST optionally populate the header row from a string array parameter `headerRow`.
- **FR-004**: The MCP server MUST expose `word_update_table_cell` using tracked changes (`changeTrackingMode: "TrackMineOnly"`).
- **FR-005**: The MCP server MUST expose `word_get_headers_footers` returning content for default, first page, and odd/even headers per section.
- **FR-006**: The MCP server MUST expose `word_set_header_footer` supporting text content for header and footer types.
- **FR-007**: The MCP server MUST expose `word_replace_selection` using tracked changes — the user accepts or rejects via the Review ribbon.
- **FR-008**: The MCP server MUST expose `word_apply_style` validating that the style exists in the document before applying.
- **FR-009**: The MCP server MUST expose `word_insert_image` accepting base64-encoded image data and inserting as an inline picture.
- **FR-010**: The MCP server MUST expose `word_insert_list` supporting `"bulleted"` and `"numbered"` list types.
- **FR-011**: ALL write tools MUST use the tracked changes pattern (save mode → `TrackMineOnly` → mutate → restore mode → return `{ tracked: true }`).
- **FR-012**: Table cell coordinates MUST be zero-based (`row: 0, column: 0` = first cell).

### Key Entities

- **Table**: A Word table identified by zero-based index within the document. Has row count, column count, and cell contents. May contain merged cells.
- **Table Cell**: Identified by table index, row (zero-based), and column (zero-based). Contains text content.
- **Header/Footer**: Content region within a document section. Has variants: default, firstPage, oddPages, evenPages. Can contain text and field codes.
- **Section**: A document division with its own page layout and header/footer configuration. Identified by zero-based index.
- **Style**: A named formatting preset available in the document (e.g., "Heading 1", "Normal", "Title"). Must be validated before application.
- **Selection**: The currently selected text range in the Word document. Used by `word_replace_selection` as the mutation target.

## Success Criteria

### Measurable Outcomes

- **SC-001**: A table with 50 rows and 10 columns is read within 3 seconds via `word_get_tables`.
- **SC-002**: A 5x5 table with a header row is inserted within 2 seconds via `word_insert_table`.
- **SC-003**: A selection replacement via `word_replace_selection` creates a tracked change visible in the Review pane.
- **SC-004**: Style application via `word_apply_style` is reflected in the document within 1 second.
- **SC-005**: All 24 existing Word MVP tests continue to pass after integration of the new tools.

## Assumptions

- Word 2019 or later (or Microsoft 365) is available on the developer machine.
- Word JS API `WordApi 1.4+` is available for `changeTrackingMode` support.
- The add-in runs in the same process space as Word.
- Headers/footers with complex field codes are preserved as-is; full field manipulation is out of scope.
- Image insertion supports PNG, JPEG, GIF, BMP, and SVG formats via base64 encoding.
- Nested table access is not supported in this phase — only top-level tables are returned.
