# Feature Specification: Phase 3 — Excel MVP

**Feature Branch**: `004-excel-mvp`

**Created**: 2026-05-14

**Status**: Draft

**Input**: User description: "Implement workbook map, read range, write range and write formula tools. Add range-size limits and formula validation. Add confirmation preview for any write operation."

## Architecture

The MCP server acts as a central hub. The **unified Office JS Add-in** auto-detects the host via `Office.onReady(info.host)`. For Excel instances, the add-in registers with host type `Excel` and instance IDs like `excel_1`, `excel_2`. Tools accept optional `instanceId` parameter to target specific instances.

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

### User Story 1 — Analyst explores workbook structure (Priority: P1)

A data analyst opens a complex Excel workbook with multiple sheets, tables, and named ranges. They ask Open WebUI to describe the workbook's structure — which sheets exist, what tables are defined, what named ranges are available, and the size of used ranges on each sheet. This gives them a map without manually clicking through tabs.

**Why this priority**: Excel workbooks can have dozens of sheets and hundreds of named ranges. A workbook map tool is essential for the model to understand where to operate before performing any read or write.

**Independent Test**: Call `excel_get_workbook_map` and verify the response lists all sheets, tables, named ranges, and used range sizes.

**Acceptance Scenarios**:

1. **Given** a workbook with 5 sheets, 3 tables, and 2 named ranges, **When** developer calls `excel_get_workbook_map`, **Then** the response includes sheet names, table names with their ranges, named range definitions, and used range dimensions for each sheet
2. **Given** an empty workbook, **When** developer calls `excel_get_workbook_map`, **Then** the response includes at least one blank sheet with zero used range
3. **Given** hidden sheets exist, **When** `includeHiddenSheets: false` (default), **Then** hidden sheets are excluded from the map

---

### User Story 2 — Analyst reads a bounded data range (Priority: P1)

An analyst wants to examine a specific data range (e.g., `B2:G24`) in the "Q3 Forecast" sheet. They ask Open WebUI to read the range values and formulas, receiving structured JSON that fits within the bounded output limit.

**Why this priority**: Reading bounded ranges is the fundamental data extraction operation. The model needs this to understand the data before suggesting analysis, transformations, or corrections.

**Independent Test**: Call `excel_read_range` with a known range, verify the returned values and formulas match the actual cell contents.

**Acceptance Scenarios**:

1. **Given** a range B2:G24 containing mixed values (numbers, text, formulas), **When** developer calls `excel_read_range` with `includeFormulas: true`, **Then** the response includes both display values and formula strings for formula cells
2. **Given** `includeNumberFormats: true`, **When** developer calls `excel_read_range`, **Then** each cell's number format string is included in the response
3. **Given** a range exceeds the maximum byte limit (256 KB), **When** developer calls `excel_read_range`, **Then** the response is truncated and includes `truncated: true` with the actual byte count

---

### User Story 3 — Analyst writes values to a bounded range with preview (Priority: P1)

An analyst asks Open WebUI to populate an empty column with computed values (e.g., a calculated total). The tool shows which cells will be affected, displays sample before/after values, and requires approval before writing.

**Why this priority**: Writing data is a destructive operation in Excel — it overwrites existing content. The confirmation preview is critical to prevent accidental data loss.

**Independent Test**: Call `excel_write_range` with new values, verify the diff preview shows correct affected cells, approve, and verify values are written.

**Acceptance Scenarios**:

1. **Given** an empty column D, **When** developer calls `excel_write_range` with computed values for D2:D24, **Then** the task pane shows a sample of before/after values and the exact cell addresses affected
2. **Given** user approves, **When** the write is applied, **Then** the specified cells contain the new values
3. **Given** the target range overlaps with existing data, **When** diff preview is shown, **Then** both old and new values are displayed for overlapping cells
4. **Given** user rejects, **When** no action is taken, **Then** all cells remain unchanged

---

### User Story 4 — Analyst writes formulas with validation (Priority: P2)

An analyst asks Open WebUI to insert a formula (e.g., `=SUM(B2:B24)`) into a range. The tool validates the formula syntax before showing the preview, and rejects obviously invalid formulas.

**Why this priority**: Formula insertion is a common Excel automation task. Invalid formulas would corrupt the worksheet, so validation is essential.

**Independent Test**: Call `excel_write_formula` with valid and invalid formulas, verify validation catches errors and previews show computed results where possible.

**Acceptance Scenarios**:

1. **Given** a valid formula `=SUM(B2:B24)`, **When** developer calls `excel_write_formula`, **Then** the tool validates the syntax and shows a preview of computed values
2. **Given** an invalid formula `=SUM(B2:B24`, **When** developer calls `excel_write_formula`, **Then** the tool returns an error with `errorCode: "INVALID_FORMULA"` and a human-readable message
3. **Given** a formula referencing a different sheet, **When** developer calls `excel_write_formula`, **Then** the formula is written with proper sheet qualification (e.g., `='Q1 Data'!A1`)

---

### User Story 5 — Analyst creates an Excel table from data (Priority: P2)

An analyst selects or specifies a data range and asks Open WebUI to convert it into a formatted Excel table with headers and auto-filtering. The table creation is shown as a diff preview before application.

**Why this priority**: Tables are a fundamental Excel feature for data organization and analysis. Converting ranges to tables is a common post-processing step after LLM-assisted data work.

**Independent Test**: Call `excel_create_table` on a specified range, verify an Excel table appears with correct headers and filtering.

**Acceptance Scenarios**:

1. **Given** a range A1:D100 with a header row, **When** developer calls `excel_create_table`, **Then** an Excel ListObject is created spanning A1:D100 with the first row as headers
2. **Given** no clear header row exists, **When** developer calls `excel_create_table`, **Then** the tool generates generic headers (Column1, Column2, etc.) and shows them in the preview
3. **Given** user approves, **When** the table is created, **Then** auto-filter buttons appear on the header row

---

### Edge Cases

- What happens when a formula references a deleted or renamed sheet? The write tool should detect this and return a recoverable error with suggestions.
- What happens when the user tries to write a range larger than 1 million cells? The tool should reject the operation with `errorCode: "RANGE_TOO_LARGE"` and a suggested smaller range.
- What happens when a cell contains an error value (#N/A, #DIV/0!)? The read tool should preserve and report error values distinctly from normal values.
- What happens when the workbook is in Formula Auditing mode or has circular reference warnings? The write tool should proceed normally but note any pre-existing warnings in the response.

## Requirements

### Functional Requirements

- **FR-001**: The MCP server MUST expose `excel_get_workbook_map` returning sheet names, table definitions (name, range), named ranges, and used range dimensions for each sheet.
- **FR-002**: The MCP server MUST expose `excel_read_range` reading cells, formulas, and optional number formats from a bounded range specified by sheet name and address.
- **FR-003**: `excel_read_range` MUST accept parameters: `sheetName` (required), `address` (required), `includeFormulas` (default: true), `includeNumberFormats` (default: false).
- **FR-004**: The MCP server MUST expose `excel_write_range` writing values to a bounded range with confirmation preview.
- **FR-005**: The MCP server MUST expose `excel_write_formula` writing formulas with syntax validation and preview.
- **FR-006**: The MCP server MUST expose `excel_create_table` creating a formatted Excel table (ListObject) from a specified range.
- **FR-007**: All write operations MUST return `requiresConfirmation: true` with a diff preview showing affected cell addresses and sample before/after values.
- **FR-008**: Range reads MUST enforce a default output limit of 32 KB and a maximum of 256 KB, returning `truncated: true` when exceeded.
- **FR-009**: Formula writes MUST validate syntax before showing the preview and return an error for invalid formulas.
- **FR-010**: Write operations MUST respect Excel's native undo — approved changes must be undoable via Ctrl+Z.

### Key Entities

- **Sheet**: A worksheet within a workbook. Identified by name. Has a used range, tables, and named ranges.
- **Range**: A rectangular set of cells identified by sheet name and address (e.g., "B2:G24"). Contains values, formulas, and formatting.
- **Table (ListObject)**: A structured Excel table with headers, auto-filtering, and programmatic access. Spans a contiguous range.
- **Named Range**: A named reference to a cell or range, scoped to the workbook or a specific sheet.
- **Formula Validation Result**: The outcome of parsing and validating an Excel formula string, including syntax errors and referenced ranges.

## Success Criteria

### Measurable Outcomes

- **SC-001**: A developer can retrieve the workbook map of a 20-sheet workbook from Open WebUI within 5 seconds.
- **SC-002**: A range read of 1000 cells returns within 2 seconds with accurate values and formulas.
- **SC-003**: Any write operation requires explicit user approval via the task pane before modifying cells.
- **SC-004**: Invalid formulas are rejected before any cell is modified, with a clear error message.
- **SC-005**: Approved writes are undoable via Excel's native Ctrl+Z undo.

## Assumptions

- Excel 2019 or later (or Microsoft 365) is available on the developer machine.
- Range addresses use A1-style notation (not R1C1).
- Named ranges are workbook-scoped by default unless sheet-scoped.
- The add-in runs in the same process space as Excel.
