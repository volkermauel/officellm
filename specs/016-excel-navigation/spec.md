# Feature Specification: Phase 16 — Excel Navigation & Named Ranges

**Feature Branch**: `016-excel-navigation`

**Created**: 2026-05-16

**Status**: Draft

**Depends on**: Phase 3 (Excel MVP), Phase 10 (Excel Analysis)

**Input**: User description: "Add freeze panes, named ranges, and data validation tools for Excel. These improve navigation, readability, and data integrity in LLM-generated workbooks."

## Architecture

Uses `Excel.run()` context batching. Freeze panes and named ranges use the worksheet/range APIs. Data validation uses `range.dataValidation`.

```
LLM Client                     MCP Server (port 3000)              Office Add-in
    │                                │                               │
    │── tools/call ─────────────────►│                               │
    │   { excel_freeze_panes         │── push command (SignalR) ────►│── worksheet.freezePanes
    │     { instanceId,              │                               │   or range.addNamedItem
    │       sheetName,               │                               │   or range.dataValidation
    │       at: "A2" }}              │◄── { frozen: true } ──────────│
    │◄── { frozen: true } ──────────│                               │
```

## User Scenarios & Testing

### User Story 1 — Freeze panes for navigation (Priority: P0)

A user asks "Freeze the header row so it stays visible when scrolling." The LLM calls `excel_freeze_panes` with the cell address to freeze above/left of.

**Why this priority**: Freeze panes is the #1 navigation aid for large spreadsheets. Without it, scrolling loses context. An LLM generating reports should freeze headers automatically.

**Independent Test**: Create a large sheet, freeze at row 2, verify header row stays visible when scrolling.

**Acceptance Scenarios**:

1. **Given** a sheet with headers in row 1 and 500 data rows, **When** developer calls `excel_freeze_panes` with `at: "A2"`, **Then** row 1 stays frozen at the top when scrolling
2. **Given** a sheet with frozen panes, **When** developer calls `excel_freeze_panes` with `action: "unfreeze"`, **Then** the freeze is removed
3. **Given** a sheet with data in columns A–F and row headers in column A, **When** developer calls `excel_freeze_panes` with `at: "B2"`, **Then** both row 1 and column A are frozen

---

### User Story 2 — Named ranges for readability (Priority: P0)

A user asks "Name the sales data range 'Q1Sales' so I can reference it easily." The LLM calls `excel_add_named_range` to create a named range. Later, `excel_get_named_ranges` lists all named ranges.

**Why this priority**: Named ranges transform spreadsheet navigation. Instead of "$B$2:$D$100", you get "Q1Sales". This makes LLM-generated formulas readable and self-documenting.

**Independent Test**: Add a named range, retrieve it, verify the name maps to the correct address.

**Acceptance Scenarios**:

1. **Given** a range B2:D100 on "SalesData" sheet, **When** developer calls `excel_add_named_range` with `name: "Q1Sales", sheetName: "SalesData", address: "B2:D100"`, **Then** the named range "Q1Sales" is created and visible in the Name Box
2. **Given** a workbook with 5 named ranges, **When** developer calls `excel_get_named_ranges`, **Then** all 5 named ranges are listed with their addresses and scope (worksheet or workbook)
3. **Given** a named range "Q1Sales" already exists, **When** developer calls `excel_add_named_range` with the same name, **Then** the range is updated to the new address

---

### User Story 3 — Data validation (dropdown lists) (Priority: P1)

A user asks "Add a dropdown to the Status column with options Active, Inactive, Pending." The LLM calls `excel_add_data_validation` to constrain cell input.

**Why this priority**: Data validation prevents data entry errors. Dropdown lists are the most common validation type. LLM-generated templates should include validation automatically.

**Independent Test**: Add a list validation, verify the dropdown appears and rejects invalid values.

**Acceptance Scenarios**:

1. **Given** a Status column (D2:D100), **When** developer calls `excel_add_data_validation` with `type: "list", formula1: "Active,Inactive,Pending"`, **Then** cells D2:D100 show a dropdown with the three options
2. **Given** a range with list validation, **When** a user types "Unknown", **Then** Excel rejects the input with a validation error
3. **Given** a Salary column (E2:E100), **When** developer calls `excel_add_data_validation` with `type: "wholeNumber", operator: "between", formula1: "0", formula2: "1000000"`, **Then** only integers between 0 and 1,000,000 are accepted

---

### User Story 4 — Remove data validation (Priority: P2)

A user asks "Remove the validation from column D." The LLM calls `excel_remove_data_validation`.

**Why this priority**: Cleanup tool. Less critical but needed for removing incorrect validation.

**Acceptance Scenarios**:

1. **Given** a range with data validation, **When** developer calls `excel_remove_data_validation`, **Then** the validation is removed and any value can be entered

---

### Edge Cases

- **Freeze panes on protected sheets**: Returns error if sheet protection prevents freeze
- **Named range scope**: Names can be workbook-scoped (default) or worksheet-scoped. Workbook-scoped names are accessible from any sheet.
- **Validation on merged cells**: Validation applies to the entire merged area
- **Circular named ranges**: Name cannot reference itself
- **Data validation formula limits**: List validation supports comma-separated values or range references (e.g., `=Sheet1!$F$1:$F$10`)

## Requirements

### Functional Requirements

- **FR-001**: The MCP server MUST expose `excel_freeze_panes` accepting `instanceId`, `sheetName`, `at` (cell address), and optional `action` ("freeze"|"unfreeze").
- **FR-002**: The MCP server MUST expose `excel_get_named_ranges` accepting `instanceId` and optional `sheetName`.
- **FR-003**: The MCP server MUST expose `excel_add_named_range` accepting `instanceId`, `name`, `sheetName`, `address`, and optional `comment`.
- **FR-004**: The MCP server MUST expose `excel_add_data_validation` accepting `instanceId`, `sheetName`, `address`, `type` ("list"|"wholeNumber"|"decimal"|"date"|"textLength"), `operator`, `formula1`, `formula2`, `showErrorMessage`, `errorTitle`, `errorMessage`.
- **FR-005**: The MCP server MUST expose `excel_remove_data_validation` accepting `instanceId`, `sheetName`, `address`.
- **FR-006**: All tools MUST be undoable via Ctrl+Z.

### Key Entities

- **Frozen Panes**: A view state that keeps specific rows/columns visible while scrolling. Identified by the cell address to freeze above/left of.
- **Named Range**: A human-readable name mapped to a cell range. Has a scope (workbook or worksheet) and optional comment.
- **Data Validation**: A constraint on cell input. Types include list (dropdown), number range, date range, and text length. Can show custom error messages.

## Success Criteria

### Measurable Outcomes

- **SC-001**: Freeze panes applied within 1 second.
- **SC-002**: Named range creation within 1 second.
- **SC-003**: Data validation applied to 1,000 cells within 2 seconds.
- **SC-004**: All existing Excel tests continue to pass.

## Assumptions

- Excel JS API `worksheet.freezePanes` is available (ExcelApi 1.7+).
- Named ranges via `workbook.names.add` or `worksheet.names.add` (ExcelApi 1.4+).
- Data validation via `range.dataValidation` (ExcelApi 1.8+).
- The add-in runs in the same process space as Excel.
