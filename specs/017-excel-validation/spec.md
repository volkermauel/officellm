# Feature Specification: Phase 17 — Excel Data Validation & Structured Data

**Feature Branch**: `017-excel-validation`

**Created**: 2026-05-16

**Status**: Draft

**Depends on**: Phase 3 (Excel MVP), Phase 10 (Excel Analysis), Phase 16 (Excel Navigation)

**Input**: User description: "Extend Excel with data validation (dropdowns, number constraints, date ranges, custom formulas), input messages, and error alerts. Enables LLM-generated structured data templates with built-in data integrity rules."

## Architecture

Uses `Excel.run()` context batching with the `range.dataValidation` API. Supports all validation types: list, whole number, decimal, date, text length, and custom formula. Input messages and error alerts provide user guidance.

```
LLM Client                     MCP Server (port 3000)              Office Add-in
    │                                │                               │
    │── tools/call ─────────────────►│                               │
    │   { excel_add_data_validation  │── push command (SignalR) ────►│
    │     { instanceId,              │                               │── range.dataValidation
    │       sheetName,               │                               │   .type = "List"
    │       address: "D2:D100",      │                               │   .rule = {list: {...}}
    │       type: "list",            │                               │── .prompt = {title, message}
    │       formula1: "Yes,No,Maybe",│                               │── .errorAlert = {title, msg}
    │       showInputMessage: true,  │◄── { applied: true } ────────│
    │       inputTitle: "Select",    │                               │
    │       inputMessage: "Pick one"│                               │
    │     }}                         │                               │
    │◄── { applied: true, cells: 99 }│                               │
```

## User Scenarios & Testing

### User Story 1 — Dropdown list validation (Priority: P0)

A user asks "Add a Status dropdown with options: Active, Inactive, Pending, Review." The LLM calls `excel_add_data_validation` with `type: "list"`.

**Why this priority**: Dropdowns are the most common validation. They prevent typos and standardize data entry.

**Independent Test**: Apply list validation, verify dropdown appears and invalid entries are rejected.

**Acceptance Scenarios**:

1. **Given** range D2:D100 on "Data" sheet, **When** developer calls `excel_add_data_validation` with `type: "list", formula1: "Active,Inactive,Pending,Review"`, **Then** cells show a dropdown with 4 options and invalid values are rejected
2. **Given** a "Categories" sheet with values in A1:A10, **When** developer calls `excel_add_data_validation` with `type: "list", formula1: "=Categories!$A$1:$A$10"`, **Then** the dropdown shows values from the referenced range
3. **Given** a range with list validation, **When** developer calls `excel_remove_data_validation`, **Then** the dropdown is removed and any value can be entered

---

### User Story 2 — Number range validation (Priority: P0)

A user asks "Make sure the Age column only accepts numbers 0–150." The LLM calls `excel_add_data_validation` with `type: "wholeNumber"`.

**Why this priority**: Number constraints prevent nonsensical data. Critical for financial and scientific data.

**Independent Test**: Apply number validation, verify out-of-range values are rejected.

**Acceptance Scenarios**:

1. **Given** range C2:C100 (Age column), **When** developer calls `excel_add_data_validation` with `type: "wholeNumber", operator: "between", formula1: "0", formula2: "150"`, **Then** only integers 0–150 are accepted
2. **Given** range E2:E100 (Salary column), **When** developer calls `excel_add_data_validation` with `type: "decimal", operator: "greaterThanOrEqual", formula1: "0"`, **Then** any non-negative decimal is accepted
3. **Given** a cell with number validation, **When** user enters text, **Then** Excel rejects the input

---

### User Story 3 — Date validation (Priority: P1)

A user asks "The Start Date must be between Jan 1 2024 and Dec 31 2025." The LLM calls `excel_add_data_validation` with `type: "date"`.

**Why this priority**: Date validation prevents impossible dates and enforces project timelines.

**Independent Test**: Apply date validation, verify out-of-range dates are rejected.

**Acceptance Scenarios**:

1. **Given** range F2:F100 (Start Date column), **When** developer calls `excel_add_data_validation` with `type: "date", operator: "between", formula1: "2024-01-01", formula2: "2025-12-31"`, **Then** only dates in 2024–2025 are accepted
2. **Given** a date validation with `operator: "greaterThan", formula1: "2024-01-01"`, **Then** any date after Jan 1, 2024 is accepted

---

### User Story 4 — Input messages and error alerts (Priority: P1)

A user asks "When someone clicks the Status cell, show a hint. When they enter something wrong, show 'Please select from the list'." The LLM specifies `showInputMessage` and `showErrorMessage`.

**Why this priority**: Input messages guide users without documentation. Error alerts provide immediate feedback.

**Acceptance Scenarios**:

1. **Given** a range with `showInputMessage: true, inputTitle: "Status", inputMessage: "Select the current status"`, **When** user selects a cell in the range, **Then** a tooltip appears with the title and message
2. **Given** a range with `showErrorMessage: true, errorTitle: "Invalid Status", errorMessage: "Please select from the dropdown list", errorStyle: "stop"`, **When** user enters invalid data, **Then** an error dialog appears and the entry is rejected
3. **Given** `errorStyle: "warning"`, **When** user enters invalid data, **Then** a warning appears but the user can choose to proceed
4. **Given** `errorStyle: "information"`, **When** user enters invalid data, **Then** an info message appears but entry is allowed

---

### User Story 5 — Custom formula validation (Priority: P2)

A user asks "Make sure column B is always greater than column A." The LLM calls `excel_add_data_validation` with `type: "custom"` and a formula.

**Why this priority**: Custom formulas enable cross-cell validation rules. Powerful but less common.

**Acceptance Scenarios**:

1. **Given** range B2:B100, **When** developer calls `excel_add_data_validation` with `type: "custom", formula1: "=B2>A2"`, **Then** only values where B > A are accepted
2. **Given** a custom formula that references another sheet, **When** validation is applied, **Then** the cross-sheet reference works correctly

---

### Edge Cases

- **Validation on merged cells**: Applies to the entire merged area
- **Validation on cells with existing data**: Validation is applied but existing invalid data is NOT flagged until edited
- **Removing validation**: `excel_remove_data_validation` clears all validation rules for the range
- **Formula limits**: List validation with hardcoded values supports up to ~255 characters in the formula string
- **Protected sheets**: Cannot apply validation to locked cells on protected sheets

## Requirements

### Functional Requirements

- **FR-001**: The MCP server MUST expose `excel_add_data_validation` accepting `instanceId`, `sheetName`, `address`, `type` ("list"|"wholeNumber"|"decimal"|"date"|"textLength"|"custom"), `operator` ("between"|"notBetween"|"equalTo"|"notEqualTo"|"greaterThan"|"lessThan"|"greaterThanOrEqual"|"lessThanOrEqual"), `formula1`, optional `formula2`, and optional `showInputMessage`, `inputTitle`, `inputMessage`, `showErrorMessage`, `errorTitle`, `errorMessage`, `errorStyle` ("stop"|"warning"|"information").
- **FR-002**: The MCP server MUST expose `excel_remove_data_validation` accepting `instanceId`, `sheetName`, `address`.
- **FR-003**: All validation rules MUST be undoable via Ctrl+Z.
- **FR-004**: Date formulas MUST accept ISO 8601 format strings ("2024-01-01") and convert to Excel serial dates.
- **FR-005**: The tool MUST support both hardcoded lists ("A,B,C") and range references ("=Sheet!$A$1:$A$10") for list-type validation.

### Key Entities

- **Data Validation Rule**: A constraint on cell input with type, operator, formulas, input message, and error alert.
- **Validation Type**: The kind of data being validated: list (dropdown), wholeNumber, decimal, date, textLength, custom (formula).
- **Error Style**: "stop" (block), "warning" (allow with warning), "information" (inform but allow).

## Success Criteria

### Measurable Outcomes

- **SC-001**: List validation applied to 1,000 cells within 2 seconds.
- **SC-002**: Input messages appear immediately on cell selection.
- **SC-003**: Error alerts fire on invalid entry with the specified title and message.
- **SC-004**: All existing Excel tests continue to pass.

## Assumptions

- Excel JS API `range.dataValidation` is available (ExcelApi 1.8+).
- Date validation formulas use Excel serial date numbers internally.
- The add-in runs in the same process space as Excel.
