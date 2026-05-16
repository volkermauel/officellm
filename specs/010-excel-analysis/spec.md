# Feature Specification: Phase 10 — Excel Analysis Tools

**Feature Branch**: `010-excel-analysis`

**Created**: 2026-05-16

**Status**: Draft

**Depends on**: Phase 3 (Excel MVP) — basic read/write must exist

**Input**: User description: "Implement worksheet management, sort, filter, chart creation/reading, conditional formatting, cell formatting, and pivot table tools for Excel. All tools follow the existing Excel.run() batching pattern and are undoable via Ctrl+Z."

## Architecture

Same as Excel MVP — uses `Excel.run()` context batching. All new tools follow the existing pattern: backend tool definition in `McpToolEngine.cs`, frontend handler in `excel-commands.ts`.

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

### User Story 1 — Manage worksheets (Priority: P0)

A user asks "Create a new sheet called Summary" or "Delete the Scratch sheet." The LLM calls `excel_add_sheet`, `excel_delete_sheet`, or `excel_rename_sheet` to organize the workbook structure without manual tab management.

**Why this priority**: Sheet management is foundational. Before performing analysis (charts, pivot tables, formatting), the user often needs to create target sheets, rename existing ones, or clean up unused tabs. Without this, every downstream tool is limited to the current sheet layout.

**Independent Test**: Call `excel_add_sheet` with a name, verify the sheet appears. Call `excel_rename_sheet`, verify the name changes. Call `excel_delete_sheet`, verify it is removed.

**Acceptance Scenarios**:

1. **Given** a workbook with 3 sheets, **When** developer calls `excel_add_sheet` with `name: "Summary"`, **Then** a new sheet named "Summary" is added at the end of the sheet tab list
2. **Given** a workbook with sheets ["Data", "Scratch", "Report"], **When** developer calls `excel_delete_sheet` with `sheetName: "Scratch"`, **Then** the "Scratch" sheet is removed and the workbook has 2 sheets
3. **Given** a workbook with sheets ["Sheet1", "Sheet2"], **When** developer calls `excel_rename_sheet` with `sheetName: "Sheet1"` and `newName: "Data"`, **Then** the sheet is renamed to "Data" and "Sheet2" remains unchanged
4. **Given** a workbook with only one sheet, **When** developer calls `excel_delete_sheet` on that sheet, **Then** the tool returns an error with `errorCode: "CANNOT_DELETE_LAST_SHEET"`
5. **Given** a workbook with a sheet named "Data", **When** developer calls `excel_rename_sheet` with `newName: "Data"`, **Then** the tool returns an error with `errorCode: "DUPLICATE_SHEET_NAME"`

---

### User Story 2 — Sort data range (Priority: P0)

A user asks "Sort by column B descending, then column A ascending." The LLM calls `excel_sort_range` with a sort criteria array supporting multi-column sorting.

**Why this priority**: Sorting is one of the most common data operations. Multi-column sort is essential for preparing data before analysis, charting, or pivot table creation.

**Independent Test**: Populate a range with mixed data, call `excel_sort_range` with two sort criteria, verify the rows are ordered correctly.

**Acceptance Scenarios**:

1. **Given** a range A1:D100 with data, **When** developer calls `excel_sort_range` with criteria `[{ column: 1, ascending: true }]`, **Then** rows are sorted by column A in ascending order
2. **Given** a range A1:D100, **When** developer calls `excel_sort_range` with criteria `[{ column: 2, ascending: false }, { column: 1, ascending: true }]`, **Then** rows are sorted by column B descending, then column A ascending as tiebreaker
3. **Given** a range that is part of an Excel table, **When** developer calls `excel_sort_range`, **Then** the tool uses the table sort API and sorts the entire table
4. **Given** a single-row range, **When** developer calls `excel_sort_range`, **Then** the tool returns an error with `errorCode: "RANGE_TOO_SMALL"`

---

### User Story 3 — Filter data range (Priority: P0)

A user asks "Show only rows where Status equals 'Active'." The LLM calls `excel_filter_range` to apply autofilter with criteria, hiding non-matching rows.

**Why this priority**: Filtering is essential for focused analysis. Users frequently need to isolate subsets of data before charting, formatting, or summarizing.

**Independent Test**: Populate a range with varied values, apply a filter, verify only matching rows are visible.

**Acceptance Scenarios**:

1. **Given** a range A1:D50 with a "Status" column (column C), **When** developer calls `excel_filter_range` with `column: 3, criteria: { value: "Active" }`, **Then** only rows where column C equals "Active" are visible
2. **Given** a filtered range, **When** developer calls `excel_filter_range` with `clearFilters: true`, **Then** all rows become visible again
3. **Given** a range that is part of an Excel table, **When** developer calls `excel_filter_range`, **Then** the filter is applied to the table's autofilter
4. **Given** a range without headers, **When** developer calls `excel_filter_range`, **Then** the tool returns an error with `errorCode: "NO_HEADER_ROW"`

---

### User Story 4 — Create chart from data (Priority: P1)

A user asks "Create a bar chart from the sales data." The LLM calls `excel_create_chart` specifying chart type, data range, and position. The chart is placed next to the data without overlapping.

**Why this priority**: Charting is the primary way users visualize data in Excel. Automating chart creation from natural language is a high-value feature for analysts.

**Independent Test**: Call `excel_create_chart` with known data, verify the chart appears with correct type and data range.

**Acceptance Scenarios**:

1. **Given** a range A1:B10 with headers "Month" and "Revenue", **When** developer calls `excel_create_chart` with `chartType: "Column"`, **Then** a column chart is created showing revenue by month
2. **Given** existing data in range A1:D50, **When** developer calls `excel_create_chart`, **Then** the chart is positioned to the right of column D, avoiding overlap with data
3. **Given** an unsupported chart type "Radar", **When** developer calls `excel_create_chart`, **Then** the tool returns an error with `errorCode: "UNSUPPORTED_CHART_TYPE"` listing supported types
4. **Given** a non-contiguous range, **When** developer calls `excel_create_chart`, **Then** the tool returns an error with `errorCode: "NON_CONTIGUOUS_RANGE"`
5. **Given** `chartType` is one of "Column", "Bar", "Line", "Pie", "Scatter", "Area", or "Doughnut", **When** developer calls `excel_create_chart`, **Then** the chart is created successfully

---

### User Story 5 — Read chart properties (Priority: P2)

A user asks "What charts are in this workbook?" The LLM calls `excel_get_charts` to list all charts with their types, data ranges, titles, and positions.

**Why this priority**: Chart inspection is needed before modifying or recreating charts. It allows the LLM to understand the current visualization state.

**Independent Test**: Create charts manually, call `excel_get_charts`, verify the returned list matches.

**Acceptance Scenarios**:

1. **Given** a workbook with 3 charts across 2 sheets, **When** developer calls `excel_get_charts`, **Then** the response lists all 3 charts with their sheet name, chart type, title, data range, and position (top/left/width/height)
2. **Given** a workbook with no charts, **When** developer calls `excel_get_charts`, **Then** the response returns an empty array `[]`
3. **Given** a chart with no title set, **When** developer calls `excel_get_charts`, **Then** the chart's `title` field is `null`

---

### User Story 6 — Apply conditional formatting (Priority: P1)

A user asks "Highlight cells over 1000 in red." The LLM calls `excel_apply_conditional_formatting` with range, rule type, and style configuration.

**Why this priority**: Conditional formatting enables visual data analysis at scale. Users frequently want to highlight outliers, apply data bars, or use color scales without manually configuring rules.

**Independent Test**: Call `excel_apply_conditional_formatting` with a cell-value rule, verify the formatting appears on matching cells.

**Acceptance Scenarios**:

1. **Given** a range B2:B100 with numeric values, **When** developer calls `excel_apply_conditional_formatting` with `ruleType: "cellValue", operator: "greaterThan", value: 1000, format: { fillColor: "red" }`, **Then** cells with values > 1000 are highlighted in red
2. **Given** a range C2:C50, **When** developer calls `excel_apply_conditional_formatting` with `ruleType: "dataBar"`, **Then** data bars are applied to the range showing relative values
3. **Given** a range D2:D50, **When** developer calls `excel_apply_conditional_formatting` with `ruleType: "colorScale", minColor: "green", maxColor: "red"`, **Then** a two-color gradient scale is applied
4. **Given** a range E2:E50, **When** developer calls `excel_apply_conditional_formatting` with `ruleType: "iconSet", iconSet: "3TrafficLights"`, **Then** traffic light icons appear next to values
5. **Given** conflicting conditional formatting rules on the same range, **When** a new rule is applied, **Then** the new rule takes precedence (last applied wins)

---

### User Story 7 — Format cell range (Priority: P1)

A user asks "Make the header row bold with a blue background." The LLM calls `excel_format_range` with font, fill, border, and alignment options.

**Why this priority**: Formatting is essential for creating professional-looking reports. Automating header formatting, number formatting, and visual styling saves significant manual effort.

**Independent Test**: Call `excel_format_range` with formatting options, verify the cells display the specified styles.

**Acceptance Scenarios**:

1. **Given** a range A1:D1 (header row), **When** developer calls `excel_format_range` with `font: { bold: true }, fill: { color: "#4472C4" }`, **Then** the header row is bold with a blue background
2. **Given** a range B2:B100 with raw numbers, **When** developer calls `excel_format_range` with `numberFormat: "#,##0.00"`, **Then** all cells display with two decimal places and thousand separators
3. **Given** a range A1:D10, **When** developer calls `excel_format_range` with `borders: { style: "thin", color: "black" }`, **Then** thin black borders appear around all cells in the range
4. **Given** a range C1:C20, **When** developer calls `excel_format_range` with `alignment: { horizontal: "center", wrapText: true }`, **Then** text is centered and wraps within cells
5. **Given** a range with mixed content, **When** developer calls `excel_format_range` with `font: { name: "Calibri", size: 11, italic: true, color: "red" }`, **Then** all cells in the range use Calibri 11pt italic red font

---

### User Story 8 — Create pivot table (Priority: P2)

A user asks "Create a pivot table summarizing sales by region." The LLM calls `excel_create_pivottable` specifying source range, row fields, column fields, and value fields with aggregation.

**Why this priority**: Pivot tables are Excel's most powerful analysis feature. Enabling LLM-driven pivot table creation unlocks automated summarization and cross-tabulation.

**Independent Test**: Call `excel_create_pivottable` with known data, verify the pivot table appears on a new sheet with correct field layout.

**Acceptance Scenarios**:

1. **Given** a data range A1:E1000 with columns "Region", "Product", "Quarter", "Sales", "Quantity", **When** developer calls `excel_create_pivottable` with `rows: ["Region"], values: [{ field: "Sales", aggregation: "sum" }]`, **Then** a pivot table is created on a new sheet showing sum of sales by region
2. **Given** source data with >1,000,000 rows, **When** developer calls `excel_create_pivottable`, **Then** the tool returns a warning `performanceWarning: true` noting potential slowdown but proceeds with creation
3. **Given** a source range that includes blank rows, **When** developer calls `excel_create_pivottable`, **Then** the pivot table is created and blank rows are excluded from aggregation
4. **Given** rows, columns, and values fields, **When** developer calls `excel_create_pivottable`, **Then** the pivot table has the correct layout with rows on the left, columns across the top, and values in the data area

---

### Edge Cases

- Sorting a range that's part of a table — the tool must detect this and use the table sort API instead of range sort.
- Creating a chart with non-contiguous data — the Excel JS API does not support non-contiguous ranges for chart source data; return error `NON_CONTIGUOUS_RANGE`.
- Pivot table with >1M rows — warn about performance but allow creation.
- Deleting the last sheet — Excel requires at least one visible sheet; return error `CANNOT_DELETE_LAST_SHEET`.
- Conditional formatting rule conflicts — last applied rule wins (document this behavior).
- Renaming a sheet to a name that already exists — return error `DUPLICATE_SHEET_NAME`.
- Formatting a merged cell range — apply formatting to the entire merged area.
- Creating a pivot table on data with no headers — return error `NO_HEADER_ROW`.

## Requirements

### Functional Requirements

- **FR-001**: The MCP server MUST expose `excel_add_sheet` accepting optional `name` (default: auto-generated) and `position` (default: end of tab list).
- **FR-002**: The MCP server MUST expose `excel_delete_sheet` accepting `sheetName` (required) and confirming the sheet exists before deleting; MUST reject deletion of the last sheet.
- **FR-003**: The MCP server MUST expose `excel_rename_sheet` accepting `sheetName` (required) and `newName` (required); MUST validate name uniqueness.
- **FR-004**: The MCP server MUST expose `excel_sort_range` accepting `sheetName`, `address`, and `criteria` (array of `{ column: number, ascending: boolean }`) supporting multi-column sort.
- **FR-005**: The MCP server MUST expose `excel_filter_range` accepting `sheetName`, `address`, and filter criteria; MUST apply autofilter to the specified range with column-value criteria.
- **FR-006**: The MCP server MUST expose `excel_create_chart` accepting `chartType` (Column, Bar, Line, Pie, Scatter, Area, Doughnut), `dataRange`, and optional `title`; MUST auto-position chart to avoid overlapping data.
- **FR-007**: The MCP server MUST expose `excel_get_charts` returning an array of charts with `sheetName`, `chartType`, `title`, `dataRange`, and `position` (top, left, width, height) for each chart.
- **FR-008**: The MCP server MUST expose `excel_apply_conditional_formatting` accepting `sheetName`, `address`, `ruleType` (dataBar, colorScale, iconSet, cellValue), and rule-specific parameters with style configuration.
- **FR-009**: The MCP server MUST expose `excel_format_range` accepting `sheetName`, `address`, and formatting options: font (name, size, bold, italic, color), fill (color, pattern), borders (style, color), alignment (horizontal, vertical, wrapText), and numberFormat.
- **FR-010**: The MCP server MUST expose `excel_create_pivottable` accepting `sourceRange`, `rows`, optional `columns`, and `values` (array of `{ field, aggregation }`); MUST create the pivot table on a new sheet.
- **FR-011**: All write tools MUST be undoable via Ctrl+Z (Excel's native undo grouped per `Excel.run()` batch).
- **FR-012**: Chart creation MUST auto-position the chart adjacent to the data range to avoid overlapping existing content.

### Key Entities

- **Worksheet**: A sheet within a workbook. Identified by name. Can be added, deleted, or renamed. At least one must exist at all times.
- **Sort Criteria**: An ordered array of column index and direction pairs applied to a range. Multi-column criteria are applied in array order.
- **Filter Criteria**: Column-value pairs applied to a range's autofilter. Non-matching rows are hidden.
- **Chart**: A visual representation of data. Has a type (Column, Bar, Line, Pie, Scatter, Area, Doughnut), data source range, title, and position on the sheet.
- **Conditional Format Rule**: A visual rule applied to a range. Types include data bars, color scales, icon sets, and cell value rules. Last applied wins on conflict.
- **Cell Format**: Styling applied to a range including font, fill, borders, alignment, and number format.
- **Pivot Table**: A summary table created from source data with row fields, column fields, and value aggregations. Created on a new sheet.

## Success Criteria

### Measurable Outcomes

- **SC-001**: Sort operation on 10,000 rows completes within 3 seconds.
- **SC-002**: Chart created and visible within 2 seconds.
- **SC-003**: Sheet operations (add/delete/rename) complete within 1 second.
- **SC-004**: Conditional formatting applied to 1,000+ cells within 2 seconds.
- **SC-005**: All 23 existing Excel tests continue to pass (no regressions).
- **SC-006**: Pivot table on 10,000 rows of source data completes within 5 seconds.
- **SC-007**: `excel_get_charts` returns results for a workbook with 10 charts within 2 seconds.

## Assumptions

- Excel MVP (Phase 3) is complete and deployed: `excel_get_workbook_map`, `excel_read_range`, `excel_write_range`, `excel_write_formula`, `excel_create_table` are available.
- Excel 2019 or later (or Microsoft 365) is available on the developer machine.
- Range addresses use A1-style notation (not R1C1).
- The add-in runs in the same process space as Excel.
- Pivot table creation uses the Excel JS API `PivotTable` capability (requires ExcelApi 1.12+).
- Chart types are limited to those supported by the Excel JavaScript API.
- Conditional formatting types are limited to: data bars, color scales, icon sets, and cell value rules.
