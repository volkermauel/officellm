# Feature Specification: Phase 13 — Document Export

**Feature Branch**: `013-document-export`

**Created**: 2026-05-16

**Status**: Draft

**Depends on**: All MVP phases complete, SignalR transport (Phase 8)

**Input**: User description: "Export the current document as PDF or native format (PPTX/DOCX/XLSX) using Office.js getFileAsync. Enables visual verification by LLMs, document backup, and file sharing."

## Architecture

Uses `Office.context.document.getFileAsync(fileType, options, callback)` to extract the entire document as slices. Slices are reassembled server-side and returned as base64. Works for PowerPoint, Word, and Excel.

```
LLM Client                     MCP Server (port 3000)              Office Add-in
    │                                │                               │
    │── tools/call ─────────────────►│                               │
    │   { office_export_document     │                               │
    │     { instanceId,              │                               │
    │       format: "pdf" }}         │── push command (SignalR) ────►│
    │                                │                               │── getFileAsync("pdf")
    │                                │                               │── getSliceAsync (loop)
    │                                │                               │── reassemble base64
    │                                │◄── result { base64, size } ───│
    │◄── { content: base64... } ─────│                               │
    │                                │                               │
    │── (optional: LLM verifies     │                               │
    │    visual output, shares       │                               │
    │    with user)                  │                               │
```

## User Scenarios & Testing

### User Story 1 — Export as PDF for sharing (Priority: P0)

A user asks "Export this as PDF so I can email it." The LLM calls `office_export_document` with `format: "pdf"` and receives the entire document as base64. The LLM can then save it or share it.

**Why this priority**: PDF export is the most common document operation after editing. It's how users share final documents.

**Independent Test**: Call `office_export_document` with `format: "pdf"`, verify the returned base64 decodes to a valid PDF.

**Acceptance Scenarios**:

1. **Given** a PowerPoint presentation with 12 slides, **When** developer calls `office_export_document` with `format: "pdf"`, **Then** the response includes `base64` data that decodes to a valid PDF with all 12 slides
2. **Given** a Word document with 15 pages, **When** developer calls `office_export_document` with `format: "pdf"`, **Then** the response includes `base64` data that decodes to a valid PDF with all 15 pages
3. **Given** an Excel workbook with 4 sheets, **When** developer calls `office_export_document` with `format: "pdf"`, **Then** the response includes `base64` data that decodes to a valid PDF showing the active sheet

---

### User Story 2 — Export as native format for backup (Priority: P0)

A user asks "Save a backup of this file before we make changes." The LLM calls `office_export_document` with `format: "native"` and receives the PPTX/DOCX/XLSX as base64.

**Why this priority**: Backup-before-edit is a critical safety pattern. The LLM should be able to snapshot the document state before risky operations.

**Independent Test**: Call `office_export_document` with `format: "native"`, verify the returned base64 decodes to a valid Office file.

**Acceptance Scenarios**:

1. **Given** a PowerPoint presentation, **When** developer calls `office_export_document` with `format: "native"`, **Then** the response includes `base64` data that decodes to a valid `.pptx` file
2. **Given** a Word document, **When** developer calls `office_export_document` with `format: "native"`, **Then** the response includes `base64` data that decodes to a valid `.docx` file
3. **Given** an Excel workbook, **When** developer calls `office_export_document` with `format: "native"`, **Then** the response includes `base64` data that decodes to a valid `.xlsx` file

---

### User Story 3 — Visual verification after edits (Priority: P1)

A user asks "Does the slide look right now?" after making edits. The LLM calls `office_export_document` with `format: "pdf"`, decodes the PDF, and analyzes the visual output to verify layout, formatting, and content placement.

**Why this priority**: This is the "eyes" the LLM has been missing. Visual verification catches formatting errors, layout issues, and content placement problems that text-only APIs cannot detect.

**Independent Test**: Make edits, export as PDF, verify the LLM can describe the visual layout accurately.

**Acceptance Scenarios**:

1. **Given** the LLM has added a table to slide 3, **When** developer calls `office_export_document` with `format: "pdf"`, **Then** the response includes the PDF with the new table visible
2. **Given** the LLM has formatted header cells bold with blue background, **When** developer calls `office_export_document`, **Then** the exported PDF reflects the formatting
3. **Given** the document is very large (100+ slides), **When** developer calls `office_export_document`, **Then** the export completes within 30 seconds

---

### Edge Cases

- **File too large** (>50MB): Return error with `errorCode: "FILE_TOO_LARGE"` and suggest paginated export or reduced format
- **Outlook items**: Outlook does not support `getFileAsync` — return `errorCode: "HOST_NOT_SUPPORTED"`
- **Concurrent exports**: Only 2 file handles allowed simultaneously — queue or reject with `errorCode: "EXPORT_BUSY"`
- **Unsaved changes**: `getFileAsync` includes unsaved changes — no need to save first
- **Slice reassembly**: Slices may arrive out of order — must reassemble by index

## Requirements

### Functional Requirements

- **FR-001**: The MCP server MUST expose `office_export_document` accepting `instanceId`, `format` ("pdf" or "native"), and optional `maxSizeMB` (default: 50).
- **FR-002**: For PowerPoint and Word: `getFileAsync` supports both `Compressed` (native) and `PDF` formats.
- **FR-003**: For Excel: `getFileAsync` supports `Compressed` (native XLSX) format. PDF export is limited to the active sheet.
- **FR-004**: The add-in MUST reassemble slices in order and return the complete file as a single base64 string.
- **FR-005**: Export MUST complete within 30 seconds for documents under 50MB.
- **FR-006**: If the document exceeds `maxSizeMB`, the tool MUST return an error with `errorCode: "FILE_TOO_LARGE"` and the actual size.
- **FR-007**: For Outlook hosts, the tool MUST return `errorCode: "HOST_NOT_SUPPORTED"` with a message explaining Outlook does not support file export.
- **FR-008**: The response MUST include `fileName` (derived from document title + extension), `sizeBytes`, `format`, `mimeType`, and `base64`.

### Key Entities

- **Export Format**: "pdf" or "native". PDF produces a fixed-layout document. Native produces the original Office format (PPTX/DOCX/XLSX).
- **File Slice**: A chunk of the document returned by `getSliceAsync`. Typically 4KB–1MB. Slices are reassembled in index order.
- **Export Result**: Base64-encoded file data with metadata (name, size, format, MIME type).

## Success Criteria

### Measurable Outcomes

- **SC-001**: A 10-slide PowerPoint exports as PDF within 5 seconds.
- **SC-002**: A 50-page Word document exports as PDF within 10 seconds.
- **SC-003**: The exported file decodes to a valid document openable in the corresponding Office application.
- **SC-002**: File size limit is enforced — documents exceeding `maxSizeMB` are rejected gracefully.
- **SC-005**: All existing tests continue to pass.

## Assumptions

- Office.js `getFileAsync` is available in all target Office versions (2016+ / Microsoft 365).
- The add-in has `ReadDocument` permission (required for `getFileAsync`).
- Slice-based extraction works reliably for documents up to 50MB.
- PDF export quality matches the "Save as PDF" feature in the Office application.
