# Feature Specification: PowerPoint Tools v2 — Full Read/Write API

**Feature Branch**: `006-powerpoint-v2`

**Created**: 2026-05-15

**Status**: Draft

**Supersedes**: `002-powerpoint-mvp` (enhances and extends Phase 1 tools)

**Input**: User description: "Return shape properties (type, position, size, styling, alignment). Add image export for slides and shapes so LLM can describe them visually. Add table reading, selection context, speaker notes read. Implement real write operations (update text, properties, add/delete shapes, add/delete/move slides). Users create backups so no confirmation flow needed for mutations."

## Architecture

Builds on the existing MCP hub + Office JS Add-in architecture from Phase 1. All tools route through the same instance registration + command polling mechanism. The add-in's `powerpoint-commands.ts` handler grows additional command types.

```
LLM Client ──► MCP Server ──► Add-in (powerpoint-commands.ts)
                    │              │
                    │              ├── Read tools (enhanced with properties + images)
                    │              ├── Write tools (direct, no confirmation gate)
                    │              └── Slide management (add/delete/move)
                    │
                    └── Images returned as base64 in tool results
                        (LLM processes via vision capabilities)
```

**Key change from Phase 1**: Mutation tools apply changes directly. Users are responsible for creating backups before enabling LLM write access. The two-phase confirmation flow is removed.

## PowerPoint JS API Shape Properties Reference

The Office JS API exposes these properties per shape. We load them via `shape.load(...)` and `ctx.sync()` before reading.

| Property                              | Type            | Description                                                                   |
| ------------------------------------- | --------------- | ----------------------------------------------------------------------------- |
| `id`                                  | string          | Unique shape identifier within slide                                          |
| `name`                                | string          | Human-readable name (e.g., "Title 1")                                         |
| `type`                                | ShapeType enum  | `Image`, `TextBox`, `Table`, `Group`, `Line`, `GeometricShape`, `Chart`, etc. |
| `left`                                | number          | X position in points                                                          |
| `top`                                 | number          | Y position in points                                                          |
| `width`                               | number          | Width in points                                                               |
| `height`                              | number          | Height in points                                                              |
| `rotation`                            | number          | Rotation in degrees                                                           |
| `fill`                                | ShapeFill       | Fill formatting (`foregroundColor`, `transparency`)                           |
| `textFrame`                           | TextFrame       | Text content (throws on non-text shapes — use `getTextFrameOrNullObject()`)   |
| `textFrame.textRange.font`            | TextRange.font  | `name`, `size`, `bold`, `italic`, `color`, `underline`, `strikethrough`       |
| `textFrame.textRange.paragraphFormat` | ParagraphFormat | `horizontalAlignment`, `bulletFormat`                                         |

**Image export APIs:**

- `slide.getImageAsBase64(options?)` → base64 PNG of rendered slide
- `shape.getImageAsBase64(options?)` → base64 PNG of rendered shape
- Options: `{ width?, height? }` — preserves aspect ratio

**Shape creation APIs:**

- `shapes.addTextBox(text, { left, top, width, height })`
- `shapes.addPicture(base64, { left, top, width, height })`
- `shapes.addTable(rowCount, colCount, { left, top, width, height })`
- `shapes.addGeometricShape(type, { left, top, width, height })`
- `shapes.addLine({ left, top, width, height })`

**Table API:**

- `shape.getTable()` → `Table` with `rowCount`, `columnCount`
- `table.getCell(row, col).textFrame.textRange.text` → cell text

## User Scenarios & Testing

### User Story 1 — LLM describes a visual slide (Priority: P0)

A user asks the LLM "What's on slide 3?" The LLM calls `powerpoint_get_slide` with full properties, then `powerpoint_get_slide_image` to see the slide visually. The LLM can describe layout, colors, images, and suggest improvements.

**Acceptance Scenarios**:

1. **Given** a slide with an image, a title, and a chart, **When** `powerpoint_get_slide` is called with `slideIndex: 2`, **Then** each shape includes `type`, `left`, `top`, `width`, `height`, `rotation`, and text formatting properties
2. **Given** a slide with a picture shape, **When** `powerpoint_get_slide` returns, **Then** the shape has `type: "Image"` with position/size but no text content
3. **Given** any slide, **When** `powerpoint_get_slide_image` is called, **Then** a base64 PNG image is returned that the LLM can process visually

### User Story 2 — LLM reads table data (Priority: P0)

A user asks "What's in the pricing table on slide 7?" The LLM reads the table content as a structured 2D array.

**Acceptance Scenarios**:

1. **Given** a slide with a 3×4 table, **When** `powerpoint_get_table` is called, **Then** response contains `rowCount: 3`, `columnCount: 4`, and `cells: string[][]` with all cell text
2. **Given** a slide with no table at the given shape ID, **When** `powerpoint_get_table` is called, **Then** an appropriate error is returned

### User Story 3 — LLM describes an image on a slide (Priority: P0)

A user asks "Describe the image on slide 2." The LLM exports that specific shape as an image and describes it.

**Acceptance Scenarios**:

1. **Given** a slide with an image shape (ID: "Picture 3"), **When** `powerpoint_get_shape_image` is called, **Then** a base64 PNG of just that shape is returned
2. **Given** a shape ID that doesn't exist, **When** `powerpoint_get_shape_image` is called, **Then** an error is returned

### User Story 4 — User asks about current selection (Priority: P1)

A user selects a paragraph in PowerPoint and asks the LLM to improve it. The LLM reads the selection to understand context.

**Acceptance Scenarios**:

1. **Given** the user has selected text on slide 5, **When** `powerpoint_get_selection` is called, **Then** response includes selected text, parent shape ID, slide index, and text formatting
2. **Given** the user has selected shapes (not text), **When** `powerpoint_get_selection` is called, **Then** response includes selected shape IDs with their properties
3. **Given** nothing is selected, **When** `powerpoint_get_selection` is called, **Then** response indicates empty selection

### User Story 5 — LLM edits shape text directly (Priority: P0)

A user asks "Change the title on slide 1 to 'Q4 Results'." The LLM calls `powerpoint_update_shape_text` and the change is applied immediately.

**Acceptance Scenarios**:

1. **Given** slide 0 has a title shape with ID "Title 1", **When** `powerpoint_update_shape_text` is called with `{ slideIndex: 0, shapeId: "Title 1", text: "Q4 Results" }`, **Then** the shape text is updated in PowerPoint
2. **Given** a shape ID that doesn't exist on the slide, **When** the tool is called, **Then** a `SHAPE_NOT_FOUND` error is returned
3. **Given** a non-text shape (image), **When** the tool is called, **Then** a `NOT_TEXT_SHAPE` error is returned

### User Story 6 — LLM modifies shape styling (Priority: P1)

A user asks "Make the title bigger and bold." The LLM updates font properties.

**Acceptance Scenarios**:

1. **Given** a text shape, **When** `powerpoint_update_shape_properties` is called with `{ font: { size: 36, bold: true } }`, **Then** the shape's font is updated
2. **Given** a shape, **When** position properties are set `{ left: 100, top: 50 }`, **Then** the shape moves to the specified position

### User Story 7 — LLM adds new shapes (Priority: P1)

A user asks "Add a text box with 'Key Insight' on slide 3." The LLM creates the shape.

**Acceptance Scenarios**:

1. **Given** slide 2, **When** `powerpoint_add_textbox` is called with text and position, **Then** a new text box appears on the slide
2. **Given** slide 2, **When** `powerpoint_add_image` is called with base64 data, **Then** a new image appears on the slide
3. **Given** slide 2, **When** `powerpoint_add_table` is called with `{ rows: 3, cols: 4 }`, **Then** a new table appears on the slide

### User Story 8 — LLM manages slides (Priority: P1)

A user asks "Delete slide 5" or "Move slide 2 to the end." The LLM rearranges the deck.

**Acceptance Scenarios**:

1. **Given** a 10-slide deck, **When** `powerpoint_delete_slide` is called with `slideIndex: 4`, **Then** the deck now has 9 slides
2. **Given** a 10-slide deck, **When** `powerpoint_add_slide` is called with `{ atIndex: 3 }`, **Then** a new blank slide is inserted at position 3
3. **Given** a 10-slide deck, **When** `powerpoint_move_slide` is called with `{ fromIndex: 2, toIndex: 9 }`, **Then** slide 2 moves to the end

### User Story 9 — LLM reads speaker notes (Priority: P1)

A user asks "What are my speaker notes for slides 3-5?" The LLM reads the notes.

**Acceptance Scenarios**:

1. **Given** slides with speaker notes, **When** `powerpoint_get_speaker_notes` is called, **Then** notes text is returned per slide
2. **Given** a slide with no notes, **When** the tool is called, **Then** empty string is returned for that slide

### User Story 10 — LLM writes speaker notes (Priority: P1)

A user asks "Generate speaker notes for slide 3 based on the content." The LLM writes notes directly.

**Acceptance Scenarios**:

1. **Given** slide 2, **When** `powerpoint_update_speaker_notes` is called with `{ slideIndex: 2, notes: "Talk about revenue growth..." }`, **Then** the notes are applied to the slide

## Requirements

### Functional Requirements — Enhanced Read

- **FR-101**: `powerpoint_get_deck_outline` MUST return per-shape: `type` (ShapeType enum string), `left`, `top`, `width`, `height`, `rotation`.
- **FR-102**: `powerpoint_get_slide` MUST return per-shape: `type`, `left`, `top`, `width`, `height`, `rotation`, `name`, `id`. For text shapes: font `name`, `size`, `bold`, `italic`, `color`, paragraph `horizontalAlignment`. For fill shapes: `fillColor`, `fillTransparency`.
- **FR-103**: `powerpoint_get_slide_image` MUST accept `slideIndex` and optional `{ width, height }` and return a base64-encoded PNG image of the rendered slide.
- **FR-104**: `powerpoint_get_shape_image` MUST accept `slideIndex`, `shapeId`, and optional `{ width, height }` and return a base64-encoded PNG image of the rendered shape.
- **FR-105**: `powerpoint_get_table` MUST accept `slideIndex` and `shapeId` and return `rowCount`, `columnCount`, and `cells` as a 2D string array.
- **FR-106**: `powerpoint_get_selection` MUST return currently selected text range (with formatting) or selected shape IDs, depending on what the user has selected.
- **FR-107**: `powerpoint_get_speaker_notes` MUST accept `slideIndex` (single) or `slideRange` (e.g., "2-5") and return notes text per slide.

### Functional Requirements — Write Operations

- **FR-201**: `powerpoint_update_shape_text` MUST write `text` to the specified shape on the specified slide via Office JS API. No confirmation gate — applies directly.
- **FR-202**: `powerpoint_update_shape_properties` MUST accept optional `left`, `top`, `width`, `height`, `rotation` and/or font properties (`fontName`, `fontSize`, `bold`, `italic`, `color`) and apply them to the specified shape.
- **FR-203**: `powerpoint_update_speaker_notes` MUST write notes text to the specified slide. No confirmation gate.
- **FR-204**: `powerpoint_add_textbox` MUST create a new text box on the specified slide with given text and position/size.
- **FR-205**: `powerpoint_add_image` MUST insert an image (from base64 data) on the specified slide with given position/size.
- **FR-206**: `powerpoint_add_table` MUST create a new table with specified rows/columns on the specified slide with given position/size.

### Functional Requirements — Shape & Slide Management

- **FR-301**: `powerpoint_delete_shape` MUST remove the specified shape from the slide.
- **FR-302**: `powerpoint_add_slide` MUST insert a new blank slide at the specified index (or at the end if no index).
- **FR-303**: `powerpoint_delete_slide` MUST remove the slide at the specified index.
- **FR-304**: `powerpoint_move_slide` MUST move a slide from one position to another via `slide.moveTo()`.

### Non-Functional Requirements

- **NFR-001**: All image exports MUST complete within 5 seconds for a single slide/shape.
- **NFR-002**: Base64 images for slide thumbnails SHOULD be capped at 800px width by default to keep token usage reasonable.
- **NFR-003**: Tool results including images MUST NOT exceed 10MB total payload size.
- **NFR-004**: Table reading MUST handle up to 50×50 tables (2500 cells) without timeout.
- **NFR-005**: All tools MUST return structured error objects with `error` field (string) on failure.

## Tool Definitions

### Read Tools

| Tool                           | Parameters                                                 | Returns                                                                                                                    |
| ------------------------------ | ---------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------- |
| `powerpoint_get_deck_outline`  | `instanceId` (required)                                    | `{ slides: [{ index, title, shapes: [{ id, name, type, left, top, width, height }] }] }`                                   |
| `powerpoint_get_slide`         | `instanceId`, `slideIndex`                                 | `{ slideIndex, title, shapes: [{ id, name, type, left, top, width, height, rotation, text?, font?, fill?, alignment? }] }` |
| `powerpoint_get_slide_image`   | `instanceId`, `slideIndex`, `width?`, `height?`            | `{ slideIndex, image: "data:image/png;base64,..." }`                                                                       |
| `powerpoint_get_shape_image`   | `instanceId`, `slideIndex`, `shapeId`, `width?`, `height?` | `{ slideIndex, shapeId, image: "data:image/png;base64,..." }`                                                              |
| `powerpoint_get_table`         | `instanceId`, `slideIndex`, `shapeId`                      | `{ slideIndex, shapeId, rowCount, columnCount, cells: string[][] }`                                                        |
| `powerpoint_get_selection`     | `instanceId`                                               | `{ type: "text" \| "shapes" \| "none", text?, shapeIds?, slideIndex, font? }`                                              |
| `powerpoint_get_speaker_notes` | `instanceId`, `slideIndex?`, `slideRange?`                 | `{ notes: [{ slideIndex, notes: string }] }`                                                                               |

### Write Tools

| Tool                                 | Parameters                                                                                                                                       | Returns                                      |
| ------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------ | -------------------------------------------- |
| `powerpoint_update_shape_text`       | `instanceId`, `slideIndex`, `shapeId`, `text`                                                                                                    | `{ slideIndex, shapeId, newText }`           |
| `powerpoint_update_shape_properties` | `instanceId`, `slideIndex`, `shapeId`, `left?`, `top?`, `width?`, `height?`, `rotation?`, `fontName?`, `fontSize?`, `bold?`, `italic?`, `color?` | `{ slideIndex, shapeId, updated: string[] }` |
| `powerpoint_update_speaker_notes`    | `instanceId`, `slideIndex`, `notes`                                                                                                              | `{ slideIndex, newNotes }`                   |
| `powerpoint_add_textbox`             | `instanceId`, `slideIndex`, `text`, `left`, `top`, `width`, `height`                                                                             | `{ slideIndex, shapeId, name }`              |
| `powerpoint_add_image`               | `instanceId`, `slideIndex`, `imageBase64`, `left`, `top`, `width?`, `height?`                                                                    | `{ slideIndex, shapeId, name }`              |
| `powerpoint_add_table`               | `instanceId`, `slideIndex`, `rows`, `columns`, `left`, `top`, `width?`, `height?`                                                                | `{ slideIndex, shapeId, name }`              |
| `powerpoint_delete_shape`            | `instanceId`, `slideIndex`, `shapeId`                                                                                                            | `{ slideIndex, shapeId, deleted: true }`     |

### Slide Management Tools

| Tool                      | Parameters                           | Returns                           |
| ------------------------- | ------------------------------------ | --------------------------------- |
| `powerpoint_add_slide`    | `instanceId`, `atIndex?`             | `{ slideIndex, slideId }`         |
| `powerpoint_delete_slide` | `instanceId`, `slideIndex`           | `{ slideIndex, deleted: true }`   |
| `powerpoint_move_slide`   | `instanceId`, `fromIndex`, `toIndex` | `{ fromIndex, toIndex, slideId }` |

## Implementation Notes

### Loading Shape Properties Efficiently

The Office JS API requires explicit `load()` calls followed by `sync()`. Properties must be loaded in batches:

```
// Step 1: Load shape items
slide.load("shapes/items/$none");
await ctx.sync();

// Step 2: Load direct properties (comma-separated works)
for (const s of slide.shapes.items) {
    s.load("id,name,type,left,top,width,height,rotation");
}
await ctx.sync();

// Step 3: Load nested properties (text formatting, fill) via slash paths
// Note: slash paths support ONE property at a time
for (const s of slide.shapes.items) {
    const tf = s.getTextFrameOrNullObject();
    ctx.load(tf, "isNullObject,textRange/text,textRange/font/name,textRange/font/size,textRange/font/bold,textRange/font/italic,textRange/font/color");
    // Fill: load separately
    ctx.load(s.fill, "foregroundColor,transparency");
}
await ctx.sync();
```

### Image Export

Both `slide.getImageAsBase64()` and `shape.getImageAsBase64()` return `ClientResult<string>` — must call `ctx.sync()` before reading `.value`.

### Table Cell Reading

Requires nested loading: `shape.getTable()` → `table.load("rowCount,columnCount")` → `table.getCell(r,c).textFrame.textRange.text` → `ctx.sync()`.

### Speaker Notes

Access via `slide.notesSlide.textFrame.textRange.text` (requires `notesSlide` to be loaded). Some slides may not have a notes slide — use `getNotesSlideOrNullObject()` pattern.

## Success Criteria

- **SC-101**: `powerpoint_get_slide` returns complete shape properties for all shape types without errors.
- **SC-102**: `powerpoint_get_slide_image` returns a valid base64 PNG that renders correctly.
- **SC-103**: `powerpoint_update_shape_text` updates text in PowerPoint within 2 seconds of tool call.
- **SC-104**: `powerpoint_add_textbox` creates a visible text box with the specified text and position.
- **SC-105**: All 16 tools (7 read + 6 write + 3 slide management) are registered and callable via MCP.
- **SC-106**: 60+ existing tests continue to pass (backward compatible).
- **SC-107**: New C# tests cover all tool definitions (parameters, required fields).
