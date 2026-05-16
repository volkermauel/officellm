# Feature Specification: Phase 14 — Word Find & Replace

**Feature Branch**: `014-word-find-replace`

**Created**: 2026-05-16

**Status**: Draft

**Depends on**: Phase 2 (Word MVP), Phase 11 (Word Structure)

**Input**: User description: "Implement regex-based find and replace in Word documents using the Word JS API search API. Supports wildcard patterns, case sensitivity, whole-word matching, and batch replacement with tracked changes."

## Architecture

Uses `Word.run()` context batching with the Word JS API `Document.search()` or `Range.search()` method. Search returns a `RangeCollection` that can be iterated and replaced. All replacements use the tracked changes pattern.

```
LLM Client                     MCP Server (port 3000)              Office Add-in
    │                                │                               │
    │── tools/call ─────────────────►│                               │
    │   { word_find_replace          │                               │
    │     { instanceId,              │── push command (SignalR) ────►│
    │       findText: "old corp",    │                               │── search(findText, options)
    │       replaceText: "New Corp", │                               │── iterate results
    │       matchCase: true }}       │                               │── replace each (tracked)
    │                                │◄── { replacements: 17 } ──────│
    │◄── { content: { replacements: │                               │
    │      17, preview: [...] }} ───│                               │
```

## User Scenarios & Testing

### User Story 1 — Simple find and replace (Priority: P0)

A user asks "Change all instances of 'Acme Corp' to 'NewCo Inc'." The LLM calls `word_find_replace` with the search and replacement text. All occurrences are replaced with tracked changes.

**Why this priority**: This is the single most common LLM editing operation for Word documents — bulk text replacement.

**Independent Test**: Populate a document with known text, call `word_find_replace`, verify all occurrences are replaced and tracked changes are visible.

**Acceptance Scenarios**:

1. **Given** a document containing "Acme Corp" 15 times, **When** developer calls `word_find_replace` with `findText: "Acme Corp", replaceText: "NewCo Inc"`, **Then** all 15 occurrences are replaced and 15 tracked changes appear in the Review pane
2. **Given** a document with "acme corp" (lowercase) and "Acme Corp" (mixed case), **When** developer calls `word_find_replace` with `matchCase: true`, **Then** only the mixed-case occurrences are replaced
3. **Given** a document with no matches, **When** developer calls `word_find_replace`, **Then** the response returns `{ replacements: 0 }` and no changes are made

---

### User Story 2 — Wildcard pattern matching (Priority: P0)

A user asks "Replace all dates in MM/DD/YYYY format with [DATE]." The LLM calls `word_find_replace` with a wildcard pattern.

**Why this priority**: Wildcard patterns enable structural replacements that simple text matching cannot handle — dates, phone numbers, part numbers, etc.

**Independent Test**: Insert formatted patterns, call with wildcard, verify matching.

**Acceptance Scenarios**:

1. **Given** a document with dates like "01/15/2024" and "12/03/2025", **When** developer calls `word_find_replace` with `findText: "[0-9]{2}/[0-9]{2}/[0-9]{4}"`, `useWildcards: true`, **Then** both dates are found and replaced
2. **Given** a document with "Figure 1", "Figure 12", and "Figure 100", **When** developer calls `word_find_replace` with `findText: "Figure [0-9]*"`, `useWildcards: true`, **Then** all three are found
3. **Given** a document with "fig 1" (lowercase), **When** developer calls with `matchCase: true`, **Then** "fig 1" is NOT matched when searching for "Figure"

---

### User Story 3 — Scoped replacement (Priority: P1)

A user asks "Replace 'old term' only in section 2." The LLM calls `word_find_replace` with a paragraph range scope.

**Why this priority**: Scoped replacement prevents unintended changes in other parts of the document.

**Independent Test**: Insert text in specific paragraphs, scope the search, verify only scoped occurrences change.

**Acceptance Scenarios**:

1. **Given** a document where "Acme" appears in paragraphs 0–5 and paragraphs 10–15, **When** developer calls `word_find_replace` with `scopeFromParagraph: 10, scopeToParagraph: 15`, **Then** only occurrences in paragraphs 10–15 are replaced
2. **Given** a document where "Acme" appears in the header and body, **When** developer calls `word_find_replace` with `scope: "body"`, **Then** only body occurrences are replaced

---

### User Story 4 — Preview before replace (Priority: P1)

A user asks "Show me all instances of 'old term' before replacing." The LLM calls `word_find_replace` with `previewOnly: true` to list all matches without making changes.

**Why this priority**: Preview builds trust. Users and LLMs can verify the scope of changes before committing.

**Independent Test**: Call with `previewOnly: true`, verify matches are listed without modifications.

**Acceptance Scenarios**:

1. **Given** a document with 8 occurrences of "Acme", **When** developer calls `word_find_replace` with `previewOnly: true`, **Then** the response lists all 8 matches with their paragraph index and surrounding context (±20 chars), and NO tracked changes appear
2. **Given** a preview response, **When** the LLM calls `word_find_replace` without `previewOnly`, **Then** exactly the same matches are replaced

---

### Edge Cases

- **Replacing with empty string** (deletion): Supported — acts as a delete operation with tracked change
- **Replacing in tables**: Word JS API search finds text in table cells — replacement works normally
- **Replacing in headers/footers**: Must explicitly scope to header/footer; body scope does not search headers
- **Very large documents** (>500 pages): Search may be slow — consider timeout of 60s
- **Special characters**: Wildcard patterns use Word's wildcard syntax (not standard regex) — `?` for single char, `*` for any string, `[0-9]` for character ranges
- **Replacing with format changes**: This phase only supports text replacement, not format changes

## Requirements

### Functional Requirements

- **FR-001**: The MCP server MUST expose `word_find_replace` accepting `instanceId`, `findText`, `replaceText`, and optional `matchCase`, `matchWholeWord`, `useWildcards`, `previewOnly`, `scopeFromParagraph`, `scopeToParagraph`.
- **FR-002**: All replacements MUST use tracked changes (`changeTrackingMode: "TrackMineOnly"`).
- **FR-003**: The response MUST include `replacements` (count), `preview` (array of `{ paragraphIndex, context }` for up to 50 matches), and `tracked: true`.
- **FR-004**: When `previewOnly: true`, the tool MUST return match information WITHOUT making any changes.
- **FR-005**: Wildcard mode MUST support Word's wildcard syntax (`?`, `*`, `[a-z]`, `[0-9]`, `<`, `>`).
- **FR-006**: When `matchWholeWord: true`, partial matches MUST be excluded.
- **FR-007**: The tool MUST support replacing with an empty string (deletion).

### Key Entities

- **Search Options**: `{ matchCase, matchWholeWord, useWildcards, matchPrefix, matchSuffix }` — mirrors Word JS API `SearchOptions`.
- **Match Preview**: `{ paragraphIndex, offset, context }` — location and surrounding text for each match.
- **Replacement Result**: Count of replacements, preview of changes, tracked change status.

## Success Criteria

### Measurable Outcomes

- **SC-001**: Find and replace on a 100-page document completes within 10 seconds.
- **SC-002**: Preview mode returns match locations without modifying the document.
- **SC-003**: All replacements appear as tracked changes in the Review pane.
- **SC-004**: All existing Word tests continue to pass.

## Assumptions

- Word JS API `search()` method is available (WordApi 1.3+).
- Wildcard syntax follows Word's native wildcard conventions, not standard regex.
- The add-in runs in the same process space as Word.
