# Gap Analysis: Office.js API Coverage & Transport Architecture

**Date**: 2026-05-16
**Status**: Investigation — no implementation

---

## 1. API Gap Analysis by Host

### 1.1 Outlook — Critical Gaps

Outlook has the **most missing features** relative to daily work needs.

#### What we have (5 tools):

| Tool                       | Capability                                     |
| -------------------------- | ---------------------------------------------- |
| `outlook_get_current_item` | Read selected email metadata + body            |
| `outlook_summarize_thread` | Thread summary (currently single-message only) |
| `outlook_draft_reply`      | Create draft reply (never auto-send)           |
| `outlook_apply_category`   | Apply color category to selected items         |
| `outlook_send_message`     | Gated send with confirmation token             |

#### What's MISSING — High Priority:

| Feature                    | API Availability                       | Risk Level | Notes                                                                                                                             |
| -------------------------- | -------------------------------------- | ---------- | --------------------------------------------------------------------------------------------------------------------------------- |
| **List folders**           | ❌ Office.js cannot enumerate folders  | HIGH       | Requires Microsoft Graph API (REST). Office.js only knows `Office.context.mailbox` — no folder enumeration.                       |
| **Read folder emails**     | ❌ Cannot list messages in a folder    | HIGH       | Same — needs Graph API. The add-in can only see the currently selected item.                                                      |
| **Calendar/Appointments**  | ⚠️ Partial                             | MEDIUM     | `Office.context.mailbox.item` works for appointments in read/compose mode. But no way to LIST upcoming appointments. Needs Graph. |
| **Shared mailboxes**       | ✅ Mailbox 1.13+                       | MEDIUM     | `getSharedPropertiesAsync()` identifies shared context. But full mailbox access still needs Graph.                                |
| **Compose new email**      | ⚠️ `displayNewMessageFormAsync`        | LOW        | Can open a compose form, but not fully programmatic. Limited control over body content.                                           |
| **Search emails**          | ❌ Not available via Office.js         | HIGH       | No search API in Office.js. Needs Graph `/messages?$search=`.                                                                     |
| **Move email to folder**   | ❌ Not available                       | MEDIUM     | Needs Graph `POST /messages/{id}/move`.                                                                                           |
| **Mark as read/unread**    | ❌ Not available                       | LOW        | Needs Graph.                                                                                                                      |
| **Forward email**          | ⚠️ `displayReplyFormAsync` can forward | LOW        | Opens UI form, not programmatic.                                                                                                  |
| **Attachments**            | ⚠️ Read metadata only                  | MEDIUM     | Can list attachment names/sizes. Cannot download content without Graph.                                                           |
| **Master categories list** | ✅ `mailbox.masterCategories` (1.8+)   | LOW        | Can read/add/remove categories from master list. We don't expose this.                                                            |
| **Multi-select support**   | ✅ Mailbox 1.13+                       | LOW        | `item.getSelectedItemsAsync` for bulk operations.                                                                                 |

#### Recommended Outlook Tool Additions:

**Tier 1 — Daily essentials (needs Graph API integration):**

1. `outlook_list_folders` — List all mail folders with counts
2. `outlook_list_emails` — List/recent emails in a folder (paginated)
3. `outlook_search_emails` — Full-text search across mailbox
4. `outlook_get_email` — Read a specific email by ID (full body + attachments)
5. `outlook_list_calendar_events` — Upcoming appointments from calendar
6. `outlook_get_calendar_event` — Read appointment details

**Tier 2 — Compose + organization:** 7. `outlook_compose_email` — Create a new email draft (not reply) 8. `outlook_move_email` — Move email to a different folder 9. `outlook_mark_read` — Toggle read/unread status 10. `outlook_get_categories` — List available master categories 11. `outlook_get_shared_mailboxes` — List accessible shared mailboxes/delegates

**Email safety gate design:**

- ALL email-sending tools MUST use a confirmation flow:
  1. LLM calls `outlook_draft_reply` or `outlook_compose_email` → creates draft in Drafts folder
  2. Draft is shown in Outlook UI for user review
  3. `outlook_send_message` requires a confirmation token
  4. Token can only be obtained by user clicking "Approve Send" in the task pane
  5. Optional: Policy filter checks recipient domains, attachment types before allowing send
- **NEVER auto-send** — this is non-negotiable for email safety

**Implementation challenge:** Most of Tier 1 requires **Microsoft Graph API** access, not just Office.js. This means:

- The add-in needs OAuth2 / SSO via nested app authentication (NAA)
- The MCP server would proxy Graph API calls through the add-in or directly with stored tokens
- Alternative: The MCP server itself could use client credentials flow to call Graph directly

---

### 1.2 Excel — Moderate Gaps

#### What we have (5 tools):

| Tool                     | Capability                        |
| ------------------------ | --------------------------------- |
| `excel_get_workbook_map` | Sheet names, tables, named ranges |
| `excel_read_range`       | Values, formulas, number formats  |
| `excel_write_range`      | Write values with diff preview    |
| `excel_write_formula`    | Formula validation + write        |
| `excel_create_table`     | Create ListObject from range      |

#### What's MISSING:

| Feature                    | API Availability | Priority | Notes                                                                            |
| -------------------------- | ---------------- | -------- | -------------------------------------------------------------------------------- |
| **Charts**                 | ✅ Full API      | HIGH     | `sheet.charts.add()`, read chart types/data. LLM needs to see and create charts. |
| **PivotTables**            | ✅ Full API      | MEDIUM   | `workbook.pivotTables`, refresh, create. Common analysis tool.                   |
| **Conditional formatting** | ✅ Full API      | MEDIUM   | `range.conditionalFormats`. Visual data highlighting.                            |
| **Sorting**                | ✅ Full API      | HIGH     | `range.sort.apply()`. Fundamental data operation.                                |
| **Filtering/AutoFilter**   | ✅ Full API      | HIGH     | `worksheet.autoFilter.apply()`. Fundamental data operation.                      |
| **Named ranges CRUD**      | ✅ Full API      | LOW      | `workbook.names.add()`. We read them but can't create.                           |
| **Sheet management**       | ✅ Full API      | HIGH     | Add/delete/rename worksheets. Currently only read sheet names.                   |
| **Cell formatting**        | ✅ Full API      | MEDIUM   | Font, fill, borders, alignment, number format.                                   |
| **Data validation**        | ✅ Full API      | LOW      | Dropdowns, input rules.                                                          |
| **Images/Shapes**          | ✅ Full API      | LOW      | Add images to worksheets.                                                        |

**Recommended Excel Tool Additions:**

**Tier 1 — Data manipulation essentials:**

1. `excel_list_sheets` — Already in workbook_map, but add `excel_add_sheet` / `excel_delete_sheet` / `excel_rename_sheet`
2. `excel_sort_range` — Sort a range by column(s)
3. `excel_filter_range` — Apply autofilter to a range
4. `excel_format_range` — Set font, fill, borders, number format on a range

**Tier 2 — Analysis + visualization:** 5. `excel_create_chart` — Create a chart from data range (type, position, series) 6. `excel_get_chart` — Read chart data/properties 7. `excel_create_pivottable` — Create pivot table from range 8. `excel_apply_conditional_formatting` — Add data bars, color scales, icon sets

---

### 1.3 Word — Moderate Gaps

#### What we have (11 tools):

| Tool                       | Capability                            |
| -------------------------- | ------------------------------------- |
| `word_get_outline`         | Document headings outline             |
| `word_get_paragraphs`      | Read paragraphs with styles           |
| `word_get_selection`       | Current selection text + context      |
| `word_search`              | Find text in document                 |
| `word_replace_text`        | Replace text in paragraph (tracked)   |
| `word_insert_text`         | Insert at various locations (tracked) |
| `word_add_comment`         | Add review comment                    |
| `word_delete_paragraph`    | Delete paragraph (tracked)            |
| `word_get_tracked_changes` | Change tracking status                |
| `word_accept_all_changes`  | Accept all tracked changes            |
| `word_reject_all_changes`  | Reject all tracked changes            |

#### What's MISSING:

| Feature                   | API Availability | Priority | Notes                                                        |
| ------------------------- | ---------------- | -------- | ------------------------------------------------------------ |
| **Tables**                | ✅ Full API      | HIGH     | Read/create/modify tables. Very common in documents.         |
| **Headers/Footers**       | ✅ Full API      | MEDIUM   | Read/write header/footer content. Document structure.        |
| **Sections**              | ✅ Full API      | MEDIUM   | Page layout, section breaks, margins.                        |
| **Styles**                | ✅ Full API      | MEDIUM   | Apply/get paragraph and character styles.                    |
| **Images**                | ✅ Full API      | MEDIUM   | Insert inline images, get image properties.                  |
| **Content controls**      | ✅ Full API      | MEDIUM   | Structured document parts, form fields.                      |
| **Lists**                 | ✅ Full API      | LOW      | Bulleted/numbered lists.                                     |
| **Page numbering**        | ⚠️ Partial       | LOW      | Via sections.                                                |
| **Table of contents**     | ⚠️ Partial       | LOW      | Can insert TOC field.                                        |
| **Bookmarks**             | ✅ Full API      | LOW      | Named ranges for navigation.                                 |
| **Selection-based write** | ⚠️ Partial       | HIGH     | We can get selection but can't write TO selection precisely. |

**Recommended Word Tool Additions:**

**Tier 1 — Document structure essentials:**

1. `word_get_tables` — Read all tables in document with cell data
2. `word_insert_table` — Create a table at a specific location
3. `word_update_table_cell` — Modify specific cell content
4. `word_get_headers_footers` — Read header/footer content per section
5. `word_set_headers_footers` — Write header/footer content
6. `word_replace_selection` — Replace the currently selected text (tracked)

**Tier 2 — Rich content:** 7. `word_insert_image` — Insert an image at a location 8. `word_apply_style` — Apply a named style to a paragraph/range 9. `word_get_sections` — Read document sections with page layout 10. `word_insert_list` — Create bulleted/numbered list

---

### 1.4 PowerPoint — Minor Gaps (most complete)

We already have 18 tools covering the v2 spec. Remaining gaps:

| Feature             | API Availability                | Priority | Notes                                         |
| ------------------- | ------------------------------- | -------- | --------------------------------------------- |
| **Duplicate slide** | ✅ `slide.load("id")` then copy | LOW      | Mentioned in original spec but not critical.  |
| **Slide layout**    | ⚠️ Limited                      | LOW      | Can read layout but not change it easily.     |
| **Animation**       | ❌ Not in API                   | N/A      | No animation support in PowerPoint JS API.    |
| **Transition**      | ❌ Not in API                   | N/A      | No transition support.                        |
| **Group shapes**    | ⚠️ Partial                      | LOW      | Can detect groups but limited manipulation.   |
| **Charts**          | ❌ Not in API                   | N/A      | Charts are opaque shapes — no read/write API. |

PowerPoint is effectively **feature-complete** for what the JS API can do.

---

## 2. Transport Architecture Analysis

### 2.1 Current Architecture (HTTP Polling)

```
LLM → MCP Server → [command queued] → Add-in polls every 2s → executes → reports result
                    ↑                                                              |
                    └──── CommandStore.WaitForResult() polls every 500ms ──────────┘
```

**Latency breakdown (best case):**
| Step | Latency |
|------|---------|
| LLM → MCP server | ~5ms (local) |
| Command queued in dictionary | ~1ms |
| Wait for add-in poll (avg 1s, max 2s) | **500-2000ms** |
| Add-in executes Office API | **100ms-5s** (varies wildly) |
| Add-in reports result | ~5ms |
| Server sees completion (polls every 500ms) | **0-500ms** |
| **Total: best ~600ms, typical 1.5-3s, worst 7s+** | |

### 2.2 WebSocket / SignalR Option

**Feasibility: ✅ YES** — WebSockets work in Office add-ins. Confirmed by:

- [GitHub Issue #70](https://github.com/OfficeDev/office-js-docs/issues/70) — explicit confirmation WebSockets/Socket.io work
- [GitHub Issue #5369](https://github.com/OfficeDev/office-js/issues/5369) — PoC with WebSocket IPC working
- [GitHub Issue #2701](https://github.com/OfficeDev/office-js/issues/2701) — SignalR working in Excel add-in

**Known issues:**

- Self-signed certs require loopback exemption on Windows
- SignalR shared runtime can delay Excel shutdown by ~2 minutes (needs graceful disconnect)
- WSS (WebSocket Secure) recommended — plain WS may be blocked in some environments

**Proposed architecture:**

```
LLM → MCP Server ←→ WebSocket ←→ Add-in (real-time bidirectional)
                     SignalR hub
```

```
                    ┌──────────────────┐
                    │   MCP Server     │
                    │  (SignalR Hub)   │
     POST /mcp     │                  │
  LLM ──────────►  │  CommandQueue    │
                    │       ↕          │
                    │  Hub.Clients.    │
                    │  Client(instance) │
                    └────────┬─────────┘
                             │ WebSocket
                    ┌────────┴─────────┐
                    │   Office Add-in   │
                    │  SignalR Client   │
                    │  hub.on("command" │
                    │  hub.invoke(...)  │
                    └──────────────────┘
```

**Latency improvement:**
| Step | Current | WebSocket |
|------|---------|-----------|
| Command delivery | 500-2000ms (poll) | **<10ms (push)** |
| Result reporting | 0-500ms (poll) | **<10ms (push)** |
| **Total overhead saved** | **500-2500ms** | **~20ms** |

**Implementation approach:**

1. Add ASP.NET Core SignalR to the MCP server (already .NET 8)
2. Create a hub with methods: `SendCommand(instanceId, command)`, `ReportResult(commandId, result)`
3. Add-in connects to hub on startup, joins group by instanceId
4. Replace `CommandStore.WaitForResult()` polling with `TaskCompletionSource` + SignalR callback
5. Keep HTTP polling as fallback for environments where WebSocket fails

**Estimated effort:** ~2-3 days

- Day 1: Add SignalR to server, create hub, wire up CommandStore
- Day 2: Update add-in to use SignalR client, fallback to HTTP polling
- Day 3: Testing across all hosts, graceful disconnect handling

### 2.3 Recommendation

**Go with SignalR** for these reasons:

1. **Latency reduction** is significant — 1.5-3s typical → near-instant command delivery
2. **Bidirectional** — server can push commands AND receive results without polling
3. **Automatic fallback** — SignalR falls back to long-polling if WebSocket fails
4. **Already .NET 8** — minimal dependency addition (SignalR is built into ASP.NET Core)
5. **Office add-ins support it** — confirmed working across hosts
6. **Keeps HTTP polling as fallback** — zero risk of regression

---

## 3. Also Discovered: Instance ID Bug

`InstanceRegistry.cs:51` always generates `powerpoint_{N}` regardless of `appName`:

```csharp
string instanceId = $"powerpoint_{_nextInstanceId++}";
```

Word, Excel, and Outlook instances all get `powerpoint_1`, `powerpoint_2`, etc.
The tool routing still works (it dispatches by instanceId lookup), but:

- Misleading in `office_get_active_apps` output
- Complicates debugging
- Should be: `excel_1`, `word_1`, `outlook_1` based on `appName`

---

## 4. Summary: Priority Roadmap

### Phase A: Fix bugs + Transport upgrade

1. Fix instance ID naming bug in `InstanceRegistry.cs`
2. Implement SignalR transport (keep HTTP fallback)
3. ~1 week

### Phase B: Outlook deep integration (via Graph API)

1. Add OAuth2/SSO (nested app authentication)
2. `outlook_list_folders`, `outlook_list_emails`, `outlook_search_emails`
3. `outlook_list_calendar_events`, `outlook_get_calendar_event`
4. `outlook_compose_email` (new draft, not reply)
5. `outlook_move_email`, `outlook_mark_read`
6. All email-sending gated behind confirmation token
7. ~2-3 weeks

### Phase C: Excel analysis tools

1. Sheet management (add/delete/rename)
2. Sort + filter
3. Charts (create + read)
4. Conditional formatting
5. Format range
6. ~1-2 weeks

### Phase D: Word document structure

1. Tables (read + create + modify)
2. Headers/footers
3. Replace selection (write to current selection)
4. Images, styles
5. ~1-2 weeks

### Phase E: Cross-cutting improvements

1. `office_get_document_context` — unified context across all hosts (active doc metadata)
2. Batch/multi-tool operations for efficiency
3. Error recovery and retry logic
4. ~1 week
