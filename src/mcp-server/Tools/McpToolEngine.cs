using System.Text.Json;
using OfficeMcpServer.Models;

namespace OfficeMcpServer.Tools;

/// <summary>
/// Handles MCP tool definition generation and command dispatch.
/// Uses instance-aware routing with a fixed tool list.
/// </summary>
public static class McpToolEngine
{
    private static InstanceRegistry _registry = new();
    private static CommandStore _commandStore = new();
    private static AuditLog _auditLog = new();

    /// <summary>
    /// Gets all MCP tool definitions (fixed list, instance selected via parameter).
    /// </summary>
    public static object[] GetToolDefinitions() => [
        // ── Shared ──────────────────────────────────────────────
        new
        {
            name = "office_get_active_apps",
            description = "Returns a list of all registered Office instances (PowerPoint, Word, Excel, Outlook) with their document names. Call this FIRST so the user can choose which document to work with. Then pass the chosen instanceId to all other tools.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>(),
                required = Array.Empty<string>()
            }
        },

        // ── Read tools ──────────────────────────────────────────
        new
        {
            name = "powerpoint_get_deck_outline",
            description = "Returns the full slide deck outline. Each slide lists its shapes with type, position, size, and text content. Use office_get_active_apps first to find the right instanceId.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps (e.g. 'powerpoint_1')." },
                    ["includeSpeakerNotes"] = new { type = "boolean", description = "Include speaker notes in the outline", @default = false },
                    ["includeHiddenSlides"] = new { type = "boolean", description = "Include hidden slides in the outline", @default = false }
                },
                required = new[] { "instanceId" }
            }
        },
        new
        {
            name = "powerpoint_get_slide",
            description = "Returns all shapes on a single slide with full properties: type, position (left/top), size (width/height), rotation, text content, font styling (name/size/bold/italic/color), paragraph alignment, and fill color. Use powerpoint_get_deck_outline first to see slide indices.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["slideIndex"] = new { type = "integer", description = "Zero-based slide index" }
                },
                required = new[] { "instanceId", "slideIndex" }
            }
        },
        new
        {
            name = "powerpoint_get_slide_image",
            description = "Renders a slide as a PNG image and returns it as base64. The LLM can use this to visually understand the slide layout and content. Optionally specify width/height to control image size (preserves aspect ratio).",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["slideIndex"] = new { type = "integer", description = "Zero-based slide index" },
                    ["width"] = new { type = "integer", description = "Max image width in pixels. Default: 800", @default = 800 },
                    ["height"] = new { type = "integer", description = "Max image height in pixels. Default: omitted (auto)" }
                },
                required = new[] { "instanceId", "slideIndex" }
            }
        },
        new
        {
            name = "powerpoint_get_shape_image",
            description = "Renders a single shape as a PNG image and returns it as base64. Useful for describing images, charts, icons, or diagrams on a slide.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["slideIndex"] = new { type = "integer", description = "Zero-based slide index" },
                    ["shapeId"] = new { type = "string", description = "Shape ID or name to render" },
                    ["width"] = new { type = "integer", description = "Max image width in pixels. Default: 400", @default = 400 },
                    ["height"] = new { type = "integer", description = "Max image height in pixels. Default: omitted (auto)" }
                },
                required = new[] { "instanceId", "slideIndex", "shapeId" }
            }
        },
        new
        {
            name = "powerpoint_get_table",
            description = "Reads all cell text from a table shape on a slide. Returns rowCount, columnCount, and a 2D string array of cell contents.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["slideIndex"] = new { type = "integer", description = "Zero-based slide index" },
                    ["shapeId"] = new { type = "string", description = "Shape ID of the table" }
                },
                required = new[] { "instanceId", "slideIndex", "shapeId" }
            }
        },
        new
        {
            name = "powerpoint_get_selection",
            description = "Returns what the user currently has selected in PowerPoint: selected text (with formatting and parent shape), selected shapes (with IDs and properties), or empty selection. Use this to understand user intent context.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." }
                },
                required = new[] { "instanceId" }
            }
        },
        new
        {
            name = "powerpoint_get_speaker_notes",
            description = "Returns speaker notes for one slide or a range of slides. Specify slideIndex for a single slide, or slideRange (e.g. '2-5') for multiple.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["slideIndex"] = new { type = "integer", description = "Zero-based slide index for a single slide" },
                    ["slideRange"] = new { type = "string", description = "Range of slides, e.g. '2-5'" }
                },
                required = new[] { "instanceId" }
            }
        },

        // ── Write tools (direct, no confirmation gate) ──────────
        new
        {
            name = "powerpoint_update_shape_text",
            description = "Updates the text content of a specific shape on a slide. Applies directly — no confirmation required. Users should create backups before enabling write access.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["slideIndex"] = new { type = "integer", description = "Zero-based slide index" },
                    ["shapeId"] = new { type = "string", description = "Shape ID or name" },
                    ["text"] = new { type = "string", description = "New text content" }
                },
                required = new[] { "instanceId", "slideIndex", "shapeId", "text" }
            }
        },
        new
        {
            name = "powerpoint_update_shape_properties",
            description = "Updates position, size, rotation, and/or font properties of a shape. Only specified properties are changed. Use this to resize, reposition, or restyle shapes.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["slideIndex"] = new { type = "integer", description = "Zero-based slide index" },
                    ["shapeId"] = new { type = "string", description = "Shape ID or name" },
                    ["left"] = new { type = "number", description = "X position in points" },
                    ["top"] = new { type = "number", description = "Y position in points" },
                    ["width"] = new { type = "number", description = "Width in points" },
                    ["height"] = new { type = "number", description = "Height in points" },
                    ["rotation"] = new { type = "number", description = "Rotation in degrees" },
                    ["fontName"] = new { type = "string", description = "Font family name (e.g. 'Arial')" },
                    ["fontSize"] = new { type = "number", description = "Font size in points" },
                    ["bold"] = new { type = "boolean", description = "Bold on/off" },
                    ["italic"] = new { type = "boolean", description = "Italic on/off" },
                    ["color"] = new { type = "string", description = "Font color as HTML hex (e.g. '#FF0000')" }
                },
                required = new[] { "instanceId", "slideIndex", "shapeId" }
            }
        },
        new
        {
            name = "powerpoint_update_speaker_notes",
            description = "Sets speaker notes for a specific slide. Replaces any existing notes. Applies directly.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["slideIndex"] = new { type = "integer", description = "Zero-based slide index" },
                    ["notes"] = new { type = "string", description = "Speaker notes text" }
                },
                required = new[] { "instanceId", "slideIndex", "notes" }
            }
        },

        // ── Shape CRUD ──────────────────────────────────────────
        new
        {
            name = "powerpoint_add_textbox",
            description = "Creates a new text box on a slide with the specified text, position, and size. Returns the new shape's ID.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["slideIndex"] = new { type = "integer", description = "Zero-based slide index" },
                    ["text"] = new { type = "string", description = "Initial text content" },
                    ["left"] = new { type = "number", description = "X position in points" },
                    ["top"] = new { type = "number", description = "Y position in points" },
                    ["width"] = new { type = "number", description = "Width in points" },
                    ["height"] = new { type = "number", description = "Height in points" }
                },
                required = new[] { "instanceId", "slideIndex", "text", "left", "top", "width", "height" }
            }
        },
        new
        {
            name = "powerpoint_add_image",
            description = "Inserts an image onto a slide from base64-encoded data. Returns the new shape's ID.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["slideIndex"] = new { type = "integer", description = "Zero-based slide index" },
                    ["imageBase64"] = new { type = "string", description = "Base64-encoded image data (PNG, JPG, etc.)" },
                    ["left"] = new { type = "number", description = "X position in points" },
                    ["top"] = new { type = "number", description = "Y position in points" },
                    ["width"] = new { type = "number", description = "Width in points" },
                    ["height"] = new { type = "number", description = "Height in points" }
                },
                required = new[] { "instanceId", "slideIndex", "imageBase64", "left", "top" }
            }
        },
        new
        {
            name = "powerpoint_add_table",
            description = "Creates a new table on a slide with specified rows and columns. Returns the new shape's ID.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["slideIndex"] = new { type = "integer", description = "Zero-based slide index" },
                    ["rows"] = new { type = "integer", description = "Number of rows" },
                    ["columns"] = new { type = "integer", description = "Number of columns" },
                    ["left"] = new { type = "number", description = "X position in points" },
                    ["top"] = new { type = "number", description = "Y position in points" },
                    ["width"] = new { type = "number", description = "Width in points" },
                    ["height"] = new { type = "number", description = "Height in points" }
                },
                required = new[] { "instanceId", "slideIndex", "rows", "columns", "left", "top" }
            }
        },
        new
        {
            name = "powerpoint_delete_shape",
            description = "Deletes a shape from a slide. Irreversible — users should create backups.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["slideIndex"] = new { type = "integer", description = "Zero-based slide index" },
                    ["shapeId"] = new { type = "string", description = "Shape ID or name to delete" }
                },
                required = new[] { "instanceId", "slideIndex", "shapeId" }
            }
        },

        // ── Slide management ────────────────────────────────────
        new
        {
            name = "powerpoint_add_slide",
            description = "Inserts a new blank slide at the specified position. If atIndex is omitted, adds at the end.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["atIndex"] = new { type = "integer", description = "Zero-based index to insert at. Default: end of deck." }
                },
                required = new[] { "instanceId" }
            }
        },
        new
        {
            name = "powerpoint_delete_slide",
            description = "Deletes a slide from the presentation. Irreversible — users should create backups.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["slideIndex"] = new { type = "integer", description = "Zero-based slide index to delete" }
                },
                required = new[] { "instanceId", "slideIndex" }
            }
        },
        new
        {
            name = "powerpoint_move_slide",
            description = "Moves a slide from one position to another. Other slides shift accordingly.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["fromIndex"] = new { type = "integer", description = "Current zero-based position of the slide" },
                    ["toIndex"] = new { type = "integer", description = "Target zero-based position" }
                },
                required = new[] { "instanceId", "fromIndex", "toIndex" }
            }
        },

        // ── Word Read tools ───────────────────────────────────────
        new
        {
            name = "word_get_outline",
            description = "Returns the document outline (headings with levels, styles, and text). Use office_get_active_apps first to find the right instanceId.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["maxDepth"] = new { type = "integer", description = "Maximum heading level to include (1-9). Default: 3", @default = 3 }
                },
                required = new[] { "instanceId" }
            }
        },
        new
        {
            name = "word_get_paragraphs",
            description = "Returns paragraphs from the document body, optionally filtered by range (startIndex, count).",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["startIndex"] = new { type = "integer", description = "Zero-based paragraph index to start from. Default: 0", @default = 0 },
                    ["count"] = new { type = "integer", description = "Maximum paragraphs to return. Default: 50", @default = 50 }
                },
                required = new[] { "instanceId" }
            }
        },
        new
        {
            name = "word_get_selection",
            description = "Returns the currently selected text in Word, with paragraph context and formatting.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." }
                },
                required = new[] { "instanceId" }
            }
        },
        new
        {
            name = "word_search",
            description = "Searches the document for text and returns matching paragraphs with their locations.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["searchText"] = new { type = "string", description = "Text to search for" },
                    ["matchCase"] = new { type = "boolean", description = "Case-sensitive search. Default: false", @default = false }
                },
                required = new[] { "instanceId", "searchText" }
            }
        },

        // ── Word Write tools ──────────────────────────────────────
        new
        {
            name = "word_replace_text",
            description = "Replaces text in a specific paragraph by index. Applies directly — users should create backups.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["paragraphIndex"] = new { type = "integer", description = "Zero-based paragraph index" },
                    ["oldText"] = new { type = "string", description = "Text to find and replace within the paragraph" },
                    ["newText"] = new { type = "string", description = "Replacement text" }
                },
                required = new[] { "instanceId", "paragraphIndex", "oldText", "newText" }
            }
        },
        new
        {
            name = "word_insert_text",
            description = "Inserts text at a specific location (before/after a paragraph, or at the end of the document).",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["text"] = new { type = "string", description = "Text to insert" },
                    ["insertLocation"] = new { type = "string", description = "Where to insert: 'end' (default), 'afterParagraph', 'beforeParagraph'", @default = "end" },
                    ["paragraphIndex"] = new { type = "integer", description = "Zero-based paragraph index (required for afterParagraph/beforeParagraph)" }
                },
                required = new[] { "instanceId", "text" }
            }
        },
        new
        {
            name = "word_add_comment",
            description = "Adds a Word comment to the current selection or a specific paragraph.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["commentText"] = new { type = "string", description = "Comment text to add" },
                    ["paragraphIndex"] = new { type = "integer", description = "Optional: zero-based paragraph to attach comment to. Default: current selection." }
                },
                required = new[] { "instanceId", "commentText" }
            }
        },
        new
        {
            name = "word_delete_paragraph",
            description = "Deletes a paragraph by index. Creates a tracked deletion if change tracking is enabled.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["paragraphIndex"] = new { type = "integer", description = "Zero-based paragraph index to delete" }
                },
                required = new[] { "instanceId", "paragraphIndex" }
            }
        },
        new
        {
            name = "word_get_tracked_changes",
            description = "Returns the current change tracking mode and count of pending revisions in the Word document.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." }
                },
                required = new[] { "instanceId" }
            }
        },
        new
        {
            name = "word_accept_all_changes",
            description = "Accepts all tracked changes in the Word document. Finalizes all pending insertions, deletions, and formatting changes.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." }
                },
                required = new[] { "instanceId" }
            }
        },
        new
        {
            name = "word_reject_all_changes",
            description = "Rejects all tracked changes in the Word document. Reverts all pending insertions, deletions, and formatting changes.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." }
                },
                required = new[] { "instanceId" }
            }
        },

        // ── Excel Read tools ──────────────────────────────────────
        new
        {
            name = "excel_get_workbook_map",
            description = "Returns the workbook structure: sheet names, used range dimensions, tables (name, range, sheet), and named ranges. Use office_get_active_apps first to find the right instanceId.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps (e.g. 'excel_1')." },
                    ["includeTables"] = new { type = "boolean", description = "Include table definitions in the response. Default: true", @default = true },
                    ["includeNamedRanges"] = new { type = "boolean", description = "Include named ranges in the response. Default: true", @default = true }
                },
                required = new[] { "instanceId" }
            }
        },
        new
        {
            name = "excel_read_range",
            description = "Reads cell values, formulas, and optional number formats from a bounded range. Specify sheet name and A1-style address (e.g. 'B2:G24'). Returns a 2D grid with row/column indices.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["sheetName"] = new { type = "string", description = "Name of the worksheet to read from" },
                    ["address"] = new { type = "string", description = "A1-style range address (e.g. 'A1:D10') or single cell (e.g. 'B2')" },
                    ["includeFormulas"] = new { type = "boolean", description = "Include formula strings for formula cells. Default: true", @default = true },
                    ["includeNumberFormats"] = new { type = "boolean", description = "Include number format strings. Default: false", @default = false }
                },
                required = new[] { "instanceId", "sheetName", "address" }
            }
        },

        // ── Excel Write tools ─────────────────────────────────────
        new
        {
            name = "excel_write_range",
            description = "Writes values to a bounded range on a worksheet. Applies directly — users should create backups. Values is a 2D array (rows of columns). The range dimensions must match the data dimensions.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["sheetName"] = new { type = "string", description = "Name of the worksheet to write to" },
                    ["address"] = new { type = "string", description = "A1-style range address to write to (e.g. 'A1:C3')" },
                    ["values"] = new { type = "array", description = "2D array of values (rows of columns). Numbers, strings, or null for empty.", items = new { type = "array", items = new { type = "string" } } }
                },
                required = new[] { "instanceId", "sheetName", "address", "values" }
            }
        },
        new
        {
            name = "excel_write_formula",
            description = "Writes a formula to a cell or range on a worksheet. Validates formula syntax (must start with '='). Applies directly — users should create backups.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["sheetName"] = new { type = "string", description = "Name of the worksheet" },
                    ["address"] = new { type = "string", description = "A1-style cell or range address for the formula (e.g. 'D2' or 'D2:D24')" },
                    ["formula"] = new { type = "string", description = "Excel formula string (must start with '='). e.g. '=SUM(B2:B24)'" }
                },
                required = new[] { "instanceId", "sheetName", "address", "formula" }
            }
        },
        new
        {
            name = "excel_create_table",
            description = "Creates a formatted Excel table (ListObject) from a specified range. Returns the table name and range. Applies directly.", 
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["sheetName"] = new { type = "string", description = "Name of the worksheet containing the data" },
                    ["address"] = new { type = "string", description = "A1-style range address including headers (e.g. 'A1:D25')" },
                    ["tableName"] = new { type = "string", description = "Optional name for the table. Default: auto-generated by Excel." },
                    ["hasHeaders"] = new { type = "boolean", description = "First row contains headers. Default: true", @default = true }
                },
                required = new[] { "instanceId", "sheetName", "address" }
            }
        },

        // ── Outlook Read tools ──────────────────────────────────────
        new
        {
            name = "outlook_get_current_item",
            description = "Returns metadata (sender, recipients, subject, received date) and bounded body text for the currently selected Outlook item. Includes attachment metadata (name, size, type) but not file contents.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps (e.g. 'outlook_1')." },
                    ["includeBody"] = new { type = "boolean", description = "Include the email body text. Default: true", @default = true },
                    ["bodyFormat"] = new { type = "string", description = "Body format: 'text' or 'html'. Default: 'text'", @default = "text" }
                },
                required = new[] { "instanceId" }
            }
        },
        new
        {
            name = "outlook_summarize_thread",
            description = "Returns a structured summary of the selected email thread with timeline, key decisions, action items, and unresolved questions. Processes up to 50 messages by default.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["maxMessages"] = new { type = "integer", description = "Maximum messages to process. Default: 50", @default = 50 }
                },
                required = new[] { "instanceId" }
            }
        },

        // ── Outlook Write tools (never auto-send) ──────────────────
        new
        {
            name = "outlook_draft_reply",
            description = "Creates a draft reply in Outlook's Drafts folder for the selected email. The draft is NEVER sent automatically — the user must review and send from Outlook. Supports tone control and key points.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["tone"] = new { type = "string", description = "Reply tone: 'concise', 'formal', 'friendly', or 'technical'. Default: 'concise'", @default = "concise", @enum = new[] { "concise", "formal", "friendly", "technical" } },
                    ["keyPoints"] = new { type = "array", description = "Key points to address in the reply", items = new { type = "string" } },
                    ["includeThreadSummary"] = new { type = "boolean", description = "Include a thread summary in the draft. Default: false", @default = false }
                },
                required = new[] { "instanceId" }
            }
        },
        new
        {
            name = "outlook_apply_category",
            description = "Applies an Outlook category (color-coded label) to the currently selected email(s). Returns available categories if the requested one doesn't exist.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["categoryName"] = new { type = "string", description = "Name of the Outlook category to apply" }
                },
                required = new[] { "instanceId", "categoryName" }
            }
        },
        new
        {
            name = "outlook_send_message",
            description = "Sends a drafted message. REQUIRES explicit user confirmation via the Outlook task pane — a confirmation token must be provided. NEVER auto-sends. Subject to policy filters.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["confirmationToken"] = new { type = "string", description = "REQUIRED. Confirmation token obtained from explicit user approval in the Outlook task pane." },
                    ["messageId"] = new { type = "string", description = "Optional. ID of the draft message to send. If omitted, sends the currently open draft." }
                },
                required = new[] { "instanceId", "confirmationToken" }
            }
        }
    ];

    /// <summary>
    /// All command names that should be dispatched to the add-in.
    /// </summary>
    private static readonly HashSet<string> AddInCommands = new(StringComparer.OrdinalIgnoreCase)
    {
        "powerpoint_get_deck_outline",
        "powerpoint_get_slide",
        "powerpoint_get_slide_image",
        "powerpoint_get_shape_image",
        "powerpoint_get_table",
        "powerpoint_get_selection",
        "powerpoint_get_speaker_notes",
        "powerpoint_update_shape_text",
        "powerpoint_update_shape_properties",
        "powerpoint_update_speaker_notes",
        "powerpoint_add_textbox",
        "powerpoint_add_image",
        "powerpoint_add_table",
        "powerpoint_delete_shape",
        "powerpoint_add_slide",
        "powerpoint_delete_slide",
        "powerpoint_move_slide",

        // Word
        "word_get_outline",
        "word_get_paragraphs",
        "word_get_selection",
        "word_search",
        "word_replace_text",
        "word_insert_text",
        "word_add_comment",
        "word_delete_paragraph",

        // Word tracked changes
        "word_get_tracked_changes",
        "word_accept_all_changes",
        "word_reject_all_changes",

        // Excel
        "excel_get_workbook_map",
        "excel_read_range",
        "excel_write_range",
        "excel_write_formula",
        "excel_create_table",

        // Outlook
        "outlook_get_current_item",
        "outlook_summarize_thread",
        "outlook_draft_reply",
        "outlook_apply_category",
        "outlook_send_message",
    };

    /// <summary>
    /// Executes an MCP tool call. Routes to the appropriate add-in instance.
    /// </summary>
    public static async Task<object> ExecuteTool(string name, JsonElement? args)
    {
        // office_get_active_apps is handled server-side
        if (name is "office_get_active_apps" or "office_get_active_app")
            return HandleGetActiveApps();

        // All other tools require an explicit instanceId
        string? instanceId = null;
        if (args.HasValue && args.Value.TryGetProperty("instanceId", out var iid))
            instanceId = iid.GetString();

        if (string.IsNullOrEmpty(instanceId))
        {
            return new
            {
                content = new[] { new { type = "text", text = "Missing required parameter: instanceId. Call office_get_active_apps first to get the list of available instances." } },
                isError = true
            };
        }

        // Check if instance exists
        var instance = _registry.GetInstance(instanceId);
        if (instance == null)
        {
            return new
            {
                content = new[] { new { type = "text", text = $"Instance '{instanceId}' is not registered or has timed out. Call office_get_active_apps to see current instances." } },
                isError = true
            };
        }

        // Validate tool name
        if (!AddInCommands.Contains(name))
        {
            return new
            {
                content = new[] { new { type = "text", text = $"Unknown tool: {name}" } },
                isError = true
            };
        }

        // Dispatch command to add-in
        var inputs = args.HasValue ? JsonSerializer.Serialize(args.Value) : "{}";
        string commandId = Guid.NewGuid().ToString("N")[..8];
        var cmd = new PendingCommand
        {
            Id = commandId,
            InstanceId = instanceId,
            Command = name,
            Args = args.HasValue ? JsonSerializer.Deserialize<object>(args.Value) : null,
            CreatedAt = DateTime.UtcNow,
        };
        _commandStore.AddCommand(cmd);

        _auditLog.Log(new AuditEntry
        {
            ToolName = name,
            InstanceId = instanceId,
            Inputs = inputs,
            Outcome = "pending"
        });

        var result = await _commandStore.WaitForResult(commandId, timeoutSeconds: IsImageTool(name) ? 120 : 60);
        return BuildToolResult(result, name, instanceId, inputs);
    }

    private static object HandleGetActiveApps()
    {
        var instances = _registry.GetActiveInstances();

        _auditLog.Log(new AuditEntry
        {
            ToolName = "office_get_active_apps",
            InstanceId = "",
            Inputs = "",
            Outcome = "success"
        });

        var appList = instances.Select(i => new
        {
            instanceId = i.InstanceId,
            appName = i.AppName,
            documentName = i.DocumentName
        }).ToList();

        return new
        {
            content = new[] { new { type = "text", text = JsonSerializer.Serialize(new { apps = appList, total = appList.Count }, new JsonSerializerOptions { WriteIndented = true }) } },
            isError = false
        };
    }

    private static object BuildToolResult(PendingCommand? result, string toolName, string instanceId, string inputs)
    {
        if (result == null)
        {
            _auditLog.Log(new AuditEntry
            {
                ToolName = toolName,
                InstanceId = instanceId,
                Inputs = inputs,
                Outcome = "timeout"
            });

            return new
            {
                content = new[] { new { type = "text", text = "Command timed out waiting for add-in response." } },
                isError = true
            };
        }

        if (result.Success)
        {
            _auditLog.Log(new AuditEntry
            {
                ToolName = toolName,
                InstanceId = instanceId,
                Inputs = inputs,
                Outcome = "success"
            });

            return new
            {
                content = new[] { new { type = "text", text = JsonSerializer.Serialize(result.Payload, new JsonSerializerOptions { WriteIndented = true }) } },
                isError = false
            };
        }

        _auditLog.Log(new AuditEntry
        {
            ToolName = toolName,
            InstanceId = instanceId,
            Inputs = inputs,
            Outcome = "error",
            Error = result.Error
        });

        return new
        {
            content = new[] { new { type = "text", text = $"Command failed: {result.Error}" } },
            isError = true
        };
    }

    /// <summary>Gets the registry for instance management endpoints.</summary>
    public static InstanceRegistry GetRegistry() => _registry;

    /// <summary>Gets the command store for dispatch/wait operations.</summary>
    public static CommandStore GetCommandStore() => _commandStore;

    /// <summary>
    /// Resets all static state. Used by tests to isolate test runs.
    /// </summary>
    public static void ResetForTesting()
    {
        _registry = new InstanceRegistry();
        _commandStore = new CommandStore();
        _auditLog = new AuditLog(Path.Combine(Path.GetTempPath(), $"audit-test-{Guid.NewGuid()}").TrimEnd('/'));
    }

    private static bool IsImageTool(string toolName) =>
        toolName is "powerpoint_get_slide_image" or "powerpoint_get_shape_image";
}
