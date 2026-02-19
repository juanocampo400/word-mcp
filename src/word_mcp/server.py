"""
FastMCP server for word-mcp.

This module provides the main MCP server instance and registers all document
lifecycle tools. The server exposes Word document manipulation capabilities
through the Model Context Protocol.

Entry point: Run with `python -m word_mcp.server` or via `word-mcp` command.
"""

from contextlib import asynccontextmanager
from mcp.server.fastmcp import FastMCP

from .logging_config import get_logger
from .com_pool import com_pool
from .document_manager import document_manager

logger = get_logger(__name__)

from .tools.document import (
    create_document,
    open_document,
    save_document,
    save_document_as,
    close_document,
    get_document_info,
    create_from_template,
    list_open_documents,
)
from .tools.text import (
    add_paragraph,
    edit_paragraph,
    delete_paragraph,
    read_document,
)
from .tools.search import (
    search_text,
    replace_text,
)
from .tools.styles import (
    apply_heading_style,
    apply_style,
)
from .tools.tracked_changes import (
    enable_tracked_changes,
    disable_tracked_changes,
    get_tracked_changes,
)
from .tools.tracked_editing import (
    tracked_add_paragraph,
    tracked_edit_paragraph,
    tracked_delete_paragraph,
)
from .tools.formatting import (
    format_text,
    get_paragraph_formatting,
)
from .tools.comments import (
    get_comments,
)
from .tools.tables import (
    create_table,
    list_tables,
    read_table,
    edit_table_cell,
    add_table_row,
    add_table_column,
)
from .tools.tables_com import (
    delete_table_row,
    delete_table_column,
    tracked_edit_table_cell,
)
from .tools.images import (
    insert_image,
    resize_image,
    list_images,
)
from .tools.images_com import (
    reposition_image,
)
from .tools.sections import (
    list_sections,
    add_section,
    modify_section_properties,
)
from .tools.headers_footers import (
    get_header,
    set_header,
    get_footer,
    set_footer,
)
from .tools.monitoring import (
    get_server_health,
)


@asynccontextmanager
async def app_lifespan(server):
    """
    Lifespan context manager for server initialization and cleanup.

    Handles:
    - Startup: Logs server initialization
    - Shutdown: Cleans up COM pool and document manager resources

    This ensures graceful shutdown with no zombie WINWORD.EXE processes
    and proper cleanup of in-memory document state.
    """
    logger.info("server_starting", name="word-mcp")
    try:
        yield {}
    finally:
        logger.info("server_shutting_down")

        # Close all COM instances
        com_pool.close_all()

        # Close all open documents
        doc_count = document_manager.close_all()
        if doc_count > 0:
            logger.info("documents_closed_on_shutdown", count=doc_count)

        logger.info("server_shutdown_complete")


# Create FastMCP server instance with lifespan
mcp = FastMCP("word-mcp", lifespan=app_lifespan)


# Register document lifecycle tools
@mcp.tool()
def create_document_tool(path: str = None) -> str:
    """
    Create a new blank Word document in memory.

    Creates a new blank document that is held in memory (not saved to disk
    until explicitly requested via save_document or save_document_as).

    ERROR-ON-OVERWRITE: If a path is provided and a file already exists at
    that location, returns an error instead of overwriting. This is a locked
    design decision to prevent accidental data loss.

    Args:
        path: Optional file path for the document. If provided:
              - Must not point to an existing file (error-on-overwrite)
              - Document will be keyed by this path in memory
              - Still requires explicit save to persist to disk
              If None:
              - Creates document with temporary "Untitled-N" name
              - Can be saved later with save_document_as

    Returns:
        Success message with document name/path, or error message

    Examples:
        Create untitled document:
        >>> create_document_tool()
        "Created new document 'Untitled-1'"

        Create with specific path (file must not exist):
        >>> create_document_tool("C:/Documents/report.docx")
        "Created new document at C:\\Documents\\report.docx"

        Error case - file exists:
        >>> create_document_tool("C:/Documents/existing.docx")
        "Error: File already exists at C:\\Documents\\existing.docx. Use a different path or open the existing document."

    Design notes:
        - No auto-save: Document exists only in memory until explicitly saved
        - Multi-document: Can have multiple documents open simultaneously
        - Path normalization: All paths converted to absolute internally
    """
    return create_document(path)


@mcp.tool()
def open_document_tool(path: str) -> str:
    """
    Open an existing Word document from disk into memory.

    Loads a .docx file from the file system into memory for editing. If the
    document is already open, returns the cached instance (no reload).

    Args:
        path: Path to .docx file to open (relative or absolute)

    Returns:
        Success message with filename and document statistics (paragraph count,
        approximate word count), or error message if file not found

    Examples:
        >>> open_document_tool("C:/Documents/report.docx")
        "Opened 'report.docx' (12 paragraphs, ~350 words)"

        >>> open_document_tool("missing.docx")
        "Error: File not found: missing.docx"

    Design notes:
        - Idempotent: Opening an already-open document returns cached instance
        - Statistics: Provides quick overview of document size upon opening
        - No auto-save: Changes remain in memory until explicitly saved
    """
    return open_document(path)


@mcp.tool()
def save_document_tool(path: str) -> str:
    """
    Save an open document to its current path.

    EXPLICIT SAVE ONLY: This is the ONLY way changes persist to disk. No
    auto-save occurs anywhere in the system. This is a locked design decision
    to give users full control over when disk writes occur.

    Args:
        path: Path or key of open document to save
              - For path-based documents: saves to that path
              - For untitled documents: error (use save_document_as instead)

    Returns:
        Success message with absolute path, or error message

    Examples:
        Save document to its path:
        >>> save_document_tool("C:/Documents/report.docx")
        "Saved document to C:\\Documents\\report.docx"

        Error - untitled document:
        >>> save_document_tool("Untitled-1")
        "Error: Cannot save untitled document without path. Use save_document_as instead."

        Error - document not open:
        >>> save_document_tool("C:/Documents/missing.docx")
        "Error: Document not open: C:/Documents/missing.docx"

    Design notes:
        - Creates parent directories automatically if needed
        - Overwrites existing file at path (since document already open from there)
        - Untitled documents must use save_document_as to specify path
    """
    return save_document(path)


@mcp.tool()
def save_document_as_tool(path: str, new_path: str) -> str:
    """
    Save an open document to a new path (save-as operation).

    Saves the document to a new location and re-keys it to the new path.
    The original file (if any) is unaffected. The document instance in memory
    is now associated with the new path.

    Args:
        path: Current path or key of open document
        new_path: New path to save document to

    Returns:
        Success message showing old and new paths, or error message

    Examples:
        Save untitled document:
        >>> save_document_as_tool("Untitled-1", "C:/Documents/report.docx")
        "Saved document to C:\\Documents\\report.docx (was: Untitled-1)"

        Save existing document to new location:
        >>> save_document_as_tool("C:/Documents/draft.docx", "C:/Documents/final.docx")
        "Saved document to C:\\Documents\\final.docx (was: C:\\Documents\\draft.docx)"

        Error - document not open:
        >>> save_document_as_tool("missing.docx", "new.docx")
        "Error: Document not open: missing.docx"

    Design notes:
        - Original file unchanged: Creates new file, doesn't affect original
        - Re-keying: Document is now referenced by new_path
        - Creates parent directories automatically if needed
    """
    return save_document_as(path, new_path)


@mcp.tool()
def close_document_tool(path: str) -> str:
    """
    Close an open document (remove from memory).

    EXPLICIT-ONLY BEHAVIOR: Unsaved changes are DISCARDED without warning or
    prompt. This is a locked design decision for simplicity. Users must save
    before closing if they want to keep changes.

    Args:
        path: Path or key of document to close

    Returns:
        Success message confirming close and change discard, or error if not open

    Examples:
        Close document:
        >>> close_document_tool("C:/Documents/report.docx")
        "Closed document 'C:\\Documents\\report.docx'. Unsaved changes were discarded."

        Close untitled document:
        >>> close_document_tool("Untitled-1")
        "Closed document 'Untitled-1'. Unsaved changes were discarded."

        Error - not open:
        >>> close_document_tool("missing.docx")
        "Error: Document not open: missing.docx"

    Design notes:
        - No confirmation: Changes lost immediately if not saved
        - Memory cleanup: Document instance removed from manager
        - Multi-document: Other open documents unaffected
    """
    return close_document(path)


@mcp.tool()
def get_document_info_tool(path: str) -> str:
    """
    Get detailed information about an open document.

    LOCKED DECISION: Must include page count, styles used, and word count.

    Provides comprehensive document statistics and metadata including:
    - File path and name
    - Paragraph, word, and character counts
    - List of all styles used in the document
    - Page count (approximate, requires document saved to disk)
    - Core properties (title, author) if available

    Args:
        path: Path or key of open document

    Returns:
        Multi-line formatted string with document information, or error message

    Examples:
        >>> get_document_info_tool("C:/Documents/report.docx")
        '''Document Information: C:\\Documents\\report.docx

        File: report.docx
        Paragraphs: 42
        Words: ~1,250
        Characters: 7,890

        Styles in use:
          - Heading 1
          - Heading 2
          - List Paragraph
          - Normal

        Page count: 5

        Title: Q4 Sales Report
        Author: John Smith
        '''

        >>> get_document_info_tool("Untitled-1")
        '''Document Information: Untitled-1

        File: Untitled-1
        Paragraphs: 0
        Words: ~0
        Characters: 0

        Styles in use:
          (none)

        Page count: unavailable (save document first)

        Title: Not set
        Author: Not set
        '''

    Design notes:
        - Word count approximate: Based on whitespace splitting
        - Page count requires save: Extracted from docProps/app.xml in .docx ZIP
        - Styles: Unique set of paragraph styles used anywhere in document
        - Core properties: May be "Not set" for newly created documents
    """
    return get_document_info(path)


@mcp.tool()
def create_from_template_tool(template_path: str, save_path: str = None) -> str:
    """
    Create a new document from a .docx or .dotx template.

    Opens the template file and creates a new document instance based on it.
    The template file itself is not modified. The new document inherits all
    content, styles, and formatting from the template.

    Args:
        template_path: Path to template file (.docx or .dotx)
        save_path: Optional path where document will be saved. If provided:
                   - Must not point to existing file (error-on-overwrite)
                   - Document keyed by this path in memory
                   - Still requires explicit save to persist to disk
                   If None:
                   - Creates with temporary "Untitled-N" name
                   - Can be saved later with save_document_as

    Returns:
        Success message with template name and new document key, or error message

    Examples:
        Create from template without path:
        >>> create_from_template_tool("C:/Templates/report.dotx")
        "Created document from template 'report.dotx' as 'Untitled-1'"

        Create from template with specific path:
        >>> create_from_template_tool("C:/Templates/report.dotx", "C:/Documents/q4-report.docx")
        "Created document from template 'report.dotx' at C:\\Documents\\q4-report.docx"

        Error - template not found:
        >>> create_from_template_tool("missing.dotx")
        "Error: Template not found: missing.dotx"

        Error - save_path exists:
        >>> create_from_template_tool("template.dotx", "existing.docx")
        "Error: File already exists at existing.docx. Use a different path."

    Design notes:
        - Template unchanged: Source template file is never modified
        - Both .docx and .dotx: Accepts both Word documents and templates
        - Error-on-overwrite: Same protection as create_document
        - No auto-save: Document exists in memory until explicitly saved
    """
    return create_from_template(template_path, save_path)


@mcp.tool()
def list_open_documents_tool() -> str:
    """
    List all currently open documents.

    Returns a formatted list of all document keys/paths currently held in memory.
    Useful for checking which documents are available for editing.

    Returns:
        Formatted numbered list of open documents, or message if none open

    Examples:
        With documents open:
        >>> list_open_documents_tool()
        '''Open documents:
          1. C:\\Documents\\report.docx
          2. C:\\Documents\\memo.docx
          3. Untitled-1
        '''

        With no documents open:
        >>> list_open_documents_tool()
        "No documents are currently open"

    Design notes:
        - Multi-document awareness: Shows all documents in memory
        - Path types: Includes both file paths and "Untitled-N" keys
        - Order: Documents listed in iteration order (insertion order in dict)
    """
    return list_open_documents()


# Register text editing tools
@mcp.tool()
def read_document_tool(
    path: str, start_index: int = None, end_index: int = None
) -> str:
    """
    Read document content as a list of indexed paragraphs.

    Returns all paragraphs with their indexes, styles, and text content.
    This is how to "see" what's in a document – use this frequently to check
    document state before editing.

    PAGINATION: Supports optional start_index/end_index for large documents.
    By default returns the full document.

    Args:
        path: Document path or key
        start_index: Optional starting paragraph index (0-based, inclusive)
        end_index: Optional ending paragraph index (0-based, inclusive)

    Returns:
        Formatted paragraph list with indexes, styles, and text previews

    Examples:
        Read full document:
        >>> read_document_tool("C:/Documents/report.docx")
        '''Document: C:\\Documents\\report.docx | Paragraphs: 3 | Showing: 0-2
        [0] (Heading 1) Introduction
        [1] (Normal) This is the first paragraph of content...
        [2] (Normal) This is the second paragraph.
        '''

        Read specific range:
        >>> read_document_tool("report.docx", start_index=10, end_index=15)
        '''Document: C:\\Documents\\report.docx | Paragraphs: 50 | Showing: 10-15
        [10] (Normal) Paragraph 10 text...
        ...
        [15] (Normal) Paragraph 15 text...
        '''

    Design notes:
        - Zero-based indexing: First paragraph is index 0
        - Text preview: Shows up to 200 chars, truncated with "..." if longer
        - Style info: Shows paragraph style name for each paragraph
        - Empty documents: Returns "Document is empty (0 paragraphs with content)"
    """
    return read_document(path, start_index, end_index)


@mcp.tool()
def add_paragraph_tool(
    path: str, text: str, position: int = None, style: str = None
) -> str:
    """
    Add a paragraph to the document.

    By default, APPENDS to the end of the document. Can optionally INSERT at
    a specific position (shifts subsequent paragraphs down).

    Args:
        path: Document path or key
        text: Paragraph text content
        position: Optional 0-based index to insert at (None = append to end)
                  - If None: appends to end
                  - If 0: inserts at beginning
                  - If N: inserts before paragraph N (shifts N and later down)
                  - If equals paragraph count: appends to end
        style: Optional paragraph style name (e.g., "Normal", "Heading 1")

    Returns:
        Success message with index where paragraph was added and updated count

    Examples:
        Append to end:
        >>> add_paragraph_tool("report.docx", "This is a new paragraph")
        "Added paragraph at index 5: 'This is a new paragraph'\nDocument now has 6 paragraphs."

        Insert at beginning:
        >>> add_paragraph_tool("report.docx", "New intro", position=0)
        "Added paragraph at index 0: 'New intro'\nDocument now has 7 paragraphs."

        Add with style:
        >>> add_paragraph_tool("report.docx", "Chapter 1", style="Heading 1")
        "Added paragraph at index 7: 'Chapter 1'\nDocument now has 8 paragraphs."

    Design notes:
        - Zero-based indexing: position=0 inserts at beginning
        - Index shifting: Insertion shifts subsequent paragraph indexes up
        - Default append: Most common use case is adding to end (position=None)
        - Style parameter: Can set style during creation
    """
    return add_paragraph(path, text, position, style)


@mcp.tool()
def edit_paragraph_tool(path: str, index: int, new_text: str) -> str:
    """
    Edit (replace) the text of an existing paragraph by index.

    INDEX REFERENCING: Use read_document first to see current paragraph indexes.
    Indexes are zero-based (first paragraph is 0).

    FORMATTING LOSS WARNING: This replaces all text runs and loses any formatting
    (bold, italic, font changes, etc.) within the paragraph. This is acceptable
    for Phase 1 which has no formatting requirements. The paragraph style
    (e.g., "Heading 1") is preserved.

    Args:
        path: Document path or key
        index: 0-based paragraph index to edit
        new_text: New text content to replace existing text

    Returns:
        Success message showing before/after text previews

    Examples:
        Edit paragraph:
        >>> edit_paragraph_tool("report.docx", 0, "Updated first paragraph")
        "Edited paragraph 0. Was: 'Original text...' -> Now: 'Updated first paragraph'\nDocument has 5 paragraphs."

        Error case - invalid index:
        >>> edit_paragraph_tool("report.docx", 100, "New text")
        "Error: Invalid paragraph index 100. Document has 5 paragraphs (valid range: 0-4)."

    Design notes:
        - Zero-based indexing: First paragraph is index 0
        - Index validation: Returns clear error if index out of bounds
        - Text previews: Shows first 50 chars of old and new text
        - Paragraph style preserved: Style like "Heading 1" stays unchanged
    """
    return edit_paragraph(path, index, new_text)


@mcp.tool()
def delete_paragraph_tool(path: str, index: int) -> str:
    """
    Delete a paragraph by index.

    INDEX SHIFT WARNING: After deletion, all subsequent paragraphs shift down
    by one index position. ALWAYS re-read the document with read_document after
    deletion before performing additional operations. This is critical for
    correct multi-step editing.

    Args:
        path: Document path or key
        index: 0-based paragraph index to delete

    Returns:
        Success message with shift warning and updated paragraph count

    Examples:
        Delete paragraph:
        >>> delete_paragraph_tool("report.docx", 2)
        "Deleted paragraph 2 ('Old paragraph text'). Remaining paragraphs have shifted -- re-read document to get updated indexes. Document now has 4 paragraphs."

        Error case - invalid index:
        >>> delete_paragraph_tool("report.docx", 10)
        "Error: Invalid paragraph index 10. Document has 5 paragraphs (valid range: 0-4)."

    Design notes:
        - Zero-based indexing: First paragraph is index 0
        - Index shifting: Paragraph 3 becomes paragraph 2 after deleting paragraph 2
        - Re-read required: Must call read_document to see new indexes
        - Text preview: Shows first 50 chars of deleted text for confirmation
    """
    return delete_paragraph(path, index)


# Register search tools
@mcp.tool()
def search_text_tool(path: str, query: str, case_sensitive: bool = False, use_regex: bool = False) -> str:
    """
    Search for text in document paragraphs.

    Supports plain text search (default) and regex search via use_regex=True.

    Returns all paragraphs containing the query text, with match counts and context.

    Args:
        path: Document path or key
        query: Text string to search for. When use_regex=True, this is a regex pattern.
        case_sensitive: If False, performs case-insensitive search (default: False)
        use_regex: If True, treats query as a regular expression pattern (default: False).
                   Returns a clear error if the pattern is invalid.

    Returns:
        Formatted search results showing paragraph indexes, match counts, and context

    Examples:
        Case-insensitive plain text search (default):
        >>> search_text_tool("report.docx", "paragraph")
        '''Found 2 match(es) in 2 paragraph(s):
        [1] 1 match(es): "This is the first paragraph of content."
        [2] 1 match(es): "This is the second paragraph of content."
        '''

        Case-sensitive search:
        >>> search_text_tool("report.docx", "Paragraph", case_sensitive=True)
        "No matches found for 'Paragraph'"

        Regex search -- whole-word match:
        >>> search_text_tool("report.docx", r"\\bparagraph\\b", use_regex=True)
        "Found 2 match(es) in 2 paragraph(s): ..."

        Regex search -- ISO date pattern:
        >>> search_text_tool("report.docx", r"\\d{4}-\\d{2}-\\d{2}", use_regex=True)
        "Found 3 match(es) in 3 paragraph(s): ..."

        Invalid regex pattern:
        >>> search_text_tool("report.docx", r"[unclosed", use_regex=True)
        "Error: Invalid regex pattern: unterminated character set at position 0"

    Design notes:
        - Context window: Shows 50 chars before/after match, or full paragraph if short
        - Multiple matches per paragraph: Shows total count per paragraph
        - Zero-based indexing: Paragraph indexes shown for easy editing
        - Regex error handling: Invalid patterns return a clear error message
    """
    return search_text(path, query, case_sensitive, use_regex)


@mcp.tool()
def replace_text_tool(
    path: str,
    find_text: str,
    replace_with: str,
    case_sensitive: bool = False,
    replace_all: bool = True,
) -> str:
    """
    Find and replace text across the document.

    FORMATTING RESET WARNING: Replacements occur at the paragraph level.
    Formatting within replaced paragraphs (bold, italic, etc.) will be reset
    to default. This is acceptable for Phase 1 which has no formatting requirements.

    REPLACE MODES:
    - replace_all=True (default): Replaces ALL occurrences throughout document
    - replace_all=False: Replaces only the FIRST occurrence in first matching paragraph, then stops

    Args:
        path: Document path or key
        find_text: Text to find (plain text, not regex)
        replace_with: Text to replace with
        case_sensitive: If False, performs case-insensitive replacement (default: False)
        replace_all: If True, replaces all occurrences; if False, first only (default: True)

    Returns:
        Summary of replacements made with counts

    Examples:
        Replace all (case-insensitive):
        >>> replace_text_tool("report.docx", "paragraph", "section")
        "Replaced 5 occurrence(s) of 'paragraph' with 'section' in 3 paragraph(s)."

        Replace first only:
        >>> replace_text_tool("report.docx", "TODO", "DONE", replace_all=False)
        "Replaced 1 occurrence(s) of 'TODO' with 'DONE' in 1 paragraph(s)."

        Case-sensitive:
        >>> replace_text_tool("report.docx", "Word", "Microsoft Word", case_sensitive=True)
        "Replaced 2 occurrence(s) of 'Word' with 'Microsoft Word' in 2 paragraph(s)."

        No matches:
        >>> replace_text_tool("report.docx", "missing", "found")
        "No occurrences of 'missing' found."

    Design notes:
        - Paragraph-level replacement: Entire paragraph text replaced
        - Formatting loss: In-paragraph formatting (bold, etc.) is reset
        - Paragraph style preserved: Style like "Heading 1" stays unchanged
        - Case-insensitive uses regex: Preserves non-matched casing in text
    """
    return replace_text(path, find_text, replace_with, case_sensitive, replace_all)


# Register style tools
@mcp.tool()
def apply_heading_style_tool(path: str, index: int, level: int) -> str:
    """
    Apply a heading style (H1-H9) to a paragraph.

    HEADING LEVELS: level 1-9 maps to "Heading 1" through "Heading 9" styles.

    Args:
        path: Document path or key
        index: 0-based paragraph index
        level: Heading level 1-9
               - 1 = "Heading 1" (typically largest, for document title or main sections)
               - 2 = "Heading 2" (subsections)
               - 3 = "Heading 3" (sub-subsections)
               - ... up to 9

    Returns:
        Success message with applied style and paragraph preview

    Examples:
        Apply H1:
        >>> apply_heading_style_tool("report.docx", 0, 1)
        "Applied 'Heading 1' style to paragraph 0: 'Introduction'"

        Apply H2:
        >>> apply_heading_style_tool("report.docx", 5, 2)
        "Applied 'Heading 2' style to paragraph 5: 'Background'"

        Error - invalid level:
        >>> apply_heading_style_tool("report.docx", 0, 10)
        "Error: Invalid heading level 10. Valid range is 1-9 (for Heading 1 through Heading 9)."

        Error - invalid index:
        >>> apply_heading_style_tool("report.docx", 100, 1)
        "Error: Invalid paragraph index 100. Document has 5 paragraphs (valid range: 0-4)."

    Design notes:
        - Zero-based indexing: First paragraph is index 0
        - Level validation: Must be 1-9
        - Text preserved: Only style changes, text content unchanged
    """
    return apply_heading_style(path, index, level)


@mcp.tool()
def apply_style_tool(path: str, index: int, style_name: str) -> str:
    """
    Apply a paragraph style by name to a paragraph.

    General-purpose style application for any paragraph style (Normal, Title,
    Heading 1, List Paragraph, Quote, etc.).

    AVAILABLE STYLES: Use get_document_info to see available paragraph styles
    in the current document. Common styles include:
    - "Normal" (default body text)
    - "Heading 1" through "Heading 9"
    - "Title"
    - "Subtitle"
    - "Quote"
    - "List Paragraph"

    Args:
        path: Document path or key
        index: 0-based paragraph index
        style_name: Exact style name (case-sensitive)

    Returns:
        Success message, or error if style not found in document

    Examples:
        Apply Normal style:
        >>> apply_style_tool("report.docx", 0, "Normal")
        "Applied 'Normal' style to paragraph 0"

        Apply Title style:
        >>> apply_style_tool("report.docx", 0, "Title")
        "Applied 'Title' style to paragraph 0"

        Error - style not found:
        >>> apply_style_tool("report.docx", 0, "InvalidStyle")
        "Error: Style 'InvalidStyle' not found. Available paragraph styles: Normal, Heading 1, Heading 2, Title, ..."

    Design notes:
        - Zero-based indexing: First paragraph is index 0
        - Style validation: Checks document's available styles
        - Text preserved: Only style changes, text content unchanged
        - Case-sensitive: Style name must match exactly
    """
    return apply_style(path, index, style_name)


# Register tracked changes tools (Phase 2)
@mcp.tool()
def enable_tracked_changes_tool(path: str, author: str = "Claude") -> str:
    """
    Enable tracked changes on a document.

    PREREQUISITE: Document must be saved to disk first (use save_document or
    save_document_as before calling this tool).

    After enabling tracked changes, use the tracked_add_paragraph, tracked_edit_paragraph,
    and tracked_delete_paragraph tools to make edits that Word records as revisions.

    IMPORTANT: Phase 1 tools (add_paragraph, edit_paragraph, delete_paragraph) still
    work after enabling tracking, but they do NOT create tracked changes. Use the
    tracked_* versions for tracked edits.

    Args:
        path: Path or key of open document
        author: Author name for all future tracked revisions (default: "Claude")

    Returns:
        Success message or error message prefixed with "Error:"

    Examples:
        Enable tracking:
        >>> enable_tracked_changes_tool("C:/Documents/report.docx", "Claude")
        "Tracked changes enabled on 'report.docx'. Author set to 'Claude'. All subsequent COM-based edits will be tracked."

        Error - not saved:
        >>> enable_tracked_changes_tool("Untitled-1")
        "Error: Document must be saved to disk first. Use save_document_as."

    Design notes:
        - Requires saved document: COM opens files from disk, not memory
        - Sets author attribution: All subsequent revisions show this author
        - Bridge pattern: Uses COM to set TrackRevisions=True, then reloads python-docx
        - Preserves existing content: Only enables tracking, doesn't modify document
    """
    return enable_tracked_changes(path, author)


@mcp.tool()
def disable_tracked_changes_tool(path: str) -> str:
    """
    Disable tracked changes on a document.

    Stops tracking new edits. Existing revisions are preserved in the document
    and remain visible/accept/rejectable in Word.

    After disabling, both Phase 1 tools (add_paragraph, etc.) and Phase 2 tools
    (tracked_add_paragraph, etc.) produce untracked edits.

    Args:
        path: Path or key of open document

    Returns:
        Success message or error message prefixed with "Error:"

    Examples:
        Disable tracking:
        >>> disable_tracked_changes_tool("C:/Documents/report.docx")
        "Tracked changes disabled on 'report.docx'. Future edits will not be tracked. Existing revisions are preserved."

    Design notes:
        - Preserves existing revisions: Only disables future tracking
        - Requires saved document: COM opens files from disk
        - Bridge pattern: Uses COM to set TrackRevisions=False, then reloads python-docx
    """
    return disable_tracked_changes(path)


@mcp.tool()
def get_tracked_changes_tool(path: str) -> str:
    """
    Read all tracked changes with metadata.

    Returns information about all revisions in the document including type
    (Insertion, Deletion, Property, etc.), author, date, and text involved.

    PREREQUISITE: Document must be saved to disk first.

    Useful for reviewing what changes have been made before accepting or rejecting
    them in Microsoft Word.

    Args:
        path: Path or key of open document

    Returns:
        Formatted list of tracked changes or error message prefixed with "Error:"

    Examples:
        Read changes:
        >>> get_tracked_changes_tool("C:/Documents/report.docx")
        '''Tracked changes in 'report.docx': 3 revision(s)

        [1] Insertion by 'Claude' on 2026-02-13 14:30:00
            Text: "inserted text here"
        [2] Deletion by 'John' on 2026-02-13 14:31:00
            Text: "deleted text here"
        [3] Property by 'Claude' on 2026-02-13 14:32:00
            Text: "formatted text"
        '''

        No changes:
        >>> get_tracked_changes_tool("C:/Documents/clean.docx")
        "No tracked changes found in 'clean.docx'."

    Design notes:
        - Read-only operation: Doesn't modify document or revisions
        - Requires saved document: COM opens files from disk
        - 1-based revision indexing: Matches Word's revision numbering
        - Bridge pattern: Uses COM to read Revisions collection
    """
    return get_tracked_changes(path)


@mcp.tool()
def tracked_add_paragraph_tool(
    path: str, text: str, position: str = "end", author: str = "Claude",
    expected_text: str = None
) -> str:
    """
    Add a paragraph that appears as an Insertion revision in Word.

    REQUIRES: Tracked changes must be enabled first (call enable_tracked_changes).

    This is different from add_paragraph: this version uses COM automation to
    create a paragraph that Word records as a tracked change. The user will see
    it as an insertion (colored/underlined) in Word and can accept or reject it.

    INDEX TRANSLATION: Internally translates python-docx body paragraph indexes to
    COM paragraph indexes, so documents with tables are handled correctly. The index
    you pass should match what read_document_tool reports.

    Args:
        path: Path or key of open document
        text: Paragraph text content
        position: "end" to append, or zero-based index string to insert before
        author: Author name for this tracked change (default: "Claude")
        expected_text: Optional content verification string. If provided and position
                       is not "end", the paragraph at that position must contain this
                       text (case-sensitive partial match) before inserting. If it does
                       not match, insert is refused with an error message including the
                       index, expected text, and actual text preview. Ignored for "end"
                       position (nothing to verify at end of document).

    Returns:
        Success message or error message prefixed with "Error:"

    Examples:
        Append tracked paragraph:
        >>> tracked_add_paragraph_tool("C:/Documents/report.docx", "New conclusion")
        "Added tracked paragraph at end: 'New conclusion'. Revision will appear as insertion by 'Claude'."

        Insert tracked paragraph:
        >>> tracked_add_paragraph_tool("C:/Documents/report.docx", "New intro", "0")
        "Added tracked paragraph at 0: 'New intro'. Revision will appear as insertion by 'Claude'."

        Insert with content verification:
        >>> tracked_add_paragraph_tool("C:/Documents/report.docx", "New section", "5", expected_text="Background")
        "Added tracked paragraph at 5: 'New section'. Revision will appear as insertion by 'Claude'."

        Error - content mismatch (indexes have shifted):
        >>> tracked_add_paragraph_tool("C:/Documents/report.docx", "New section", "5", expected_text="Background")
        "Error: Content verification failed for paragraph 5. Expected text containing 'Background' but found: 'Introduction to the project...'. The paragraph may have shifted -- re-read the document."

        Error - tracking not enabled:
        >>> tracked_add_paragraph_tool("C:/Documents/report.docx", "Text")
        "Error: Tracked changes are not enabled. Call enable_tracked_changes first."

    Design notes:
        - Zero-based indexing: position="0" inserts at beginning, consistent with add_paragraph
        - Requires tracking enabled: Returns error if TrackRevisions=False
        - Bridge pattern: Uses COM to add text, creates Insertion revision
        - Author attribution: Sets UserName in Word before adding
        - Index translation: COM paragraph indexes adjusted to skip table cell paragraphs
    """
    return tracked_add_paragraph(path, text, position, author, expected_text)


@mcp.tool()
def tracked_edit_paragraph_tool(
    path: str, index: int, new_text: str, author: str = "Claude",
    expected_text: str = None
) -> str:
    """
    Replace paragraph text creating tracked Deletion + Insertion revisions.

    REQUIRES: Tracked changes must be enabled first (call enable_tracked_changes).

    This is different from edit_paragraph: this version uses COM automation to
    replace text so Word records it as tracked changes. The user will see the old
    text as a deletion (strikethrough) and new text as an insertion (colored/underlined).

    INDEX TRANSLATION: Internally translates python-docx body paragraph indexes to
    COM paragraph indexes, so documents with tables are handled correctly. The index
    you pass should match what read_document_tool reports. This fixes the bug where
    edits after a table would land on a table cell paragraph instead of the intended
    body paragraph.

    Args:
        path: Path or key of open document
        index: Zero-based paragraph index to edit (use the index from read_document_tool)
        new_text: New text content to replace existing text
        author: Author name for these tracked changes (default: "Claude")
        expected_text: Optional content verification string. If provided, the target
                       paragraph must contain this text (case-sensitive partial match)
                       before the edit proceeds. If it does not match, the edit is
                       refused with an error message including the index, expected text,
                       and actual text preview (first ~80 chars). Use this to prevent
                       silent data corruption when paragraph indexes may have shifted.

    Returns:
        Success message or error message prefixed with "Error:"

    Examples:
        Edit paragraph with tracking:
        >>> tracked_edit_paragraph_tool("C:/Documents/report.docx", 0, "Updated text")
        "Edited tracked paragraph 0. Was: 'Original text' -> Now: 'Updated text'. Changes tracked as revisions by 'Claude'."

        Edit with content verification (safe editing):
        >>> tracked_edit_paragraph_tool("C:/Documents/report.docx", 40, "New section 4 text", expected_text="Section 4")
        "Edited tracked paragraph 40. Was: 'Section 4 of Task Order...' -> Now: 'New section 4 text'. Changes tracked as revisions by 'Claude'."

        Error - content mismatch (table cell was targeted instead of body paragraph):
        >>> tracked_edit_paragraph_tool("C:/Documents/report.docx", 40, "New text", expected_text="Section 4")
        "Error: Content verification failed for paragraph 40. Expected text containing 'Section 4' but found: 'DELIVERABLE'. The paragraph may have shifted -- re-read the document."

        Error - tracking not enabled:
        >>> tracked_edit_paragraph_tool("C:/Documents/report.docx", 0, "Text")
        "Error: Tracked changes are not enabled. Call enable_tracked_changes first."

    Design notes:
        - Zero-based indexing: Same as edit_paragraph and read_document for consistency
        - Requires tracking enabled: Returns error if TrackRevisions=False
        - Bridge pattern: Uses COM Range.Text replacement, creates Deletion + Insertion
        - Author attribution: Sets UserName in Word before editing
        - Index translation: COM paragraph indexes adjusted to skip table cell paragraphs
        - expected_text guard: Recommended for all edits in documents with tables
    """
    return tracked_edit_paragraph(path, index, new_text, author, expected_text)


@mcp.tool()
def tracked_delete_paragraph_tool(
    path: str, index: int, author: str = "Claude", expected_text: str = None
) -> str:
    """
    Delete a paragraph creating a tracked Deletion revision (strikethrough in Word).

    REQUIRES: Tracked changes must be enabled first (call enable_tracked_changes).

    This is different from delete_paragraph: this version uses COM automation to
    delete so Word records it as a tracked change. The user will see the deleted
    text as a strikethrough and can accept or reject the deletion.

    INDEX TRANSLATION: Internally translates python-docx body paragraph indexes to
    COM paragraph indexes, so documents with tables are handled correctly. The index
    you pass should match what read_document_tool reports.

    INDEX SHIFT WARNING: After deletion, remaining paragraphs shift down. Same
    warning as delete_paragraph - re-read document before additional operations.

    Args:
        path: Path or key of open document
        index: Zero-based paragraph index to delete (use the index from read_document_tool)
        author: Author name for this tracked change (default: "Claude")
        expected_text: Optional content verification string. If provided, the target
                       paragraph must contain this text (case-sensitive partial match)
                       before the delete proceeds. If it does not match, the delete is
                       refused with an error message including the index, expected text,
                       and actual text preview (first ~80 chars). Use this to prevent
                       accidentally deleting the wrong paragraph when indexes may have shifted.

    Returns:
        Success message or error message prefixed with "Error:"

    Examples:
        Delete paragraph with tracking:
        >>> tracked_delete_paragraph_tool("C:/Documents/report.docx", 2)
        "Deleted tracked paragraph 2 ('Old text'). Deletion tracked as revision by 'Claude'. Remaining paragraphs have shifted -- re-read document to get updated indexes."

        Delete with content verification (safe deletion):
        >>> tracked_delete_paragraph_tool("C:/Documents/report.docx", 5, expected_text="Obsolete section header")
        "Deleted tracked paragraph 5 ('Obsolete section header'). Deletion tracked as revision by 'Claude'. Remaining paragraphs have shifted -- re-read document to get updated indexes."

        Error - content mismatch (wrong paragraph would be deleted):
        >>> tracked_delete_paragraph_tool("C:/Documents/report.docx", 5, expected_text="Obsolete section header")
        "Error: Content verification failed for paragraph 5. Expected text containing 'Obsolete section header' but found: 'Introduction text here'. The paragraph may have shifted -- re-read the document."

        Error - tracking not enabled:
        >>> tracked_delete_paragraph_tool("C:/Documents/report.docx", 2)
        "Error: Tracked changes are not enabled. Call enable_tracked_changes first."

    Design notes:
        - Zero-based indexing: Same as delete_paragraph and read_document for consistency
        - Requires tracking enabled: Returns error if TrackRevisions=False
        - Bridge pattern: Uses COM Range.Delete(), creates Deletion revision
        - Index shift warning: Same behavior as Phase 1's delete_paragraph
        - Author attribution: Sets UserName in Word before deleting
        - Index translation: COM paragraph indexes adjusted to skip table cell paragraphs
        - expected_text guard: Recommended for all deletions in documents with tables
    """
    return tracked_delete_paragraph(path, index, author, expected_text)


@mcp.tool()
def tracked_edit_table_cell_tool(
    path: str, table_index: int, row_index: int, col_index: int,
    new_text: str, author: str = "Claude"
) -> str:
    """
    Edit a table cell creating tracked Deletion + Insertion revisions in Word.

    REQUIRES: Tracked changes must be enabled first (call enable_tracked_changes).

    This fills the gap between edit_table_cell_tool (untracked) and the tracked
    paragraph tools. It uses COM automation so Word records the cell edit as a
    tracked change -- the old text appears as a deletion (strikethrough) and the
    new text as an insertion (colored/underlined). The user can accept or reject
    the change in Word.

    REQUIRES COM AUTOMATION: Document must be saved to disk first (use save_document
    or save_document_as before calling this tool).

    Args:
        path: Path or key of open document
        table_index: Zero-based table index in the document
        row_index: Zero-based row index within the table
        col_index: Zero-based column index within the table
        new_text: New text content to replace existing cell text
        author: Author name for the tracked changes (default: "Claude")

    Returns:
        Success message with before/after preview, or error message prefixed with "Error:"

    Examples:
        Edit table cell with tracking:
        >>> tracked_edit_table_cell_tool("C:/Documents/report.docx", 0, 1, 0, "Alice")
        "Edited tracked table 0, cell (1, 0). Was: 'John' -> Now: 'Alice'. Changes tracked as revisions by 'Claude'."

        Edit with different author:
        >>> tracked_edit_table_cell_tool("C:/Documents/report.docx", 0, 0, 2, "March 31, 2027", "Juan Ocampo")
        "Edited tracked table 0, cell (0, 2). Was: 'February 23, 2026' -> Now: 'March 31, 2027'. Changes tracked as revisions by 'Juan Ocampo'."

        Error - tracking not enabled:
        >>> tracked_edit_table_cell_tool("C:/Documents/report.docx", 0, 0, 0, "text")
        "Error: Tracked changes are not enabled on this document. Call enable_tracked_changes first."

        Error - document not saved:
        >>> tracked_edit_table_cell_tool("Untitled-1", 0, 0, 0, "text")
        "Error: Document must be saved to disk before tracked editing. Use save_document_as first."

        Error - invalid table index:
        >>> tracked_edit_table_cell_tool("C:/Documents/report.docx", 5, 0, 0, "text")
        "Error: Invalid table index 5. Document has 2 table(s) (valid range: 0-1)."

    Design notes:
        - Requires COM automation: Document must be saved to disk
        - Requires tracked changes enabled: Returns error if TrackRevisions=False
        - Bridge pattern: Uses COM for tracked cell edit, then reloads python-docx
        - Zero-based indexing: All table/row/col indexes are 0-based
        - Cell end marker: COM cell ranges end with \\r\\x07; range is trimmed before text replacement
        - Author attribution: Sets UserName in Word before editing
        - Completes tracked workflow: Use alongside tracked_edit_paragraph_tool for documents with tables
    """
    return tracked_edit_table_cell(path, table_index, row_index, col_index, new_text, author)


# Register formatting tools (Phase 3)
@mcp.tool()
def format_text_tool(
    path: str,
    paragraph_index: int,
    bold: bool = None,
    italic: bool = None,
    underline: bool = None,
    font_name: str = None,
    font_size: float = None,
    font_color: str = None,
    run_index: int = None
) -> str:
    """
    Apply text formatting to runs in a paragraph.

    CRITICAL: Works at Run level, preserving existing formatting. Does NOT use
    paragraph.text which would destroy formatting. This is run-level formatting
    to comply with FMT-04 (preserve existing formatting on other runs).

    Args:
        path: Document path or key
        paragraph_index: 0-based paragraph index
        bold: Optional bool (True/False/None). None means "don't change".
        italic: Optional bool (True/False/None). None means "don't change".
        underline: Optional bool (True/False/None). None means "don't change".
        font_name: Optional str (e.g., "Arial", "Times New Roman"). None means "don't change".
        font_size: Optional float in points (e.g., 12.0 for 12pt). None means "don't change".
        font_color: Optional str in hex format "#RRGGBB" (e.g., "#FF0000" for red). None means "don't change".
        run_index: Optional int (0-based). If None, applies to ALL runs in paragraph.
                   If specified, applies only to that run.

    Returns:
        Success message listing what was changed, or error message

    Examples:
        Apply bold to all runs in paragraph 0:
        >>> format_text_tool("report.docx", 0, bold=True)
        "Applied formatting to paragraph 0: bold=True (applied to all runs)"

        Format specific run:
        >>> format_text_tool("report.docx", 1, font_name="Arial", font_size=12.0, run_index=0)
        "Applied formatting to paragraph 1, run 0: font_name='Arial', font_size=12.0pt"

        Change text color:
        >>> format_text_tool("report.docx", 2, font_color="#FF0000")
        "Applied formatting to paragraph 2: font_color='#FF0000' (applied to all runs)"

    Design notes:
        - Run-level formatting: Operates on Run.font properties only
        - Preserves other runs: When run_index specified, other runs unchanged
        - Tri-state booleans: None means "don't change", preserves python-docx tri-state
        - Color format: Must be "#RRGGBB" hex format
    """
    return format_text(path, paragraph_index, bold, italic, underline, font_name, font_size, font_color, run_index)


@mcp.tool()
def get_paragraph_formatting_tool(path: str, paragraph_index: int) -> str:
    """
    Get detailed formatting information for all runs in a paragraph.

    Returns per-run formatting details including bold, italic, underline, font name,
    size, and color. Useful for inspecting existing formatting before modification.

    Args:
        path: Document path or key
        paragraph_index: 0-based paragraph index

    Returns:
        Formatted string showing run-by-run formatting details, or error message

    Examples:
        >>> get_paragraph_formatting_tool("report.docx", 0)
        '''Paragraph 0 formatting (3 runs):
        Run 0: "Hello " - bold=True, italic=False, underline=False, font='Arial', size=12.0pt, color=None
        Run 1: "world" - bold=True, italic=True, underline=False, font='Arial', size=12.0pt, color='#FF0000'
        Run 2: "!" - bold=False, italic=False, underline=False, font='Arial', size=12.0pt, color=None
        '''

    Design notes:
        - Read-only operation: Doesn't modify document
        - Per-run breakdown: Shows formatting for each run separately
        - Tri-state values: Shows None for unset properties
    """
    return get_paragraph_formatting(path, paragraph_index)


# Register comment tools (Phase 3)
@mcp.tool()
def get_comments_tool(path: str) -> str:
    """
    Read all comments in the document with metadata.

    Returns comment information including id, author, date, text, and initials.

    LIMITATION: python-docx 1.2.0 does not expose comment location/range information.
    This is a known API limitation. Comments are returned in document order but
    without position data.

    Args:
        path: Document path or key

    Returns:
        Formatted list of comments with metadata, or error message

    Examples:
        >>> get_comments_tool("report.docx")
        '''Comments in 'report.docx': 3 comment(s)

        [1] By 'John Smith' (JS) on 2026-02-13 14:30:00
            "This needs revision"

        [2] By 'Jane Doe' (JD) on 2026-02-13 15:00:00
            "Approved"

        [3] By 'Claude' (C) on 2026-02-13 15:30:00
            "Updated per feedback"
        '''

        >>> get_comments_tool("no-comments.docx")
        "No comments found in 'no-comments.docx'."

    Design notes:
        - Read-only: Does not modify comments or document
        - No location data: python-docx API limitation (may require COM in future)
        - 1-based comment indexing: Matches Word's comment numbering
    """
    return get_comments(path)


# Register table tools (Phase 3)
@mcp.tool()
def create_table_tool(
    path: str,
    rows: int,
    cols: int,
    data: list = None,
    style: str = None
) -> str:
    """
    Create a new table in the document.

    The table is appended at the end of the document. Optionally populate with
    data and apply a table style.

    Args:
        path: Document path or key
        rows: Number of rows (must be > 0)
        cols: Number of columns (must be > 0)
        data: Optional 2D list to populate table. If provided, len(data) must equal
              rows and each inner list must have cols elements.
        style: Optional table style name (e.g., "Table Grid", "Light Shading")

    Returns:
        Success message with table index, or error message

    Examples:
        Create empty table:
        >>> create_table_tool("report.docx", 3, 2)
        "Created table 0 (3 rows x 2 cols) at end of document. Document now has 1 table(s)."

        Create table with data:
        >>> create_table_tool("report.docx", 2, 2, data=[["A", "B"], ["C", "D"]])
        "Created table 0 (2 rows x 2 cols) with data at end of document. Document now has 1 table(s)."

        Create styled table:
        >>> create_table_tool("report.docx", 3, 2, style="Table Grid")
        "Created table 0 (3 rows x 2 cols) with style 'Table Grid' at end of document. Document now has 1 table(s)."

    Design notes:
        - Zero-based indexing: Table index returned is 0-based
        - Appends to end: Table added after all existing content
        - Data validation: Ensures data dimensions match rows x cols
        - Style validation: Returns available styles if requested style not found
    """
    return create_table(path, rows, cols, data, style)


@mcp.tool()
def list_tables_tool(path: str) -> str:
    """
    List all tables in the document with dimensions and preview.

    Returns a summary of all tables showing index, dimensions, and first cell text.

    Args:
        path: Document path or key

    Returns:
        Formatted list of tables, or error message

    Examples:
        >>> list_tables_tool("report.docx")
        '''Tables in 'report.docx': 3 table(s)

        Table 0: 3 rows x 2 cols
          First cell: "Name"

        Table 1: 5 rows x 4 cols
          First cell: "Q1"

        Table 2: 2 rows x 3 cols
          First cell: "Total"
        '''

        >>> list_tables_tool("no-tables.docx")
        "No tables found in 'no-tables.docx'."

    Design notes:
        - Zero-based indexing: Table indexes shown are 0-based
        - Read-only: Does not modify document
        - Preview: Shows first cell text for quick identification
    """
    return list_tables(path)


@mcp.tool()
def read_table_tool(
    path: str,
    table_index: int,
    start_row: int = None,
    end_row: int = None
) -> str:
    """
    Read table content as a formatted grid.

    Returns table data in a grid format with row labels. Supports optional row
    range for large tables.

    Args:
        path: Document path or key
        table_index: Zero-based table index in document
        start_row: Optional starting row (0-based, inclusive)
        end_row: Optional ending row (0-based, inclusive)

    Returns:
        Formatted table grid, or error message

    Examples:
        Read full table:
        >>> read_table_tool("report.docx", 0)
        '''Table 0 (3 rows x 2 cols):
        Row 0: | Name | Age |
        Row 1: | John | 30  |
        Row 2: | Jane | 25  |
        '''

        Read row range:
        >>> read_table_tool("report.docx", 0, start_row=1, end_row=2)
        '''Table 0 (3 rows x 2 cols) - showing rows 1-2:
        Row 1: | John | 30  |
        Row 2: | Jane | 25  |
        '''

    Design notes:
        - Zero-based indexing: All indexes are 0-based
        - Formatted output: Pipe-separated grid for readability
        - Cell truncation: Long cell text truncated to 40 chars
        - Merged cells: Inaccessible cells show "?" placeholder
    """
    return read_table(path, table_index, start_row, end_row)


@mcp.tool()
def edit_table_cell_tool(
    path: str,
    table_index: int,
    row_index: int,
    col_index: int,
    text: str
) -> str:
    """
    Edit (replace) the text content of a table cell.

    FORMATTING LOSS: Uses cell.text assignment which loses any run-level formatting
    within the cell. This is acceptable for table cells which typically don't have
    complex run formatting.

    Args:
        path: Document path or key
        table_index: Zero-based table index
        row_index: Zero-based row index
        col_index: Zero-based column index
        text: New text content for the cell

    Returns:
        Success message with before/after preview, or error message

    Examples:
        >>> edit_table_cell_tool("report.docx", 0, 1, 0, "Alice")
        "Edited table 0, cell (1, 0). Was: 'John' -> Now: 'Alice'"

    Design notes:
        - Zero-based indexing: All indexes are 0-based
        - Text replacement: Uses cell.text (acceptable formatting loss)
        - Index validation: Checks all indexes before modification
    """
    return edit_table_cell(path, table_index, row_index, col_index, text)


@mcp.tool()
def add_table_row_tool(
    path: str,
    table_index: int,
    data: list = None
) -> str:
    """
    Add a row to the end of a table.

    Appends a new row to the table. Optionally populate with data.

    Args:
        path: Document path or key
        table_index: Zero-based table index
        data: Optional list of cell values. Length must match table column count.

    Returns:
        Success message with updated row count, or error message

    Examples:
        Add empty row:
        >>> add_table_row_tool("report.docx", 0)
        "Added row to table 0. Table now has 4 rows."

        Add row with data:
        >>> add_table_row_tool("report.docx", 0, data=["Bob", "35"])
        "Added row to table 0 with data: ['Bob', '35']. Table now has 5 rows."

    Design notes:
        - Zero-based indexing: Table index is 0-based
        - Appends to end: New row added after existing rows
        - Data validation: Ensures data length matches column count
    """
    return add_table_row(path, table_index, data)


@mcp.tool()
def add_table_column_tool(
    path: str,
    table_index: int,
    width: float = 1.0,
    data: list = None
) -> str:
    """
    Add a column to the end of a table.

    Appends a new column to the table. Optionally populate with data.

    Args:
        path: Document path or key
        table_index: Zero-based table index
        width: Column width in inches (default: 1.0)
        data: Optional list of cell values. Length must match table row count.

    Returns:
        Success message with updated column count, or error message

    Examples:
        Add column with default width:
        >>> add_table_column_tool("report.docx", 0)
        "Added column (width=1.0in) to table 0. Table now has 3 columns."

        Add column with data:
        >>> add_table_column_tool("report.docx", 0, width=1.5, data=["City", "NYC", "LA"])
        "Added column (width=1.5in) to table 0 with data: ['City', 'NYC', 'LA']. Table now has 4 columns."

    Design notes:
        - Zero-based indexing: Table index is 0-based
        - Default width: 1 inch is reasonable default (table.add_column requires width)
        - Appends to end: New column added after existing columns
        - Data validation: Ensures data length matches row count
    """
    return add_table_column(path, table_index, width, data)


@mcp.tool()
def delete_table_row_tool(path: str, table_index: int, row_index: int) -> str:
    """
    Delete a row from a table using COM automation.

    REQUIRES COM AUTOMATION: Document must be saved to disk first (use save_document
    or save_document_as before calling this tool). python-docx does not support
    row deletion, so COM is required.

    Args:
        path: Path or key of open document
        table_index: Zero-based table index
        row_index: Zero-based row index to delete

    Returns:
        Success message with updated row count, or error message

    Examples:
        >>> delete_table_row_tool("C:/Documents/report.docx", 0, 2)
        "Deleted row 2 from table 0. Table now has 4 rows."

        >>> delete_table_row_tool("Untitled-1", 0, 0)
        "Error: Document must be saved to disk before COM operations. Use save_document_as first."

    Design notes:
        - Requires COM automation: Document must be saved to disk
        - Zero-based indexing: All indexes are 0-based
        - Bridge pattern: Uses COM to delete, then reloads python-docx
        - Index validation: Checks table and row exist before deleting
    """
    return delete_table_row(path, table_index, row_index)


@mcp.tool()
def delete_table_column_tool(path: str, table_index: int, col_index: int) -> str:
    """
    Delete a column from a table using COM automation.

    REQUIRES COM AUTOMATION: Document must be saved to disk first (use save_document
    or save_document_as before calling this tool). python-docx does not support
    column deletion, so COM is required.

    Args:
        path: Path or key of open document
        table_index: Zero-based table index
        col_index: Zero-based column index to delete

    Returns:
        Success message with updated column count, or error message

    Examples:
        >>> delete_table_column_tool("C:/Documents/report.docx", 0, 1)
        "Deleted column 1 from table 0. Table now has 3 columns."

        >>> delete_table_column_tool("Untitled-1", 0, 0)
        "Error: Document must be saved to disk before COM operations. Use save_document_as first."

    Design notes:
        - Requires COM automation: Document must be saved to disk
        - Zero-based indexing: All indexes are 0-based
        - Bridge pattern: Uses COM to delete, then reloads python-docx
        - Index validation: Checks table and column exist before deleting
    """
    return delete_table_column(path, table_index, col_index)


# Register image tools (Phase 3)
@mcp.tool()
def insert_image_tool(
    path: str,
    image_path: str,
    width: float = None,
    height: float = None
) -> str:
    """
    Insert an inline image into the document.

    The image is inserted at the end of the document as an InlineShape. Optional
    width and height allow resizing on insert (aspect ratio is preserved if only
    one dimension specified).

    Args:
        path: Document path or key
        image_path: Path to image file (PNG, JPG, etc.)
        width: Optional width in inches
        height: Optional height in inches

    Returns:
        Success message with image index, or error message

    Examples:
        Insert image with original size:
        >>> insert_image_tool("report.docx", "C:/Images/chart.png")
        "Inserted image from 'chart.png' at index 0. Document now has 1 inline image(s)."

        Insert with width (preserves aspect ratio):
        >>> insert_image_tool("report.docx", "C:/Images/logo.png", width=2.0)
        "Inserted image from 'logo.png' at index 1 (width=2.0in). Document now has 2 inline image(s)."

        Insert with both dimensions:
        >>> insert_image_tool("report.docx", "C:/Images/photo.jpg", width=3.0, height=2.0)
        "Inserted image from 'photo.jpg' at index 2 (width=3.0in, height=2.0in). Document now has 3 inline image(s)."

    Design notes:
        - Inline images: Added as InlineShape (flows with text)
        - Aspect ratio: Preserved when only one dimension specified
        - File validation: Returns error if image file not found
        - Zero-based indexing: Image index is 0-based
    """
    return insert_image(path, image_path, width, height)


@mcp.tool()
def resize_image_tool(
    path: str,
    image_index: int,
    width: float = None,
    height: float = None,
    preserve_aspect_ratio: bool = True
) -> str:
    """
    Resize an existing inline image.

    By default, aspect ratio is preserved when only one dimension is specified.
    Pass preserve_aspect_ratio=False to change only the supplied dimension and
    leave the other unchanged (the old behavior).

    Args:
        path: Document path or key
        image_index: Zero-based index into inline images collection
        width: Optional new width in inches
        height: Optional new height in inches
        preserve_aspect_ratio: If True (default), auto-computes the missing dimension
                               to keep the original aspect ratio. If False, only the
                               provided dimension(s) are applied.

    Returns:
        Success message with new dimensions, or error message

    Examples:
        Width only -- height auto-computed to preserve aspect ratio (default):
        >>> resize_image_tool("report.docx", 0, width=4.0)
        "Resized image 0 to width=4.00in, height=2.67in (aspect ratio preserved)."

        Height only -- width auto-computed:
        >>> resize_image_tool("report.docx", 0, height=3.0)
        "Resized image 0 to width=4.50in, height=3.00in (aspect ratio preserved)."

        Both dimensions -- applied as-is:
        >>> resize_image_tool("report.docx", 0, width=3.0, height=2.0)
        "Resized image 0 to width=3.00in, height=2.00in."

        Width only, no aspect ratio preservation:
        >>> resize_image_tool("report.docx", 0, width=5.0, preserve_aspect_ratio=False)
        "Resized image 0 to width=5.00in, height=2.00in."

    Design notes:
        - Zero-based indexing: Image index is 0-based
        - Aspect ratio default: preserve_aspect_ratio=True prevents unintentional distortion
        - Inline images only: Works on InlineShape collection
        - At least one dimension required: Must specify width, height, or both
    """
    return resize_image(path, image_index, width, height, preserve_aspect_ratio)


@mcp.tool()
def list_images_tool(path: str) -> str:
    """
    List all inline images in the document.

    Returns summary of all inline images showing index and dimensions.

    Args:
        path: Document path or key

    Returns:
        Formatted list of images, or error message

    Examples:
        >>> list_images_tool("report.docx")
        '''Inline images in 'report.docx': 3 image(s)

        Image 0: 3.00in x 2.00in
        Image 1: 2.50in x 2.50in
        Image 2: 4.00in x 3.00in
        '''

        >>> list_images_tool("no-images.docx")
        "No inline images found in 'no-images.docx'."

    Design notes:
        - Inline images only: Shows InlineShape collection
        - Zero-based indexing: Image indexes are 0-based
        - Dimensions: Shows width x height in inches
        - Read-only: Does not modify document
    """
    return list_images(path)


@mcp.tool()
def reposition_image_tool(
    path: str,
    image_index: int,
    left: float = None,
    top: float = None,
    width: float = None,
    height: float = None
) -> str:
    """
    Reposition an inline image to an absolute position on the page.

    WARNING: Converts inline image to floating shape (one-way). Image will no longer
    appear in inline_shapes collection after this operation. This is a permanent
    conversion from InlineShape to Shape.

    REQUIRES COM AUTOMATION: Document must be saved to disk first (use save_document
    or save_document_as before calling this tool). python-docx does not support
    absolute positioning, so COM is required.

    Args:
        path: Path or key of open document
        image_index: Zero-based index into inline images collection
        left: Optional horizontal position in inches (from left edge of page)
        top: Optional vertical position in inches (from top edge of page)
        width: Optional width in inches (for resizing during reposition)
        height: Optional height in inches (for resizing during reposition)

    Returns:
        Success message with position details, or error message

    Examples:
        Position image:
        >>> reposition_image_tool("C:/Documents/report.docx", 0, left=1.0, top=2.0)
        "Repositioned image 0 to absolute position (left=1.0in, top=2.0in). Image converted from inline to floating shape."

        Position and resize:
        >>> reposition_image_tool("C:/Documents/report.docx", 0, left=1.5, top=1.5, width=3.0, height=2.0)
        "Repositioned image 0 to absolute position (left=1.5in, top=1.5in). Resized to 3.0in x 2.0in. Image converted from inline to floating shape."

    Design notes:
        - Requires COM automation: Document must be saved to disk
        - One-way conversion: InlineShape becomes Shape (cannot revert)
        - Zero-based indexing: Image index is 0-based
        - Bridge pattern: Uses COM to reposition, then reloads python-docx
        - Coordinate units: Input in inches, converted to points internally
    """
    return reposition_image(path, image_index, left, top, width, height)


# Register section and header/footer tools (Phase 4)
@mcp.tool()
def list_sections_tool(path: str) -> str:
    """
    List all sections in the document with their properties.

    Returns detailed information about each section including break type,
    orientation, page dimensions, margins, and header/footer link status.

    Args:
        path: Document path or key

    Returns:
        Formatted list of sections with properties, or error message

    Examples:
        >>> list_sections_tool("report.docx")
        '''Sections in 'report.docx': 2 section(s)

        Section 0:
          Break type: new_page
          Orientation: portrait
          Page dimensions: 8.5in x 11.0in
          Margins: top=1.0in, bottom=1.0in, left=1.0in, right=1.0in
          Header distance: 0.5in
          Footer distance: 0.5in
          Header linked to previous: False
          Footer linked to previous: False

        Section 1:
          Break type: new_page
          Orientation: landscape
          Page dimensions: 11.0in x 8.5in
          Margins: top=1.0in, bottom=1.0in, left=1.0in, right=1.0in
          Header distance: 0.5in
          Footer distance: 0.5in
          Header linked to previous: True
          Footer linked to previous: True
        '''

        >>> list_sections_tool("single-section.docx")
        '''Sections in 'single-section.docx': 1 section(s)

        Section 0:
          Break type: new_page
          Orientation: portrait
          Page dimensions: 8.5in x 11.0in
          ...
        '''

    Design notes:
        - Zero-based indexing: Section indexes are 0-based
        - All sections: Every document has at least one section
        - Link status: Shows whether headers/footers inherit from previous section
        - Dimensions in inches: All measurements shown in inches (converted from internal units)
    """
    return list_sections(path)


@mcp.tool()
def add_section_tool(path: str, break_type: str = "new_page") -> str:
    """
    Add a new section to the document.

    Creates a new section with a specified break type. The new section is added
    at the current end of the document (after all existing content).

    Args:
        path: Document path or key
        break_type: Type of section break (default: "new_page")
                    Valid values:
                    - "new_page" (default): Starts new section on next page
                    - "continuous": Starts new section on same page
                    - "even_page": Starts new section on next even-numbered page
                    - "odd_page": Starts new section on next odd-numbered page
                    - "new_column": Starts new section at next column (multi-column layouts)

    Returns:
        Success message with section count, or error message

    Examples:
        Add section with default break type:
        >>> add_section_tool("report.docx")
        "Added section with break type 'new_page'. Document now has 2 section(s)."

        Add continuous section:
        >>> add_section_tool("report.docx", "continuous")
        "Added section with break type 'continuous'. Document now has 3 section(s)."

        Error - invalid break type:
        >>> add_section_tool("report.docx", "invalid")
        "Error: Invalid break_type 'invalid'. Valid values: new_page, continuous, even_page, odd_page, new_column."

    Design notes:
        - Appends to end: New section added after all existing content
        - Inherits settings: New section initially inherits properties from previous section
        - Break type validation: Returns clear error for invalid break types
        - Use with modify_section_properties: Create section first, then customize properties
    """
    return add_section(path, break_type)


@mcp.tool()
def modify_section_properties_tool(
    path: str,
    section_index: int,
    orientation: str = None,
    page_width: float = None,
    page_height: float = None,
    top_margin: float = None,
    bottom_margin: float = None,
    left_margin: float = None,
    right_margin: float = None,
    header_distance: float = None,
    footer_distance: float = None
) -> str:
    """
    Modify properties of an existing section.

    Allows changing page orientation, dimensions, margins, and header/footer
    distances. Auto-swaps page dimensions when orientation changes (unless
    explicit dimensions provided).

    Args:
        path: Document path or key
        section_index: Zero-based section index
        orientation: Optional "portrait" or "landscape"
        page_width: Optional page width in inches
        page_height: Optional page height in inches
        top_margin: Optional top margin in inches
        bottom_margin: Optional bottom margin in inches
        left_margin: Optional left margin in inches
        right_margin: Optional right margin in inches
        header_distance: Optional distance from top of page to header in inches
        footer_distance: Optional distance from bottom of page to footer in inches

    Returns:
        Success message listing changes, or error message

    Examples:
        Change orientation to landscape (auto-swaps dimensions):
        >>> modify_section_properties_tool("report.docx", 0, orientation="landscape")
        "Modified section 0: orientation='landscape', page_width=11.0in, page_height=8.5in (auto-swapped)."

        Change margins:
        >>> modify_section_properties_tool("report.docx", 0, top_margin=1.5, bottom_margin=1.5)
        "Modified section 0: top_margin=1.5in, bottom_margin=1.5in."

        Custom page size:
        >>> modify_section_properties_tool("report.docx", 1, page_width=8.0, page_height=10.0)
        "Modified section 1: page_width=8.0in, page_height=10.0in."

        Error - invalid section index:
        >>> modify_section_properties_tool("report.docx", 5, orientation="landscape")
        "Error: Invalid section_index 5. Document has 2 section(s) (valid range: 0-1)."

        Error - invalid orientation:
        >>> modify_section_properties_tool("report.docx", 0, orientation="sideways")
        "Error: Invalid orientation 'sideways'. Valid values: portrait, landscape."

    Design notes:
        - Zero-based indexing: Section index is 0-based
        - Auto-dimension swap: When orientation changes without explicit width/height,
          automatically swaps dimensions to prevent layout corruption
        - All units in inches: Input values are in inches (converted to internal units)
        - Partial updates: Can change any subset of properties (None means "don't change")
        - Header/footer distance: Controls space between page edge and header/footer content
    """
    return modify_section_properties(
        path,
        section_index,
        orientation,
        page_width,
        page_height,
        top_margin,
        bottom_margin,
        left_margin,
        right_margin,
        header_distance,
        footer_distance
    )


@mcp.tool()
def get_header_tool(path: str, section_index: int = 0, header_type: str = "primary") -> str:
    """
    Read header content from a section.

    Retrieves the text content from the specified header type in the given section.
    Also reports whether the header is linked to the previous section.

    Args:
        path: Document path or key
        section_index: Zero-based section index (default: 0)
        header_type: Type of header to read (default: "primary")
                     Valid values:
                     - "primary": Regular header (appears on all pages except first/even if configured)
                     - "first_page": Header for first page of section (requires different_first_page_header_footer enabled)
                     - "even_page": Header for even-numbered pages (requires even/odd page settings)

    Returns:
        Header content and link status, or error message

    Examples:
        Read primary header:
        >>> get_header_tool("report.docx", 0)
        "Header (primary) in section 0 (linked: False):\nAnnual Report 2026"

        Read first page header:
        >>> get_header_tool("report.docx", 0, "first_page")
        "Header (first_page) in section 0 (linked: False, first_page_enabled: True):\nCOVER PAGE"

        Empty header:
        >>> get_header_tool("report.docx", 1)
        "Header (primary) in section 1 (linked: True):\n(empty)"

        Error - invalid section:
        >>> get_header_tool("report.docx", 10)
        "Error: Invalid section_index 10. Document has 2 section(s) (valid range: 0-1)."

    Design notes:
        - Zero-based indexing: Section index is 0-based
        - Link status: Reports if header inherits from previous section
        - First page content: Only visible if different_first_page_header_footer is True
        - Even page headers: Require document-level even/odd page setting (rarely used)
    """
    return get_header(path, section_index, header_type)


@mcp.tool()
def set_header_tool(path: str, text: str, section_index: int = 0, header_type: str = "primary") -> str:
    """
    Set header content for a section.

    Writes text to the specified header type in the given section. Automatically
    unlinks the header from the previous section and enables different_first_page_header_footer
    if setting a first_page header.

    Args:
        path: Document path or key
        text: Header text content
        section_index: Zero-based section index (default: 0)
        header_type: Type of header to set (default: "primary")
                     Valid values:
                     - "primary": Regular header (appears on all pages except first/even if configured)
                     - "first_page": Header for first page of section (auto-enables different_first_page_header_footer)
                     - "even_page": Header for even-numbered pages

    Returns:
        Success message with unlink/enablement status, or error message

    Examples:
        Set primary header:
        >>> set_header_tool("report.docx", "Company Confidential", 0)
        "Set header (primary) in section 0 (auto-unlinked from previous). Content: 'Company Confidential'"

        Set first page header:
        >>> set_header_tool("report.docx", "COVER PAGE", 0, "first_page")
        "Set header (first_page) in section 0 (auto-unlinked from previous, enabled different_first_page). Content: 'COVER PAGE'"

        Error - invalid section:
        >>> set_header_tool("report.docx", "Header", 10)
        "Error: Invalid section_index 10. Document has 2 section(s) (valid range: 0-1)."

    Design notes:
        - Auto-unlinking: Always sets is_linked_to_previous=False to prevent circular inheritance bugs
        - First page auto-enablement: Setting first_page header automatically enables different_first_page_header_footer
        - Uses existing paragraph: Modifies first paragraph instead of adding new ones (prevents formatting issues)
        - Zero-based indexing: Section index is 0-based
    """
    return set_header(path, text, section_index, header_type)


@mcp.tool()
def get_footer_tool(path: str, section_index: int = 0, footer_type: str = "primary") -> str:
    """
    Read footer content from a section.

    Retrieves the text content from the specified footer type in the given section.
    Also reports whether the footer is linked to the previous section.

    Args:
        path: Document path or key
        section_index: Zero-based section index (default: 0)
        footer_type: Type of footer to read (default: "primary")
                     Valid values:
                     - "primary": Regular footer (appears on all pages except first/even if configured)
                     - "first_page": Footer for first page of section (requires different_first_page_header_footer enabled)
                     - "even_page": Footer for even-numbered pages (requires even/odd page settings)

    Returns:
        Footer content and link status, or error message

    Examples:
        Read primary footer:
        >>> get_footer_tool("report.docx", 0)
        "Footer (primary) in section 0 (linked: False):\nPage 1"

        Read first page footer:
        >>> get_footer_tool("report.docx", 0, "first_page")
        "Footer (first_page) in section 0 (linked: False, first_page_enabled: True):\n(empty)"

        Empty footer:
        >>> get_footer_tool("report.docx", 1)
        "Footer (primary) in section 1 (linked: True):\n(empty)"

        Error - invalid section:
        >>> get_footer_tool("report.docx", 10)
        "Error: Invalid section_index 10. Document has 2 section(s) (valid range: 0-1)."

    Design notes:
        - Zero-based indexing: Section index is 0-based
        - Link status: Reports if footer inherits from previous section
        - First page content: Only visible if different_first_page_header_footer is True
        - Even page footers: Require document-level even/odd page setting (rarely used)
    """
    return get_footer(path, section_index, footer_type)


@mcp.tool()
def set_footer_tool(path: str, text: str, section_index: int = 0, footer_type: str = "primary") -> str:
    """
    Set footer content for a section.

    Writes text to the specified footer type in the given section. Automatically
    unlinks the footer from the previous section and enables different_first_page_header_footer
    if setting a first_page footer.

    Args:
        path: Document path or key
        text: Footer text content
        section_index: Zero-based section index (default: 0)
        footer_type: Type of footer to set (default: "primary")
                     Valid values:
                     - "primary": Regular footer (appears on all pages except first/even if configured)
                     - "first_page": Footer for first page of section (auto-enables different_first_page_header_footer)
                     - "even_page": Footer for even-numbered pages

    Returns:
        Success message with unlink/enablement status, or error message

    Examples:
        Set primary footer:
        >>> set_footer_tool("report.docx", "Page Footer - Confidential", 0)
        "Set footer (primary) in section 0 (auto-unlinked from previous). Content: 'Page Footer - Confidential'"

        Set first page footer:
        >>> set_footer_tool("report.docx", "Draft Version", 0, "first_page")
        "Set footer (first_page) in section 0 (auto-unlinked from previous, enabled different_first_page). Content: 'Draft Version'"

        Error - invalid section:
        >>> set_footer_tool("report.docx", "Footer", 10)
        "Error: Invalid section_index 10. Document has 2 section(s) (valid range: 0-1)."

    Design notes:
        - Auto-unlinking: Always sets is_linked_to_previous=False to prevent circular inheritance bugs
        - First page auto-enablement: Setting first_page footer automatically enables different_first_page_header_footer
        - Uses existing paragraph: Modifies first paragraph instead of adding new ones (prevents formatting issues)
        - Zero-based indexing: Section index is 0-based
    """
    return set_footer(path, text, section_index, footer_type)


# Register monitoring tools (Phase 5)
@mcp.tool()
def get_server_health_tool() -> str:
    """
    Get server health status and resource metrics.

    Returns production metrics including memory usage, COM pool status,
    and open document count. Use this to check if the server is operating
    normally before performing resource-intensive operations.

    Returns:
        Formatted health report with status (HEALTHY/DEGRADED/UNHEALTHY),
        memory metrics, COM pool metrics, and any active alerts.

    Examples:
        >>> get_server_health_tool()
        '''Server Health: HEALTHY

        Process Memory: 45.2 MB
        System Memory: 52.3%
        Open Documents: 2

        COM Pool:
          Active instances: 0
          Total created: 5
          Total failed: 0
          Pool size limit: 3
        '''

    Design notes:
        - Read-only: Does not modify server state
        - Status thresholds: Memory >80% = unhealthy, >70% = degraded
        - COM pool tracking: Shows lifetime operation counts
        - Alerts: Lists any active warnings (high memory, failed operations)
    """
    return get_server_health()


def main():
    """
    Main entry point for word-mcp server.

    Starts the FastMCP server and begins listening for MCP protocol messages.
    """
    mcp.run()


if __name__ == "__main__":
    main()
