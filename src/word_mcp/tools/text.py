"""Text editing tools for Word documents.

Provides paragraph CRUD operations (create, read, update, delete) for Word documents.
All functions use zero-based indexing for paragraph operations.
"""

from typing import Optional
from ..document_manager import document_manager
from ..logging_config import get_logger

logger = get_logger(__name__)


def read_document(
    path: str,
    start_index: Optional[int] = None,
    end_index: Optional[int] = None
) -> str:
    """Read document content as a list of indexed paragraphs.

    Returns all paragraphs with their indexes, styles, and text content.
    Supports optional pagination via start_index/end_index (inclusive range).

    Args:
        path: Document path or key
        start_index: Optional starting paragraph index (0-based, inclusive)
        end_index: Optional ending paragraph index (0-based, inclusive)

    Returns:
        Formatted string with paragraph list or error message

    Example output:
        Document: Untitled-1 | Paragraphs: 3 | Showing: 0-2
        [0] (Heading 1) Introduction
        [1] (Normal) This is the first paragraph of content...
        [2] (Normal) This is the second paragraph.
    """
    doc = document_manager.get_document(path)
    if doc is None:
        logger.error("document_not_open", tool="read_document", path=path)
        return f"Error: No document open at '{path}'. Use open_document or create_document first."

    paragraphs = doc.paragraphs
    total_count = len(paragraphs)

    # Check if document is empty (no paragraphs with content)
    if total_count == 0:
        return "Document is empty (0 paragraphs with content)"

    # Count non-empty paragraphs
    content_count = sum(1 for p in paragraphs if p.text.strip())
    if content_count == 0:
        return "Document is empty (0 paragraphs with content)"

    # Handle pagination
    if start_index is None:
        start_index = 0
    if end_index is None:
        end_index = total_count - 1

    # Validate range
    if start_index < 0 or start_index >= total_count:
        return f"Error: Invalid start_index {start_index}. Document has {total_count} paragraphs (valid range: 0-{total_count-1})."
    if end_index < 0 or end_index >= total_count:
        return f"Error: Invalid end_index {end_index}. Document has {total_count} paragraphs (valid range: 0-{total_count-1})."
    if start_index > end_index:
        return f"Error: start_index ({start_index}) cannot be greater than end_index ({end_index})."

    # Build header
    lines = [f"Document: {path} | Paragraphs: {total_count} | Showing: {start_index}-{end_index}"]

    # Build paragraph list
    for i in range(start_index, end_index + 1):
        para = paragraphs[i]
        text = para.text
        style_name = para.style.name if para.style else "Normal"

        # Text preview: full text up to 200 chars, truncated with char count if longer
        if len(text) > 200:
            text_preview = text[:200] + f"... ({len(text)} chars)"
        else:
            text_preview = text

        lines.append(f"[{i}] ({style_name}) {text_preview}")

    return "\n".join(lines)


def add_paragraph(
    path: str,
    text: str,
    position: Optional[int] = None,
    style: Optional[str] = None
) -> str:
    """Add a paragraph to the document.

    By default, appends to the end. Can insert at a specific position.

    Args:
        path: Document path or key
        text: Paragraph text content
        position: Optional 0-based index to insert at (None = append to end)
        style: Optional paragraph style name (e.g., "Normal", "Heading 1")

    Returns:
        Success message with index and updated count, or error message

    Examples:
        add_paragraph(key, "New paragraph")  # Appends to end
        add_paragraph(key, "Inserted", position=1)  # Inserts at index 1
        add_paragraph(key, "Title", style="Heading 1")  # Appends with style
    """
    doc = document_manager.get_document(path)
    if doc is None:
        logger.error("document_not_open", tool="add_paragraph", path=path)
        return f"Error: No document open at '{path}'. Use open_document or create_document first."

    para_count = len(doc.paragraphs)

    # Determine insertion behavior
    if position is None:
        # Append to end
        new_para = doc.add_paragraph(text, style=style)
        idx = para_count
    elif position == para_count:
        # Position equals count: append to end
        new_para = doc.add_paragraph(text, style=style)
        idx = para_count
    else:
        # Insert at specific position
        if position < 0 or position > para_count:
            return f"Error: Invalid paragraph position {position}. Document has {para_count} paragraphs (valid range: 0-{para_count})."

        new_para = doc.paragraphs[position].insert_paragraph_before(text)
        if style:
            new_para.style = style
        idx = position

    # Update count
    new_count = len(doc.paragraphs)

    # Preview: first 50 chars
    text_preview = text[:50] if len(text) <= 50 else text[:50] + "..."

    return f"Added paragraph at index {idx}: '{text_preview}'\nDocument now has {new_count} paragraphs."


def edit_paragraph(path: str, index: int, new_text: str) -> str:
    """Edit (replace) the text of an existing paragraph by index.

    Preserves the formatting of the first run (bold, italic, font size, color,
    underline, font name) on the replacement text. If the paragraph has no runs
    (empty paragraph), falls back to direct text assignment.

    Args:
        path: Document path or key
        index: 0-based paragraph index
        new_text: New text content

    Returns:
        Success message showing before/after text, or error message

    Example:
        edit_paragraph(key, 0, "Updated first paragraph")
    """
    doc = document_manager.get_document(path)
    if doc is None:
        logger.error("document_not_open", tool="edit_paragraph", path=path)
        return f"Error: No document open at '{path}'. Use open_document or create_document first."

    para_count = len(doc.paragraphs)

    # Validate index
    if index < 0 or index >= para_count:
        return f"Error: Invalid paragraph index {index}. Document has {para_count} paragraphs (valid range: 0-{para_count-1})."

    para = doc.paragraphs[index]
    old_text = para.text

    # Replace text at run level to preserve formatting
    runs = list(para.runs)
    if not runs:
        # No runs: empty paragraph -- fall back to direct assignment
        para.text = new_text
    else:
        # Set first run's text to full new text; clear remaining runs
        # This preserves the first run's font properties on the new text
        runs[0].text = new_text
        for run in runs[1:]:
            run.text = ""

    # Previews: first 50 chars each
    old_preview = old_text[:50] if len(old_text) <= 50 else old_text[:50] + "..."
    new_preview = new_text[:50] if len(new_text) <= 50 else new_text[:50] + "..."

    return f"Edited paragraph {index}. Was: '{old_preview}' -> Now: '{new_preview}'\nDocument has {para_count} paragraphs."


def delete_paragraph(path: str, index: int) -> str:
    """Delete a paragraph by index.

    WARNING: After deletion, remaining paragraphs shift indexes.
    Re-read the document to get updated indexes before further operations.

    Args:
        path: Document path or key
        index: 0-based paragraph index to delete

    Returns:
        Success message with shift warning, or error message

    Example:
        delete_paragraph(key, 2)  # Deletes paragraph at index 2
    """
    doc = document_manager.get_document(path)
    if doc is None:
        logger.error("document_not_open", tool="delete_paragraph", path=path)
        return f"Error: No document open at '{path}'. Use open_document or create_document first."

    para_count = len(doc.paragraphs)

    # Validate index
    if index < 0 or index >= para_count:
        return f"Error: Invalid paragraph index {index}. Document has {para_count} paragraphs (valid range: 0-{para_count-1})."

    para = doc.paragraphs[index]
    text = para.text

    # Text preview
    text_preview = text[:50] if len(text) <= 50 else text[:50] + "..."

    # Delete paragraph (python-docx has no native delete API)
    para._element.getparent().remove(para._element)

    # Update count
    new_count = len(doc.paragraphs)

    return f"Deleted paragraph {index} ('{text_preview}'). Remaining paragraphs have shifted -- re-read document to get updated indexes. Document now has {new_count} paragraphs."
