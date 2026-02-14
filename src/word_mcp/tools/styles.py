"""Style application tools for Word documents.

Provides paragraph style application including heading styles H1-H9.
"""

from docx.oxml.shared import OxmlElement
from ..document_manager import document_manager
from ..logging_config import get_logger

logger = get_logger(__name__)


def apply_heading_style(path: str, index: int, level: int) -> str:
    """Apply a heading style (H1-H9) to a paragraph.

    Args:
        path: Document path or key
        index: 0-based paragraph index
        level: Heading level 1-9 (maps to "Heading 1" through "Heading 9")

    Returns:
        Success message with paragraph preview, or error message

    Example:
        apply_heading_style(key, 0, 1)  # Applies "Heading 1" to paragraph 0
        apply_heading_style(key, 5, 3)  # Applies "Heading 3" to paragraph 5
    """
    doc = document_manager.get_document(path)
    if doc is None:
        return f"Error: No document open at '{path}'. Use open_document or create_document first."

    # Validate level
    if level < 1 or level > 9:
        return f"Error: Invalid heading level {level}. Valid range is 1-9 (for Heading 1 through Heading 9)."

    paragraphs = doc.paragraphs
    para_count = len(paragraphs)

    # Validate index
    if index < 0 or index >= para_count:
        return f"Error: Invalid paragraph index {index}. Document has {para_count} paragraphs (valid range: 0-{para_count-1})."

    para = paragraphs[index]
    text = para.text

    # Apply heading style
    style_name = f"Heading {level}"
    para.style = style_name

    # Text preview
    text_preview = text[:50] if len(text) <= 50 else text[:50] + "..."

    return f"Applied '{style_name}' style to paragraph {index}: '{text_preview}'"


def apply_style(path: str, index: int, style_name: str) -> str:
    """Apply a paragraph style by name.

    General-purpose style application for any paragraph style (Normal, Title, Heading 1, etc.).
    To see available styles in a document, use get_document_info().

    Args:
        path: Document path or key
        index: 0-based paragraph index
        style_name: Style name (e.g., "Normal", "Title", "Heading 1")

    Returns:
        Success message, or error if style not found

    Example:
        apply_style(key, 0, "Title")
        apply_style(key, 1, "Normal")
    """
    doc = document_manager.get_document(path)
    if doc is None:
        return f"Error: No document open at '{path}'. Use open_document or create_document first."

    paragraphs = doc.paragraphs
    para_count = len(paragraphs)

    # Validate index
    if index < 0 or index >= para_count:
        return f"Error: Invalid paragraph index {index}. Document has {para_count} paragraphs (valid range: 0-{para_count-1})."

    para = paragraphs[index]

    # Try to apply style
    try:
        para.style = style_name
    except KeyError:
        # Style not found - list available paragraph styles
        from docx.enum.style import WD_STYLE_TYPE
        available_styles = [s.name for s in doc.styles if s.type == WD_STYLE_TYPE.PARAGRAPH]
        logger.error("tool_operation_failed", tool="apply_style", error=f"Style '{style_name}' not found", error_type="KeyError")
        return f"Error: Style '{style_name}' not found. Available paragraph styles: {', '.join(available_styles)}"

    return f"Applied '{style_name}' style to paragraph {index}"
