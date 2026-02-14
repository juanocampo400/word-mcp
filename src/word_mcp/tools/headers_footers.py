"""Header and footer tools for Word documents.

Provides header/footer read and write operations for Word document sections.
All functions use zero-based indexing for sections.

CRITICAL PATTERNS (from research):
- Always unlink headers/footers before editing (Pitfall 1)
- Enable different_first_page_header_footer before setting first-page content (Pitfall 2)
- Use paragraphs[0].text for initial content, not add_paragraph() (Pitfall 5)
"""

from ..document_manager import document_manager
from ..logging_config import get_logger

logger = get_logger(__name__)


def get_header(path: str, section_index: int = 0, header_type: str = "primary") -> str:
    """Get header content from a section.

    Args:
        path: Document path or key
        section_index: Zero-based section index (default: 0)
        header_type: Header type to read. Valid options:
                     - "primary": Main header (odd pages or all pages)
                     - "first_page": First page header
                     - "even_page": Even page header

    Returns:
        Formatted header content with metadata, or error message

    Example output:
        Header (primary) for section 0:
        Annual Report - Confidential
        Linked to previous: False
    """
    try:
        doc = document_manager.get_document(path)
    except ValueError as e:
        logger.error("tool_operation_failed", tool="get_header", error=str(e), error_type=type(e).__name__)
        return f"Error: {str(e)}"

    # Validate section_index
    section_count = len(doc.sections)
    if section_count == 0:
        return "Error: No sections found in document."
    if section_index < 0 or section_index >= section_count:
        return f"Error: Invalid section_index {section_index}. Document has {section_count} section(s) (valid range: 0-{section_count-1})."

    section = doc.sections[section_index]

    # Map header_type to section property
    header_type_lower = header_type.lower()
    if header_type_lower == "primary":
        header = section.header
    elif header_type_lower == "first_page":
        header = section.first_page_header
    elif header_type_lower == "even_page":
        header = section.even_page_header
    else:
        return f"Error: Invalid header_type '{header_type}'. Valid options: primary, first_page, even_page"

    # Read all paragraphs from header
    if header.paragraphs:
        header_text = "\n".join([p.text for p in header.paragraphs])
    else:
        header_text = "(empty)"

    # Build response with metadata
    lines = [f"Header ({header_type_lower}) for section {section_index}:"]
    lines.append(header_text)
    lines.append(f"Linked to previous: {header.is_linked_to_previous}")

    # For first_page header, also report if different_first_page is enabled
    if header_type_lower == "first_page":
        enabled = section.different_first_page_header_footer
        lines.append(f"Different first page enabled: {enabled} (content is ignored if False)")

    return "\n".join(lines)


def set_header(path: str, text: str, section_index: int = 0, header_type: str = "primary") -> str:
    """Set header content for a section.

    IMPORTANT: This function automatically:
    - Unlinks the header from previous section (creates unique definition)
    - Enables different_first_page_header_footer if header_type is "first_page"
    - Uses paragraphs[0].text (no empty paragraph above content)

    Args:
        path: Document path or key
        text: Header text to set
        section_index: Zero-based section index (default: 0)
        header_type: Header type to set. Valid options:
                     - "primary": Main header (odd pages or all pages)
                     - "first_page": First page header
                     - "even_page": Even page header

    Returns:
        Success message with preview of set text, or error message

    Example:
        set_header(key, "Confidential Report", 0, "primary")
        set_header(key, "Title Page", 0, "first_page")  # Auto-enables different first page
    """
    try:
        doc = document_manager.get_document(path)
    except ValueError as e:
        logger.error("tool_operation_failed", tool="set_header", error=str(e), error_type=type(e).__name__)
        return f"Error: {str(e)}"

    # Validate section_index
    section_count = len(doc.sections)
    if section_count == 0:
        return "Error: No sections found in document."
    if section_index < 0 or section_index >= section_count:
        return f"Error: Invalid section_index {section_index}. Document has {section_count} section(s) (valid range: 0-{section_count-1})."

    section = doc.sections[section_index]

    # Map header_type to section property
    header_type_lower = header_type.lower()
    if header_type_lower == "primary":
        header = section.header
    elif header_type_lower == "first_page":
        # CRITICAL: Enable different_first_page_header_footer BEFORE setting content (Pitfall 2)
        section.different_first_page_header_footer = True
        header = section.first_page_header
    elif header_type_lower == "even_page":
        header = section.even_page_header
    else:
        return f"Error: Invalid header_type '{header_type}'. Valid options: primary, first_page, even_page"

    # Check if header was previously linked
    was_linked = header.is_linked_to_previous

    # CRITICAL: Unlink header before editing (Pitfall 1)
    # If linked, editing would modify the source section's header
    header.is_linked_to_previous = False

    # CRITICAL: Use paragraphs[0].text for initial content (Pitfall 5)
    # add_paragraph() would leave empty paragraph above
    header.paragraphs[0].text = text

    # Build response message
    text_preview = text[:50] if len(text) <= 50 else text[:50] + "..."
    response = f"Set {header_type_lower} header for section {section_index}: '{text_preview}'"

    if was_linked:
        response += " (Header unlinked from previous section.)"

    return response


def get_footer(path: str, section_index: int = 0, footer_type: str = "primary") -> str:
    """Get footer content from a section.

    Args:
        path: Document path or key
        section_index: Zero-based section index (default: 0)
        footer_type: Footer type to read. Valid options:
                     - "primary": Main footer (odd pages or all pages)
                     - "first_page": First page footer
                     - "even_page": Even page footer

    Returns:
        Formatted footer content with metadata, or error message

    Example output:
        Footer (primary) for section 0:
        Page 1 of 10
        Linked to previous: False
    """
    try:
        doc = document_manager.get_document(path)
    except ValueError as e:
        logger.error("tool_operation_failed", tool="get_footer", error=str(e), error_type=type(e).__name__)
        return f"Error: {str(e)}"

    # Validate section_index
    section_count = len(doc.sections)
    if section_count == 0:
        return "Error: No sections found in document."
    if section_index < 0 or section_index >= section_count:
        return f"Error: Invalid section_index {section_index}. Document has {section_count} section(s) (valid range: 0-{section_count-1})."

    section = doc.sections[section_index]

    # Map footer_type to section property
    footer_type_lower = footer_type.lower()
    if footer_type_lower == "primary":
        footer = section.footer
    elif footer_type_lower == "first_page":
        footer = section.first_page_footer
    elif footer_type_lower == "even_page":
        footer = section.even_page_footer
    else:
        return f"Error: Invalid footer_type '{footer_type}'. Valid options: primary, first_page, even_page"

    # Read all paragraphs from footer
    if footer.paragraphs:
        footer_text = "\n".join([p.text for p in footer.paragraphs])
    else:
        footer_text = "(empty)"

    # Build response with metadata
    lines = [f"Footer ({footer_type_lower}) for section {section_index}:"]
    lines.append(footer_text)
    lines.append(f"Linked to previous: {footer.is_linked_to_previous}")

    # For first_page footer, also report if different_first_page is enabled
    if footer_type_lower == "first_page":
        enabled = section.different_first_page_header_footer
        lines.append(f"Different first page enabled: {enabled} (content is ignored if False)")

    return "\n".join(lines)


def set_footer(path: str, text: str, section_index: int = 0, footer_type: str = "primary") -> str:
    """Set footer content for a section.

    IMPORTANT: This function automatically:
    - Unlinks the footer from previous section (creates unique definition)
    - Enables different_first_page_header_footer if footer_type is "first_page"
    - Uses paragraphs[0].text (no empty paragraph above content)

    Args:
        path: Document path or key
        text: Footer text to set
        section_index: Zero-based section index (default: 0)
        footer_type: Footer type to set. Valid options:
                     - "primary": Main footer (odd pages or all pages)
                     - "first_page": First page footer
                     - "even_page": Even page footer

    Returns:
        Success message with preview of set text, or error message

    Example:
        set_footer(key, "Page 1", 0, "primary")
        set_footer(key, "Cover Page", 0, "first_page")  # Auto-enables different first page
    """
    try:
        doc = document_manager.get_document(path)
    except ValueError as e:
        logger.error("tool_operation_failed", tool="set_footer", error=str(e), error_type=type(e).__name__)
        return f"Error: {str(e)}"

    # Validate section_index
    section_count = len(doc.sections)
    if section_count == 0:
        return "Error: No sections found in document."
    if section_index < 0 or section_index >= section_count:
        return f"Error: Invalid section_index {section_index}. Document has {section_count} section(s) (valid range: 0-{section_count-1})."

    section = doc.sections[section_index]

    # Map footer_type to section property
    footer_type_lower = footer_type.lower()
    if footer_type_lower == "primary":
        footer = section.footer
    elif footer_type_lower == "first_page":
        # CRITICAL: Enable different_first_page_header_footer BEFORE setting content (Pitfall 2)
        section.different_first_page_header_footer = True
        footer = section.first_page_footer
    elif footer_type_lower == "even_page":
        footer = section.even_page_footer
    else:
        return f"Error: Invalid footer_type '{footer_type}'. Valid options: primary, first_page, even_page"

    # Check if footer was previously linked
    was_linked = footer.is_linked_to_previous

    # CRITICAL: Unlink footer before editing (Pitfall 1)
    # If linked, editing would modify the source section's footer
    footer.is_linked_to_previous = False

    # CRITICAL: Use paragraphs[0].text for initial content (Pitfall 5)
    # add_paragraph() would leave empty paragraph above
    footer.paragraphs[0].text = text

    # Build response message
    text_preview = text[:50] if len(text) <= 50 else text[:50] + "..."
    response = f"Set {footer_type_lower} footer for section {section_index}: '{text_preview}'"

    if was_linked:
        response += " (Footer unlinked from previous section.)"

    return response
