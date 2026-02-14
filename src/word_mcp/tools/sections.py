"""Section management tools for Word documents.

Provides section CRUD operations (list, add, modify properties) for Word documents.
All functions use zero-based indexing for sections.
"""

from docx.enum.section import WD_SECTION_START, WD_ORIENTATION
from docx.shared import Inches
from ..document_manager import document_manager
from ..logging_config import get_logger

logger = get_logger(__name__)


def list_sections(path: str) -> str:
    """List all sections in the document.

    Shows section index, break type, orientation, page dimensions, margins,
    header link status, and different first page setting.

    Args:
        path: Document path or key

    Returns:
        Formatted list of sections or error message

    Example output:
        Sections in 'document.docx': 2 section(s)

        Section 0: NEW_PAGE | Portrait | 8.5in x 11.0in
          Margins: top=1.0in, bottom=1.0in, left=1.0in, right=1.0in
          Header linked to previous: False | Different first page: False

        Section 1: CONTINUOUS | Landscape | 11.0in x 8.5in
          Margins: top=0.5in, bottom=0.5in, left=0.5in, right=0.5in
          Header linked to previous: True | Different first page: False
    """
    try:
        doc = document_manager.get_document(path)
    except ValueError as e:
        logger.error("tool_operation_failed", tool="list_sections", error=str(e), error_type=type(e).__name__)
        return f"Error: {str(e)}"

    sections = doc.sections
    section_count = len(sections)

    # Get filename for display
    from pathlib import Path
    filename = Path(path).name if not path.startswith("Untitled-") else path

    if section_count == 0:
        return f"No sections found in '{filename}' (rare but possible)."

    # Build header
    lines = [f"Sections in '{filename}': {section_count} section(s)"]

    # List each section
    for idx, section in enumerate(sections):
        # Section break type (convert enum to human-readable name)
        start_type = section.start_type
        start_type_name = _section_start_to_string(start_type)

        # Orientation
        orientation = "Portrait" if section.orientation == WD_ORIENTATION.PORTRAIT else "Landscape"

        # Page dimensions (convert EMU to inches)
        page_width_inches = section.page_width.inches
        page_height_inches = section.page_height.inches

        # Margins (convert EMU to inches)
        top_margin = section.top_margin.inches
        bottom_margin = section.bottom_margin.inches
        left_margin = section.left_margin.inches
        right_margin = section.right_margin.inches

        # Header link status
        header_linked = section.header.is_linked_to_previous

        # Different first page setting
        different_first_page = section.different_first_page_header_footer

        # Format section info
        lines.append("")  # Blank line between sections
        lines.append(f"Section {idx}: {start_type_name} | {orientation} | {page_width_inches:.1f}in x {page_height_inches:.1f}in")
        lines.append(f"  Margins: top={top_margin:.1f}in, bottom={bottom_margin:.1f}in, left={left_margin:.1f}in, right={right_margin:.1f}in")
        lines.append(f"  Header linked to previous: {header_linked} | Different first page: {different_first_page}")

    return "\n".join(lines)


def add_section(path: str, break_type: str = "new_page") -> str:
    """Add a new section to the document.

    The section is added at the end of the document.

    Args:
        path: Document path or key
        break_type: Section break type. Valid options:
                    - "new_page": Start section on new page (default)
                    - "continuous": Continue on same page
                    - "even_page": Start on next even page
                    - "odd_page": Start on next odd page
                    - "new_column": Start in next column (multi-column layouts)

    Returns:
        Success message with section index and count, or error message

    Example:
        add_section(key, "new_page")  # Standard page break
        add_section(key, "continuous")  # No page break
        add_section(key, "odd_page")  # Ensure odd page start
    """
    try:
        doc = document_manager.get_document(path)
    except ValueError as e:
        logger.error("tool_operation_failed", tool="add_section", error=str(e), error_type=type(e).__name__)
        return f"Error: {str(e)}"

    # Map break_type string to WD_SECTION_START enum
    break_type_lower = break_type.lower()
    break_type_map = {
        "new_page": WD_SECTION_START.NEW_PAGE,
        "continuous": WD_SECTION_START.CONTINUOUS,
        "even_page": WD_SECTION_START.EVEN_PAGE,
        "odd_page": WD_SECTION_START.ODD_PAGE,
        "new_column": WD_SECTION_START.NEW_COLUMN
    }

    if break_type_lower not in break_type_map:
        valid_types = ", ".join(break_type_map.keys())
        return f"Error: Invalid break_type '{break_type}'. Valid options: {valid_types}"

    start_type = break_type_map[break_type_lower]

    # Add section
    doc.add_section(start_type)

    # Reset different_first_page_header_footer on the new section.
    # python-docx copies this from the previous section's sectPr, which causes
    # new sections to unexpectedly inherit first-page header/footer behavior.
    new_section = doc.sections[-1]
    new_section.different_first_page_header_footer = False

    # Get section index (zero-based)
    section_idx = len(doc.sections) - 1
    section_count = len(doc.sections)

    return f"Added section {section_idx} with break type '{break_type}'. Document now has {section_count} section(s)."


def modify_section_properties(
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
    """Modify properties of an existing section.

    All dimension/margin properties are in inches.

    IMPORTANT: When changing orientation without explicitly providing page_width/page_height,
    dimensions are automatically swapped (e.g., portrait 8.5x11 becomes landscape 11x8.5).
    This prevents Word layout corruption (Research Pitfall 7).

    Args:
        path: Document path or key
        section_index: Zero-based section index
        orientation: Page orientation ("portrait" or "landscape", case-insensitive)
        page_width: Page width in inches
        page_height: Page height in inches
        top_margin: Top margin in inches
        bottom_margin: Bottom margin in inches
        left_margin: Left margin in inches
        right_margin: Right margin in inches
        header_distance: Distance from page edge to header (inches)
        footer_distance: Distance from page edge to footer (inches)

    Returns:
        Success message listing modified properties, or error message

    Example:
        modify_section_properties(key, 0, orientation="landscape")
        # Auto-swaps dimensions: 8.5x11 -> 11x8.5

        modify_section_properties(key, 1, top_margin=0.5, bottom_margin=0.5)
        # Sets custom margins

        modify_section_properties(key, 0, orientation="landscape",
                                   page_width=11.0, page_height=8.5)
        # Explicit dimensions (no auto-swap)
    """
    try:
        doc = document_manager.get_document(path)
    except ValueError as e:
        logger.error("tool_operation_failed", tool="unknown", error=str(e), error_type=type(e).__name__)
        return f"Error: {str(e)}"

    # Validate section_index
    section_count = len(doc.sections)
    if section_count == 0:
        return "Error: No sections found in document."
    if section_index < 0 or section_index >= section_count:
        return f"Error: Invalid section_index {section_index}. Document has {section_count} section(s) (valid range: 0-{section_count-1})."

    section = doc.sections[section_index]

    # Track what was changed for response message
    changes = []

    # Handle orientation change
    if orientation is not None:
        orientation_lower = orientation.lower()
        if orientation_lower not in ["portrait", "landscape"]:
            return "Error: orientation must be 'portrait' or 'landscape'."

        new_orientation = WD_ORIENTATION.PORTRAIT if orientation_lower == "portrait" else WD_ORIENTATION.LANDSCAPE

        # Auto-swap dimensions if orientation changes and dimensions not explicitly provided
        if section.orientation != new_orientation and page_width is None and page_height is None:
            # Swap current dimensions
            current_width = section.page_width
            current_height = section.page_height
            section.page_width = current_height
            section.page_height = current_width
            changes.append(f"orientation={orientation_lower}, page_width={current_height.inches:.1f}in (auto-swapped), page_height={current_width.inches:.1f}in (auto-swapped)")
        else:
            changes.append(f"orientation={orientation_lower}")

        section.orientation = new_orientation

    # Handle explicit page dimensions
    if page_width is not None:
        section.page_width = Inches(page_width)
        if orientation is None:  # Only add if not already included above
            changes.append(f"page_width={page_width:.1f}in")

    if page_height is not None:
        section.page_height = Inches(page_height)
        if orientation is None:  # Only add if not already included above
            changes.append(f"page_height={page_height:.1f}in")

    # Handle margins
    if top_margin is not None:
        section.top_margin = Inches(top_margin)
        changes.append(f"top_margin={top_margin:.1f}in")

    if bottom_margin is not None:
        section.bottom_margin = Inches(bottom_margin)
        changes.append(f"bottom_margin={bottom_margin:.1f}in")

    if left_margin is not None:
        section.left_margin = Inches(left_margin)
        changes.append(f"left_margin={left_margin:.1f}in")

    if right_margin is not None:
        section.right_margin = Inches(right_margin)
        changes.append(f"right_margin={right_margin:.1f}in")

    # Handle header/footer distance
    if header_distance is not None:
        section.header_distance = Inches(header_distance)
        changes.append(f"header_distance={header_distance:.1f}in")

    if footer_distance is not None:
        section.footer_distance = Inches(footer_distance)
        changes.append(f"footer_distance={footer_distance:.1f}in")

    # Check if any properties were specified
    if not changes:
        return "Error: No properties specified to modify."

    return f"Modified section {section_index}: {', '.join(changes)}"


def _section_start_to_string(start_type) -> str:
    """Convert WD_SECTION_START enum value to human-readable string."""
    # WD_SECTION_START values: CONTINUOUS=0, NEW_COLUMN=1, NEW_PAGE=2, EVEN_PAGE=3, ODD_PAGE=4
    start_type_names = {
        0: "CONTINUOUS",
        1: "NEW_COLUMN",
        2: "NEW_PAGE",
        3: "EVEN_PAGE",
        4: "ODD_PAGE"
    }
    return start_type_names.get(start_type, f"UNKNOWN({start_type})")
