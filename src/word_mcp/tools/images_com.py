"""
COM-based image positioning tools for word-mcp.

This module provides image repositioning using win32com automation.
python-docx does not support absolute positioning of images, so COM automation is required.

All functions use the bridge pattern:
1. Validate document open in DocumentManager
2. Validate document saved to disk (COM needs file on disk)
3. Open via COM (WordApplication context manager)
4. Perform COM-based operation
5. Save and close via COM
6. Reload python-docx document to sync state
"""

from pathlib import Path
from docx import Document
from ..document_manager import document_manager
from ..com_pool import com_pool
from ..logging_config import get_logger

logger = get_logger(__name__)


def reposition_image(
    path: str,
    image_index: int,
    left: float = None,
    top: float = None,
    width: float = None,
    height: float = None
) -> str:
    """
    Reposition an inline image to an absolute position on the page (IMG-03).

    python-docx does not support absolute image positioning, so this function uses
    win32com to convert an InlineShape to a floating Shape with absolute coordinates.

    WARNING: This is a ONE-WAY conversion. Once an InlineShape becomes a Shape, it
    cannot be converted back to InlineShape. The image will no longer appear in
    doc.inline_shapes after this operation.

    PREREQUISITE: Document must be saved to disk first (COM opens files from disk).

    Args:
        path: Path or key of open document
        image_index: Zero-based index into InlineShapes collection
        left: Optional horizontal position in inches (from left edge of page)
        top: Optional vertical position in inches (from top edge of page)
        width: Optional width in inches (for resizing during reposition)
        height: Optional height in inches (for resizing during reposition)

    Returns:
        Success message with position details, or error message prefixed with "Error:"

    Examples:
        >>> reposition_image("C:/Documents/report.docx", 0, left=1.0, top=2.0)
        "Repositioned image 0 to absolute position (left=1.0in, top=2.0in). Image converted from inline to floating shape."

        >>> reposition_image("C:/Documents/report.docx", 0, left=1.5, top=1.5, width=3.0, height=2.0)
        "Repositioned image 0 to absolute position (left=1.5in, top=1.5in). Resized to 3.0in x 2.0in. Image converted from inline to floating shape."

        >>> reposition_image("Untitled-1", 0, left=1.0, top=1.0)
        "Error: Document must be saved to disk before COM operations. Use save_document_as first."

    Design notes:
        - Requires COM automation: Document must be saved to disk
        - One-way conversion: InlineShape becomes Shape (cannot revert)
        - Bridge pattern: Uses COM to reposition, then reloads python-docx
        - Zero-based indexing: Converts to 1-based for COM internally
        - Coordinate units: Input in inches, converted to points (1 inch = 72 points)
        - Optional parameters: At least one position parameter (left or top) recommended
    """
    try:
        # Validate document is open in DocumentManager
        key = path if path.startswith("Untitled-") else str(Path(path).resolve())
        doc = document_manager.get_document(key)

        # Check file exists on disk (COM requires saved file)
        if key.startswith("Untitled-"):
            return "Error: Document must be saved to disk before COM operations. Use save_document_as first."

        if not Path(key).exists():
            return "Error: Document must be saved to disk before COM operations. Use save_document first."

        # Validate image_index is within bounds (using python-docx for validation)
        if image_index < 0 or image_index >= len(doc.inline_shapes):
            return f"Error: Invalid image index {image_index}. Document has {len(doc.inline_shapes)} inline image(s) (valid range: 0-{len(doc.inline_shapes) - 1})."

        # Use COM to reposition image
        try:
            with com_pool.get_word_app() as word:
                com_doc = word.Documents.Open(key)

                # Convert 0-based to 1-based for COM
                com_image_index = image_index + 1

                # Validate image exists in COM document
                if com_image_index > com_doc.InlineShapes.Count:
                    return f"Error: Invalid image index {image_index}. Document has {com_doc.InlineShapes.Count} inline image(s)."

                # Get inline shape and convert to floating shape
                inline_shape = com_doc.InlineShapes(com_image_index)
                shape = inline_shape.ConvertToShape()

                # Apply position (convert inches to points: 1 inch = 72 points)
                if left is not None:
                    shape.Left = left * 72
                if top is not None:
                    shape.Top = top * 72

                # Apply size if specified (convert inches to points)
                if width is not None:
                    shape.Width = width * 72
                if height is not None:
                    shape.Height = height * 72

                # Save and close
                com_doc.Save()
                com_doc.Close()

        except Exception as e:
            logger.error("tool_operation_failed", tool="unknown", error=str(e), error_type=type(e).__name__)
            return f"Error: COM automation failed: {str(e)}. Ensure Microsoft Word is installed."

        # Reload python-docx document to sync in-memory state
        document_manager._documents[key] = Document(key)

        # Build success message
        msg_parts = [f"Repositioned image {image_index} to absolute position"]
        position_parts = []
        if left is not None:
            position_parts.append(f"left={left}in")
        if top is not None:
            position_parts.append(f"top={top}in")
        if position_parts:
            msg_parts.append(f"({', '.join(position_parts)})")

        if width is not None or height is not None:
            size_parts = []
            if width is not None:
                size_parts.append(f"{width}in")
            if height is not None:
                size_parts.append(f"{height}in")
            msg_parts.append(f". Resized to {' x '.join(size_parts)}")

        msg_parts.append(". Image converted from inline to floating shape.")

        return "".join(msg_parts)

    except ValueError:
        return f"Error: Document not open: {path}"
    except Exception as e:
        logger.error("tool_operation_failed", tool="unknown", error=str(e), error_type=type(e).__name__)
        return f"Error: {str(e)}"
