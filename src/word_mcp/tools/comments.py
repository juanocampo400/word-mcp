"""Comment reading tools for Word documents.

Provides read-only access to document comments with metadata (author, date, text).

NOTE: python-docx 1.2.0 does not expose comment location/range information. Comments are
returned with metadata (id, author, date, text, initials) but without information about
which text range they are anchored to.
"""

from pathlib import Path
from ..document_manager import document_manager
from ..logging_config import get_logger

logger = get_logger(__name__)


def get_comments(path: str) -> str:
    """Get all comments in the document with metadata.

    Returns comment metadata including id, author, date, text, and initials (if present).
    Comments are listed in the order they appear in the document structure.

    NOTE: python-docx 1.2.0 does not expose comment location/range information. Comments
    are returned with metadata but without information about which text range they are
    anchored to.

    Args:
        path: Document path or key

    Returns:
        Formatted string with comment list and metadata, or error message

    Example output:
        Comments in 'document.docx': 3 comment(s)
        [1] John Doe (initials: JD) (2026-02-13 10:30): This needs revision
        [2] Jane Smith (2026-02-13 11:45): Agreed, let's update this section
        [3] Bob Wilson (no date): Please review
    """
    try:
        doc = document_manager.get_document(path)
    except ValueError:
        return f"Error: No document open at '{path}'. Use open_document or create_document first."

    # Get filename for display
    if path.startswith("Untitled-"):
        filename = path
    else:
        filename = Path(path).name

    # Access comments collection
    try:
        comments = list(doc.comments)
    except AttributeError:
        # python-docx version doesn't support comments API
        return f"Error: Document comments API not available. Requires python-docx 1.2.0 or later."

    # Check if there are any comments
    if len(comments) == 0:
        return f"No comments found in '{filename}'."

    # Build header
    lines = [f"Comments in '{filename}': {len(comments)} comment(s)"]

    # Build comment list
    for comment in comments:
        # Extract comment properties
        comment_id = comment.id if hasattr(comment, 'id') else "?"
        author = comment.author if hasattr(comment, 'author') and comment.author else "Unknown"
        text = comment.text if hasattr(comment, 'text') and comment.text else ""
        initials = comment.initials if hasattr(comment, 'initials') and comment.initials else None

        # Format date
        if hasattr(comment, 'date') and comment.date is not None:
            try:
                # Format datetime as "YYYY-MM-DD HH:MM"
                date_str = comment.date.strftime("%Y-%m-%d %H:%M")
            except (AttributeError, ValueError):
                date_str = "no date"
        else:
            date_str = "no date"

        # Build comment line
        if initials:
            author_str = f"{author} (initials: {initials})"
        else:
            author_str = author

        lines.append(f"[{comment_id}] {author_str} ({date_str}): {text}")

    return "\n".join(lines)
