"""
Tracked changes tools for word-mcp.

This module provides MCP tool functions for tracked changes operations:
enable tracking, disable tracking, and read existing revisions.

These tools bridge python-docx (Phase 1 DocumentManager) with win32com
(COM automation). Documents must be saved to disk before COM operations
can be performed, as COM opens files from disk rather than memory.
"""

from pathlib import Path
from docx import Document
from ..document_manager import document_manager
from ..com_pool import com_pool
from ..logging_config import get_logger

logger = get_logger(__name__)


# WdRevisionType enum mapping (all 22 types from Word COM API)
REVISION_TYPES = {
    0: "NoRevision",
    1: "Insertion",
    2: "Deletion",
    3: "Property",
    4: "ParagraphNumber",
    5: "DisplayField",
    6: "ReconcileField",
    7: "ConflictField",
    8: "Style",
    9: "Replace",
    10: "ParagraphProperty",
    11: "TableProperty",
    12: "SectionProperty",
    13: "StyleDefinition",
    14: "MovedFrom",
    15: "MovedTo",
    16: "CellInsertion",
    17: "CellDeletion",
    18: "CellMerge",
    19: "ConflictInsertion",
    20: "ConflictDeletion",
    21: "ConflictDelete",
}


def enable_tracked_changes(path: str, author: str = "Claude") -> str:
    """
    Enable tracked changes on a document (TRACK-01).

    This tool bridges python-docx DocumentManager with win32com COM automation.
    The document MUST be saved to disk before enabling tracked changes, as COM
    opens files from disk rather than from python-docx memory.

    Args:
        path: Path or key of open document
        author: Author name for tracked changes (default: "Claude")

    Returns:
        Success message or error message prefixed with "Error:"

    Examples:
        >>> enable_tracked_changes("C:/Documents/report.docx", "Claude")
        "Tracked changes enabled on 'report.docx'. Author set to 'Claude'. All subsequent COM-based edits will be tracked."
    """
    try:
        # Validate document is open in DocumentManager
        key = path if path.startswith("Untitled-") else str(Path(path).resolve())
        doc = document_manager.get_document(key)

        # Check file exists on disk (COM requires saved file)
        if key.startswith("Untitled-"):
            return "Error: Document must be saved to disk before enabling tracked changes. Use save_document_as first."

        if not Path(key).exists():
            return "Error: Document must be saved to disk before enabling tracked changes. Use save_document first."

        # Use COM to enable tracked changes
        try:
            with com_pool.get_word_app() as word:
                com_doc = word.Documents.Open(key)

                # Set author for new revisions
                word.UserName = author

                # Enable tracked changes
                com_doc.TrackRevisions = True
                com_doc.ShowRevisions = True

                # Save and close
                com_doc.Save()
                com_doc.Close()

        except Exception as e:
            logger.error("tool_operation_failed", tool="enable_tracked_changes", error=str(e), error_type=type(e).__name__)
            return f"Error: COM automation failed: {str(e)}. Ensure Microsoft Word is installed."

        # Reload python-docx document to sync in-memory state
        document_manager._documents[key] = Document(key)

        filename = Path(key).name
        return f"Tracked changes enabled on '{filename}'. Author set to '{author}'. All subsequent COM-based edits will be tracked."

    except ValueError:
        return f"Error: Document not open: {path}"
    except Exception as e:
        logger.error("tool_operation_failed", tool="unknown", error=str(e), error_type=type(e).__name__)
        return f"Error: {str(e)}"


def disable_tracked_changes(path: str) -> str:
    """
    Disable tracked changes on a document (TRACK-02).

    Existing revisions are preserved but future edits will not be tracked.

    Args:
        path: Path or key of open document

    Returns:
        Success message or error message prefixed with "Error:"

    Examples:
        >>> disable_tracked_changes("C:/Documents/report.docx")
        "Tracked changes disabled on 'report.docx'. Future edits will not be tracked. Existing revisions are preserved."
    """
    try:
        # Validate document is open in DocumentManager
        key = path if path.startswith("Untitled-") else str(Path(path).resolve())
        doc = document_manager.get_document(key)

        # Check file exists on disk (COM requires saved file)
        if key.startswith("Untitled-"):
            return "Error: Document must be saved to disk before disabling tracked changes. Use save_document_as first."

        if not Path(key).exists():
            return "Error: Document must be saved to disk before disabling tracked changes. Use save_document first."

        # Use COM to disable tracked changes
        try:
            with com_pool.get_word_app() as word:
                com_doc = word.Documents.Open(key)

                # Disable tracked changes
                com_doc.TrackRevisions = False

                # Save and close
                com_doc.Save()
                com_doc.Close()

        except Exception as e:
            logger.error("tool_operation_failed", tool="disable_tracked_changes", error=str(e), error_type=type(e).__name__)
            return f"Error: COM automation failed: {str(e)}. Ensure Microsoft Word is installed."

        # Reload python-docx document to sync in-memory state
        document_manager._documents[key] = Document(key)

        filename = Path(key).name
        return f"Tracked changes disabled on '{filename}'. Future edits will not be tracked. Existing revisions are preserved."

    except ValueError:
        return f"Error: Document not open: {path}"
    except Exception as e:
        logger.error("tool_operation_failed", tool="unknown", error=str(e), error_type=type(e).__name__)
        return f"Error: {str(e)}"


def get_tracked_changes(path: str) -> str:
    """
    Read all tracked changes with metadata (TRACK-04).

    Returns information about all revisions in the document including type,
    author, date, and the text involved in each change.

    Args:
        path: Path or key of open document

    Returns:
        Formatted list of tracked changes or error message prefixed with "Error:"

    Examples:
        >>> get_tracked_changes("C:/Documents/report.docx")
        '''Tracked changes in 'report.docx': 3 revision(s)

        [1] Insertion by 'Claude' on 2026-02-13 14:30:00
            Text: "inserted text here"
        [2] Deletion by 'John' on 2026-02-13 14:31:00
            Text: "deleted text here"
        [3] Property by 'Claude' on 2026-02-13 14:32:00
            Text: "formatted text"
        '''
    """
    try:
        # Validate document is open in DocumentManager
        key = path if path.startswith("Untitled-") else str(Path(path).resolve())
        doc = document_manager.get_document(key)

        # Check file exists on disk (COM requires saved file)
        if key.startswith("Untitled-"):
            return "Error: Document must be saved to disk before reading tracked changes. Use save_document_as first."

        if not Path(key).exists():
            return "Error: Document must be saved to disk before reading tracked changes. Use save_document first."

        # Use COM to read tracked changes (read-only operation)
        revisions = []
        try:
            with com_pool.get_word_app() as word:
                com_doc = word.Documents.Open(key)

                # Iterate revisions (1-based indexing in COM!)
                for i in range(1, com_doc.Revisions.Count + 1):
                    rev = com_doc.Revisions(i)

                    # Extract metadata
                    rev_type = REVISION_TYPES.get(rev.Type, f"Unknown({rev.Type})")
                    author = rev.Author
                    text = rev.Range.Text

                    # Format date
                    try:
                        date = rev.Date.strftime('%Y-%m-%d %H:%M:%S')
                    except (AttributeError, ValueError):
                        date = str(rev.Date)

                    revisions.append({
                        'index': i,
                        'type': rev_type,
                        'author': author,
                        'date': date,
                        'text': text
                    })

                # Close without saving (read-only operation)
                com_doc.Close(SaveChanges=0)

        except Exception as e:
            logger.error("tool_operation_failed", tool="unknown", error=str(e), error_type=type(e).__name__)
            return f"Error: COM automation failed: {str(e)}. Ensure Microsoft Word is installed."

        # Format output
        filename = Path(key).name
        if not revisions:
            return f"No tracked changes found in '{filename}'."

        result = f"Tracked changes in '{filename}': {len(revisions)} revision(s)\n\n"
        for rev in revisions:
            result += f"[{rev['index']}] {rev['type']} by '{rev['author']}' on {rev['date']}\n"
            result += f"    Text: \"{rev['text']}\"\n"

        return result.rstrip()

    except ValueError:
        return f"Error: Document not open: {path}"
    except Exception as e:
        logger.error("tool_operation_failed", tool="unknown", error=str(e), error_type=type(e).__name__)
        return f"Error: {str(e)}"
