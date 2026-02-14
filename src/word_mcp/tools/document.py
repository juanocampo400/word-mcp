"""
Document lifecycle tools for word-mcp.

This module provides MCP tool functions for document management operations:
create, open, save, save-as, close, info, from-template, and list.

All functions follow explicit-save-only semantics (no auto-save) and
error-on-overwrite for document creation operations.
"""

import os
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Optional

from ..document_manager import document_manager
from ..logging_config import get_logger
from ..errors import validate_document_size, DocumentTooLargeError, format_size

logger = get_logger(__name__)


def create_document(path: Optional[str] = None) -> str:
    """
    Create a new blank Word document in memory.

    LOCKED DECISION: Error-on-overwrite - if path provided and file exists,
    returns error instead of overwriting.

    Args:
        path: Optional file path for the document. If provided, must not exist.
              If None, creates with temporary "Untitled-N" name.

    Returns:
        Success message with document name/path, or error message prefixed with "Error:"

    Examples:
        >>> create_document()
        "Created new document 'Untitled-1'"

        >>> create_document("C:/Documents/report.docx")
        "Created new document at C:\\Documents\\report.docx"

        >>> create_document("C:/Documents/existing.docx")
        "Error: File already exists at C:\\Documents\\existing.docx. Use a different path or open the existing document."
    """
    try:
        if path:
            abs_path = str(Path(path).resolve())

        key, doc = document_manager.create_document(path)

        logger.info("document_created", key=key, has_path=not key.startswith("Untitled-"))

        if key.startswith("Untitled-"):
            return f"Created new document '{key}'"
        else:
            return f"Created new document at {key}"

    except FileExistsError as e:
        logger.error("document_creation_failed", tool="create_document", error=str(e), error_type=type(e).__name__, path=path)
        return f"Error: File already exists at {path}. Use a different path or open the existing document."
    except Exception as e:
        logger.error("document_creation_failed", tool="create_document", error=str(e), error_type=type(e).__name__, path=path)
        return f"Error: {str(e)}"


def open_document(path: str) -> str:
    """
    Open an existing Word document from disk into memory.

    If document is already open, returns the cached instance. Provides basic
    document statistics upon opening.

    Args:
        path: Path to .docx file to open

    Returns:
        Success message with filename and document stats, or error message

    Examples:
        >>> open_document("C:/Documents/report.docx")
        "Opened 'report.docx' (12 paragraphs, ~350 words)"
    """
    try:
        abs_path = str(Path(path).resolve())

        # Validate document size before opening
        try:
            validate_document_size(abs_path)
        except DocumentTooLargeError as e:
            logger.error("document_too_large", tool="open_document", path=abs_path, size_bytes=e.size_bytes, max_bytes=e.max_bytes)
            return f"Error: Document exceeds 10MB size limit ({format_size(e.size_bytes)}). Large documents may cause memory issues."

        doc = document_manager.open_document(abs_path)

        # Calculate basic stats
        filename = Path(abs_path).name
        para_count = len(doc.paragraphs)
        word_count = sum(len(para.text.split()) for para in doc.paragraphs)

        logger.info("document_opened", path=abs_path, paragraphs=para_count, words=word_count)

        return f"Opened '{filename}' ({para_count} paragraphs, ~{word_count} words)"

    except FileNotFoundError:
        logger.error("document_open_failed", tool="open_document", error="File not found", error_type="FileNotFoundError", path=path)
        return f"Error: File not found: {path}"
    except Exception as e:
        logger.error("document_open_failed", tool="open_document", error=str(e), error_type=type(e).__name__, path=path)
        return f"Error: {str(e)}"


def save_document(path: str) -> str:
    """
    Save an open document to its current path.

    LOCKED DECISION: Explicit save only - this is the ONLY way changes persist.
    No auto-save occurs anywhere in the system.

    Args:
        path: Path or key of open document to save

    Returns:
        Success message with absolute path, or error message

    Examples:
        >>> save_document("Untitled-1")
        "Error: Cannot save untitled document without path. Use save_document_as instead."

        >>> save_document("C:/Documents/report.docx")
        "Saved document to C:\\Documents\\report.docx"
    """
    try:
        abs_path = str(Path(path).resolve()) if not path.startswith("Untitled-") else path
        document_manager.save_document(abs_path)

        logger.info("document_saved", path=abs_path)

        return f"Saved document to {abs_path}"

    except ValueError as e:
        # Untitled documents need save_as
        logger.error("document_save_failed", tool="save_document", error=str(e), error_type=type(e).__name__, path=path)
        if "Cannot save untitled" in str(e):
            return f"Error: Cannot save untitled document without path. Use save_document_as instead."
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error("document_save_failed", tool="save_document", error=str(e), error_type=type(e).__name__, path=path)
        return f"Error: {str(e)}"


def save_document_as(path: str, new_path: str) -> str:
    """
    Save an open document to a new path (save-as operation).

    Document is re-keyed to the new path after saving. The old path is no longer
    associated with this document instance.

    Args:
        path: Current path or key of open document
        new_path: New path to save document to

    Returns:
        Success message showing old and new paths, or error message

    Examples:
        >>> save_document_as("Untitled-1", "C:/Documents/report.docx")
        "Saved document to C:\\Documents\\report.docx (was: Untitled-1)"

        >>> save_document_as("C:/Documents/draft.docx", "C:/Documents/final.docx")
        "Saved document to C:\\Documents\\final.docx (was: C:\\Documents\\draft.docx)"
    """
    try:
        old_key = str(Path(path).resolve()) if not path.startswith("Untitled-") else path
        abs_new = str(Path(new_path).resolve())

        document_manager.save_document(old_key, save_as=abs_new)

        logger.info("document_saved_as", old_path=old_key, new_path=abs_new)

        return f"Saved document to {abs_new} (was: {old_key})"

    except Exception as e:
        logger.error("document_save_as_failed", tool="save_document_as", error=str(e), error_type=type(e).__name__, path=path, new_path=new_path)
        return f"Error: {str(e)}"


def close_document(path: str) -> str:
    """
    Close an open document (remove from memory).

    LOCKED DECISION: Explicit-only behavior - unsaved changes are DISCARDED.
    No warning or prompt. User must save before closing if they want to keep changes.

    Args:
        path: Path or key of document to close

    Returns:
        Success message, or error message if document not open

    Examples:
        >>> close_document("C:/Documents/report.docx")
        "Closed document 'C:\\Documents\\report.docx'. Unsaved changes were discarded."
    """
    try:
        key = str(Path(path).resolve()) if not path.startswith("Untitled-") else path
        document_manager.close_document(key)

        logger.info("document_closed", path=key)

        return f"Closed document '{key}'. Unsaved changes were discarded."

    except ValueError:
        logger.error("document_close_failed", tool="close_document", error="Document not open", error_type="ValueError", path=path)
        return f"Error: Document not open: {path}"
    except Exception as e:
        logger.error("document_close_failed", tool="close_document", error=str(e), error_type=type(e).__name__, path=path)
        return f"Error: {str(e)}"


def get_document_info(path: str) -> str:
    """
    Get detailed information about an open document.

    LOCKED DECISION: Must include page count, styles used, and word count.

    Provides comprehensive document statistics including:
    - File path and name
    - Paragraph, word, and character counts
    - Styles in use throughout the document
    - Page count (approximate, requires document to be saved first)
    - Core properties (title, author) if available

    Args:
        path: Path or key of open document

    Returns:
        Multi-line formatted string with document information, or error message

    Examples:
        >>> get_document_info("C:/Documents/report.docx")
        '''Document Information: C:\\Documents\\report.docx

        File: report.docx
        Paragraphs: 42
        Words: ~1,250
        Characters: 7,890

        Styles in use:
          - Normal
          - Heading 1
          - Heading 2
          - List Paragraph

        Page count: 5

        Title: Q4 Sales Report
        Author: John Smith
        '''
    """
    try:
        key = str(Path(path).resolve()) if not path.startswith("Untitled-") else path
        doc = document_manager.get_document(key)

        # Basic counts
        para_count = len(doc.paragraphs)
        word_count = sum(len(para.text.split()) for para in doc.paragraphs)
        char_count = sum(len(para.text) for para in doc.paragraphs)

        # Styles in use
        styles = set()
        for para in doc.paragraphs:
            if para.style and para.style.name:
                styles.add(para.style.name)

        # Page count (requires saved document)
        page_count = _get_page_count(key)

        # Core properties
        title = doc.core_properties.title or "Not set"
        author = doc.core_properties.author or "Not set"

        # Format output
        result = f"Document Information: {key}\n\n"
        result += f"File: {Path(key).name if not key.startswith('Untitled-') else key}\n"
        result += f"Paragraphs: {para_count}\n"
        result += f"Words: ~{word_count:,}\n"
        result += f"Characters: {char_count:,}\n\n"

        if styles:
            result += "Styles in use:\n"
            for style in sorted(styles):
                result += f"  - {style}\n"
            result += "\n"

        result += f"Page count: {page_count}\n\n"
        result += f"Title: {title}\n"
        result += f"Author: {author}"

        return result

    except ValueError:
        logger.error("get_document_info_failed", tool="get_document_info", error="Document not open", error_type="ValueError", path=path)
        return f"Error: Document not open: {path}"
    except Exception as e:
        logger.error("get_document_info_failed", tool="get_document_info", error=str(e), error_type=type(e).__name__, path=path)
        return f"Error: {str(e)}"


def _get_page_count(path: str) -> str:
    """
    Extract page count from saved .docx file.

    Page count is stored in docProps/app.xml inside the .docx ZIP archive.
    Only available for documents that have been saved to disk.

    Args:
        path: Absolute path to saved .docx file

    Returns:
        Page count as string, or error message if unavailable
    """
    # Untitled documents haven't been saved yet
    if path.startswith("Untitled-"):
        return "unavailable (save document first)"

    # Check if file exists on disk
    if not Path(path).exists():
        return "unavailable (save document first)"

    try:
        # Extract from docProps/app.xml in the .docx ZIP
        with zipfile.ZipFile(path, 'r') as docx_zip:
            app_xml = docx_zip.read('docProps/app.xml')

        # Parse XML and find <Pages> element
        root = ET.fromstring(app_xml)

        # Handle namespace in app.xml
        ns = {'ep': 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'}
        pages_elem = root.find('.//ep:Pages', ns)

        if pages_elem is not None and pages_elem.text:
            return pages_elem.text
        else:
            return "unavailable"

    except KeyError:
        # app.xml doesn't exist in the ZIP
        return "unavailable"
    except Exception:
        # Any other parsing error
        return "unavailable"


def create_from_template(
    template_path: str,
    save_path: Optional[str] = None
) -> str:
    """
    Create a new document from a .docx or .dotx template.

    Opens the template file and creates a new document instance based on it.
    The template file itself is not modified.

    Args:
        template_path: Path to template file (.docx or .dotx)
        save_path: Optional path where document will be saved. If None, creates
                   with temporary "Untitled-N" name.

    Returns:
        Success message with template name and new document key, or error message

    Examples:
        >>> create_from_template("C:/Templates/report.dotx")
        "Created document from template 'report.dotx' as 'Untitled-1'"

        >>> create_from_template("C:/Templates/report.dotx", "C:/Documents/q4-report.docx")
        "Created document from template 'report.dotx' at C:\\Documents\\q4-report.docx"
    """
    try:
        abs_template = str(Path(template_path).resolve())
        template_name = Path(abs_template).name

        # Validate template size before opening
        try:
            validate_document_size(abs_template)
        except DocumentTooLargeError as e:
            logger.error("template_too_large", tool="create_from_template", template_path=abs_template, size_bytes=e.size_bytes, max_bytes=e.max_bytes)
            return f"Error: Template exceeds 10MB size limit ({format_size(e.size_bytes)}). Large documents may cause memory issues."

        key, doc = document_manager.create_from_template(abs_template, save_path)

        logger.info("document_created_from_template", template_path=abs_template, key=key, has_save_path=not key.startswith("Untitled-"))

        if key.startswith("Untitled-"):
            return f"Created document from template '{template_name}' as '{key}'"
        else:
            return f"Created document from template '{template_name}' at {key}"

    except FileNotFoundError:
        logger.error("create_from_template_failed", tool="create_from_template", error="Template not found", error_type="FileNotFoundError", template_path=template_path)
        return f"Error: Template not found: {template_path}"
    except FileExistsError:
        logger.error("create_from_template_failed", tool="create_from_template", error="File already exists", error_type="FileExistsError", template_path=template_path, save_path=save_path)
        return f"Error: File already exists at {save_path}. Use a different path."
    except Exception as e:
        logger.error("create_from_template_failed", tool="create_from_template", error=str(e), error_type=type(e).__name__, template_path=template_path)
        return f"Error: {str(e)}"


def list_open_documents() -> str:
    """
    List all currently open documents.

    Returns a formatted list of all document keys/paths currently held in memory.

    Returns:
        Formatted list of open documents, or message if none open

    Examples:
        >>> list_open_documents()
        '''Open documents:
          1. C:\\Documents\\report.docx
          2. C:\\Documents\\memo.docx
          3. Untitled-1
        '''

        >>> list_open_documents()  # When no documents open
        "No documents are currently open"
    """
    docs = document_manager.list_documents()

    if not docs:
        return "No documents are currently open"

    result = "Open documents:\n"
    for i, doc_path in enumerate(docs, 1):
        result += f"  {i}. {doc_path}\n"

    return result.rstrip()
