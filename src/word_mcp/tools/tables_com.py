"""
COM-based table editing tools for word-mcp.

This module provides table row and column deletion using win32com automation.
python-docx does not support row/column deletion, so COM automation is required.

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


def delete_table_row(path: str, table_index: int, row_index: int) -> str:
    """
    Delete a row from an existing table using COM automation (TBL-05).

    python-docx does not support row deletion, so this function uses win32com
    to delete a row and maintain synchronization with the in-memory Document.

    PREREQUISITE: Document must be saved to disk first (COM opens files from disk).

    Args:
        path: Path or key of open document
        table_index: Zero-based table index in the document
        row_index: Zero-based row index within the table

    Returns:
        Success message with updated row count, or error message prefixed with "Error:"

    Examples:
        >>> delete_table_row("C:/Documents/report.docx", 0, 2)
        "Deleted row 2 from table 0. Table now has 4 rows."

        >>> delete_table_row("Untitled-1", 0, 0)
        "Error: Document must be saved to disk before COM operations. Use save_document_as first."

    Design notes:
        - Requires COM automation: Document must be saved to disk
        - Bridge pattern: Uses COM to delete row, then reloads python-docx
        - Zero-based indexing: Converts to 1-based for COM internally
        - Index validation: Checks table and row exist before deleting
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

        # Validate table_index is within bounds (using python-docx for validation)
        if table_index < 0 or table_index >= len(doc.tables):
            return f"Error: Invalid table index {table_index}. Document has {len(doc.tables)} table(s) (valid range: 0-{len(doc.tables) - 1})."

        # Use COM to delete row
        try:
            with com_pool.get_word_app() as word:
                com_doc = word.Documents.Open(key)

                # Convert 0-based to 1-based for COM
                com_table_index = table_index + 1
                com_row_index = row_index + 1

                # Validate table exists in COM document
                if com_table_index > com_doc.Tables.Count:
                    return f"Error: Invalid table index {table_index}. Document has {com_doc.Tables.Count} table(s)."

                # Get table
                table = com_doc.Tables(com_table_index)

                # Validate row exists
                if com_row_index < 1 or com_row_index > table.Rows.Count:
                    return f"Error: Invalid row index {row_index}. Table {table_index} has {table.Rows.Count} row(s) (valid range: 0-{table.Rows.Count - 1})."

                # Delete row
                table.Rows(com_row_index).Delete()

                # Get updated row count
                updated_row_count = table.Rows.Count

                # Save and close
                com_doc.Save()
                com_doc.Close()

        except Exception as e:
            logger.error("tool_operation_failed", tool="unknown", error=str(e), error_type=type(e).__name__)
            return f"Error: COM automation failed: {str(e)}. Ensure Microsoft Word is installed."

        # Reload python-docx document to sync in-memory state
        document_manager._documents[key] = Document(key)

        return f"Deleted row {row_index} from table {table_index}. Table now has {updated_row_count} rows."

    except ValueError:
        return f"Error: Document not open: {path}"
    except Exception as e:
        logger.error("tool_operation_failed", tool="unknown", error=str(e), error_type=type(e).__name__)
        return f"Error: {str(e)}"


def delete_table_column(path: str, table_index: int, col_index: int) -> str:
    """
    Delete a column from an existing table using COM automation (TBL-07).

    python-docx does not support column deletion, so this function uses win32com
    to delete a column and maintain synchronization with the in-memory Document.

    PREREQUISITE: Document must be saved to disk first (COM opens files from disk).

    Args:
        path: Path or key of open document
        table_index: Zero-based table index in the document
        col_index: Zero-based column index within the table

    Returns:
        Success message with updated column count, or error message prefixed with "Error:"

    Examples:
        >>> delete_table_column("C:/Documents/report.docx", 0, 1)
        "Deleted column 1 from table 0. Table now has 3 columns."

        >>> delete_table_column("Untitled-1", 0, 0)
        "Error: Document must be saved to disk before COM operations. Use save_document_as first."

    Design notes:
        - Requires COM automation: Document must be saved to disk
        - Bridge pattern: Uses COM to delete column, then reloads python-docx
        - Zero-based indexing: Converts to 1-based for COM internally
        - Index validation: Checks table and column exist before deleting
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

        # Validate table_index is within bounds (using python-docx for validation)
        if table_index < 0 or table_index >= len(doc.tables):
            return f"Error: Invalid table index {table_index}. Document has {len(doc.tables)} table(s) (valid range: 0-{len(doc.tables) - 1})."

        # Use COM to delete column
        try:
            with com_pool.get_word_app() as word:
                com_doc = word.Documents.Open(key)

                # Convert 0-based to 1-based for COM
                com_table_index = table_index + 1
                com_col_index = col_index + 1

                # Validate table exists in COM document
                if com_table_index > com_doc.Tables.Count:
                    return f"Error: Invalid table index {table_index}. Document has {com_doc.Tables.Count} table(s)."

                # Get table
                table = com_doc.Tables(com_table_index)

                # Validate column exists
                if com_col_index < 1 or com_col_index > table.Columns.Count:
                    return f"Error: Invalid column index {col_index}. Table {table_index} has {table.Columns.Count} column(s) (valid range: 0-{table.Columns.Count - 1})."

                # Delete column
                table.Columns(com_col_index).Delete()

                # Get updated column count
                updated_col_count = table.Columns.Count

                # Save and close
                com_doc.Save()
                com_doc.Close()

        except Exception as e:
            logger.error("tool_operation_failed", tool="unknown", error=str(e), error_type=type(e).__name__)
            return f"Error: COM automation failed: {str(e)}. Ensure Microsoft Word is installed."

        # Reload python-docx document to sync in-memory state
        document_manager._documents[key] = Document(key)

        return f"Deleted column {col_index} from table {table_index}. Table now has {updated_col_count} columns."

    except ValueError:
        return f"Error: Document not open: {path}"
    except Exception as e:
        logger.error("tool_operation_failed", tool="unknown", error=str(e), error_type=type(e).__name__)
        return f"Error: {str(e)}"
