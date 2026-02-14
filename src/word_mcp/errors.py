"""Error utilities and validation for word-mcp.

Provides document size validation and custom exception types for production use.
"""

import os
from typing import Optional


# Maximum document size: 10 MB
MAX_DOCUMENT_SIZE = 10 * 1024 * 1024  # 10MB in bytes


class DocumentTooLargeError(ValueError):
    """Raised when a document exceeds the maximum allowed size.

    Attributes:
        path: Path to the oversized document
        size_bytes: Actual size of the document in bytes
        max_bytes: Maximum allowed size in bytes
    """

    def __init__(self, path: str, size_bytes: int, max_bytes: int):
        self.path = path
        self.size_bytes = size_bytes
        self.max_bytes = max_bytes
        super().__init__(
            f"Document {path} ({format_size(size_bytes)}) exceeds "
            f"maximum size limit of {format_size(max_bytes)}"
        )


def format_size(bytes_count: int) -> str:
    """Format byte count as human-readable size string.

    Args:
        bytes_count: Number of bytes

    Returns:
        Human-readable size string (e.g., "5.2 MB", "1.5 KB")
    """
    if bytes_count < 1024:
        return f"{bytes_count} B"
    elif bytes_count < 1024 * 1024:
        return f"{bytes_count / 1024:.1f} KB"
    elif bytes_count < 1024 * 1024 * 1024:
        return f"{bytes_count / (1024 * 1024):.1f} MB"
    else:
        return f"{bytes_count / (1024 * 1024 * 1024):.1f} GB"


def validate_document_size(path: str) -> None:
    """Validate that a document does not exceed the maximum size limit.

    Args:
        path: Path to the document file

    Raises:
        DocumentTooLargeError: If the document exceeds MAX_DOCUMENT_SIZE
        OSError: If the file cannot be accessed
    """
    size = os.path.getsize(path)
    if size > MAX_DOCUMENT_SIZE:
        raise DocumentTooLargeError(path, size, MAX_DOCUMENT_SIZE)
