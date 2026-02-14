"""Search and replace tools for Word documents.

Provides text search and find-and-replace functionality with case sensitivity control.
Plain text search only (no regex) for v1.
"""

import re
from ..document_manager import document_manager
from ..logging_config import get_logger

logger = get_logger(__name__)


def search_text(path: str, query: str, case_sensitive: bool = False) -> str:
    """Search for text in document paragraphs.

    Returns all paragraphs containing the query text with match counts and context.
    Plain text search only (no regex support in v1).

    Args:
        path: Document path or key
        query: Text to search for
        case_sensitive: If False, performs case-insensitive search (default: False)

    Returns:
        Formatted search results with paragraph indexes and context, or "No matches found"

    Example output:
        Found 2 match(es) in 2 paragraph(s):
        [1] 1 match: "This is the first paragraph of content."
        [2] 1 match: "This is the second paragraph of content."
    """
    doc = document_manager.get_document(path)
    if doc is None:
        return f"Error: No document open at '{path}'. Use open_document or create_document first."

    paragraphs = doc.paragraphs
    matches = []
    total_matches = 0

    for i, para in enumerate(paragraphs):
        text = para.text

        # Count matches in this paragraph
        if case_sensitive:
            match_count = text.count(query)
        else:
            match_count = text.lower().count(query.lower())

        if match_count > 0:
            total_matches += match_count

            # Create context: show 50 chars before/after first match, or full paragraph if short
            if case_sensitive:
                first_match_pos = text.find(query)
            else:
                # Case-insensitive: find position
                text_lower = text.lower()
                query_lower = query.lower()
                first_match_pos = text_lower.find(query_lower)

            # Build context window
            if len(text) <= 150:
                context = text
            else:
                start = max(0, first_match_pos - 50)
                end = min(len(text), first_match_pos + len(query) + 50)
                context = text[start:end]
                if start > 0:
                    context = "..." + context
                if end < len(text):
                    context = context + "..."

            matches.append((i, match_count, context))

    # Format results
    if total_matches == 0:
        return f"No matches found for '{query}'"

    lines = [f"Found {total_matches} match(es) in {len(matches)} paragraph(s):"]
    for idx, count, context in matches:
        lines.append(f"[{idx}] {count} match(es): \"{context}\"")

    return "\n".join(lines)


def replace_text(
    path: str,
    find_text: str,
    replace_with: str,
    case_sensitive: bool = False,
    replace_all: bool = True
) -> str:
    """Find and replace text across the document.

    Limitation: Replaces at paragraph level. Formatting within replaced paragraphs
    will be reset to default (acceptable for Phase 1).

    Args:
        path: Document path or key
        find_text: Text to find
        replace_with: Text to replace with
        case_sensitive: If False, performs case-insensitive replacement (default: False)
        replace_all: If True, replaces all occurrences; if False, replaces only first (default: True)

    Returns:
        Summary of replacements made, or "No occurrences found"

    Example:
        replace_text(key, "paragraph", "section")  # Replace all, case-insensitive
        replace_text(key, "TODO", "DONE", case_sensitive=True, replace_all=False)  # Replace first only
    """
    doc = document_manager.get_document(path)
    if doc is None:
        return f"Error: No document open at '{path}'. Use open_document or create_document first."

    paragraphs = doc.paragraphs
    total_replacements = 0
    paragraphs_modified = 0
    stopped = False

    for para in paragraphs:
        if stopped:
            break

        text = para.text

        # Check if find_text exists
        if case_sensitive:
            has_match = find_text in text
        else:
            has_match = find_text.lower() in text.lower()

        if has_match:
            # Perform replacement
            if case_sensitive:
                if replace_all:
                    new_text = text.replace(find_text, replace_with)
                    count = text.count(find_text)
                else:
                    new_text = text.replace(find_text, replace_with, 1)
                    count = 1
            else:
                # Case-insensitive replacement using regex
                if replace_all:
                    new_text = re.sub(re.escape(find_text), replace_with, text, flags=re.IGNORECASE)
                    count = len(re.findall(re.escape(find_text), text, flags=re.IGNORECASE))
                else:
                    new_text = re.sub(re.escape(find_text), replace_with, text, count=1, flags=re.IGNORECASE)
                    count = 1

            # Apply replacement (loses formatting)
            para.text = new_text
            total_replacements += count
            paragraphs_modified += 1

            # If not replace_all, stop after first replacement
            if not replace_all:
                stopped = True

    # Format result
    if total_replacements == 0:
        return f"No occurrences of '{find_text}' found."

    return f"Replaced {total_replacements} occurrence(s) of '{find_text}' with '{replace_with}' in {paragraphs_modified} paragraph(s)."
