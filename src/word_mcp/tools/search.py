"""Search and replace tools for Word documents.

Provides text search and find-and-replace functionality with case sensitivity control.
Supports plain text search (default) and optional regex search via use_regex=True.
"""

import re
from ..document_manager import document_manager
from ..logging_config import get_logger

logger = get_logger(__name__)


def search_text(path: str, query: str, case_sensitive: bool = False, use_regex: bool = False) -> str:
    """Search for text in document paragraphs.

    Returns all paragraphs containing the query text with match counts and context.
    Supports both plain text search (default) and regex search via use_regex=True.

    Args:
        path: Document path or key
        query: Text to search for. When use_regex=True, treated as a regex pattern.
        case_sensitive: If False, performs case-insensitive search (default: False)
        use_regex: If True, treats query as a regular expression pattern (default: False).
                   Returns a clear error message if the pattern is invalid.

    Returns:
        Formatted search results with paragraph indexes and context, or "No matches found"

    Example output:
        Found 2 match(es) in 2 paragraph(s):
        [1] 1 match: "This is the first paragraph of content."
        [2] 1 match: "This is the second paragraph of content."

    Regex example:
        search_text(key, r"\\bword\\b", use_regex=True)  # Whole-word match
        search_text(key, r"\\d{4}-\\d{2}-\\d{2}", use_regex=True)  # ISO date pattern
    """
    doc = document_manager.get_document(path)
    if doc is None:
        return f"Error: No document open at '{path}'. Use open_document or create_document first."

    # Validate regex pattern up-front to give a clear error before touching paragraphs
    if use_regex:
        flags = 0 if case_sensitive else re.IGNORECASE
        try:
            compiled = re.compile(query, flags)
        except re.error as exc:
            return f"Error: Invalid regex pattern: {exc}"

    paragraphs = doc.paragraphs
    matches = []
    total_matches = 0

    for i, para in enumerate(paragraphs):
        text = para.text

        if use_regex:
            # Regex path
            found = compiled.findall(text)
            match_count = len(found)
            if match_count > 0:
                total_matches += match_count
                m = compiled.search(text)
                first_match_pos = m.start() if m else 0
                match_len = len(m.group(0)) if m else len(query)

                if len(text) <= 150:
                    context = text
                else:
                    start = max(0, first_match_pos - 50)
                    end = min(len(text), first_match_pos + match_len + 50)
                    context = text[start:end]
                    if start > 0:
                        context = "..." + context
                    if end < len(text):
                        context = context + "..."

                matches.append((i, match_count, context))
        else:
            # Plain text path (original behavior)
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
                    text_lower = text.lower()
                    query_lower = query.lower()
                    first_match_pos = text_lower.find(query_lower)

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

    Preserves the formatting of the first run in each modified paragraph (bold,
    italic, font size, color, underline, font name). If a paragraph has no runs,
    falls back to direct text assignment. Consistent with edit_paragraph behavior.

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

        # Build full paragraph text from runs for matching
        runs = list(para.runs)
        if runs:
            text = "".join(r.text for r in runs)
        else:
            text = para.text

        # Check if find_text exists
        if case_sensitive:
            has_match = find_text in text
        else:
            has_match = find_text.lower() in text.lower()

        if has_match:
            # Perform replacement on full text string
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

            # Apply replacement at run level to preserve formatting
            if not runs:
                # No runs: fall back to direct assignment
                para.text = new_text
            else:
                # Set first run's text to full new text; clear remaining runs
                # This preserves the first run's font properties on the new text
                runs[0].text = new_text
                for run in runs[1:]:
                    run.text = ""

            total_replacements += count
            paragraphs_modified += 1

            # If not replace_all, stop after first replacement
            if not replace_all:
                stopped = True

    # Format result
    if total_replacements == 0:
        return f"No occurrences of '{find_text}' found."

    return f"Replaced {total_replacements} occurrence(s) of '{find_text}' with '{replace_with}' in {paragraphs_modified} paragraph(s)."
