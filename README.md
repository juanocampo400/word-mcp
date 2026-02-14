# Word MCP Server

Edit Word documents with Claude with full tracked changes support.

## Why This Exists

I wanted a Word MCP server that could do tracked changes. In a professional setting you need to see what changed, who changed it, and accept or reject edits.

Other approaches I looked at:
- **python-docx based servers** – Can read and write .docx files, but no tracked changes support. Changes are invisible in Word's review workflow.
- **Microsoft's official Word MCP** – Cloud-based, read/create only, no direct editing of local files, and no tracked changes.

This server:
- **Full tracked changes** – Enable revision tracking, make edits that appear as insertions/deletions in Word, review with author and date metadata
- **Direct local editing** – Changes happen in your .docx files on disk, no cloud dependencies
- **45 tools** – Documents, text, tables, images, formatting, sections, headers/footers, and more

## Getting Started

### Prerequisites

- **Windows** with **Microsoft Word** installed (required for COM-based tracked changes)
- **Python 3.10+** – Check with `python --version`

### Install

#### 1. Clone the repo:
```bash
git clone https://github.com/juanocampo400/word-mcp.git
cd word-mcp
```

#### 2. Install dependencies:
```bash
pip install -e .
```

#### 3. Add to Claude Code:

```bash
claude mcp add word-mcp --scope user -- python -m word_mcp.server
```

<details>
<summary><strong>Manual configuration</strong></summary>

**Why `--scope user`?** Makes the server available globally. Without it, the server only works in the project directory.

**Manual config** -- Edit `~/.claude.json`:

```json
{
  "mcpServers": {
    "word-mcp": {
      "type": "stdio",
      "command": "python",
      "args": ["-m", "word_mcp.server"]
    }
  }
}
```

If `python` doesn't resolve, use the full path:

```json
{
  "mcpServers": {
    "word-mcp": {
      "type": "stdio",
      "command": "C:/Users/yourname/AppData/Local/Programs/Python/Python312/python.exe",
      "args": ["-m", "word_mcp.server"]
    }
  }
}
```
</details>

### Try It Out

Open Claude Code and say:

> "Open my document at C:/Documents/report.docx and enable tracked changes, then fix the typos"

Claude will open the document, enable tracking, and make edits that appear as tracked changes in Word. Open the file in Word to review and accept/reject.

## How It Works

The server uses two engines:

- **python-docx** for most operations -- text, styles, tables, images, sections. Fast, doesn't need Word running.
- **pywin32 COM automation** for things python-docx can't do -- tracked changes, row/column deletion, image repositioning. Requires Word installed.

When a COM operation runs, the document is saved to disk, edited through Word's COM interface, then reloaded into python-docx. This bridge pattern keeps both engines in sync.

## Notes

- Windows only (COM automation requires Microsoft Word)
- Multiple documents can be open simultaneously
- Documents are held in memory until explicitly saved – no auto-save
- COM operations (tracked changes, some table/image ops) require the document to be saved to disk first

## Available Tools

### Document Management
| Tool | Description |
|------|-------------|
| `create_document` | Create a new blank document |
| `open_document` | Open a .docx file from disk |
| `save_document` | Save to current path |
| `save_document_as` | Save to a new path |
| `close_document` | Close and discard unsaved changes |
| `get_document_info` | Get paragraphs, words, pages, styles in use |
| `create_from_template` | Create from a .docx/.dotx template |
| `list_open_documents` | List all open documents |

### Content Editing
| Tool | Description |
|------|-------------|
| `read_document` | Read paragraphs with indexes and styles |
| `add_paragraph` | Add or insert paragraph |
| `edit_paragraph` | Replace paragraph text by index |
| `delete_paragraph` | Delete paragraph by index |
| `search_text` | Search for text across paragraphs |
| `replace_text` | Find and replace text |

### Styles & Formatting
| Tool | Description |
|------|-------------|
| `apply_heading_style` | Apply heading style (H1-H9) |
| `apply_style` | Apply any paragraph style by name |
| `format_text` | Bold, italic, underline, font, size, color |
| `get_paragraph_formatting` | Inspect per-run formatting details |

### Track Changes
| Tool | Description |
|------|-------------|
| `enable_tracked_changes` | Enable revision tracking with author name |
| `disable_tracked_changes` | Disable tracking (preserves existing revisions) |
| `get_tracked_changes` | List revisions with type, author, date, text |
| `tracked_add_paragraph` | Add paragraph as tracked insertion |
| `tracked_edit_paragraph` | Edit creating tracked deletion + insertion |
| `tracked_delete_paragraph` | Delete as tracked deletion (strikethrough) |

### Tables
| Tool | Description |
|------|-------------|
| `create_table` | Create table with optional data and style |
| `list_tables` | List all tables with dimensions |
| `read_table` | Read table as formatted grid |
| `edit_table_cell` | Edit cell by row/column index |
| `add_table_row` | Append row with optional data |
| `add_table_column` | Append column with optional data |
| `delete_table_row` | Delete row (requires saved document) |
| `delete_table_column` | Delete column (requires saved document) |

### Images
| Tool | Description |
|------|-------------|
| `insert_image` | Insert inline image with optional resize |
| `resize_image` | Resize existing image |
| `list_images` | List all images with dimensions |
| `reposition_image` | Convert to floating and position absolutely |

### Sections & Headers/Footers
| Tool | Description |
|------|-------------|
| `list_sections` | List sections with orientation, margins, dimensions |
| `add_section` | Add section (new page, continuous, etc.) |
| `modify_section_properties` | Change orientation, dimensions, margins |
| `get_header` / `set_header` | Read/write header content per section |
| `get_footer` / `set_footer` | Read/write footer content per section |

### Comments
| Tool | Description |
|------|-------------|
| `get_comments` | Read all comments with author and date |

### Monitoring
| Tool | Description |
|------|-------------|
| `get_server_health` | Memory usage, COM pool status, open documents |

## License

MIT
