# Windows Office MCP Servers

Three MCP (Model Context Protocol) servers for Windows that enable AI agents to edit **live, open Microsoft Office documents** in real time via COM automation. No file round-trips — edits appear instantly on screen.

Built for [Hermes Agent](https://hermes-agent.nousresearch.com), [Claude Desktop](https://claude.ai/desktop), or any MCP-compatible client.

## Servers

| Server | Tools | Description |
|--------|-------|-------------|
| **PowerPoint** | 17 | Edit live decks: slides, shapes, text, formatting, screenshots |
| **Word** | 16 | Edit live documents: paragraphs, tables, find-replace, formatting |
| **Excel** | 17 | Edit live workbooks: cells, ranges, tables, formatting, formulas |

## Requirements

- **Windows** + Microsoft Office installed (PowerPoint, Word, Excel)
- Python 3.11+
- `pip install pywin32 mcp`
- pywin32 post-install: `python Scripts/pywin32_postinstall.py -install`

## Installation

### Hermes Agent

Add to `~/.hermes/config.yaml`:

```yaml
mcp_servers:
  powerpoint:
    command: "C:/path/to/python.exe"
    args: ["C:/path/to/office-mcp/powerpoint-mcp/server.py"]
    timeout: 90

  word:
    command: "C:/path/to/python.exe"
    args: ["C:/path/to/office-mcp/word-mcp/server.py"]
    timeout: 90

  excel:
    command: "C:/path/to/python.exe"
    args: ["C:/path/to/office-mcp/excel-mcp/server.py"]
    timeout: 90
```

### Claude Desktop

Add to `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "python",
      "args": ["C:/path/to/office-mcp/powerpoint-mcp/server.py"]
    },
    "word": {
      "command": "python",
      "args": ["C:/path/to/office-mcp/word-mcp/server.py"]
    },
    "excel": {
      "command": "python",
      "args": ["C:/path/to/office-mcp/excel-mcp/server.py"]
    }
  }
}
```

## Tool Reference

### PowerPoint (17 tools)

| Tool | Description |
|------|-------------|
| `list_decks` | All open presentations |
| `get_active_deck` | Active presentation details |
| `goto_slide` | Navigate to slide (1-indexed) |
| `get_slide_content` | All shapes: text, position, type |
| `get_speaker_notes` | Speaker notes for a slide |
| `get_selection` | Current selection |
| `set_shape_text` | Set/append text on a shape |
| `add_text_box` | Add textbox at position |
| `add_slide` | New slide (multiple layouts) |
| `duplicate_slide` | Duplicate a slide |
| `delete_slide` | Delete a slide |
| `set_speaker_notes` | Set/append notes |
| `replace_text` | Find-replace (slide or deck scope) |
| `set_shape_fill_color` | Shape fill colour (hex) |
| `set_font` | Font properties on a shape |
| `screenshot_slide` | Export slide to PNG |
| `screenshot_deck` | Export all slides to PNGs |

### Word (16 tools)

| Tool | Description |
|------|-------------|
| `list_documents` | All open documents |
| `get_active_document` | Active document details |
| `create_new_document` | Create a blank document |
| `save_document` | Save to file path |
| `save_as_pdf` | Export as PDF |
| `get_document_content` | Get all text in document |
| `get_paragraphs` | List paragraphs (1-indexed, with text) |
| `get_sections` | List document sections |
| `get_tables` | List tables with dimensions |
| `get_selection` | Current selection text |
| `find_text` | Find text occurrences |
| `insert_paragraph` | Insert paragraph (before/after position) |
| `insert_text` | Insert text at various positions |
| `replace_text` | Find and replace all occurrences |
| `set_paragraph_alignment` | Set alignment (left/center/right/justify) |
| `set_font_properties` | Set font name, size, bold, italic, color |

### Excel (17 tools)

| Tool | Description |
|------|-------------|
| `list_workbooks` | All open workbooks |
| `get_active_workbook` | Active workbook details |
| `list_sheets` | List all worksheets |
| `read_cell` | Read cell value, formula, format |
| `read_range` | Read a range (A1:B10 format) |
| `read_table` | Read an Excel Table (ListObject) |
| `get_sheet_content` | Read all used cells as 2D array |
| `write_cell` | Write value to a cell |
| `write_range` | Write value to a range |
| `clear_cell` | Clear cell content |
| `clear_range` | Clear range content |
| `set_cell_format` | Set cell font properties |
| `set_range_format` | Set range font properties |
| `set_cell_fill` | Set cell fill color (hex) |
| `add_sheet` | Add a new worksheet |
| `delete_sheet` | Delete a sheet |
| `export_to_pdf` | Export workbook to PDF |

## Usage Examples

```
"List my open PowerPoint decks"
→ mcp_powerpoint_list_decks()

"Replace 'old term' with 'new term' in the Word document"
→ mcp_word_replace_text(find="old term", replace="new term")

"Read cells A1:B10 from the active Excel sheet"
→ mcp_excel_read_range(range="A1:B10")

"What's in cell C5?"
→ mcp_excel_read_cell(cell="C5")

"Add a slide after slide 5 titled 'Quarterly Results'"
→ mcp_powerpoint_add_slide(layout="title_only", title="Quarterly Results", position=6)

"Set the heading font to 18pt bold"
→ mcp_word_set_font_properties(size=18, bold=true)
```

## Architecture

```
┌──────────────────┐     stdio MCP      ┌──────────────────┐
│   AI Agent       │ ←─────────────────→ │  MCP Server      │
│  (Hermes/Claude) │    JSON-RPC         │  (Python)        │
└──────────────────┘                     └────────┬─────────┘
                                                  │
                                     ┌────────────┼────────────┐
                               PowerPoint COM  Word COM    Excel COM
                               (pywin32)       (pywin32)   (pywin32)
```

All three servers share the same design:
- **COM automation** via `pywin32` — attaches to the running Office application
- **FastMCP** with stdio transport for declarative tool registration
- **Per-call COM init** — each tool call re-initializes COM for thread safety
- **RGB helper** — `office_rgb(r, g, b)` prevents R/B channel swap (`R + G*256 + B*65536`)
- **1-indexed** — matches what users see in Office UI
- **Stateless** — each tool call is independent, no state held between calls

## License

MIT
