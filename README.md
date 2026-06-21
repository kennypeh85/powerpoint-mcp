# PowerPoint Live MCP

Edit a **live, open PowerPoint deck** in real time via COM automation. Read slides, set text, add slides/textboxes, find-replace, format shapes, export screenshots — all without closing or reopening the file.

This is an MCP (Model Context Protocol) server designed for AI agents like Hermes Agent, Claude Desktop, or any MCP-compatible client. It attaches to the running `PowerPoint.Application` instance via `pywin32` and exposes 17 tools over stdio MCP.

## Requirements

- **Windows** + Microsoft PowerPoint installed and running
- `pip install pywin32 mcp`
- pywin32 post-install: `python Scripts/pywin32_postinstall.py -install`

## Installation

### Hermes Agent

Add to `~/.hermes/config.yaml`:

```yaml
mcp_servers:
  powerpoint:
    command: "C:/path/to/python.exe"
    args: ["C:/path/to/powerpoint-mcp/server.py"]
    connect_timeout: 30
    timeout: 90
```

Restart Hermes: `/reset` to load the new tools.

### Claude Desktop

Add to `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "python",
      "args": ["C:/path/to/powerpoint-mcp/server.py"]
    }
  }
}
```

### Any MCP client

The server uses stdio transport and is compatible with any MCP-compatible client. Tools are registered automatically on connection.

## Tools (17 total)

| Tool | Description |
|------|-------------|
| `list_decks` | All open presentations (name, path, slide count, active flag) |
| `get_active_deck` | Details of the active presentation + current slide |
| `goto_slide(slide_index)` | Navigate the window to a slide (1-indexed) |
| `get_slide_content(slide_index)` | All shapes on a slide: text, position, type, name, index |
| `get_speaker_notes(slide_index)` | Speaker notes text for a slide |
| `get_selection` | Current selection (text range or shapes) |
| `set_shape_text(slide, shape, text, append?)` | Set/append text on an existing shape |
| `add_text_box(slide, text, left, top, width, height, ...)` | Add a textbox at position (points) |
| `add_slide(layout, position?, title?, body?)` | New slide (blank, title, content, etc.) |
| `duplicate_slide(slide_index)` | Duplicate (copy inserted after original) |
| `delete_slide(slide_index)` | Delete — irreversible |
| `set_speaker_notes(slide, notes, append?)` | Set/append notes |
| `replace_text(find, replace, scope, slide_index?)` | Find-replace (slide or deck scope) |
| `set_shape_fill_color(slide, shape, rgb)` | Fill colour as hex (e.g. "1E2761") |
| `set_font(slide, shape, size?, bold?, italic?, color_rgb?, font_name?)` | Font properties |
| `screenshot_slide(slide_index?, width?)` | Export one slide to PNG for visual QA |
| `screenshot_deck(width?, out_dir?)` | Export all slides to PNGs |

## Usage Examples

```
"List my open PowerPoint decks"
→ mcp_powerpoint_list_decks()

"Add a slide after slide 5 titled 'Quarterly Results'"
→ mcp_powerpoint_add_slide(layout="title_only", title="Quarterly Results", position=6)

"Replace 'old term' with 'new term' across the whole deck"
→ mcp_powerpoint_replace_text(find="old term", replace="new term", scope="deck")

"Show me what slide 4 looks like right now"
→ mcp_powerpoint_screenshot_slide(slide_index=4)
```

## Key Design Decisions

- **1-indexed slides/shapes** — matches what users see in PowerPoint
- **Forgiving shape lookup** — reference shapes by index, name, or text content
- **COM RGB helper** — `pptx_rgb(r, g, b)` prevents R/B channel swap (COM uses `R + G*256 + B*65536`)
- **No state between calls** — each tool call is independent, server is stateless

## License

MIT
