# Windows MCP Servers

Two MCP (Model Context Protocol) servers for Windows, designed for AI agents like Hermes Agent, Claude Desktop, or any MCP-compatible client.

## Servers

### 1. PowerPoint Live MCP (`powerpoint-mcp/`)
Edit a **live, open PowerPoint deck** in real time via COM automation. Read slides, set text, add slides/textboxes, find-replace, format shapes, export screenshots — all without closing or reopening the file.

**17 tools** including: `list_decks`, `get_slide_content`, `set_shape_text`, `add_slide`, `replace_text`, `set_font`, `set_shape_fill_color`, `screenshot_slide`, and more.

### 2. Tesseract OCR MCP (`tesseract-ocr-mcp/`)
OCR capabilities exposed over MCP. Extract text from images, screenshots, PDFs, and specific screen regions. Supports 100+ languages, table extraction, and confidence scoring.

**6 tools**: `extract_text`, `extract_text_region`, `extract_tables`, `screen_ocr`, `extract_text_from_pdf`, `get_available_languages`.

## Requirements

### PowerPoint MCP
- **Windows** + Microsoft PowerPoint installed and running
- `pip install pywin32 mcp`
- pywin32 post-install: `python Scripts/pywin32_postinstall.py -install`

### Tesseract OCR MCP
- **Tesseract OCR** installed: `winget install --id UB-Mannheim.TesseractOCR`
- `pip install mcp pytesseract Pillow`
- Optional: `pip install mss` (screen capture), `pip install pdf2image` (PDF OCR)
- Language packs: Download from [tessdata](https://github.com/tesseract-ocr/tessdata) and place in Tesseract's `tessdata` folder

## Installation

### Option A: Hermes Agent

Add to `~/.hermes/config.yaml`:

```yaml
mcp_servers:
  powerpoint:
    command: "C:/path/to/python.exe"
    args: ["C:/path/to/powerpoint-mcp/server.py"]
    connect_timeout: 30
    timeout: 90

  tesseract:
    command: "C:/path/to/python.exe"
    args: ["C:/path/to/tesseract-ocr-mcp/server.py"]
    env:
      TESSERACT_CMD: "C:/Program Files/Tesseract-OCR/tesseract.exe"
```

Restart Hermes: `/reset` to load the new tools.

### Option B: Claude Desktop

Add to `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "python",
      "args": ["C:/path/to/powerpoint-mcp/server.py"]
    },
    "tesseract": {
      "command": "python",
      "args": ["C:/path/to/tesseract-ocr-mcp/server.py"],
      "env": {
        "TESSERACT_CMD": "C:/Program Files/Tesseract-OCR/tesseract.exe"
      }
    }
  }
}
```

### Option C: Any MCP client

Both servers use stdio transport and are compatible with any MCP-compatible client. Each server registers its tools automatically on connection.

## Usage Examples

Once connected, the AI agent can:

```
"List my open PowerPoint decks"
→ mcp_powerpoint_list_decks()

"Add a slide after slide 5 titled 'Quarterly Results'"
→ mcp_powerpoint_add_slide(layout="title_only", title="Quarterly Results", position=6)

"Extract text from this screenshot: /path/to/screenshot.png"
→ tesseract_extract_text(image_path="/path/to/screenshot.png")

"What text is in the top-right corner of this image?"
→ tesseract_extract_text_region(image_path="/path/to/img.png", left=800, top=0, right=1200, bottom=200)

"Read this PDF and extract all text"
→ tesseract_extract_text_from_pdf(pdf_path="/path/to/document.pdf")
```

## Architecture

```
┌──────────────────┐     stdio MCP      ┌──────────────────┐
│   AI Agent       │ ←─────────────────→ │  MCP Server      │
│  (Hermes/Claude) │    JSON-RPC         │  (Python)        │
└──────────────────┘                     └────────┬─────────┘
                                                  │
                                          ┌───────┴───────┐
                                  PowerPoint COM   Tesseract CLI
                                  (pywin32)        (pytesseract)
```

Both servers use the `mcp` Python SDK with `FastMCP` for declarative tool registration. COM connections are re-established per tool call for robustness. Image downloads include retry logic for async generation APIs.

## Key Design Decisions

- **1-indexed slides/shapes** — matches what users see in PowerPoint
- **Forgiving shape lookup** — reference shapes by index, name, or text content
- **COM RGB helper** — `pptx_rgb(r, g, b)` prevents R/B channel swap (COM uses `R + G*256 + B*65536`)
- **Tesseract auto-discovery** — searches common paths, falls back to PATH
- **No state between calls** — each tool call is independent, servers are stateless

## License

MIT
