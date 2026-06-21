"""
Tesseract OCR MCP Server
========================
Exposes Tesseract OCR capabilities over stdio MCP for AI agents.
Extracts text from images, screenshots, PDFs, and screen regions.

Requirements
------------
- Tesseract OCR installed (https://github.com/UB-Mannheim/tesseract/wiki)
  Windows: winget install --id UB-Mannheim.TesseractOCR
- Python packages: pip install mcp pytesseract Pillow pdf2image

Register in ~/.hermes/config.yaml or Claude Desktop config:
    mcp_servers:
      tesseract:
        command: "python"
        args: ["path/to/tesseract_mcp/server.py"]
        env:
          TESSERACT_CMD: "C:/Program Files/Tesseract-OCR/tesseract.exe"
"""

from __future__ import annotations

import base64
import io
import json
import os
import subprocess
import sys
import tempfile
from functools import wraps
from typing import Any

from mcp.server.fastmcp import FastMCP
from PIL import Image

# ─── Tesseract path resolution ────────────────────────────────────────────

def _get_tesseract_cmd() -> str:
    """Find tesseract executable."""
    # 1. Explicit env var
    cmd = os.environ.get("TESSERACT_CMD")
    if cmd and os.path.exists(cmd):
        return cmd
    # 2. Common Windows paths
    for path in [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        os.path.expanduser(r"~\AppData\Local\Tesseract-OCR\tesseract.exe"),
    ]:
        if os.path.exists(path):
            return path
    # 3. System PATH
    try:
        subprocess.run(["tesseract", "--version"], capture_output=True, check=True)
        return "tesseract"
    except (FileNotFoundError, subprocess.CalledProcessError):
        pass
    raise RuntimeError(
        "Tesseract not found. Set TESSERACT_CMD env var or install: "
        "winget install --id UB-Mannheim.TesseractOCR"
    )


TESSERACT_CMD = _get_tesseract_cmd()
import pytesseract
pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD

# ─── Image loading helpers ────────────────────────────────────────────────

def _load_image(source: str) -> Image.Image:
    """Load image from file path, URL, or base64 data URI."""
    if source.startswith("data:"):
        # data:image/png;base64,xxxx
        header, data = source.split(",", 1)
        img_data = base64.b64decode(data)
        return Image.open(io.BytesIO(img_data))
    elif source.startswith("http"):
        import requests
        resp = requests.get(source, timeout=30)
        resp.raise_for_status()
        return Image.open(io.BytesIO(resp.content))
    elif os.path.exists(source):
        return Image.open(source)
    else:
        raise FileNotFoundError(f"Image not found: {source}")


def _ocr_tool(fn):
    """Decorator with error capture."""
    @wraps(fn)
    def wrapper(*args, **kwargs):
        try:
            return fn(*args, **kwargs)
        except Exception as e:
            import traceback
            return json.dumps({
                "error": f"{type(e).__name__}: {e}",
                "traceback": traceback.format_exc(),
            })
    return wrapper


# ─── MCP server + tools ───────────────────────────────────────────────────

mcp = FastMCP("tesseract-ocr")


@mcp.tool()
@_ocr_tool
def extract_text(
    image_path: str,
    lang: str = "eng",
    psm: int = 3,
    oem: int = 3,
    whitelist: str = "",
    dpi: int = 300,
) -> str:
    """Extract text from an image file, URL, or base64 data URI.

    Args:
        image_path: Path to image, HTTP URL, or data: URI
        lang: Language code (eng, chi_sim, chi_tra, jpn, etc.). Multi-lang: "eng+chi_sim"
        psm: Page segmentation mode (0=auto, 1=auto+OSD, 3=auto default, 4=single column,
             6=single block, 7=single line, 8=single word, 11=sparse text, 12=sparse+OSD)
        oem: OCR engine mode (0=legacy, 1=LSTM only, 2=legacy+LSTM, 3=default)
        whitelist: Optional character whitelist (e.g. "0123456789." for numbers only)
        dpi: DPI hint for the OCR engine

    Returns JSON with extracted text, confidence, and metadata.
    """
    img = _load_image(image_path)

    config_parts = [f"--psm {psm}", "--oem" if oem else f"--oem {oem}"]
    if whitelist:
        config_parts.append(f'-c tessedit_char_whitelist={whitelist}')
    config = " ".join(config_parts)

    # Get text
    text = pytesseract.image_to_string(img, lang=lang, config=config)

    # Get confidence data
    data = pytesseract.image_to_data(
        img, lang=lang, config=config, output_type=pytesseract.Output.DICT
    )

    # Calculate average confidence (excluding -1 entries)
    confidences = [int(c) for c in data["conf"] if int(c) > 0]
    avg_conf = sum(confidences) / len(confidences) if confidences else 0

    # Word count
    words = [w for w in data["text"] if w.strip()]

    return json.dumps({
        "text": text.strip(),
        "confidence": round(avg_conf, 1),
        "word_count": len(words),
        "lang": lang,
        "psm": psm,
        "image_size": f"{img.width}x{img.height}",
    }, indent=2)


@mcp.tool()
@_ocr_tool
def extract_text_region(
    image_path: str,
    left: int,
    top: int,
    right: int,
    bottom: int,
    lang: str = "eng",
    psm: int = 6,
) -> str:
    """Extract text from a specific rectangular region of an image.

    Useful for extracting text from a specific UI element, table cell, or screen region.

    Args:
        image_path: Path to image, URL, or data: URI
        left, top, right, bottom: Bounding box coordinates in pixels
        lang: Language code
        psm: Page segmentation mode (6=single block is default for regions)
    """
    img = _load_image(image_path)
    region = img.crop((left, top, right, bottom))
    config = f"--psm {psm}"
    text = pytesseract.image_to_string(region, lang=lang, config=config)

    return json.dumps({
        "text": text.strip(),
        "region": {"left": left, "top": top, "right": right, "bottom": bottom},
        "region_size": f"{right-left}x{bottom-top}",
        "lang": lang,
    }, indent=2)


@mcp.tool()
@_ocr_tool
def extract_tables(
    image_path: str,
    lang: str = "eng",
) -> str:
    """Extract tabular data from an image. Returns structured row/column data.

    Uses Tesseract's TSV output to reconstruct table-like structures.
    Best for clean, well-separated tables with visible grid lines.
    """
    img = _load_image(image_path)
    config = "--psm 6"  # single block for tables

    data = pytesseract.image_to_data(
        img, lang=lang, config=config, output_type=pytesseract.Output.DICT
    )

    # Group words by row (top position) and column (left position)
    rows = {}
    for i, word in enumerate(data["text"]):
        word = word.strip()
        if not word or int(data["conf"][i]) < 30:
            continue
        top_key = round(data["top"][i] / 10) * 10  # cluster nearby rows
        left_val = data["left"][i]
        conf = int(data["conf"][i])

        if top_key not in rows:
            rows[top_key] = []
        rows[top_key].append({
            "text": word,
            "left": left_val,
            "conf": conf,
        })

    # Sort rows by position, build row arrays
    table = []
    for row_top in sorted(rows.keys()):
        words_in_row = sorted(rows[row_top], key=lambda w: w["left"])
        # Cluster by columns (gap detection)
        columns = []
        current_col = [words_in_row[0]]
        for w in words_in_row[1:]:
            if w["left"] - current_col[-1]["left"] > 50:  # new column
                columns.append(" ".join(c["text"] for c in current_col))
                current_col = [w]
            else:
                current_col.append(w)
        columns.append(" ".join(c["text"] for c in current_col))
        table.append(columns)

    return json.dumps({
        "rows": len(table),
        "columns": max(len(r) for r in table) if table else 0,
        "table": table,
    }, indent=2)


@mcp.tool()
@_ocr_tool
def get_available_languages() -> str:
    """List all installed Tesseract language packs."""
    langs = pytesseract.get_languages(config="")
    return json.dumps({
        "languages": langs,
        "count": len(langs),
        "common_combos": {
            "English + Simplified Chinese": "eng+chi_sim",
            "English + Traditional Chinese": "eng+chi_tra",
            "English + Japanese": "eng+jpn",
            "English + Korean": "eng+kor",
        },
    }, indent=2)


@mcp.tool()
@_ocr_tool
def screen_ocr(
    lang: str = "eng",
    psm: int = 3,
    monitor: int = 0,
) -> str:
    """Capture the current screen and run OCR on it.

    Useful for reading text from applications that don't expose accessibility APIs.

    Args:
        lang: Language code
        psm: Page segmentation mode
        monitor: Monitor index (0=primary)
    """
    try:
        import mss
    except ImportError:
        return json.dumps({"error": "mss package not installed. Run: pip install mss"})

    with mss.mss() as sct:
        monitor_info = sct.monitors[monitor + 1] if monitor < len(sct.monitors) else sct.monitors[1]
        screenshot = sct.grab(monitor_info)
        img = Image.frombytes("RGB", screenshot.size, screenshot.bgra, "raw", "BGRX")

    config = f"--psm {psm}"
    text = pytesseract.image_to_string(img, lang=lang, config=config)

    data = pytesseract.image_to_data(
        img, lang=lang, config=config, output_type=pytesseract.Output.DICT
    )
    confidences = [int(c) for c in data["conf"] if int(c) > 0]
    avg_conf = sum(confidences) / len(confidences) if confidences else 0

    return json.dumps({
        "text": text.strip(),
        "confidence": round(avg_conf, 1),
        "screen_size": f"{img.width}x{img.height}",
        "lang": lang,
    }, indent=2)


@mcp.tool()
@_ocr_tool
def extract_text_from_pdf(
    pdf_path: str,
    lang: str = "eng",
    dpi: int = 300,
    max_pages: int = 10,
) -> str:
    """Extract text from a PDF file by rendering pages to images then OCR.

    Args:
        pdf_path: Path to PDF file
        lang: Language code
        dpi: DPI for rendering (higher = better quality but slower)
        max_pages: Maximum pages to process
    """
    try:
        from pdf2image import convert_from_path
    except ImportError:
        return json.dumps({
            "error": "pdf2image not installed. Run: pip install pdf2image"
        })

    if not os.path.exists(pdf_path):
        return json.dumps({"error": f"PDF not found: {pdf_path}"})

    pages = convert_from_path(pdf_path, dpi=dpi, first_page=1, last_page=max_pages)

    results = []
    for i, page_img in enumerate(pages, 1):
        text = pytesseract.image_to_string(page_img, lang=lang)
        results.append({
            "page": i,
            "text": text.strip(),
            "char_count": len(text.strip()),
        })

    return json.dumps({
        "pdf": pdf_path,
        "pages_processed": len(results),
        "pages": results,
    }, indent=2)


# ─── Run ──────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    mcp.run(transport="stdio")
