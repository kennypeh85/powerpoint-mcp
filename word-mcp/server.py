"""
Microsoft Word MCP Server
=========================
Attaches to a *running* Microsoft Word instance via COM (pywin32) and
exposes read/edit/screenshot tools over stdio MCP so Hermes Agent can edit the
document live on screen.

Requirements
------------
- Windows + Microsoft Word installed.
- pywin32 (py -m pip install pywin32)  -- run `python Scripts/pywin32_postinstall.py -install` once.
- mcp  (py -m pip install mcp)

Register in ~/.hermes/config.yaml:
    mcp_servers:
      word:
        command: "C:/Users/Kenny Peh/AppData/Local/hermes/hermes-agent/venv/Scripts/python.exe"
        args: ["C:/Users/Kenny Peh/OneDrive/Documents/Claude/Projects/office-tools.ai/windows-mcp-servers/word-mcp/server.py"]
        connect_timeout: 30
        timeout: 90

Notes
-----
- Word COM is 1-indexed. Paragraph/Table indices in this server are 1-indexed
  to match what the user sees on screen.
- COM objects are apartment-threaded. Each tool call re-initialises COM and
  re-dispatches Word.Application, which transparently reconnects to the
  already-running instance.
- All tools return JSON-serialisable dicts.
"""

from __future__ import annotations

import base64
import json
import os
import sys
import tempfile
import traceback
from contextlib import contextmanager
from functools import wraps
from typing import Any

import pythoncom
import win32com.client
from mcp.server.fastmcp import FastMCP


# Word-specific enums (common values used across Word versions).
wdExportFormatPDF = 17
wdExportOptimizeForPrint = 0

wdParagraphAlignment = {
    "left": 0,           # wdAlignParagraphLeft
    "center": 1,         # wdAlignParagraphCenter
    "right": 2,          # wdAlignParagraphRight
    "justify": 3,        # wdAlignParagraphJustify
}

wdLineStyle = {
    "single": 1,         # wdLineStyleSingle
}

wdUnits = {
    "paragraph": 4,      # wdParagraph
    "word": 2,           # wdWord
    "character": 1,      # wdCharacter
    "line": 5,           # wdLine
    "story": 6,          # wdStory
}

wdGoTo = {
    "page": 1,           # wdGoToPage
    "section": 0,        # wdGoToSection
}

wdGoToDirection = {
    "start": 1,          # wdGoToStart
    "end": 2,            # wdGoToEnd
}


def word_rgb(r: int, g: int, b: int) -> int:
    """Convert hex RGB components to Word COM color value.

    COM stores color as: R + G*256 + B*65536  (NOT B + G*256 + R*65536!).
    Getting this wrong swaps red and blue channels — peach becomes light blue.
    Always use this helper instead of manual hex arithmetic.
    """
    return r + g * 256 + b * 65536


@contextmanager
def com_session():
    """Initialise COM for the current thread and yield a fresh app handle."""
    pythoncom.CoInitialize()
    try:
        # Dispatch returns the running instance if Word is already open,
        # otherwise it launches Word (visible if Interactive).
        app = win32com.client.Dispatch("Word.Application")
        try:
            yield app
        finally:
            # Do NOT quit the app — the user owns the Word process.
            del app
    finally:
        pythoncom.CoUninitialize()


def word_tool(fn):
    """Decorator: run a tool body inside a COM session with robust error capture."""
    @wraps(fn)
    def wrapper(*args, **kwargs):
        try:
            with com_session() as app:
                return fn(app, *args, **kwargs)
        except Exception as e:
            tb = traceback.format_exc()
            return {"error": f"{type(e).__name__}: {e}", "traceback": tb}
    return wrapper


def _resolve_doc(app, name: str | None):
    """Return the requested Document. If name is None, return ActiveDocument."""
    if name:
        for i in range(1, app.Documents.Count + 1):
            doc = app.Documents.Item(i)
            # Match by Name (no extension) or full Path.
            if doc.Name == name or doc.Name == name + ".docx" or doc.Path and doc.FullName == name:
                return doc
        raise ValueError(f"No open document matching '{name}'.")
    doc = app.ActiveDocument
    if doc is None:
        raise RuntimeError("No active document. Open a document in Word first.")
    return doc


def _element_summary(el) -> dict[str, Any]:
    """Extract a compact, JSON-safe description of a document element."""
    info: dict[str, Any] = {
        "index": None,  # filled by caller
        "name": getattr(el, "Name", "unnamed"),
        "type": str(type(el).__name__),
    }
    try:
        # Paragraph-specific
        if hasattr(el, "Range"):
            r = el.Range
            info["start"] = r.Start
            info["end"] = r.End
            info["length"] = r.Length
            info["text"] = r.Text.replace("\r", "\n").rstrip("\n")
        # Table-specific
        if hasattr(el, "Rows"):
            info["rows"] = el.Rows.Count
            info["columns"] = el.Columns.Count
        # Section-specific
        if hasattr(el, "Start") and hasattr(el, "End"):
            info["start"] = el.Start
            info["end"] = el.End
    except Exception:
        pass
    return info


# ---------------------------------------------------------------------------
# MCP server + tools
# ---------------------------------------------------------------------------

mcp = FastMCP("word-live")


# --- Session / document discovery -----------------------------------------

@mcp.tool()
@word_tool
def list_documents(app) -> str:
    """List all documents currently open in Word.

    Returns name, full path, paragraph count, and whether each is the active document.
    """
    docs = []
    active = None
    try:
        active = app.ActiveDocument
    except Exception:
        active = None
    active_name = active.Name if active else None
    for i in range(1, app.Documents.Count + 1):
        doc = app.Documents.Item(i)
        docs.append({
            "name": doc.Name,
            "path": doc.FullName if doc.Path else "(unsaved)",
            "paragraph_count": doc.Paragraphs.Count,
            "active": doc.Name == active_name,
        })
    return json.dumps({"documents": docs, "active": active_name, "count": len(docs)})


@mcp.tool()
@word_tool
def get_active_document(app) -> str:
    """Return details of the currently active document."""
    doc = app.ActiveDocument
    if doc is None:
        return json.dumps({"error": "No active document."})
    return json.dumps({
        "name": doc.Name,
        "path": doc.FullName if doc.Path else "(unsaved)",
        "paragraph_count": doc.Paragraphs.Count,
        "word_count": len(doc.Words),
    })


@mcp.tool()
@word_tool
def create_new_document(app, name: str | None = None) -> str:
    """Create a new blank document. If name is given, saves it with that name."""
    doc = app.Documents.Add()
    if name:
        doc.SaveAs2(name)
    return json.dumps({"ok": True, "document": doc.Name, "path": doc.FullName if doc.Path else "(unsaved)"})


@mcp.tool()
@word_tool
def save_document(app, file_path: str, name: str | None = None) -> str:
    """Save the active document (or specified document) to a file path."""
    doc = _resolve_doc(app, name)
    doc.SaveAs2(file_path)
    return json.dumps({"ok": True, "document": doc.Name, "path": file_path})


@mcp.tool()
@word_tool
def save_as_pdf(app, file_path: str, name: str | None = None) -> str:
    """Save the active document (or specified document) as a PDF."""
    doc = _resolve_doc(app, name)
    doc.ExportAsFixedFormat(
        OutputFileName=file_path,
        ExportFormat=wdExportFormatPDF,
        OptimizeFor=wdExportOptimizeForPrint,
        OpenAfterExport=False,
    )
    return json.dumps({"ok": True, "document": doc.Name, "path": file_path})


# --- Read -----------------------------------------------------------------

@mcp.tool()
@word_tool
def get_document_content(app, name: str | None = None) -> str:
    """Return all text in the active document."""
    doc = _resolve_doc(app, name)
    full_text = ""
    try:
        full_text = doc.Content.Text.replace("\r", "\n").rstrip("\n")
    except Exception:
        pass
    return json.dumps({
        "document": doc.Name,
        "paragraph_count": doc.Paragraphs.Count,
        "word_count": len(doc.Words),
        "text": full_text,
    })


@mcp.tool()
@word_tool
def get_paragraphs(app, start: int = 1, end: int | None = None, name: str | None = None) -> str:
    """Return a summary of paragraphs in the active document (or specified document).

    start and end are 1-indexed paragraph indices. If end is None, returns all paragraphs from start.
    """
    doc = _resolve_doc(app, name)
    if not (1 <= start <= doc.Paragraphs.Count):
        return json.dumps({"error": f"start paragraph {start} out of range (1..{doc.Paragraphs.Count})"})
    if end is None:
        end = doc.Paragraphs.Count
    elif not (1 <= end <= doc.Paragraphs.Count):
        return json.dumps({"error": f"end paragraph {end} out of range (1..{doc.Paragraphs.Count})"})
    paragraphs = []
    for i in range(start, end + 1):
        p = doc.Paragraphs.Item(i)
        info = _element_summary(p)
        info["index"] = i
        paragraphs.append(info)
    return json.dumps({
        "document": doc.Name,
        "start": start,
        "end": end,
        "paragraph_count": len(paragraphs),
        "paragraphs": paragraphs,
    }, indent=2)


@mcp.tool()
@word_tool
def get_sections(app, name: str | None = None) -> str:
    """Return all sections in the active document."""
    doc = _resolve_doc(app, name)
    sections = []
    for i in range(1, doc.Sections.Count + 1):
        sec = doc.Sections.Item(i)
        info = _element_summary(sec)
        info["index"] = i
        info["page_start"] = sec.Start
        info["page_end"] = sec.End
        sections.append(info)
    return json.dumps({
        "document": doc.Name,
        "section_count": len(sections),
        "sections": sections,
    }, indent=2)


@mcp.tool()
@word_tool
def get_tables(app, start: int = 1, end: int | None = None, name: str | None = None) -> str:
    """Return all tables in the active document (or specified document).

    start and end are 1-indexed table indices. If end is None, returns all tables from start.
    """
    doc = _resolve_doc(app, name)
    if not (1 <= start <= doc.Tables.Count):
        return json.dumps({"error": f"start table {start} out of range (1..{doc.Tables.Count})"})
    if end is None:
        end = doc.Tables.Count
    elif not (1 <= end <= doc.Tables.Count):
        return json.dumps({"error": f"end table {end} out of range (1..{doc.Tables.Count})"})
    tables = []
    for i in range(start, end + 1):
        t = doc.Tables.Item(i)
        info = _element_summary(t)
        info["index"] = i
        info["cell_count"] = t.Cells.Count
        tables.append(info)
    return json.dumps({
        "document": doc.Name,
        "start": start,
        "end": end,
        "table_count": len(tables),
        "tables": tables,
    }, indent=2)


@mcp.tool()
@word_tool
def get_selection(app) -> str:
    """Describe the current selection (text or range)."""
    if not app.ActiveDocument:
        return json.dumps({"error": "No active document."})
    sel = app.Selection
    if not sel:
        return json.dumps({"error": "No selection."})
    kind = None
    try:
        kind = sel.Type  # wdSelectionNone=0, wdSelectionInlineShape=3, wdSelectionShape=2, wdSelectionIP=1
    except Exception:
        pass
    out: dict[str, Any] = {"selection_type": kind}
    try:
        if kind in (1, 3, 4, 5, 6, 7, 8):  # text or shapes
            out["text"] = sel.Text.replace("\r", "\n")
            out["start"] = sel.Start
            out["end"] = sel.End
    except Exception as e:
        out["warning"] = f"Could not read selection detail: {e}"
    return json.dumps(out)


@mcp.tool()
@word_tool
def find_text(app, find: str, replace: str | None = None, case_sensitive: bool = False,
              whole_word: bool = False, name: str | None = None) -> str:
    """Find text in the active document and optionally replace it.

    Returns information about all matches and replacements made.
    """
    doc = _resolve_doc(app, name)
    find_obj = doc.Content.Find
    find_obj.ClearFormatting()
    find_obj.Text = find
    find_obj.MatchCase = case_sensitive
    find_obj.MatchWholeWord = whole_word
    
    if replace is not None:
        find_obj.Replacement.ClearFormatting()
        find_obj.Replacement.Text = replace
    
    found = 0
    replaced = 0
    
    if replace is not None:
        # Execute replace
        success = find_obj.Execute(Replace=2)  # wdReplaceAll
        if success:
            replaced = 1
        # Note: Word's Execute returns False if no replacements found, not 0
    else:
        # Count matches (FindWrap=0 = wdFindStop, so we just count what we find)
        while find_obj.Execute():
            found += 1
            if not find_obj.Find.Found:
                break
        find_obj.Execute(Replace=0)  # Reset
    
    return json.dumps({
        "document": doc.Name,
        "find": find,
        "replace": replace,
        "found": found,
        "replaced": replaced,
    })


# --- Edit -----------------------------------------------------------------

@mcp.tool()
@word_tool
def insert_paragraph(app, text: str, after: int | None = None, before: int | None = None,
                     align: str | None = None, name: str | None = None) -> str:
    """Insert a new paragraph with the given text.

    after and before are 1-indexed paragraph indices. If both are None, appends at end.
    align is one of: left, center, right, justify.
    """
    doc = _resolve_doc(app, name)
    
    # Determine insertion point
    if after is not None:
        if not (1 <= after <= doc.Paragraphs.Count):
            return json.dumps({"error": f"after paragraph {after} out of range"})
        para = doc.Paragraphs.Item(after + 1)
        para.Range.InsertParagraphAfter()
    elif before is not None:
        if not (1 <= before <= doc.Paragraphs.Count):
            return json.dumps({"error": f"before paragraph {before} out of range"})
        para = doc.Paragraphs.Item(before)
        para.Range.InsertParagraphBefore()
    else:
        # Append at end
        para = doc.Content.End - 1
    
    # Set text and alignment
    try:
        # Insert text
        para.Range.Text = text.replace("\n", "\r")
        
        # Set alignment
        if align:
            align_val = wdParagraphAlignment.get(align.lower())
            if align_val is None:
                return json.dumps({"error": f"Invalid alignment '{align}'. Available: {sorted(wdParagraphAlignment.keys())}"})
            para.Format.Alignment = align_val
    except Exception as e:
        return json.dumps({"error": f"Failed to insert paragraph: {e}"})
    
    return json.dumps({
        "ok": True,
        "document": doc.Name,
        "paragraph_index": para.Index,
        "chars": len(text),
    })


@mcp.tool()
@word_tool
def insert_text(app, text: str, at: str = "start", name: str | None = None) -> str:
    """Insert text at a position in the active document.

    at can be: 'start', 'end', 'paragraph:<index>', 'range:<start>-<end>', 'selection'.
    Paragraph and range indices are 1-indexed.
    """
    doc = _resolve_doc(app, name)
    sel = app.Selection
    
    if at == "start":
        doc.Content.SetRange(0, 0)
    elif at == "end":
        doc.Content.SetRange(doc.Content.End, doc.Content.End)
    elif at.startswith("paragraph:"):
        idx = int(at.split(":")[1])
        if not (1 <= idx <= doc.Paragraphs.Count):
            return json.dumps({"error": f"Paragraph {idx} out of range"})
        doc.Paragraphs.Item(idx).Range.SetRange(0, 0)
    elif at.startswith("range:"):
        parts = at.split(":")[1].split("-")
        if len(parts) != 2:
            return json.dumps({"error": "Invalid range format. Use 'range:start-end' (1-indexed)"})
        start = int(parts[0]) - 1
        end = int(parts[1])
        if not (0 <= start <= doc.Content.End and 0 <= end <= doc.Content.End):
            return json.dumps({"error": "Range indices out of document bounds"})
        doc.Content.SetRange(start, end)
    elif at == "selection":
        # Already at selection
        pass
    else:
        return json.dumps({"error": f"Invalid position '{at}'. Use 'start', 'end', 'paragraph:<index>', 'range:start-end', or 'selection'"})
    
    # Insert text
    sel.TypeText(text)
    return json.dumps({
        "ok": True,
        "document": doc.Name,
        "position": at,
        "chars": len(text),
    })


@mcp.tool()
@word_tool
def replace_text(app, find: str, replace: str, case_sensitive: bool = False,
                 whole_word: bool = False, name: str | None = None) -> str:
    """Find and replace text in the active document.

    Returns information about replacements made.
    """
    doc = _resolve_doc(app, name)
    find_obj = doc.Content.Find
    find_obj.ClearFormatting()
    find_obj.Text = find
    find_obj.MatchCase = case_sensitive
    find_obj.MatchWholeWord = whole_word
    
    find_obj.Replacement.ClearFormatting()
    find_obj.Replacement.Text = replace
    
    success = find_obj.Execute(Replace=2)  # wdReplaceAll
    
    return json.dumps({
        "ok": True,
        "document": doc.Name,
        "find": find,
        "replace": replace,
        "case_sensitive": case_sensitive,
        "whole_word": whole_word,
        "replaced": 1 if success else 0,
    })


@mcp.tool()
@word_tool
def set_paragraph_alignment(app, align: str, paragraph_index: int | None = None,
                            name: str | None = None) -> str:
    """Set alignment for a paragraph.

    align is one of: left, center, right, justify.
    paragraph_index is 1-indexed. If None, sets alignment for the selection or current paragraph.
    """
    doc = _resolve_doc(app, name)
    align_val = wdParagraphAlignment.get(align.lower())
    if align_val is None:
        return json.dumps({"error": f"Invalid alignment '{align}'. Available: {sorted(wdParagraphAlignment.keys())}"})
    
    if paragraph_index is not None:
        if not (1 <= paragraph_index <= doc.Paragraphs.Count):
            return json.dumps({"error": f"Paragraph {paragraph_index} out of range"})
        doc.Paragraphs.Item(paragraph_index).Format.Alignment = align_val
    else:
        app.Selection.ParagraphFormat.Alignment = align_val
    
    return json.dumps({
        "ok": True,
        "document": doc.Name,
        "align": align,
        "paragraph_index": paragraph_index,
    })


@mcp.tool()
@word_tool
def set_font_properties(app, font_name: str | None = None, size: int | None = None,
                        bold: bool | None = None, italic: bool | None = None,
                        color_rgb: str | None = None, name: str | None = None) -> str:
    """Set font properties on the current selection or all text in the active document.

    If no selection, applies to the entire document.
    """
    doc = _resolve_doc(app, name)
    font = app.Selection.Font
    
    if font_name is not None:
        font.Name = font_name
    if size is not None:
        font.Size = size
    if bold is not None:
        font.Bold = -1 if bold else 0
    if italic is not None:
        font.Italic = -1 if italic else 0
    if color_rgb is not None:
        hex_clean = color_rgb.lstrip("#")
        font.Color.RGB = word_rgb(int(hex_clean[0:2], 16), int(hex_clean[2:4], 16), int(hex_clean[4:6], 16))
    
    return json.dumps({
        "ok": True,
        "document": doc.Name,
        "font_name": font_name,
        "size": size,
        "bold": bold,
        "italic": italic,
    })


# --- Run -------------------------------------------------------------------

if __name__ == "__main__":
    mcp.run(transport="stdio")
