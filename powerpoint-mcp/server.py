"""
PowerPoint Live MCP Server
===========================
Attaches to a *running* Microsoft PowerPoint instance via COM (pywin32) and
exposes read/edit/screenshot tools over stdio MCP so Hermes Agent can edit the
deck live on screen.

Requirements
------------
- Windows + Microsoft PowerPoint installed.
- pywin32 (py -m pip install pywin32)  -- run `python Scripts/pywin32_postinstall.py -install` once.
- mcp  (py -m pip install mcp)

Register in ~/.hermes/config.yaml:
    mcp_servers:
      powerpoint:
        command: "C:/Users/Kenny Peh/AppData/Local/hermes/hermes-agent/venv/Scripts/python.exe"
        args: ["C:/Users/Kenny Peh/OneDrive/Documents/Claude/Projects/office-tools.ai/ppt-live-mcp/server.py"]
        connect_timeout: 30
        timeout: 90

Notes
-----
- PowerPoint COM is 1-indexed. Slide/shape indices in this server are 1-indexed
  to match what the user sees on screen.
- COM objects are apartment-threaded. Each tool call re-initialises COM and
  re-dispatches PowerPoint.Application, which transparently reconnects to the
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

# PowerPoint layout enum (ppSlideLayout) — most useful subset.
# Full list: https://learn.microsoft.com/office/vba/api/powerpoint.ppslidelayout


def pptx_rgb(r: int, g: int, b: int) -> int:
    """Convert hex RGB components to PowerPoint COM color value.

    COM stores color as: R + G*256 + B*65536  (NOT B + G*256 + R*65536!).
    Getting this wrong swaps red and blue channels — peach becomes light blue.
    Always use this helper instead of manual hex arithmetic.
    """
    return r + g * 256 + b * 65536
LAYOUTS = {
    "blank": 12,            # ppLayoutBlank
    "title": 1,             # ppLayoutTitle
    "title_only": 11,       # ppLayoutTitleOnly
    "title_and_content": 1, # ambiguous alias kept for clarity; use 'content'
    "content": 2,           # ppLayoutText
    "two_content": 3,       # ppLayoutTwoColumnText
    "section_header": 2,    # ppLayoutSectionHeader is 2 in modern builds; overridden below
    "comparison": 5,        # ppLayoutComparison
    "content_with_caption": 7,  # ppLayoutContentWithCaption
    "picture_with_caption": 8,
    "title_and_vertical_text": 3,
    "vertical_title_and_text": 4,
    "table": 4,
    "chart": 6,
    "object": 7,
    "text": 1,
}
# Prefer explicit, well-known values used by modern PowerPoint 2016+:
LAYOUTS.update({
    "blank": 12,
    "title": 1,           # ppLayoutTitle (a single title slide)
    "title_only": 11,
    "content": 2,         # ppLayoutText — title + single content placeholder
    "two_content": 3,     # ppLayoutTwoColumnText
    "section_header": 33, # ppLayoutSectionHeader (PP 2010+)
    "comparison": 5,
    "content_with_caption": 7,
    "picture_with_caption": 8,
})


@contextmanager
def com_session():
    """Initialise COM for the current thread and yield a fresh app handle."""
    pythoncom.CoInitialize()
    try:
        # Dispatch returns the running instance if PowerPoint is already open,
        # otherwise it launches PowerPoint (visible if Interactive).
        app = win32com.client.Dispatch("PowerPoint.Application")
        try:
            yield app
        finally:
            # Do NOT quit the app — the user owns the PowerPoint process.
            del app
    finally:
        pythoncom.CoUninitialize()


def pptx_tool(fn):
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


def _resolve_deck(app, name: str | None):
    """Return the requested Presentation. If name is None, return ActivePresentation."""
    if name:
        for i in range(1, app.Presentations.Count + 1):
            pres = app.Presentations.Item(i)
            # Match by Name (no extension) or full Path.
            if pres.Name == name or pres.Name == name + ".pptx" or pres.Path and pres.FullName == name:
                return pres
        raise ValueError(f"No open presentation matching '{name}'.")
    pres = app.ActivePresentation
    if pres is None:
        raise RuntimeError("No active presentation. Open a deck in PowerPoint first.")
    return pres


def _shape_summary(shape) -> dict[str, Any]:
    """Extract a compact, JSON-safe description of a shape."""
    info: dict[str, Any] = {
        "index": None,  # filled by caller
        "name": shape.Name,
        "type": str(shape.Type),
        "left": round(shape.Left, 1),
        "top": round(shape.Top, 1),
        "width": round(shape.Width, 1),
        "height": round(shape.Height, 1),
    }
    try:
        if shape.HasTextFrame:
            tf = shape.TextFrame
            if tf.HasText:
                info["text"] = tf.TextRange.Text.replace("\r", "\n").rstrip("\n")
    except Exception:
        pass
    return info


# ---------------------------------------------------------------------------
# MCP server + tools
# ---------------------------------------------------------------------------

mcp = FastMCP("powerpoint-live")


# --- Session / deck discovery ---------------------------------------------

@mcp.tool()
@pptx_tool
def list_decks(app) -> str:
    """List all presentations currently open in PowerPoint.

    Returns name, full path, slide count, and whether each is the active deck.
    """
    decks = []
    active = None
    try:
        active = app.ActivePresentation
    except Exception:
        active = None
    active_name = active.Name if active else None
    for i in range(1, app.Presentations.Count + 1):
        pres = app.Presentations.Item(i)
        decks.append({
            "name": pres.Name,
            "path": pres.FullName if pres.Path else "(unsaved)",
            "slide_count": pres.Slides.Count,
            "active": pres.Name == active_name,
        })
    return json.dumps({"decks": decks, "active": active_name, "count": len(decks)})


@mcp.tool()
@pptx_tool
def get_active_deck(app) -> str:
    """Return details of the currently active presentation."""
    pres = app.ActivePresentation
    if pres is None:
        return json.dumps({"error": "No active presentation."})
    # Detect the current slide if a window is open.
    current_slide = None
    try:
        if app.ActiveWindow and app.ActiveWindow.View:
            current_slide = app.ActiveWindow.View.Slide.SlideIndex
    except Exception:
        current_slide = None
    return json.dumps({
        "name": pres.Name,
        "path": pres.FullName if pres.Path else "(unsaved)",
        "slide_count": pres.Slides.Count,
        "current_slide": current_slide,
    })


@mcp.tool()
@pptx_tool
def goto_slide(app, slide_index: int, deck: str | None = None) -> str:
    """Navigate the active window to a slide (1-indexed)."""
    pres = _resolve_deck(app, deck)
    if not (1 <= slide_index <= pres.Slides.Count):
        return json.dumps({"error": f"slide_index {slide_index} out of range (1..{pres.Slides.Count})"})
    if app.ActiveWindow:
        app.ActiveWindow.View.GotoSlide(slide_index)
    return json.dumps({"ok": True, "slide": slide_index, "deck": pres.Name})


# --- Read -----------------------------------------------------------------

@mcp.tool()
@pptx_tool
def get_slide_content(app, slide_index: int | None = None, deck: str | None = None) -> str:
    """Return all shapes on a slide with text, position, and type.

    If slide_index is None, uses the current slide in the active window.
    """
    pres = _resolve_deck(app, deck)
    if slide_index is None:
        if not app.ActiveWindow:
            return json.dumps({"error": "No active window; pass slide_index."})
        slide = app.ActiveWindow.View.Slide
        slide_index = slide.SlideIndex
    else:
        if not (1 <= slide_index <= pres.Slides.Count):
            return json.dumps({"error": f"slide_index {slide_index} out of range (1..{pres.Slides.Count})"})
        slide = pres.Slides.Item(slide_index)
    shapes = []
    for i in range(1, slide.Shapes.Count + 1):
        info = _shape_summary(slide.Shapes.Item(i))
        info["index"] = i
        shapes.append(info)
    layout = None
    try:
        layout = slide.Layout
    except Exception:
        pass
    return json.dumps({
        "deck": pres.Name,
        "slide_index": slide_index,
        "slide_id": slide.SlideID,
        "layout": layout,
        "shape_count": len(shapes),
        "shapes": shapes,
    }, indent=2)


@mcp.tool()
@pptx_tool
def get_speaker_notes(app, slide_index: int | None = None, deck: str | None = None) -> str:
    """Return the speaker notes for a slide."""
    pres = _resolve_deck(app, deck)
    if slide_index is None:
        if not app.ActiveWindow:
            return json.dumps({"error": "No active window; pass slide_index."})
        slide = app.ActiveWindow.View.Slide
        slide_index = slide.SlideIndex
    else:
        slide = pres.Slides.Item(slide_index)
    notes = ""
    try:
        notes = slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
    except Exception:
        # Fallback: scan shapes for the notes body.
        for i in range(1, slide.NotesPage.Shapes.Count + 1):
            sh = slide.NotesPage.Shapes.Item(i)
            try:
                if sh.HasTextFrame and sh.TextFrame.HasText:
                    t = sh.TextFrame.TextRange.Text
                    if t and t.strip():
                        notes = t
                        break
            except Exception:
                continue
    return json.dumps({"slide": slide_index, "notes": notes.replace("\r", "\n").rstrip()})


@mcp.tool()
@pptx_tool
def get_selection(app) -> str:
    """Describe the current selection (shape range or text)."""
    if not app.ActiveWindow:
        return json.dumps({"error": "No active window."})
    sel = app.ActiveWindow.Selection
    kind = None
    try:
        kind = sel.Type  # ppSelectionNone=0, Text=3, Shapes=2, Slide=1
    except Exception:
        pass
    out: dict[str, Any] = {"selection_type": kind}
    try:
        if kind == 3:  # text
            out["text"] = sel.TextRange.Text.replace("\r", "\n")
        elif kind == 2:  # shapes
            names = []
            for i in range(1, sel.ShapeRange.Count + 1):
                sh = sel.ShapeRange.Item(i)
                names.append({"name": sh.Name, "type": str(sh.Type)})
            out["shapes"] = names
    except Exception as e:
        out["warning"] = f"Could not read selection detail: {e}"
    return json.dumps(out)


# --- Edit -----------------------------------------------------------------

@mcp.tool()
@pptx_tool
def set_shape_text(
    app,
    slide_index: int,
    shape_name_or_index: str | int,
    text: str,
    deck: str | None = None,
    append: bool = False,
) -> str:
    """Set or append the text of an existing shape on a slide.

    shape_name_or_index may be a shape name (string) or 1-indexed shape number.
    """
    pres = _resolve_deck(app, deck)
    if not (1 <= slide_index <= pres.Slides.Count):
        return json.dumps({"error": f"slide_index out of range"})
    slide = pres.Slides.Item(slide_index)
    shape = _find_shape(slide, shape_name_or_index)
    if shape is None:
        return json.dumps({"error": f"Shape '{shape_name_or_index}' not found on slide {slide_index}"})
    if not shape.HasTextFrame:
        return json.dumps({"error": f"Shape '{shape.Name}' has no text frame"})
    tf = shape.TextFrame
    if append and tf.HasText:
        existing = tf.TextRange.Text.replace("\r", "\n")
        new_text = existing.rstrip("\n") + "\n" + text
    else:
        new_text = text
    tf.TextRange.Text = new_text.replace("\n", "\r")
    return json.dumps({
        "ok": True, "deck": pres.Name, "slide": slide_index,
        "shape": shape.Name, "chars": len(new_text),
    })


@mcp.tool()
@pptx_tool
def add_text_box(
    app,
    slide_index: int,
    text: str,
    left: float = 100.0,
    top: float = 100.0,
    width: float = 500.0,
    height: float = 50.0,
    deck: str | None = None,
    font_size: float | None = None,
    bold: bool | None = None,
) -> str:
    """Add a text box to a slide at the given position (points)."""
    pres = _resolve_deck(app, deck)
    slide = pres.Slides.Item(slide_index)
    box = slide.Shapes.AddTextbox(Orientation=1, Left=left, Top=top, Width=width, Height=height)  # 1 = horizontal
    tf = box.TextFrame
    tf.TextRange.Text = text.replace("\n", "\r")
    if font_size is not None:
        tf.TextRange.Font.Size = font_size
    if bold is not None:
        tf.TextRange.Font.Bold = -1 if bold else 0
    return json.dumps({
        "ok": True, "deck": pres.Name, "slide": slide_index,
        "shape_name": box.Name, "left": left, "top": top,
    })


@mcp.tool()
@pptx_tool
def add_slide(
    app,
    layout: str = "content",
    position: int | None = None,
    title: str | None = None,
    body: str | None = None,
    deck: str | None = None,
) -> str:
    """Add a new slide. layout is one of: blank, title, title_only, content,
    two_content, section_header, comparison, content_with_caption, picture_with_caption.
    position is 1-indexed; None = at end."""
    pres = _resolve_deck(app, deck)
    layout_id = LAYOUTS.get(layout.lower())
    if layout_id is None:
        return json.dumps({"error": f"Unknown layout '{layout}'. Available: {sorted(LAYOUTS)}"})
    if position is None:
        position = pres.Slides.Count + 1
    # PowerPoint COM: Slides.Add(Index, Layout). AddSlide() also exists but
    # takes Layout differently across versions; .Add() is the stable API.
    slide = pres.Slides.Add(Index=position, Layout=layout_id)
    # Populate placeholders if requested.
    if title is not None:
        try:
            slide.Shapes.Placeholders(1).TextFrame.TextRange.Text = title.replace("\n", "\r")
        except Exception:
            pass
    if body is not None:
        try:
            slide.Shapes.Placeholders(2).TextFrame.TextRange.Text = body.replace("\n", "\r")
        except Exception:
            pass
    # Jump to it so the user sees it.
    try:
        if app.ActiveWindow:
            app.ActiveWindow.View.GotoSlide(position)
    except Exception:
        pass
    return json.dumps({"ok": True, "deck": pres.Name, "slide_index": position, "layout": layout})


@mcp.tool()
@pptx_tool
def duplicate_slide(app, slide_index: int, deck: str | None = None) -> str:
    """Duplicate a slide; the copy is inserted immediately after the original."""
    pres = _resolve_deck(app, deck)
    slide = pres.Slides.Item(slide_index)
    new_range = slide.Duplicate()
    new_index = new_range.SlideIndex
    return json.dumps({"ok": True, "deck": pres.Name, "original": slide_index, "new_index": new_index})


@mcp.tool()
@pptx_tool
def delete_slide(app, slide_index: int, deck: str | None = None) -> str:
    """Delete a slide (1-indexed). Irreversible — warn the user first."""
    pres = _resolve_deck(app, deck)
    if pres.Slides.Count <= 1:
        return json.dumps({"error": "Cannot delete the only slide in the deck."})
    pres.Slides.Item(slide_index).Delete()
    return json.dumps({"ok": True, "deck": pres.Name, "deleted": slide_index, "remaining": pres.Slides.Count})


@mcp.tool()
@pptx_tool
def set_speaker_notes(app, slide_index: int, notes: str, deck: str | None = None, append: bool = False) -> str:
    """Set or append speaker notes for a slide."""
    pres = _resolve_deck(app, deck)
    slide = pres.Slides.Item(slide_index)
    notes_shape = slide.NotesPage.Shapes.Placeholders(2)
    if append:
        existing = notes_shape.TextFrame.TextRange.Text.replace("\r", "\n").rstrip()
        notes_shape.TextFrame.TextRange.Text = (existing + "\n" + notes).replace("\n", "\r")
    else:
        notes_shape.TextFrame.TextRange.Text = notes.replace("\n", "\r")
    return json.dumps({"ok": True, "deck": pres.Name, "slide": slide_index, "chars": len(notes)})


@mcp.tool()
@pptx_tool
def replace_text(
    app,
    find: str,
    replace: str,
    scope: str = "slide",   # "slide" | "deck"
    slide_index: int | None = None,
    match_case: bool = False,
    whole_word: bool = False,
    deck: str | None = None,
) -> str:
    """Find-and-replace text. scope='deck' replaces everywhere; 'slide' limits to one slide."""
    pres = _resolve_deck(app, deck)
    replaced = 0
    if scope == "deck":
        # Use Find on the presentation's slides via the TextRange approach.
        for i in range(1, pres.Slides.Count + 1):
            replaced += _find_replace_in_slide(pres.Slides.Item(i), find, replace, match_case, whole_word)
    else:
        if slide_index is None:
            if app.ActiveWindow:
                slide_index = app.ActiveWindow.View.Slide.SlideIndex
            else:
                return json.dumps({"error": "slide_index required when scope='slide' and no active window"})
        replaced = _find_replace_in_slide(pres.Slides.Item(slide_index), find, replace, match_case, whole_word)
    return json.dumps({"ok": True, "deck": pres.Name, "replacements": replaced, "scope": scope})


def _find_replace_in_slide(slide, find, replace, match_case, whole_word) -> int:
    """Iterate all text-bearing shapes on a slide and replace text."""
    n = 0
    for i in range(1, slide.Shapes.Count + 1):
        sh = slide.Shapes.Item(i)
        try:
            if sh.HasTextFrame and sh.TextFrame.HasText:
                tr = sh.TextFrame.TextRange
                found = tr.Find(FindWhat=find, MatchCase=match_case, WholeWords=whole_word)
                if found and found.Text:
                    n += 1
                    found.Text = replace
        except Exception:
            continue
    return n


# --- Formatting (lightweight) ---------------------------------------------

@mcp.tool()
@pptx_tool
def set_shape_fill_color(
    app,
    slide_index: int,
    shape_name_or_index: str | int,
    rgb: str,
    deck: str | None = None,
) -> str:
    """Set the fill colour of a shape. rgb is a hex string like '1E2761'."""
    pres = _resolve_deck(app, deck)
    slide = pres.Slides.Item(slide_index)
    shape = _find_shape(slide, shape_name_or_index)
    if shape is None:
        return json.dumps({"error": f"Shape '{shape_name_or_index}' not found"})
    hex_clean = rgb.lstrip("#")
    shape.Fill.ForeColor.RGB = pptx_rgb(int(hex_clean[0:2], 16), int(hex_clean[2:4], 16), int(hex_clean[4:6], 16))
    shape.Fill.Visible = -1
    return json.dumps({"ok": True, "shape": shape.Name, "rgb": hex_clean})


@mcp.tool()
@pptx_tool
def set_font(
    app,
    slide_index: int,
    shape_name_or_index: str | int,
    size: float | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    color_rgb: str | None = None,
    font_name: str | None = None,
    deck: str | None = None,
) -> str:
    """Set font properties on all text in a shape."""
    pres = _resolve_deck(app, deck)
    slide = pres.Slides.Item(slide_index)
    shape = _find_shape(slide, shape_name_or_index)
    if shape is None:
        return json.dumps({"error": f"Shape '{shape_name_or_index}' not found"})
    if not shape.HasTextFrame or not shape.TextFrame.HasText:
        return json.dumps({"error": "Shape has no text"})
    font = shape.TextFrame.TextRange.Font
    if size is not None:
        font.Size = size
    if bold is not None:
        font.Bold = -1 if bold else 0
    if italic is not None:
        font.Italic = -1 if italic else 0
    if font_name is not None:
        font.Name = font_name
    if color_rgb is not None:
        hex_clean = color_rgb.lstrip("#")
        font.Color.RGB = pptx_rgb(int(hex_clean[0:2], 16), int(hex_clean[2:4], 16), int(hex_clean[4:6], 16))
    return json.dumps({"ok": True, "shape": shape.Name})


def _find_shape(slide, name_or_index):
    """Resolve a shape. Tries in order: 1-indexed int → exact name → text
    content (exact) → text content (substring). Returns the Shape or None."""
    # 1. Numeric index
    if isinstance(name_or_index, int):
        if 1 <= name_or_index <= slide.Shapes.Count:
            return slide.Shapes.Item(name_or_index)
        return None
    key = str(name_or_index)
    # 2. Exact name match
    for i in range(1, slide.Shapes.Count + 1):
        sh = slide.Shapes.Item(i)
        if sh.Name == key:
            return sh
    # 3. Exact text-content match (users often refer to shapes by what they say)
    for i in range(1, slide.Shapes.Count + 1):
        sh = slide.Shapes.Item(i)
        try:
            if sh.HasTextFrame and sh.TextFrame.HasText:
                if sh.TextFrame.TextRange.Text.replace("\r", "\n").strip() == key.strip():
                    return sh
        except Exception:
            continue
    # 4. Substring text-content match (fallback for partial references)
    for i in range(1, slide.Shapes.Count + 1):
        sh = slide.Shapes.Item(i)
        try:
            if sh.HasTextFrame and sh.TextFrame.HasText:
                if key.lower() in sh.TextFrame.TextRange.Text.replace("\r", "\n").lower():
                    return sh
        except Exception:
            continue
    return None


# --- Visual / export -------------------------------------------------------

@mcp.tool()
@pptx_tool
def screenshot_slide(
    app,
    slide_index: int | None = None,
    width: int = 1600,
    deck: str | None = None,
    out_dir: str | None = None,
) -> str:
    """Export a slide to PNG and return the file path. Useful for visual QA.

    Saves to a temp file by default, or under out_dir if given.
    """
    pres = _resolve_deck(app, deck)
    if slide_index is None:
        if not app.ActiveWindow:
            return json.dumps({"error": "No active window; pass slide_index."})
        slide = app.ActiveWindow.View.Slide
        slide_index = slide.SlideIndex
    else:
        slide = pres.Slides.Item(slide_index)
    height = int(width * 9 / 16)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
        path = os.path.join(out_dir, f"slide-{slide_index:02d}.png")
    else:
        fd, path = tempfile.mkstemp(prefix=f"pptx_slide{slide_index:02d}_", suffix=".png")
        os.close(fd)
    slide.Export(path, "PNG", width, height)
    size_kb = round(os.path.getsize(path) / 1024, 1)
    return json.dumps({"ok": True, "path": path, "slide": slide_index, "width": width, "size_kb": size_kb})


@mcp.tool()
@pptx_tool
def screenshot_deck(
    app,
    width: int = 1280,
    deck: str | None = None,
    out_dir: str | None = None,
) -> str:
    """Export ALL slides in the deck to PNGs. Returns a list of file paths."""
    pres = _resolve_deck(app, deck)
    if out_dir is None:
        out_dir = tempfile.mkdtemp(prefix="pptx_deck_")
    else:
        os.makedirs(out_dir, exist_ok=True)
    height = int(width * 9 / 16)
    paths = []
    for i in range(1, pres.Slides.Count + 1):
        p = os.path.join(out_dir, f"slide-{i:03d}.png")
        pres.Slides.Item(i).Export(p, "PNG", width, height)
        paths.append(p)
    return json.dumps({"ok": True, "deck": pres.Name, "count": len(paths), "out_dir": out_dir, "slides": paths})


# --- Run -------------------------------------------------------------------

if __name__ == "__main__":
    mcp.run(transport="stdio")
