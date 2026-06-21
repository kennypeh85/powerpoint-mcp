"""Full end-to-end smoke test for the PowerPoint Live MCP server."""
import sys, os, json, tempfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pythoncom, win32com.client

pythoncom.CoInitialize()
app = win32com.client.Dispatch("PowerPoint.Application")
app.Visible = True

pres = app.Presentations.Add(WithWindow=True)
s1 = pres.Slides.Add(Index=1, Layout=2)
s1.Shapes.Placeholders(1).TextFrame.TextRange.Text = "Test Slide 1"
s1.Shapes.Placeholders(2).TextFrame.TextRange.Text = "Original body text\rLine two"
s2 = pres.Slides.Add(Index=2, Layout=12)
print(f"Created test deck: {pres.Name}, {pres.Slides.Count} slides")

import server

def call(fn, *args, **kwargs):
    return json.loads(fn.__wrapped__(app, *args, **kwargs))

results = {}
results["list_decks"]      = call(server.list_decks)
results["get_active_deck"] = call(server.get_active_deck)
results["get_slide_1"]     = call(server.get_slide_content, 1)
results["get_notes_1"]     = call(server.get_speaker_notes, 1)
results["set_text"]        = call(server.set_shape_text, 1, 2, "EDITED by Hermes!\rLive edit OK.", append=False)
results["add_textbox"]     = call(server.add_text_box, 2, "New textbox", left=50, top=50, width=600, height=40, font_size=18, bold=True)
results["add_slide"]       = call(server.add_slide, layout="title_only", title="Slide 3", position=3)
results["set_notes"]       = call(server.set_speaker_notes, 1, "Notes set by Hermes.", append=False)
results["replace"]         = call(server.replace_text, "EDITED", "REPLACED", scope="slide", slide_index=1)
results["set_color"]       = call(server.set_shape_fill_color, 2, "New textbox", "1E2761")
results["set_font"]        = call(server.set_font, 2, "New textbox", color_rgb="FFFFFF")
results["screenshot"]      = call(server.screenshot_slide, 1, width=1280, out_dir=tempfile.gettempdir())
results["verify_slide1"]   = call(server.get_slide_content, 1)
results["verify_notes1"]   = call(server.get_speaker_notes, 1)
results["verify_deck"]     = call(server.get_active_deck)
results["goto"]            = call(server.goto_slide, 2)

print("\n" + "="*60)
print("SMOKE TEST RESULTS")
print("="*60)
for name, r in results.items():
    ok = r.get("ok", True) if isinstance(r, dict) else True
    ok = ok and "error" not in r
    status = "PASS" if ok else "FAIL"
    short = {k: v for k, v in r.items() if k != "shapes"} if isinstance(r, dict) else r
    print(f"{status:4} {name:18} {json.dumps(short)[:110]}")

ss = results["screenshot"]
if "path" in ss:
    exists = os.path.exists(ss["path"])
    size = os.path.getsize(ss["path"]) if exists else 0
    print(f"\nScreenshot: {ss['path']} ({size} bytes, exists={exists})")

# Verify text edit actually landed
s1_body = results["verify_slide1"]["shapes"]
edited_shape = [s for s in s1_body if s.get("index") == 2]
if edited_shape:
    txt = edited_shape[0].get("text", "")
    print(f"Slide 1 body text after edit: {txt[:80]}")
    print("Edit confirmed" if "REPLACED" in txt or "EDITED" in txt else "WARN: expected edited text")

pres.Saved = True
pres.Close()
print("\nTest deck closed.")
