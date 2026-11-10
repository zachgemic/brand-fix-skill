#!/usr/bin/env python3
"""Self-test for brand_fix.py. Builds one deck with
  * OK_*  shapes that comply with the Atlas design system (must NEVER be flagged or changed), and
  * BAD_* shapes with seeded violations (must be flagged; auto-fixable ones must be fixed).
Run before trusting a report:  python selftest.py   (prints SELFTEST PASSED / FAILED)
"""
import os, subprocess, sys, tempfile
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Emu, Inches, Pt

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPT = os.path.join(HERE, "brand_fix.py")


def build(path, bad=True):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(12192000), Emu(6858000)
    s = prs.slides.add_slide(prs.slide_layouts[6])

    def tb(name, x, y, w, txt, pt, bold=False, ls=1.17, font="Inter", margins0=True):
        t = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(0.4))
        t.name = name
        tf = t.text_frame
        if margins0:
            tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        r = tf.paragraphs[0].add_run()
        r.text, r.font.size, r.font.bold, r.font.name = txt, Pt(pt), bold, font
        tf.paragraphs[0].line_spacing = ls

    def rect(name, x, y, rgb, margin=None):
        r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(0.9), Inches(0.3))
        r.name = name
        r.fill.solid(); r.fill.fore_color.rgb = RGBColor(*rgb); r.line.fill.background()
        if margin is not None:
            tf = r.text_frame
            tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Inches(margin)
            run = tf.paragraphs[0].add_run(); run.text = "LABEL"
            run.font.size, run.font.bold, run.font.name = Pt(10), True, "Inter"

    def table(name, x, y, pad_emu):
        t = s.shapes.add_table(2, 2, Inches(x), Inches(y), Inches(3), Inches(0.8))
        t.name = name
        for row in t.table.rows:
            for c in row.cells:
                c.text = "x"
                tcPr = c._tc.get_or_add_tcPr()
                for a in ("marL", "marR", "marT", "marB"):
                    tcPr.set(a, str(pad_emu))
                for p in c.text_frame.paragraphs:
                    for r in p.runs:
                        r.font.name, r.font.size = "Inter", Pt(10)

    # compliant (Atlas) -----------------------------------------------------------
    tb("OK_headline", 0.845, 0.96, 5.7, "Headline sentence", 20)          # 12-col edge 1
    tb("OK_body", 2.825, 1.96, 3.0, "Body copy", 12)                       # 12-col edge 3
    tb("OK_fivecol", 3.221, 2.6, 2.0, "Five-column text", 10)              # 5-col track 2
    tb("OK_heading", 5.597, 2.6, 2.0, "Heading", 16, bold=True)            # 5-col track 3
    tb("OK_display", 0.845, 3.4, 6.0, "Display", 38)                        # display size
    rect("OK_label", 0.845, 5.2, (0xA2, 0x4F, 0x32), margin=0.1)           # orange-700 label
    rect("OK_taupe", 2.0, 5.2, (0x84, 0x7D, 0x77))                         # taupe
    table("OK_table", 5.597, 4.6, 91440)                                   # 0.10 in padding
    if not bad:
        prs.save(path)
        return
    # seeded violations ------------------------------------------------------------
    tb("BAD_font", 8.0, 1.0, 3.0, "Arial text", 12, font="Arial")
    tb("BAD_size", 8.0, 1.5, 3.0, "24 point text", 24)
    tb("BAD_spacing", 8.0, 2.0, 3.0, "Spacing 1.0", 12, ls=1.0)
    tb("BAD_pad", 8.0, 2.6, 3.0, "Has padding", 12, margins0=False)
    rect("BAD_black", 8.0, 3.4, (0, 0, 0))                                 # pure black -> autofix
    rect("BAD_offpalette", 9.3, 3.4, (0x12, 0xAB, 0x34))                   # not in palette
    table("BAD_table", 8.0, 4.6, 182880)                                   # 0.20 in padding
    tb("BAD_grid", 1.3, 6.2, 1.5, "Off grid", 12)                          # 0.455 in from any edge: flag only
    prs.save(path)


def main():
    ok = True
    def check(name, cond, detail=""):
        nonlocal ok
        ok &= bool(cond)
        print("%-4s %s %s" % ("ok" if cond else "BAD", name, detail))

    with tempfile.TemporaryDirectory() as td:
        deck = os.path.join(td, "t.pptx")
        build(deck)
        r = subprocess.run([sys.executable, SCRIPT, deck], capture_output=True, text=True, cwd=td)
        check("analyze runs", r.returncode == 0, r.stderr[-200:])
        audit = open(os.path.join(td, "t_audit.md"), encoding="utf-8").read()
        check("no compliant shape flagged", "`OK_" not in audit and "[OK_" not in audit,
              [l for l in audit.splitlines() if "OK_" in l][:3])
        for name, why in [("BAD_font", "font"), ("BAD_size", "size"), ("BAD_spacing", "spacing"),
                          ("BAD_pad", "padding"), ("BAD_black", "pure black"), ("BAD_offpalette", "palette"),
                          ("BAD_table", "table padding"), ("BAD_grid", "grid")]:
            check("flags %s (%s)" % (name, why), name in audit)
        out = os.path.join(td, "fixed.pptx")
        r = subprocess.run([sys.executable, SCRIPT, deck, "--fix", "--out", out, "--no-annotate", "--no-before-after"],
                           capture_output=True, text=True, cwd=td)
        check("fix runs", r.returncode == 0 and os.path.exists(out), r.stderr[-200:])
        if os.path.exists(out):
            before = {sh.name: sh for sh in Presentation(deck).slides[0].shapes}
            after = {sh.name: sh for sh in Presentation(out).slides[0].shapes}
            for n in [k for k in before if k.startswith("OK_")]:
                same = (before[n].left, before[n].top) == (after[n].left, after[n].top)
                if before[n].has_text_frame and before[n].text_frame.paragraphs[0].runs:
                    same &= before[n].text_frame.paragraphs[0].line_spacing == after[n].text_frame.paragraphs[0].line_spacing
                check("fix leaves %s untouched" % n, same)
            check("BAD_font -> Inter", after["BAD_font"].text_frame.paragraphs[0].runs[0].font.name == "Inter")
            check("BAD_spacing -> 1.17", abs(after["BAD_spacing"].text_frame.paragraphs[0].line_spacing - 1.17) < 0.01)
            check("BAD_pad -> 0", after["BAD_pad"].text_frame.margin_top == 0)
            check("BAD_black -> brand black", tuple(after["BAD_black"].fill.fore_color.rgb) == (31, 31, 36))
            tc = after["BAD_table"].table.cell(0, 0)._tc.tcPr
            check("BAD_table padding -> 0.10 in", tc.get("marL") == "91440")
            check("BAD_grid not auto-moved (too far to snap)", after["BAD_grid"].left == before["BAD_grid"].left)
            check("BAD_size still flagged-only (not auto-changed)",
                  after["BAD_size"].text_frame.paragraphs[0].runs[0].font.size == Pt(24))
    print("\nSELFTEST", "PASSED" if ok else "FAILED")
    return ok


if __name__ == "__main__":
    sys.exit(0 if main() else 1)
