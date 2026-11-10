#!/usr/bin/env python3
"""
brand_fix.py — Gemic Brand Compliance Tool
==========================================
Automates low-level design cleanup on PowerPoint files built from the Gemic
Atlas template. Frees designers from mechanical checking so they can focus on
high-impact slide work.

SAFE: Never modifies text content, data, or slide meaning.
      Only touches: font names, text box margins, fill colors, line spacing.

USAGE:
    python brand_fix.py deck.pptx              # analyze only (default)
    python brand_fix.py deck.pptx --fix        # apply auto-fixes
    python brand_fix.py deck.pptx --fix --out cleaned.pptx

OUTPUT:
    [filename]_audit.md     — full report (always produced)
    [filename]_fixed.pptx   — corrected file (only with --fix)
"""

import sys
import os
import math
import shutil
import argparse
import copy
from datetime import datetime
from collections import defaultdict

from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn


# =============================================================================
# SECTION 1: BRAND CONFIGURATION
# All spec values live here. Update this section when the spec changes.
# =============================================================================

BRAND_FONT = "Inter"

# Atlas rule (typography tokens): Inter Regular 400 and Inter Bold 700 only.
# PowerPoint can only represent weight as a binary bold flag, so Bold == bold=True.
# Any named weight heavier than Regular (Medium, SemiBold, Bold, ExtraBold, Black)
# collapses to Bold. Anything lighter (Thin, ExtraLight, Light) collapses to Regular.
# Italic is reserved for new terms (see Atlas punctuation/emphasis rules); it is not
# checked here, only preserved.
INTER_WEIGHT_TO_BOLD = {
    "thin": False, "extralight": False, "light": False, "regular": False,
    "medium": True, "semibold": True, "bold": True,
    "extrabold": True, "black": True,
}


def resolve_inter_variant(font_name):
    """
    Parse an Inter-family font name into the two booleans PowerPoint can
    represent (bold, italic). Returns (is_inter_family, bold, italic).
    is_inter_family is False if font_name isn't Inter-derived at all —
    the caller falls back to a plain "not brand font" correction for those.
    """
    if not font_name or not font_name.lower().startswith("inter"):
        return False, False, False

    lower = font_name.lower()
    italic = "italic" in lower

    bold = False
    for weight_word, is_bold in INTER_WEIGHT_TO_BOLD.items():
        if weight_word in lower:
            bold = is_bold
            break
    # Bare "Inter" or "Inter Italic" with no weight word -> Regular (bold=False)

    return True, bold, italic

# Brand color palette — (R, G, B), name.
# SOURCE OF TRUTH: Atlas Design System tokens/colors.css (all 81 color tokens: core, extended,
# every 100–900 ramp step, and the Figma utility colors). Taupe is #847D77; #615F61 is "Grey".
# Regenerate from tokens/colors.css when the design system changes.
BRAND_COLORS = [
    (( 31,  31,  36), "Black"             ),  # #1F1F24
    ((231, 231, 235), "Black 100"         ),  # #E7E7EB
    ((194, 196, 205), "Black 300"         ),  # #C2C4CD
    ((134, 137, 154), "Black 500"         ),  # #86899A
    (( 80,  83,  97), "Black 700"         ),  # #505361
    (( 50,  52,  60), "Black 900"         ),  # #32343C
    ((253, 249, 246), "Light Beige"       ),  # #FDF9F6
    ((240, 221, 206), "Light Beige 100"   ),  # #F0DDCE
    ((223, 181, 149), "Light Beige 300"   ),  # #DFB595
    ((191, 114,  57), "Light Beige 500"   ),  # #BF7239
    (( 95,  57,  29), "Light Beige 700"   ),  # #5F391D
    (( 38,  22,  11), "Light Beige 900"   ),  # #26160B
    ((251, 243, 238), "Beige"             ),  # #FBF3EE
    ((242, 215, 198), "Beige 100"         ),  # #F2D7C6
    ((228, 173, 139), "Beige 300"         ),  # #E4AD8B
    ((198, 105,  47), "Beige 500"         ),  # #C6692F
    (( 99,  52,  23), "Beige 700"         ),  # #633417
    (( 40,  21,   9), "Beige 900"         ),  # #281509
    ((223, 211, 203), "Dark Beige"        ),  # #DFD3CB
    ((207, 187, 177), "Dark Beige 100"    ),  # #CFBBB1
    ((182, 153, 137), "Dark Beige 300"    ),  # #B69989
    (( 66,  49,  41), "Dark Beige 600"    ),  # #423129
    (( 26,  20,  16), "Dark Beige 900"    ),  # #1A1410
    ((132, 125, 119), "Taupe"             ),  # #847D77
    ((236, 234, 232), "Taupe 100"         ),  # #ECEAE8
    ((216, 213, 210), "Taupe 300"         ),  # #D8D5D2
    ((197, 193, 189), "Taupe 500"         ),  # #C5C1BD
    ((121, 114, 106), "Taupe 700"         ),  # #79726A
    (( 80,  76,  71), "Taupe 900"         ),  # #504C47
    (( 46, 103, 157), "Blue"              ),  # #2E679D
    ((208, 225, 241), "Blue 100"          ),  # #D0E1F1
    ((160, 195, 227), "Blue 300"          ),  # #A0C3E3
    ((113, 164, 213), "Blue 500"          ),  # #71A4D5
    (( 35,  77, 118), "Blue 700"          ),  # #234D76
    (( 23,  52,  78), "Blue 900"          ),  # #17344E
    ((213, 108,  74), "Orange"            ),  # #D56C4A
    ((242, 226, 219), "Orange 100"        ),  # #F2E2DB
    ((231, 198, 184), "Orange 300"        ),  # #E7C6B8
    ((220, 169, 149), "Orange 500"        ),  # #DCA995
    ((162,  79,  50), "Orange 700"        ),  # #A24F32
    ((108,  53,  32), "Orange 900"        ),  # #6C3520
    ((228, 127, 142), "Pink"              ),  # #E47F8E
    ((246, 230, 232), "Pink 100"          ),  # #F6E6E8
    ((237, 204, 209), "Pink 300"          ),  # #EDCCD1
    ((229, 180, 187), "Pink 500"          ),  # #E5B4BB
    ((195,  68,  80), "Pink 700"          ),  # #C34450
    ((133,  41,  51), "Pink 900"          ),  # #852933
    ((223, 170,  71), "Yellow"            ),  # #DFAA47
    ((246, 238, 220), "Yellow 100"        ),  # #F6EEDC
    ((239, 222, 186), "Yellow 300"        ),  # #EFDEBA
    ((230, 205, 152), "Yellow 500"        ),  # #E6CD98
    ((179, 135,  55), "Yellow 700"        ),  # #B38737
    ((120,  90,  36), "Yellow 900"        ),  # #785A24
    (( 45, 148, 153), "Teal"              ),  # #2D9499
    ((212, 238, 240), "Teal 100"          ),  # #D4EEF0
    ((171, 221, 226), "Teal 300"          ),  # #ABDDE2
    ((133, 205, 211), "Teal 500"          ),  # #85CDD3
    (( 56, 108, 113), "Teal 700"          ),  # #386C71
    (( 37,  73,  75), "Teal 900"          ),  # #25494B
    ((179, 136, 194), "Purple"            ),  # #B388C2
    ((238, 231, 241), "Purple 100"        ),  # #EEE7F1
    ((221, 208, 228), "Purple 300"        ),  # #DDD0E4
    ((205, 184, 216), "Purple 500"        ),  # #CDB8D8
    ((135,  86, 159), "Purple 700"        ),  # #87569F
    (( 89,  57, 106), "Purple 900"        ),  # #59396A
    (( 80, 186, 139), "Green"             ),  # #50BA8B
    ((223, 240, 232), "Green 100"         ),  # #DFF0E8
    ((192, 226, 209), "Green 300"         ),  # #C0E2D1
    ((163, 211, 186), "Green 500"         ),  # #A3D3BA
    (( 80, 140, 106), "Green 700"         ),  # #508C6A
    (( 53,  94,  71), "Green 900"         ),  # #355E47
    ((217, 104, 104), "Red"               ),  # #D96868
    ((247, 224, 224), "Red 100"           ),  # #F7E0E0
    ((239, 194, 194), "Red 300"           ),  # #EFC2C2
    ((232, 164, 163), "Red 500"           ),  # #E8A4A3
    ((192,  49,  48), "Red 700"           ),  # #C03130
    ((128,  32,  32), "Red 900"           ),  # #802020
    (( 97,  95,  97), "Grey"              ),  # #615F61
    ((153, 153, 161), "Light Grey"        ),  # #9999A1
    ((240, 230, 223), "Beige Medium"      ),  # #F0E6DF
    ((167, 194, 208), "Blue Light"        ),  # #A7C2D0
]

# Colors to auto-remap (paste artifacts)
COLOR_AUTOFIX = {
    (0,   0,   0):   (31,  31,  36),   # pure black → brand Black
    (255, 255, 255): (253, 249, 246),   # pure white → Light Beige
}

COLOR_TOLERANCE = 15        # Euclidean RGB distance threshold for "on-brand"

# Atlas layout grid — SOURCE OF TRUTH: engine/layout-rules.json (inches are authoritative).
# 12 columns x 0.75 in, 11 gutters x 0.24 in (11.64 in wide), margins 0.845 in left/right,
# content top 0.96 in, content bottom 6.79 in (0.71 in bottom margin).
EMU_PER_IN = 914_400
SLIDE_W_EMU = 12_192_000
SLIDE_H_EMU = 6_858_000
MARGIN_LEFT_EMU  = round(0.845 * EMU_PER_IN)   # 772,668
MARGIN_RIGHT_EMU = round(0.845 * EMU_PER_IN)
MARGIN_TOP_EMU   = round(0.96  * EMU_PER_IN)   # content top (headline starts here)
MARGIN_BOT_EMU   = round(0.71  * EMU_PER_IN)   # 649,224
COLUMN_W_EMU     = 685_800                      # 0.75 in
GUTTER_EMU       = 219_456                      # 0.24 in
FIVE_COL_W_EMU   = round(2.136 * EMU_PER_IN)    # five-column track
FOOTER_Y_EMU     = round(6.79 * EMU_PER_IN)     # content bottom: footer band is 6.79–7.5 in

# 12-column left edges plus the five-column track left edges (both are valid Atlas grids)
COLUMN_EDGES_12_EMU = [
    MARGIN_LEFT_EMU + i * (COLUMN_W_EMU + GUTTER_EMU) for i in range(12)
]
COLUMN_EDGES_5_EMU = [
    MARGIN_LEFT_EMU + i * (FIVE_COL_W_EMU + GUTTER_EMU) for i in range(5)
]
COLUMN_EDGES_EMU = sorted(set(COLUMN_EDGES_12_EMU) | set(COLUMN_EDGES_5_EMU))
COLUMN_EDGE_LABELS = {}
for _i, _e in enumerate(COLUMN_EDGES_12_EMU):
    COLUMN_EDGE_LABELS[_e] = "12-col edge %d" % (_i + 1)
for _i, _e in enumerate(COLUMN_EDGES_5_EMU):
    COLUMN_EDGE_LABELS[_e] = (COLUMN_EDGE_LABELS[_e] + " / 5-col track %d" % (_i + 1)
                              if _e in COLUMN_EDGE_LABELS else "5-col track %d" % (_i + 1))
COLUMN_EDGES_PX = [e / 914_400 * 96 for e in COLUMN_EDGES_EMU]

# Typography — SOURCE OF TRUTH: Atlas tokens/typography.css and layout-rules.json "roles".
# Roles (pt): overtitle 8 · body small 10 · body 12 · subhead 12 Bold · heading 16 Bold ·
# slide title 20 · display 38 (cover, section divider, closing). Labels are body-small Bold.
ATLAS_TYPE_SIZES      = (8, 10, 12, 16, 20, 38)
TYPE_SIZE_MINIMUM_PT  = 8.0     # smallest Atlas role (overtitle)
TYPE_HEADLINE_PT      = 20.0    # slide title / headline
TYPE_DISPLAY_PT       = 38.0    # display size; slides titled at this size are "statement" slides
LINE_SPACING          = 1.17    # Multiple 1.17 for all slide text
# Atlas has 8 distinct (size, weight) styles in total: 8R, 10R, 10B (label), 12R, 12B, 16B,
# 20R, 38R. The design system sets no per-slide limit, so the per-slide cap equals the
# whole system.
MAX_STYLES_PER_SLIDE  = 8
MAX_STYLES_PER_DECK   = 8

GRID_SNAP_THRESHOLD_EMU  = int(GUTTER_EMU * 0.5)  # ~12px — auto-snap only this close
GRID_NOISE_FLOOR_EMU = 47_625  # ~5px — closer than this is pixel-rounding noise, don't flag
# Shape types safe to auto-snap (text containers, not diagram elements)
SNAPPABLE_SHAPE_TYPES = {
    MSO_SHAPE_TYPE.TEXT_BOX,          # 17
    MSO_SHAPE_TYPE.PLACEHOLDER,       # 14
}
# Grid alignment is only meaningful for genuine content shapes. Decorative
# and diagram pieces (dots, chevrons, triangles, connectors, icon groups)
# are positioned relative to each other or to nearby text, not to the
# slide's column grid — checking them against the grid produces noise the
# designer never asked for. TABLE is checked via its GraphicFrame shape.
GRID_CHECKABLE_TYPES = {
    MSO_SHAPE_TYPE.TEXT_BOX,
    MSO_SHAPE_TYPE.PLACEHOLDER,
    MSO_SHAPE_TYPE.PICTURE,
    MSO_SHAPE_TYPE.TABLE,
}
TABLE_CELL_MARGIN_EMU = 91_440   # 0.10 in (Atlas table rule 16: cell padding 0.1 in)

# Rule top padding — Atlas rule 5: a rule sits exactly on the top edge of its text box and the
# box has 0.15 in top padding (0.20 in for display-size text).
DIVIDER_PADDING_IN         = 0.15
DIVIDER_PADDING_DISPLAY_IN = 0.20
DIVIDER_PADDING_TOL_IN     = 0.01   # Atlas position tolerance is 0.005 in; allow PowerPoint rounding

# Shape name fragments that mark protected brand assets
PROTECTED_NAME_FRAGMENTS = ("logo", "gemic", "Logo", "Gemic")

# System placeholder types to skip entirely
SYSTEM_PLACEHOLDER_TYPES = {
    PP_PLACEHOLDER.SLIDE_NUMBER,
    PP_PLACEHOLDER.DATE,
    PP_PLACEHOLDER.FOOTER,
}

# Full-width threshold — shapes this wide or wider are backgrounds/banners
FULL_WIDTH_THRESHOLD_EMU = int(SLIDE_W_EMU * 0.85)


# =============================================================================
# SECTION 2: UTILITIES
# =============================================================================

def rgb_distance(c1, c2):
    """Euclidean distance between two (R,G,B) tuples."""
    return math.sqrt(sum((a - b) ** 2 for a, b in zip(c1, c2)))


def nearest_brand_color(rgb_tuple):
    """Return (distance, name, rgb) of the closest brand color."""
    best = min(BRAND_COLORS, key=lambda bc: rgb_distance(rgb_tuple, bc[0]))
    return rgb_distance(rgb_tuple, best[0]), best[1], best[0]


def has_visible_fill(shape):
    """
    True if the shape has a real fill color (solid, gradient, pattern, picture).
    False for no-fill (None) and BACKGROUND (explicitly transparent).
    Padding only makes sense when this returns True.
    """
    try:
        fill = shape.fill
        if fill.type is None:
            return False
        # MSO_FILL.BACKGROUND (5) = transparent / no-fill
        from pptx.enum.dml import MSO_FILL
        if fill.type == MSO_FILL.BACKGROUND:
            return False
        return True
    except Exception:
        return False


def get_fill_rgb(shape):
    """
    Return the (R,G,B) tuple of a shape's solid fill, or None if
    unavailable (gradient, theme color, no fill, etc.).
    """
    try:
        fill = shape.fill
        if fill.type is None:
            return None
        rgb = fill.fore_color.rgb
        return (rgb[0], rgb[1], rgb[2])  # RGBColor is a tuple; it has no .red/.green/.blue
    except Exception:
        return None


def get_line_rgb(shape):
    """
    Return (R,G,B) of a shape's outline color, or None if no line is defined.

    CRITICAL: reads the raw XML directly rather than via shape.line.color.
    python-pptx's LineFormat.color property has a side effect — accessing it
    on a shape with no existing <a:ln> element silently INSERTS an empty
    <a:ln><a:solidFill/></a:ln> into the XML as part of the lazy-create
    pattern used for write access. That empty element then renders as a
    visible default stroke in PowerPoint/LibreOffice. Reading via find()
    is side-effect-free.
    """
    try:
        spPr = shape._element.spPr
        if spPr is None:
            return None
        a_ns = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
        ln = spPr.find(f'{a_ns}ln')
        if ln is None:
            return None  # no line defined at all — nothing to check
        srgb = ln.find(f'.//{a_ns}srgbClr')
        if srgb is None:
            return None  # line exists but no explicit RGB (theme color, or noFill)
        val = srgb.get('val')
        if val is None or len(val) != 6:
            return None
        return tuple(int(val[i:i+2], 16) for i in (0, 2, 4))
    except Exception:
        return None


def is_system_placeholder(shape):
    """True if shape is a slide-number, date, or footer placeholder."""
    try:
        ph = shape.placeholder_format
        if ph is not None and ph.type in SYSTEM_PLACEHOLDER_TYPES:
            return True
    except Exception:
        pass
    return False


def is_protected(shape):
    """True if shape should never be touched (logo, brand mark, etc.)."""
    name = shape.name or ""
    return any(frag in name for frag in PROTECTED_NAME_FRAGMENTS)


def is_title_shape(shape):
    """True if shape is the slide title placeholder."""
    try:
        ph = shape.placeholder_format
        if ph is not None and ph.type in (PP_PLACEHOLDER.TITLE,
                                           PP_PLACEHOLDER.CENTER_TITLE):
            return True
    except Exception:
        pass
    name = (shape.name or "").lower()
    return "title" in name and "subtitle" not in name


def is_full_width(shape):
    """True if shape spans nearly the full slide (background/banner)."""
    try:
        return shape.width >= FULL_WIDTH_THRESHOLD_EMU
    except Exception:
        return False


def is_in_footer_zone(shape):
    """True if shape sits in the bottom footer area."""
    try:
        return shape.top >= FOOTER_Y_EMU
    except Exception:
        return False


def find_divider_lines(slide):
    """
    Return a list of (left, top, width, height) for every line/connector
    shape on the slide. Used to detect when a text box's top padding is
    being used deliberately to create the gap below a divider line —
    see has_adjacent_divider_line().
    """
    lines = []
    for shape in slide.shapes:
        try:
            is_line = (shape.shape_type == MSO_SHAPE_TYPE.LINE
                      or "Connector" in (shape.name or ""))
            if is_line:
                lines.append((shape.left, shape.top, shape.width, shape.height))
        except Exception:
            continue
    return lines


def has_adjacent_divider_line(shape, divider_lines, tolerance_emu=45_720):
    """
    True if a divider line sits at (approximately) the same top edge as
    this text box, with horizontal overlap. This is the deck's own pattern
    for '0.2in divider line above text': the text box's top edge is left
    flush with the line, and the box's internal top padding pushes the
    actual text down to create the visual gap. Zeroing padding on such a
    box would collapse that gap to nothing — so these boxes must be
    excluded from the automatic padding fix and flagged instead.

    tolerance_emu default is 45,720 EMU (0.05in) — tight enough to only
    match genuine flush-aligned pairs, not coincidental nearby shapes.
    """
    try:
        s_left, s_top, s_width = shape.left, shape.top, shape.width
    except Exception:
        return False
    if s_left is None or s_top is None:
        return False

    for l_left, l_top, l_width, l_height in divider_lines:
        if abs(s_top - l_top) > tolerance_emu:
            continue
        # Check horizontal overlap between shape and line
        s_right = s_left + (s_width or 0)
        l_right = l_left + max(l_width or 0, 1)
        overlap = min(s_right, l_right) - max(s_left, l_left)
        if overlap > 0:
            return True
    return False


def is_off_slide(shape):
    """True if shape is in the staging area outside the slide canvas."""
    try:
        # Left of slide or above slide
        if shape.left < 0 or shape.top < 0:
            return True
        # Right of slide (shapes parked in the scratch area to the right)
        if shape.left > SLIDE_W_EMU:
            return True
        return False
    except Exception:
        return False


def size_pt_of_run(run):
    """Return run font size in pt, or None if inherited."""
    try:
        if run.font.size is None:
            return None
        return run.font.size.pt
    except Exception:
        return None


def is_atlas_size(size_pt):
    """True if size rounds to one of the Atlas type-role sizes (8, 10, 12, 16, 20, 38)."""
    return any(abs(size_pt - s) < 0.05 for s in ATLAS_TYPE_SIZES)


def nearest_atlas_size(size_pt):
    """Return the nearest Atlas type-role size."""
    return min(ATLAS_TYPE_SIZES, key=lambda s: abs(s - size_pt))


def max_run_size_pt(shape):
    """Largest explicitly-set run size in a shape, or 0 if all inherited."""
    best = 0
    try:
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                s = size_pt_of_run(run)
                if s is not None:
                    best = max(best, s)
    except Exception:
        pass
    return best


def collect_established_type_scale(prs):
    """
    Atlas type roles already present in this deck (informational; recommendations now
    always come from the fixed Atlas scale, since the design system defines exactly which
    sizes exist).
    """
    scale = set()
    for slide in prs.slides:
        if is_before_slide(slide):
            continue
        for shape in slide.shapes:
            if (shape.name or "").startswith("BFX_Flag_") or not shape.has_text_frame:
                continue
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    s = size_pt_of_run(run)
                    if s is not None and is_atlas_size(s):
                        scale.add(round(s))
    return sorted(scale)


def recommend_size(size_pt, established_sizes):
    """Recommend the nearest Atlas type-role size."""
    return nearest_atlas_size(size_pt), "Atlas type roles"


def emu_to_pt(emu):
    return emu / 12_700


def emu_to_in(emu):
    return emu / 914_400


def format_style(size_pt, bold):
    """Unambiguous style label: '12pt' or '12pt Bold' — never a cryptic
    suffix like '12ptB' that requires the reader to decode it."""
    return f"{size_pt}pt Bold" if bold else f"{size_pt}pt"


def emu_to_px(emu):
    return emu / 914_400 * 96


# =============================================================================
# SECTION 3: SHAPE ITERATION
# =============================================================================

def iter_shapes(slide, for_grid=False):
    """
    Yield shapes from a slide with appropriate skip rules applied.

    for_grid=True: additional skips for grid reporting
                   (full-width, footer zone, off-slide, pictures, lines).
    for_grid=False: skip rules for text/font/padding operations.

    GroupShapes: children are yielded individually for text/padding ops
    but the group itself is reported for grid (children have relative coords).
    """
    for shape in slide.shapes:
        # Defensive backstop: never scan our own flag-marker shapes as
        # content. Normally these are already stripped before analysis
        # runs, but this guards against any future call path that skips
        # that step.
        if (shape.name or "").startswith("BFX_Flag_"):
            yield ("SKIP", shape, "flag marker shape")
            continue

        if is_protected(shape):
            yield ("SKIP", shape, "protected brand shape")
            continue

        if is_system_placeholder(shape):
            yield ("SKIP", shape, "system placeholder")
            continue

        shape_type = shape.shape_type

        # OLE objects — skip everything
        if shape_type == MSO_SHAPE_TYPE.OLE_CONTROL_OBJECT:
            yield ("SKIP", shape, "OLE object")
            continue

        # Zero-size shapes — invisible anchors
        try:
            if shape.width == 0 or shape.height == 0:
                yield ("SKIP", shape, "zero-size shape")
                continue
        except Exception:
            pass

        # Off-slide staging shapes — skip for grid, but process text
        if for_grid and is_off_slide(shape):
            yield ("SKIP", shape, "off-slide position")
            continue

        if for_grid:
            # Full-width backgrounds
            if is_full_width(shape):
                yield ("SKIP", shape, "full-width background")
                continue
            # Footer zone
            if is_in_footer_zone(shape):
                yield ("SKIP", shape, "footer zone")
                continue
            # Line/connector shapes (grid position not meaningful)
            if shape_type in (MSO_SHAPE_TYPE.LINE,):
                yield ("SKIP", shape, "line shape")
                continue

        # Groups — recurse for text ops, report whole group for grid
        if shape_type == MSO_SHAPE_TYPE.GROUP:
            if for_grid:
                # Report the group's position, not children
                yield ("OK", shape, None)
            else:
                # Recurse into children for font/padding
                try:
                    for child in shape.shapes:
                        child_type = child.shape_type
                        if child_type == MSO_SHAPE_TYPE.PICTURE:
                            continue
                        if child_type == MSO_SHAPE_TYPE.OLE_CONTROL_OBJECT:
                            continue
                        yield ("OK", child, None)
                except Exception:
                    yield ("SKIP", shape, "group recursion error")
            continue

        # Charts — skip all text operations; skip grid too (position is data-linked)
        if shape_type == MSO_SHAPE_TYPE.CHART:
            yield ("SKIP", shape, "chart")
            continue

        # Pictures — skip text operations
        if shape_type == MSO_SHAPE_TYPE.PICTURE and not for_grid:
            yield ("SKIP", shape, "picture")
            continue

        yield ("OK", shape, None)


# =============================================================================
# SECTION 4: CHANGELOG
# =============================================================================

# Properties whose problems the tool corrects itself under --fix
AUTOFIX_PROPS = {"font.name", "padding", "line_spacing",
                 "cell_padding", "fill.color", "grid_alignment"}


def entry_is_autofix(e):
    """
    True if a log entry is something --fix changes (CHANGED), or, in analyze mode, a
    FLAGGED entry the tool says it "will" correct. Anything else flagged (off-palette
    fills, off-grid shapes too far to snap, padding under a rule, font sizes, headline
    sizes) needs a designer's decision.
    """
    if e["action"] == "CHANGED":
        return True
    note = (e["note"] or "").lower()
    return (e["action"] == "FLAGGED" and e["property"] in AUTOFIX_PROPS
            and "will" in note and "review" not in note)


class Changelog:
    def __init__(self):
        self.entries = []

    def record(self, action, slide_num, shape_name, prop, was, now, note="", shape_ref=None):
        """
        action: "CHANGED" | "FLAGGED" | "SKIPPED"
        shape_ref: the actual python-pptx shape object, when available.
                   Used only internally to place in-deck flag markers —
                   never rendered in the markdown report.
        """
        self.entries.append({
            "action":     action,
            "slide":      slide_num,
            "shape":      shape_name,
            "property":   prop,
            "was":        str(was),
            "now":        str(now),
            "note":       note,
            "_shape_ref": shape_ref,
        })

    def changed(self, slide_num, shape_name, prop, was, now, note="", shape_ref=None):
        self.record("CHANGED", slide_num, shape_name, prop, was, now, note, shape_ref)

    def flagged(self, slide_num, shape_name, prop, was, note="", shape_ref=None):
        self.record("FLAGGED", slide_num, shape_name, prop, was, "—", note, shape_ref)

    def skipped(self, slide_num, shape_name, reason):
        self.record("SKIPPED", slide_num, shape_name, "—", "—", "—", reason)

    def counts(self):
        c = defaultdict(int)
        for e in self.entries:
            c[e["action"]] += 1
        return c


# =============================================================================
# SECTION 5: FIX / FLAG OPERATIONS
# =============================================================================

def op_fonts(prs, slide_num, slide, log, fix_mode):
    """
    OP-FONT: Enforce Inter as the only typeface, with exactly two weights:
    Regular (400) and Bold (700, bold=True), each optionally italic. Any Inter
    weight variant named in the file (Medium, Bold, ExtraBold, Black,
    Light, Thin, ExtraLight — with or without Italic) is auto-corrected to
    the nearest of those two, with italic preserved. Non-Inter fonts are
    replaced with Inter, leaving whatever bold/italic they already had.
    """
    for status, shape, reason in iter_shapes(slide):
        if status == "SKIP":
            continue
        if not shape.has_text_frame:
            continue

        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                font_name = run.font.name
                if font_name is None:
                    continue  # inherited — correct, do not touch

                # PowerPoint theme font references — these resolve to Inter via the theme.
                # +mn-lt = minor Latin (body font = Inter)
                # +mj-lt = major Latin (heading font = Inter)
                # Treat identically to None — do not touch.
                if font_name in ("+mn-lt", "+mj-lt", "+mn-ea", "+mj-ea",
                                  "+mn-cs", "+mj-cs"):
                    continue

                is_inter, resolved_bold, resolved_italic = resolve_inter_variant(font_name)

                if is_inter and font_name != BRAND_FONT:
                    # Named Inter weight/italic variant — collapse to
                    # Regular or Bold per Atlas rule, preserving italic.
                    if fix_mode:
                        run.font.name = BRAND_FONT
                        run.font.bold = resolved_bold
                        run.font.italic = run.font.italic or resolved_italic
                        weight_label = "Bold" if resolved_bold else "Regular"
                        italic_label = " Italic" if resolved_italic else ""
                        log.changed(slide_num, shape.name, "font.name",
                                    font_name, f"Inter ({weight_label}{italic_label})",
                                    "Inter weight variant auto-corrected")
                    else:
                        log.flagged(slide_num, shape.name, "font.name", font_name,
                                    "Inter weight variant → will be corrected to "
                                    "Regular/Bold + italic",
                                    shape_ref=shape)
                    continue

                if font_name != BRAND_FONT:
                    if fix_mode:
                        run.font.name = BRAND_FONT
                        log.changed(slide_num, shape.name, "font.name",
                                    font_name, BRAND_FONT)
                    else:
                        log.flagged(slide_num, shape.name, "font.name",
                                    font_name, f"Will be replaced with Inter",
                                    shape_ref=shape)


def op_padding(prs, slide_num, slide, log, fix_mode):
    """
    OP-PADDING: Zero out text box internal margins on unfilled shapes.
    Shapes with a visible fill (cards, callouts) are left alone — padding
    is intentional there. Only transparent/no-fill shapes are touched.

    EXCEPTION: text boxes whose top edge is flush with a divider line
    (see has_adjacent_divider_line) are excluded from the auto-fix. On
    this deck, the top padding on such boxes is the mechanism that
    creates the line's required gap above the text — zeroing it would
    collapse that gap to nothing. These are flagged for manual review
    instead, since the correct fix (move the line down, or set an
    explicit gap another way) is a design decision.
    """
    divider_lines = find_divider_lines(slide)

    for status, shape, reason in iter_shapes(slide):
        if status == "SKIP":
            continue
        if not shape.has_text_frame:
            continue
        # Tables handled separately
        if shape.shape_type == MSO_SHAPE_TYPE.TABLE:
            continue

        if has_visible_fill(shape):
            continue  # filled shape — padding is intentional

        tf = shape.text_frame
        margins = {
            "margin_left":   tf.margin_left,
            "margin_right":  tf.margin_right,
            "margin_top":    tf.margin_top,
            "margin_bottom": tf.margin_bottom,
        }

        # Check if any margin is non-zero (None = default PowerPoint margin ≠ 0)
        needs_fix = any(v != 0 for v in margins.values())
        if not needs_fix:
            continue

        was_str = (f"L:{emu_to_pt(margins['margin_left'] or 91440):.1f}pt "
                   f"R:{emu_to_pt(margins['margin_right'] or 91440):.1f}pt "
                   f"T:{emu_to_pt(margins['margin_top'] or 45720):.1f}pt "
                   f"B:{emu_to_pt(margins['margin_bottom'] or 45720):.1f}pt")

        # Protect the divider-line-spacing pattern: if this box's top edge
        # is flush with a line and has non-zero top padding, that padding
        # is the mechanism creating the line-to-text gap.
        #
        # Atlas: top padding must be exactly 0.15 in (0.20 in for display text).
        # Emit nothing when it matches; flag otherwise.
        if has_adjacent_divider_line(shape, divider_lines) and margins["margin_top"] != 0:
            top_padding_emu = margins["margin_top"] or 45720
            top_gap_in = emu_to_in(top_padding_emu)
            expected_in = (DIVIDER_PADDING_DISPLAY_IN
                           if max_run_size_pt(shape) >= TYPE_DISPLAY_PT
                           else DIVIDER_PADDING_IN)

            if abs(top_gap_in - expected_in) <= DIVIDER_PADDING_TOL_IN:
                continue
            else:
                log.flagged(slide_num, shape.name, "padding", was_str,
                            f"Top edge is flush with a rule. Padding is "
                            f"{top_gap_in:.2f}in; Atlas requires {expected_in:.2f}in"
                            f"{' (display text)' if expected_in == DIVIDER_PADDING_DISPLAY_IN else ''}"
                            f" — review before changing.",
                            shape_ref=shape)
            continue

        if fix_mode:
            tf.margin_left   = 0
            tf.margin_right  = 0
            tf.margin_top    = 0
            tf.margin_bottom = 0
            log.changed(slide_num, shape.name, "padding", was_str, "0pt all sides")
        else:
            log.flagged(slide_num, shape.name, "padding", was_str,
                        "Unfilled shape — will be zeroed")


def op_colors(prs, slide_num, slide, log, fix_mode):
    """
    OP-COLORS: Check fill and line colors against the brand palette.
    Auto-fixes pure black (#000000) and pure white (#FFFFFF) as paste artifacts.
    All other off-palette colors are flagged for designer review.
    """
    for status, shape, reason in iter_shapes(slide):
        if status == "SKIP":
            continue

        # --- Check fill color ---
        fill_rgb = get_fill_rgb(shape)
        if fill_rgb is not None:
            # Auto-fix paste artifacts
            if fill_rgb in COLOR_AUTOFIX:
                corrected = COLOR_AUTOFIX[fill_rgb]
                if fix_mode:
                    shape.fill.solid()
                    shape.fill.fore_color.rgb = RGBColor(*corrected)
                    log.changed(slide_num, shape.name, "fill.color",
                                "#{:02X}{:02X}{:02X}".format(*fill_rgb),
                                "#{:02X}{:02X}{:02X}".format(*corrected),
                                "Auto-remapped paste artifact")
                else:
                    log.flagged(slide_num, shape.name, "fill.color",
                                "#{:02X}{:02X}{:02X}".format(*fill_rgb),
                                "Paste artifact → will remap to "
                                "#{:02X}{:02X}{:02X}".format(*corrected),
                                shape_ref=shape)
                continue

            # Check palette distance
            dist, near_name, near_rgb = nearest_brand_color(fill_rgb)
            if dist > COLOR_TOLERANCE:
                log.flagged(slide_num, shape.name, "fill.color",
                            "#{:02X}{:02X}{:02X}".format(*fill_rgb),
                            f"Off-palette (dist={dist:.1f}). "
                            f"Nearest: {near_name} "
                            f"(#{near_rgb[0]:02X}{near_rgb[1]:02X}{near_rgb[2]:02X})",
                            shape_ref=shape)

        # --- Check line/stroke color ---
        line_rgb = get_line_rgb(shape)
        if line_rgb is not None:
            if line_rgb in COLOR_AUTOFIX:
                # Don't auto-remap line colors — too risky (could affect dividers)
                dist, near_name, near_rgb = nearest_brand_color(line_rgb)
                log.flagged(slide_num, shape.name, "line.color",
                            "#{:02X}{:02X}{:02X}".format(*line_rgb),
                            f"Pure black/white stroke — review. "
                            f"Nearest brand: {near_name}",
                            shape_ref=shape)
                continue

            dist, near_name, near_rgb = nearest_brand_color(line_rgb)
            if dist > COLOR_TOLERANCE:
                log.flagged(slide_num, shape.name, "line.color",
                            "#{:02X}{:02X}{:02X}".format(*line_rgb),
                            f"Off-palette stroke (dist={dist:.1f}). "
                            f"Nearest: {near_name} "
                            f"(#{near_rgb[0]:02X}{near_rgb[1]:02X}{near_rgb[2]:02X})",
                            shape_ref=shape)


def op_typography(prs, slide_num, slide, log,
                  deck_styles, slide_styles, established_sizes):
    """
    OP-TYPOGRAPHY: Check font sizes against the Atlas type roles (8/10/12/16/20/38 pt).
    Collects style data for the per-slide and per-deck counts (checked later).

    Recommendations prefer a size already established elsewhere in this
    deck over the mathematically-nearest multiple of 4 — reusing an
    existing size never grows the type-style-count problem.

    All violations are flagged only — no auto-correction. The right
    direction to round is a design judgment call.
    """
    for status, shape, reason in iter_shapes(slide):
        if status == "SKIP":
            continue
        if not shape.has_text_frame:
            continue

        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                size_pt = size_pt_of_run(run)
                if size_pt is None:
                    continue  # inherited — skip

                # Round to 1 decimal for cleaner display
                size_display = round(size_pt, 1)
                bold = bool(run.font.bold)
                style_key = (round(size_pt), bold)

                # Collect for count checks
                slide_styles.add(style_key)
                deck_styles.add(style_key)

                # Check minimum
                if size_pt < TYPE_SIZE_MINIMUM_PT:
                    log.flagged(slide_num, shape.name, "font.size",
                                f"{size_display}pt",
                                f"Below minimum ({TYPE_SIZE_MINIMUM_PT}pt)",
                                shape_ref=shape)
                    continue

                # Check against the Atlas type roles (8, 10, 12, 16, 20, 38)
                if not is_atlas_size(size_pt):
                    recommended, source = recommend_size(size_pt, established_sizes)
                    log.flagged(slide_num, shape.name, "font.size",
                                f"{size_display}pt",
                                f"Not an Atlas type size (8/10/12/16/20/38). Recommend {recommended}pt "
                                f"({source}).",
                                shape_ref=shape)


def op_line_spacing(prs, slide_num, slide, log, fix_mode):
    """
    OP-LINE-SPACING: Atlas line spacing is Multiple 1.17 for all slide text
    (tokens/typography.css, layout-rules.json typography.lineSpacing).
    Auto-fixes explicit values that differ. Paragraphs that inherit spacing from the
    layout are left alone: the Atlas layouts carry 1.17, so inheriting is correct.
    """
    for status, shape, reason in iter_shapes(slide):
        if status == "SKIP":
            continue
        if not shape.has_text_frame:
            continue

        for para in shape.text_frame.paragraphs:
            if not para.runs:
                continue
            current_spacing = para.line_spacing
            if current_spacing is None:
                continue  # inherited from the layout (1.17) — correct
            # Absolute (point-based) spacing uses a different unit system —
            # leave it alone rather than guessing a conversion.
            if isinstance(current_spacing, int):
                continue
            if abs(current_spacing - LINE_SPACING) < 0.005:
                continue

            was = f"{current_spacing:.2f}"
            if fix_mode:
                para.line_spacing = LINE_SPACING
                log.changed(slide_num, shape.name, "line_spacing",
                            was, str(LINE_SPACING), "Atlas: Multiple 1.17")
            else:
                log.flagged(slide_num, shape.name, "line_spacing",
                            was, f"Will be set to {LINE_SPACING} (Atlas: Multiple 1.17)")


def op_headline_consistency(prs, log, fix_mode):
    """
    OP-HEADLINE-CONSISTENCY: Flag when content-slide title sizes are
    inconsistent across the deck. FLAG ONLY — never auto-fixed.

    Reverted from auto-fix: a majority-vote size change was found to
    visually break slides where a smaller title was an intentional choice
    (e.g. slides with dense tables/small supporting visuals, where a
    smaller title keeps the balance right). This is a judgment call that
    requires seeing the actual slide, not a mechanical rule.

    Statement slides (display size, 38pt) are excluded — different slide type entirely.
    Content-slide headlines must be 20pt (Atlas slide-title role).
    """
    STATEMENT_THRESHOLD_PT = TYPE_DISPLAY_PT  # 38 pt display titles (cover, divider, closing)

    headline_sizes = {}
    for slide_num, slide in enumerate(prs.slides, 1):
        if is_before_slide(slide):
            continue
        for shape in slide.shapes:
            if not is_title_shape(shape) or not shape.has_text_frame:
                continue
            sizes = set()
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    s = size_pt_of_run(run)
                    if s is not None:
                        sizes.add(round(s))
            if sizes:
                headline_sizes[slide_num] = sizes

    content_slides = {n: s for n, s in headline_sizes.items()
                      if all(sz < STATEMENT_THRESHOLD_PT for sz in s)}

    all_content_sizes = set()
    for sizes in content_slides.values():
        all_content_sizes.update(sizes)

    # Atlas: content-slide headlines are 20 pt
    for n, sizes in sorted(content_slides.items()):
        off = sorted(s for s in sizes if abs(s - TYPE_HEADLINE_PT) > 0.5)
        if off:
            log.flagged(n, "TITLE", "headline_size",
                        "/".join(f"{s}pt" for s in off),
                        f"Atlas slide title is {TYPE_HEADLINE_PT:.0f}pt — confirm this is "
                        f"an intentional exception")

    if len(all_content_sizes) <= 1:
        return headline_sizes, False  # already consistent

    breakdown = ", ".join(
        f"Slide {n}: {'/'.join(str(s) for s in sorted(sizes))}pt"
        for n, sizes in sorted(content_slides.items())
    )
    log.flagged(0, "DECK", "headline_size_consistency",
                f"{len(all_content_sizes)} different sizes in content slides",
                f"Review each slide before standardising — smaller titles "
                f"may be intentional for dense-content slides. {breakdown}")

    return headline_sizes, True


def op_style_counts(log, deck_styles, per_slide_styles):
    """
    OP-STYLE-COUNTS: Check per-slide (max 5) and per-deck (max 8) type style limits.
    A 'style' is a unique (font_size_pt_rounded, bold) pair.
    """
    # Per slide
    for slide_num, styles in per_slide_styles.items():
        if len(styles) > MAX_STYLES_PER_SLIDE:
            style_list = ", ".join(
                format_style(s, b) for s, b in sorted(styles)
            )
            log.flagged(slide_num, "SLIDE", "type_style_count",
                        f"{len(styles)} styles found",
                        f"Found {len(styles)} type styles on this slide "
                        f"(max allowed: {MAX_STYLES_PER_SLIDE}). "
                        f"Styles: {style_list}")

    # Per deck
    if len(deck_styles) > MAX_STYLES_PER_DECK:
        style_list = ", ".join(
            format_style(s, b) for s, b in sorted(deck_styles)
        )
        log.flagged(0, "DECK", "type_style_count",
                    f"{len(deck_styles)} styles across deck",
                    f"Found {len(deck_styles)} type styles across the deck "
                    f"(max allowed: {MAX_STYLES_PER_DECK}). "
                    f"Styles: {style_list}")


def find_filled_containers(slide):
    """
    Return bounding boxes (left, top, right, bottom) in EMU of all filled
    auto-shapes on the slide — rectangles, rounded rects, and other panel
    shapes used as card backgrounds. Shapes inside these containers are
    positioned relative to the card, not the slide column grid, so they
    must be excluded from grid checking.
    """
    containers = []
    for shape in slide.shapes:
        if (shape.name or "").startswith("BFX_Flag_"):
            continue
        # Only auto-shapes (panels, cards) — not placeholders, groups, etc.
        if shape.shape_type != MSO_SHAPE_TYPE.AUTO_SHAPE:
            continue
        if not has_visible_fill(shape):
            continue
        # Ignore shapes that are themselves text-primary (small label boxes)
        # by requiring a minimum area (~0.5in²)
        try:
            area = shape.width * shape.height
        except Exception:
            continue
        if area < 914400 * 0.5 * 914400 * 0.5:  # 0.5in × 0.5in minimum
            continue
        try:
            containers.append((
                shape.left,
                shape.top,
                shape.left + shape.width,
                shape.top + shape.height,
            ))
        except Exception:
            pass
    return containers


def is_inside_container(shape, containers, overlap_threshold=0.7):
    """
    True if more than overlap_threshold of the shape's area falls within
    any filled container. A shape that is 70%+ inside a card is positioned
    relative to that card, not the slide grid.
    """
    try:
        sl, st = shape.left, shape.top
        sr, sb = sl + shape.width, st + shape.height
        shape_area = (sr - sl) * (sb - st)
    except Exception:
        return False

    if shape_area <= 0:
        return False

    for (cl, ct, cr, cb) in containers:
        overlap_l = max(sl, cl)
        overlap_t = max(st, ct)
        overlap_r = min(sr, cr)
        overlap_b = min(sb, cb)
        if overlap_r > overlap_l and overlap_b > overlap_t:
            overlap_area = (overlap_r - overlap_l) * (overlap_b - overlap_t)
            if overlap_area / shape_area >= overlap_threshold:
                return True
    return False


def op_grid(prs, slide_num, slide, log, fix_mode):
    """
    OP-GRID: Check and fix content grid alignment — TEXT_BOX, PLACEHOLDER,
    PICTURE, and TABLE shapes only. This is a deliberate narrowing: only
    genuine content is checked against the 12-column grid. Decorative and
    diagram pieces (chevrons, triangles, dots, connectors, icon groups) are
    positioned relative to each other or to nearby text, not the slide grid
    — checking them produces noise, not useful signal, so they are skipped
    entirely regardless of position.

    Auto-fix: text boxes and placeholders within GRID_SNAP_THRESHOLD_EMU of a
    column edge are snapped to that column. These are clear drift cases.

    Flag only: everything else checkable that is meaningfully off-grid, no
    matter how far off. A shape 50px off the grid is exactly the case a
    designer most needs to see, not one to skip silently.

    Skipped: shapes substantially inside a filled container (card, panel) —
    they are positioned relative to that container, not the slide grid.
    """
    filled_containers = find_filled_containers(slide)

    for status, shape, reason in iter_shapes(slide, for_grid=True):
        if status == "SKIP":
            continue

        # Only genuine content shapes are grid-relevant. Decorative and
        # diagram pieces are never checked, at any distance.
        if shape.shape_type not in GRID_CHECKABLE_TYPES:
            continue

        # Skip shapes that live inside a filled card/panel.
        if is_inside_container(shape, filled_containers):
            continue

        try:
            left_emu = shape.left
            if left_emu is None:
                continue
        except Exception:
            continue

        left_px = emu_to_px(left_emu)

        # Find nearest column edge in EMU and px
        col_distances = [
            (abs(left_emu - col_emu), COLUMN_EDGE_LABELS[col_emu], col_emu, col_px)
            for i, (col_emu, col_px) in enumerate(
                zip(COLUMN_EDGES_EMU, COLUMN_EDGES_PX)
            )
        ]
        min_dist_emu, nearest_col, nearest_col_emu, nearest_col_px = min(
            col_distances, key=lambda x: x[0]
        )
        min_dist_px = emu_to_px(min_dist_emu)

        # Trivially aligned — not worth reporting at all
        if min_dist_emu <= GRID_NOISE_FLOOR_EMU:
            continue

        is_snappable = shape.shape_type in SNAPPABLE_SHAPE_TYPES

        if min_dist_emu <= GRID_SNAP_THRESHOLD_EMU and is_snappable:
            # Clear, small drift on a text shape — safe to snap automatically
            if fix_mode:
                shape.left = nearest_col_emu
                log.changed(slide_num, shape.name, "grid_alignment",
                            f"left={left_px:.1f}px",
                            f"left={nearest_col_px:.1f}px ({nearest_col})",
                            f"Snapped {min_dist_px:.1f}px drift")
            else:
                log.flagged(slide_num, shape.name, "grid_alignment",
                            f"left={left_px:.1f}px",
                            f"Will snap to {nearest_col} "
                            f"({nearest_col_px:.1f}px), {min_dist_px:.1f}px drift",
                            shape_ref=shape)
        else:
            # Everything else that's off-grid gets flagged, no matter how
            # far — this is the case that matters most to a designer.
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                shape_kind = "picture"
            elif shape.shape_type == MSO_SHAPE_TYPE.TABLE:
                shape_kind = "table"
            else:
                shape_kind = "text shape"
            log.flagged(slide_num, shape.name, "grid_alignment",
                        f"left={left_px:.1f}px",
                        f"Off-grid {shape_kind} — review manually. "
                        f"Nearest: {nearest_col} ({nearest_col_px:.1f}px), "
                        f"{min_dist_px:.1f}px off",
                        shape_ref=shape)


# =============================================================================
# SECTION 6: TABLE OPERATIONS
# =============================================================================

def op_tables(prs, slide_num, slide, log, fix_mode):
    """
    OP-TABLES: For each table on the slide:
      - Auto-fix: Inter font in all cells
      - Auto-fix: cell padding to 0.10" (91,440 EMU), per the Atlas table rules
      - Flag: off-palette fill colors
      - Flag: off-palette border colors
    """
    for shape in slide.shapes:
        if shape.shape_type != MSO_SHAPE_TYPE.TABLE:
            continue

        table = shape.table
        for row_idx, row in enumerate(table.rows):
            for col_idx, cell in enumerate(row.cells):
                cell_ref = f"Table[{shape.name}] R{row_idx}C{col_idx}"

                # --- Font ---
                if cell.text_frame:
                    for para in cell.text_frame.paragraphs:
                        for run in para.runs:
                            fn = run.font.name
                            if fn is None:
                                continue
                            # Theme font references — resolve to Inter via theme
                            if fn in ("+mn-lt", "+mj-lt", "+mn-ea", "+mj-ea",
                                      "+mn-cs", "+mj-cs"):
                                continue
                            is_inter, resolved_bold, resolved_italic = resolve_inter_variant(fn)
                            if is_inter and fn != BRAND_FONT:
                                if fix_mode:
                                    run.font.name = BRAND_FONT
                                    run.font.bold = resolved_bold
                                    run.font.italic = run.font.italic or resolved_italic
                                    weight_label = "Bold" if resolved_bold else "Regular"
                                    italic_label = " Italic" if resolved_italic else ""
                                    log.changed(slide_num, cell_ref, "font.name",
                                                fn, f"Inter ({weight_label}{italic_label})",
                                                "Inter weight variant auto-corrected")
                                else:
                                    log.flagged(slide_num, cell_ref, "font.name",
                                                fn, "Inter weight variant → will be corrected to "
                                                "Regular/Bold + italic",
                                                shape_ref=shape)
                            elif fn != BRAND_FONT:
                                if fix_mode:
                                    run.font.name = BRAND_FONT
                                    log.changed(slide_num, cell_ref, "font.name",
                                                fn, BRAND_FONT)
                                else:
                                    log.flagged(slide_num, cell_ref, "font.name",
                                                fn, "Will be replaced with Inter",
                                                shape_ref=shape)

                # --- Cell padding ---
                tc = cell._tc
                tcPr = tc.find(
                    '{http://schemas.openxmlformats.org/drawingml/2006/main}tcPr'
                )
                if tcPr is not None:
                    margins = {
                        'marL': int(tcPr.get('marL', 91440)),
                        'marR': int(tcPr.get('marR', 91440)),
                        'marT': int(tcPr.get('marT', 45720)),
                        'marB': int(tcPr.get('marB', 45720)),
                    }
                    needs_fix = any(
                        v != TABLE_CELL_MARGIN_EMU for v in margins.values()
                    )
                    if needs_fix:
                        was_str = (f"L:{emu_to_pt(margins['marL']):.1f}pt "
                                   f"R:{emu_to_pt(margins['marR']):.1f}pt "
                                   f"T:{emu_to_pt(margins['marT']):.1f}pt "
                                   f"B:{emu_to_pt(margins['marB']):.1f}pt")
                        if fix_mode:
                            for attr in ('marL', 'marR', 'marT', 'marB'):
                                tcPr.set(attr, str(TABLE_CELL_MARGIN_EMU))
                            log.changed(slide_num, cell_ref, "cell_padding",
                                        was_str, "0.10in all sides")
                        else:
                            log.flagged(slide_num, cell_ref, "cell_padding",
                                        was_str, "Will be set to 0.10in all sides",
                                        shape_ref=shape)

                # --- Cell fill color ---
                try:
                    fill = cell.fill
                    if fill.type is not None:
                        rgb = fill.fore_color.rgb
                        fill_rgb = (rgb[0], rgb[1], rgb[2])
                        dist, near_name, near_rgb = nearest_brand_color(fill_rgb)
                        if dist > COLOR_TOLERANCE:
                            log.flagged(slide_num, cell_ref, "cell_fill",
                                        "#{:02X}{:02X}{:02X}".format(*fill_rgb),
                                        f"Off-palette (dist={dist:.1f}). "
                                        f"Nearest: {near_name}",
                                        shape_ref=shape)
                except Exception:
                    pass


# =============================================================================
# SECTION 7: REPORT GENERATION
# =============================================================================

# =============================================================================
# SECTION 6.6: BEFORE/AFTER SLIDE DUPLICATION
# Replicates the designer's own manual workflow: duplicate the original
# slide, hide it, edit the copy. For every slide that receives at least one
# auto-fix, a hidden copy of its PRE-FIX content is inserted immediately
# before it — [hidden original] -> [edited version] — so before/after is
# visible just by scrolling, no separate diff tool needed.
# =============================================================================

R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
R_ID_ATTRS = (f'{{{R_NS}}}id', f'{{{R_NS}}}embed', f'{{{R_NS}}}link')

BEFORE_SLIDE_SENTINEL = (
    "BFX_BEFORE_SLIDE — original content, preserved for comparison. "
    "Safe to delete once reviewed, or run --clean-flags to remove all "
    "before-slides at once."
)


def is_before_slide(slide):
    """True if this slide was inserted by brand_fix.py as a 'before'
    reference copy. These are permanently excluded from all analysis —
    re-fixing a before-slide would destroy the very thing that makes it
    useful as a historical reference."""
    try:
        if not slide.has_notes_slide:
            return False
        return BEFORE_SLIDE_SENTINEL in (slide.notes_slide.notes_text_frame.text or "")
    except Exception:
        return False


def _remap_relationship_ids(element, rid_map):
    """Rewrite every r:id / r:embed / r:link attribute in a copied XML
    subtree from the source part's relationship IDs to the destination
    part's newly-created relationship IDs. Without this, a duplicated
    picture's <a:blip r:embed="rIdN"> points at a relationship ID that
    doesn't exist in the new slide's part, and the image breaks."""
    for el in element.iter():
        for attr in R_ID_ATTRS:
            val = el.get(attr)
            if val is not None and val in rid_map and rid_map[val] is not None:
                el.set(attr, rid_map[val])


def _move_slide_to_index(prs, from_index, to_index):
    """Reposition a slide within the deck by moving its <p:sldId> element
    inside <p:sldIdLst>. python-pptx has no native slide-reorder API."""
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    xml_slides.remove(slides[from_index])
    xml_slides.insert(to_index, slides[from_index])


def _duplicate_slide_from_snapshot(prs, target_index, snapshot_element,
                                   source_part_rels, hidden=True):
    """
    Insert a new slide at target_index, populated with the shapes from
    snapshot_element (a deep-copied <p:sld> XML tree taken before any
    fixes were applied). Uses the CURRENT live slide's relationships as
    the source for images/media, since brand_fix.py never adds, removes,
    or retargets a relationship — only text/fill/line/position properties
    — so the relationships captured at snapshot-time are still identical
    to the live slide's relationships now.

    Returns the new Slide object, or None if duplication failed (logged
    to console; the edited slide is left in place either way — this
    feature degrading gracefully must never block the actual fixes).
    """
    try:
        target_slide = prs.slides[target_index]
        new_slide = prs.slides.add_slide(target_slide.slide_layout)

        # Remove the default placeholder shapes add_slide() populated from
        # the layout — we want an exact copy of the snapshot's content instead.
        for shp in list(new_slide.shapes):
            shp._element.getparent().remove(shp._element)

        # Relate the new slide's part to the SAME target parts (images etc.)
        # the source relies on, building an old-rId -> new-rId map. Reusing
        # the target part means image bytes are never duplicated on disk.
        #
        # CRITICAL: notesSlide and slideLayout relationships must NEVER be
        # copied this way. Each slide needs its OWN unique notes part —
        # relating two different slides to the SAME notes part makes them
        # share (and silently overwrite) each other's notes text, which is
        # exactly how the before/after marking broke on the first attempt.
        # slideLayout is already set correctly by add_slide() itself.
        SKIP_RELTYPES = {
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
        }
        rid_map = {}
        for rel_id, rel in list(source_part_rels.items()):
            try:
                if rel.is_external:
                    continue
                if rel.reltype in SKIP_RELTYPES:
                    continue
                new_rid = new_slide.part.relate_to(rel.target_part, rel.reltype)
                rid_map[rel_id] = new_rid
            except Exception:
                continue

        # Deep-copy every real shape element from the snapshot into the
        # new slide, remapping relationship IDs as we go.
        snap_cSld = snapshot_element.find(qn('p:cSld'))
        snap_spTree = snap_cSld.find(qn('p:spTree'))
        dst_spTree = new_slide.shapes._spTree
        skip_tags = {qn('p:nvGrpSpPr'), qn('p:grpSpPr')}
        for child in list(snap_spTree):
            if child.tag in skip_tags:
                continue
            new_child = copy.deepcopy(child)
            _remap_relationship_ids(new_child, rid_map)
            dst_spTree.append(new_child)

        # Copy a slide-level background override, if the source had one.
        snap_bg = snap_cSld.find(qn('p:bg'))
        if snap_bg is not None:
            dst_cSld = new_slide._element.find(qn('p:cSld'))
            old_bg = dst_cSld.find(qn('p:bg'))
            if old_bg is not None:
                dst_cSld.remove(old_bg)
            new_bg = copy.deepcopy(snap_bg)
            _remap_relationship_ids(new_bg, rid_map)
            dst_cSld.insert(0, new_bg)

        if hidden:
            new_slide._element.set('show', '0')

        # Mark it so it's never re-analyzed on a future run, and so
        # --clean-flags can find and remove it safely.
        notes_tf = new_slide.notes_slide.notes_text_frame
        notes_tf.text = BEFORE_SLIDE_SENTINEL

        # Move from the end (where add_slide appends) to just before the
        # edited slide.
        new_index = len(prs.slides) - 1
        _move_slide_to_index(prs, new_index, target_index)

        return new_slide
    except Exception as e:
        print(f"  Warning: could not create before-slide for slide "
              f"{target_index + 1} ({e}). Skipping — edited slide is fine.")
        return None


def insert_before_slides(prs, log, snapshots):
    """
    For every slide that received at least one CHANGED (auto-fixed) entry,
    insert a hidden pre-fix snapshot immediately before it. Processes in
    reverse slide order so earlier insertions don't shift the indices of
    slides not yet processed.
    """
    changed_slide_indices = sorted({
        e["slide"] - 1  # log slide numbers are 1-based
        for e in log.entries
        if e["action"] == "CHANGED" and e["slide"] > 0
    }, reverse=True)

    n_inserted = 0
    for idx in changed_slide_indices:
        if idx not in snapshots:
            continue
        target_slide = prs.slides[idx]
        result = _duplicate_slide_from_snapshot(
            prs, idx, snapshots[idx], target_slide.part.rels, hidden=True
        )
        if result is not None:
            n_inserted += 1
    return n_inserted


def remove_before_slides(prs):
    """Delete every slide previously inserted as a hidden 'before' copy.
    Only matches slides carrying our exact sentinel note — never touches
    a slide the designer hid themselves for unrelated reasons."""
    indices_to_remove = [
        i for i, slide in enumerate(prs.slides) if is_before_slide(slide)
    ]
    for idx in sorted(indices_to_remove, reverse=True):
        xml_slides = prs.slides._sldIdLst
        slides = list(xml_slides)
        xml_slides.remove(slides[idx])
    return len(indices_to_remove)


# =============================================================================

from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

FLAG_MARKER_PREFIX = "BFX_Flag_"
FLAG_BLOCK_START = ">>> BFX_FLAGS_START"
FLAG_BLOCK_END = "<<< BFX_FLAGS_END"
MARKER_DIAMETER_EMU = 137_160  # ~0.15in
BRAND_RED = RGBColor(0xD9, 0x68, 0x68)  # brand semantic "error" color

CATEGORY_LABELS = {
    "font.name":                 "Font",
    "padding":                   "Padding",
    "fill.color":                "Color",
    "line.color":                "Color",
    "cell_fill":                 "Table Color",
    "cell_padding":              "Table Padding",
    "font.size":                 "Type Scale",
    "type_style_count":          "Type Styles",
    "headline_size_consistency": "Headlines",
    "grid_alignment":            "Grid",
}


def _category_label(prop):
    return CATEGORY_LABELS.get(prop, prop)


def _strip_slide_annotation(slide):
    """Remove any BFX marker shapes and flag block from a prior run,
    leaving any other speaker notes content the user wrote untouched."""
    to_remove = [s for s in slide.shapes
                 if (s.name or "").startswith(FLAG_MARKER_PREFIX)]
    for s in to_remove:
        s._element.getparent().remove(s._element)

    try:
        if slide.has_notes_slide:
            tf = slide.notes_slide.notes_text_frame
            text = tf.text or ""
            if FLAG_BLOCK_START in text:
                start = text.find(FLAG_BLOCK_START)
                end = text.find(FLAG_BLOCK_END)
                if end != -1:
                    end += len(FLAG_BLOCK_END)
                    text = (text[:start] + text[end:]).strip("\n")
                    tf.text = text
    except Exception:
        pass


def _place_marker(slide, shape, number):
    """Add a small numbered red circle at a shape's top-left corner."""
    try:
        left, top = shape.left, shape.top
        if left is None or top is None:
            return
    except Exception:
        return

    half = MARKER_DIAMETER_EMU // 2
    marker = slide.shapes.add_shape(
        MSO_SHAPE.OVAL,
        max(left - half, 0), max(top - half, 0),
        MARKER_DIAMETER_EMU, MARKER_DIAMETER_EMU,
    )
    marker.name = f"{FLAG_MARKER_PREFIX}{number}"
    marker.fill.solid()
    marker.fill.fore_color.rgb = BRAND_RED
    marker.line.fill.background()
    try:
        marker.shadow.inherit = False
    except Exception:
        pass

    tf = marker.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.word_wrap = False
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = str(number)
    run.font.size = Pt(9)
    run.font.bold = True
    run.font.name = BRAND_FONT
    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)


def annotate_flags_in_deck(prs, log, established_sizes=None):
    """
    Write flagged items directly into the deck: a numbered marker on each
    flagged shape, and a matching list in that slide's speaker notes.
    Deck-wide flags (headline consistency, type style counts) have no
    single shape — they go into Slide 1's notes under their own heading.

    Slide 1 also always shows the deck's detected type scale (the sizes
    already in use that match the Atlas type roles), so font-size
    recommendations elsewhere in the deck are transparent about where they
    came from — not a black box.
    """
    by_slide = {}
    deck_wide = []
    for e in log.entries:
        if e["action"] != "FLAGGED":
            continue
        if e["slide"] == 0:
            deck_wide.append(e)
        else:
            by_slide.setdefault(e["slide"], []).append(e)

    # Deck-wide notes go on the first REAL (non-before) slide. Usually
    # that's slide 1, but on a re-run where slide 1 itself was fixed, a
    # hidden before-copy may now occupy position 1 — attach to whichever
    # slide is actually the first genuine content slide instead.
    first_real_idx = next(
        (i for i, s in enumerate(prs.slides, 1) if not is_before_slide(s)), 1
    )

    for slide_idx, slide in enumerate(prs.slides, 1):
        _strip_slide_annotation(slide)

        entries = by_slide.get(slide_idx, [])
        extra = deck_wide if slide_idx == first_real_idx else []
        show_scale_note = (slide_idx == first_real_idx and established_sizes)

        if not entries and not extra and not show_scale_note:
            continue

        # Group entries by shape (same shape object -> one marker, one
        # numbered note entry listing every issue found on it).
        groups = []  # list of (shape_obj_or_None, [entries])
        for e in entries:
            shp = e.get("_shape_ref")
            match = next((g for g in groups if g[0] is shp and shp is not None), None)
            if match:
                match[1].append(e)
            else:
                groups.append((shp, [e]))

        block = [FLAG_BLOCK_START]

        if show_scale_note:
            scale_str = ", ".join(f"{s}pt" for s in established_sizes)
            block.append(f"DETECTED TYPE SCALE (sizes already used in this "
                        f"deck that comply with the ×4 rule): {scale_str}")
            block.append("Font-size recommendations below prefer these sizes "
                        "over introducing new ones.")
            block.append("")

        if extra:
            block.append(f"DECK-WIDE ISSUES ({len(extra)}):")
            for e in extra:
                block.append(f"  • [{_category_label(e['property'])}] {e['note']}")
            block.append("")

        if groups:
            block.append(f"FLAGGED ON THIS SLIDE ({len(groups)}):")
            marker_num = 1
            for shp, group_entries in groups:
                cats = sorted(set(_category_label(e["property"]) for e in group_entries))
                shape_label = group_entries[0]["shape"]
                block.append(f"  {marker_num}. {shape_label} [{', '.join(cats)}]")
                # Collapse identical repeated notes (e.g. same issue on
                # every run in a text box) into one line with a count.
                note_counts = {}
                note_order = []
                for e in group_entries:
                    if e["note"] not in note_counts:
                        note_order.append(e["note"])
                    note_counts[e["note"]] = note_counts.get(e["note"], 0) + 1
                for note_text in note_order:
                    n = note_counts[note_text]
                    suffix = f" (×{n})" if n > 1 else ""
                    block.append(f"     - {note_text}{suffix}")
                if shp is not None:
                    _place_marker(slide, shp, marker_num)
                marker_num += 1

        block.append(FLAG_BLOCK_END)

        notes_slide = slide.notes_slide  # creates if needed
        tf = notes_slide.notes_text_frame
        existing = (tf.text or "").strip()
        new_text = "\n".join(block)
        tf.text = new_text + ("\n\n" + existing if existing else "")


def strip_all_flag_annotations(prs):
    """Remove every BFX marker and flag note block from the whole deck.
    Used by --clean-flags to produce a final, review-free copy."""
    for slide in prs.slides:
        _strip_slide_annotation(slide)


def build_report(input_path, log, fix_mode,
                 deck_styles, per_slide_styles,
                 headline_sizes, headline_inconsistent):
    """Produce a structured markdown audit report."""
    counts = log.counts()
    changed = counts.get("CHANGED", 0)
    flagged = counts.get("FLAGGED", 0)
    skipped = counts.get("SKIPPED", 0)

    lines = []
    lines.append("# Gemic Brand Audit Report")
    lines.append(f"**File:** `{os.path.basename(input_path)}`")
    lines.append(f"**Mode:** {'FIX (auto-fixes applied)' if fix_mode else 'ANALYZE (no changes made)'}")
    lines.append(f"**Date:** {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    lines.append("")
    lines.append("---")
    lines.append("")
    # Auto-fixable props vs flag-only props
    AUTOFIX_PROPS = {"font.name", "padding", "line_spacing",
                     "cell_padding", "fill.color", "grid_alignment"}
    FLAG_ONLY_PROPS = {"line.color", "cell_fill", "font.size",
                       "type_style_count", "headline_size_consistency"}

    n_autofixable = sum(1 for e in log.entries if entry_is_autofix(e))
    n_flag_only = sum(1 for e in log.entries
                      if e["action"] == "FLAGGED" and not entry_is_autofix(e))

    lines.append("## Summary")
    lines.append(f"| | Count |")
    lines.append(f"|---|---|")
    if fix_mode:
        lines.append(f"| ✓ Auto-fixes applied | **{changed}** |")
    else:
        lines.append(f"| ✓ Auto-fixable (re-run with --fix) | **{n_autofixable}** |")
    lines.append(f"| ⚠ Flagged for designer review | **{n_flag_only}** |")
    lines.append(f"| — Skipped | **{skipped}** |")
    lines.append("")

    # Deck-level checks
    lines.append("---")
    lines.append("")
    lines.append("## Deck-Level Checks")
    lines.append("")

    # Type style count
    n_deck_styles = len(deck_styles)
    deck_ok = n_deck_styles <= MAX_STYLES_PER_DECK
    style_list = ", ".join(
        format_style(s, b) for s, b in sorted(deck_styles)
    )
    lines.append(f"**Type styles across deck:** "
                 f"{n_deck_styles}/{MAX_STYLES_PER_DECK} "
                 f"{'✓' if deck_ok else '✗ EXCEEDS LIMIT'}")
    lines.append(f"  Styles found: `{style_list}`")
    lines.append("")

    # Headline consistency
    if headline_inconsistent:
        lines.append("**Headline size consistency:** ✗ INCONSISTENT")
        for slide_num, sizes in sorted(headline_sizes.items()):
            lines.append(f"  Slide {slide_num}: {'/'.join(str(s) for s in sorted(sizes))}pt")
    else:
        if headline_sizes:
            common_size = next(iter(next(iter(headline_sizes.values()))))
            lines.append(f"**Headline size consistency:** ✓ All headlines {common_size}pt")
    lines.append("")

    # Per-slide style counts
    over_limit = {
        n: styles for n, styles in per_slide_styles.items()
        if len(styles) > MAX_STYLES_PER_SLIDE
    }
    if over_limit:
        lines.append(f"**Slides exceeding {MAX_STYLES_PER_SLIDE} type styles:**")
        for slide_num, styles in sorted(over_limit.items()):
            sl = ", ".join(format_style(s, b) for s, b in sorted(styles))
            lines.append(f"  Slide {slide_num}: {len(styles)} styles — `{sl}`")
        lines.append("")

    # Detailed entries grouped by category
    categories = {
        "font.name":               ("Font Violations",         "CHANGED"),
        "padding":                 ("Padding Violations",      "CHANGED"),
        "line_spacing":            ("Line Spacing",            "CHANGED"),
        "cell_padding":            ("Table Cell Padding",      "CHANGED"),
        "grid_alignment":          ("Grid Alignment",          "CHANGED"),
        "fill.color":              ("Fill Color Issues",       "FLAGGED"),
        "line.color":              ("Stroke Color Issues",     "FLAGGED"),
        "cell_fill":               ("Table Fill Issues",       "FLAGGED"),
        "font.size":               ("Font Size Issues",        "FLAGGED"),
        "type_style_count":        ("Type Style Count",        "FLAGGED"),
        "headline_size_consistency": ("Headline Consistency",  "FLAGGED"),
        "headline_size":           ("Headline Size (Atlas 20 pt)", "FLAGGED"),
    }

    for prop_key, (section_title, expected_action) in categories.items():
        entries = [e for e in log.entries
                   if e["property"] == prop_key
                   and e["action"] in ("CHANGED", "FLAGGED")]
        if not entries:
            continue

        lines.append("---")
        lines.append("")
        n_auto = sum(1 for e in entries if entry_is_autofix(e))
        prefix = "✓" if n_auto == len(entries) else "⚠"
        lines.append(f"## {prefix} {section_title} ({len(entries)})")
        lines.append("")
        lines.append("| Slide | Shape | Was | Action | Note |")
        lines.append("|---|---|---|---|---|")
        for e in entries:
            slide_label = f"Slide {e['slide']}" if e['slide'] > 0 else "Deck"
            act = ("Fixed" if e["action"] == "CHANGED" else "Auto-fixable") if entry_is_autofix(e) else "Designer review"
            lines.append(f"| {slide_label} | `{e['shape']}` | `{e['was']}` | {act} | {e['note']} |")
        lines.append("")

    return "\n".join(lines)


# =============================================================================
# SECTION 8: MAIN
# =============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Gemic brand compliance tool for PowerPoint files."
    )
    parser.add_argument("input", help="Path to the .pptx file to process")
    parser.add_argument(
        "--fix", action="store_true",
        help="Apply auto-fixes. Default is analyze-only."
    )
    parser.add_argument(
        "--out", default=None,
        help="Output path for fixed file (default: input_fixed.pptx)"
    )
    parser.add_argument(
        "--no-annotate", action="store_true",
        help="Skip in-deck flag markers and speaker notes (--fix only)."
    )
    parser.add_argument(
        "--clean-flags", action="store_true",
        help="Strip all flag markers/notes from an already-annotated file "
             "and save a clean copy. Use once review is done."
    )
    parser.add_argument(
        "--no-before-after", action="store_true",
        help="Skip inserting hidden 'before' copies ahead of edited slides "
             "(--fix only). On by default, matching the duplicate-and-hide "
             "review workflow."
    )
    args = parser.parse_args()

    input_path = args.input

    if not os.path.exists(input_path):
        print(f"ERROR: File not found: {input_path}")
        sys.exit(1)

    input_base = os.path.splitext(os.path.basename(input_path))[0]
    out_dir = os.path.dirname(args.out) if args.out else os.getcwd()

    # -------------------------------------------------------------------------
    # --clean-flags mode: strip annotations + before-slides, exit, skip analysis
    # -------------------------------------------------------------------------
    if args.clean_flags:
        clean_path = args.out or os.path.join(out_dir, f"{input_base}_clean.pptx")
        print(f"\nGemic Brand Fix — CLEAN-FLAGS mode")
        print(f"Input:  {input_path}")
        print(f"Output: {clean_path}\n")
        shutil.copy2(input_path, clean_path)
        prs = Presentation(clean_path)
        strip_all_flag_annotations(prs)
        n_removed = remove_before_slides(prs)
        prs.save(clean_path)
        print(f"Flags removed. {n_removed} before-slide(s) removed. "
              f"Clean file saved: {clean_path}")
        if sys.platform == "darwin":
            import subprocess
            subprocess.Popen(["open", "-a", "Microsoft PowerPoint", clean_path])
        return

    fix_mode   = args.fix
    audit_path = os.path.join(out_dir, f"{input_base}_audit.md")
    fixed_path = args.out or os.path.join(os.getcwd(), f"{input_base}_fixed.pptx")

    print(f"\nGemic Brand Fix — {'FIX' if fix_mode else 'ANALYZE'} mode")
    print(f"Input:  {input_path}")
    if fix_mode:
        print(f"Output: {fixed_path}")
    print(f"Audit:  {audit_path}")
    print()

    # Safety: work on a copy when in fix mode
    if fix_mode:
        shutil.copy2(input_path, fixed_path)
        work_path = fixed_path
    else:
        work_path = input_path

    prs = Presentation(work_path)
    log = Changelog()

    # Strip any flag annotations from a previous run BEFORE any analysis
    # sees the file — otherwise marker shapes (which have their own text
    # frames) get scanned as if they were real content and re-flagged,
    # compounding on every re-run.
    for slide in prs.slides:
        _strip_slide_annotation(slide)

    # Build the deck's own established type scale BEFORE font fixes run,
    # so recommendations reflect what the deck already uses. Font-name
    # fixes don't change sizes, so running this first vs. after op_fonts
    # makes no difference to the scale — but doing it here keeps all
    # analysis prep together in one place.
    established_sizes = collect_established_type_scale(prs)

    # Snapshot every real (non-before) slide's XML before any fix touches
    # it. Cheap in-memory deep copy — used later to build the hidden
    # "before" reference slides, without needing to undo any edits.
    snapshots = {}
    if fix_mode and not args.no_before_after:
        for slide_idx, slide in enumerate(prs.slides):
            if is_before_slide(slide):
                continue
            snapshots[slide_idx] = copy.deepcopy(slide._element)

    # Deck-level accumulators
    deck_styles    = set()
    per_slide_styles = {}

    # -------------------------------------------------------------------------
    # Slide-level passes
    # -------------------------------------------------------------------------
    n_slides = len(prs.slides)
    for slide_idx, slide in enumerate(prs.slides):
        slide_num = slide_idx + 1

        # Before-slides from a prior run are frozen historical references —
        # never analyze or fix them again.
        if is_before_slide(slide):
            continue

        slide_styles = set()

        print(f"  Processing slide {slide_num}/{n_slides}...", end="\r")

        op_fonts(prs, slide_num, slide, log, fix_mode)
        op_padding(prs, slide_num, slide, log, fix_mode)
        op_colors(prs, slide_num, slide, log, fix_mode)
        op_typography(prs, slide_num, slide, log, deck_styles, slide_styles,
                     established_sizes)
        op_line_spacing(prs, slide_num, slide, log, fix_mode)
        op_tables(prs, slide_num, slide, log, fix_mode)
        op_grid(prs, slide_num, slide, log, fix_mode)

        per_slide_styles[slide_num] = slide_styles

    print(f"  Processing complete ({n_slides} slides)       ")

    # -------------------------------------------------------------------------
    # Deck-level passes
    # -------------------------------------------------------------------------
    headline_sizes, headline_inconsistent = op_headline_consistency(
        prs, log, fix_mode
    )
    op_style_counts(log, deck_styles, per_slide_styles)

    # -------------------------------------------------------------------------
    # In-deck flag annotation (markers + speaker notes)
    # -------------------------------------------------------------------------
    if fix_mode and not args.no_annotate:
        annotate_flags_in_deck(prs, log, established_sizes)
        n_annotated_slides = len({e["slide"] for e in log.entries
                                   if e["action"] == "FLAGGED" and e["slide"] > 0})
        print(f"Flag markers + notes added to {n_annotated_slides} slide(s). "
              f"Run with --clean-flags once reviewed to remove them.")

    # -------------------------------------------------------------------------
    # Before/after slide duplication — insert hidden pre-fix copies ahead
    # of every slide that received at least one auto-fix
    # -------------------------------------------------------------------------
    if fix_mode and not args.no_before_after and snapshots:
        n_before = insert_before_slides(prs, log, snapshots)
        if n_before:
            print(f"Inserted {n_before} hidden 'before' slide(s) for "
                  f"comparison. Run --clean-flags once reviewed to remove them.")

    # -------------------------------------------------------------------------
    # Save fixed file and auto-open
    # -------------------------------------------------------------------------
    if fix_mode:
        prs.save(work_path)
        print(f"\nFixed file saved: {fixed_path}")
        if sys.platform == "darwin":
            import subprocess
            subprocess.Popen(["open", "-a", "Microsoft PowerPoint", work_path])
            print("Opening in PowerPoint...")

    # -------------------------------------------------------------------------
    # Write audit report
    # -------------------------------------------------------------------------
    report = build_report(
        input_path, log, fix_mode,
        deck_styles, per_slide_styles,
        headline_sizes, headline_inconsistent,
    )
    with open(audit_path, "w", encoding="utf-8") as f:
        f.write(report)

    # -------------------------------------------------------------------------
    # Console summary
    # -------------------------------------------------------------------------
    counts = log.counts()
    changed = sum(1 for e in log.entries if entry_is_autofix(e))
    flagged = sum(1 for e in log.entries
                  if e["action"] == "FLAGGED" and not entry_is_autofix(e))
    skipped = counts.get("SKIPPED", 0)

    print(f"\n{'='*50}")
    action = "Applied" if fix_mode else "Will apply (re-run with --fix)"
    print(f"  Auto-fixes {action}: {changed}")
    print(f"  Flagged for designer review: {flagged}")
    print(f"  Skipped:                  {skipped}")
    print(f"{'='*50}")
    print(f"\nFull report: {audit_path}")


if __name__ == "__main__":
    main()
