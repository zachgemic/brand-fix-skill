# Rules checked by brand_fix.py

Source of truth: the Atlas Design System (`tokens/colors.css`, `tokens/typography.css`, `engine/layout-rules.json`). Values below are what the script enforces.

| Rule | Atlas value | Action |
|---|---|---|
| Typeface | Inter. Regular 400 and Bold 700 only; Medium/SemiBold/ExtraBold/Black collapse to Bold, Light/Thin to Regular; italic preserved | Auto-fix (non-Inter fonts become Inter) |
| Text-box padding | 0 on all sides for unfilled shapes (filled shapes such as labels are left alone) | Auto-fix |
| Padding under a rule | Exactly 0.15 in top padding when a rule sits on the box's top edge; 0.20 in for 38 pt text (tolerance 0.01 in) | Flag |
| Colors | All 81 design system tokens (core, extended, 100-900 ramps, utility colors) are on-palette; anything farther than 15 RGB units from every token is off-palette. Taupe is #847D77 (#615F61 is "Grey") | Flag; pure #000000 fill becomes #1F1F24, pure white fill becomes Light Beige (auto-fix) |
| Type sizes | 8, 10, 12, 16, 20, 38 pt only. Minimum 8 pt | Flag, with the nearest Atlas size suggested |
| Line spacing | Multiple 1.17 for all text. Spacing inherited from the layout is left alone | Auto-fix explicit values |
| Table cell padding | 0.10 in (91,440 EMU) on all sides | Auto-fix |
| Table fonts and fill colors | Inter; palette colors | Auto-fix fonts; flag fills |
| Grid | 12 columns x 0.75 in, gutter 0.24 in, margins 0.845 in; five-column tracks of 2.136 in. Only left edges of text boxes, placeholders, pictures and tables are checked | Auto-snap text boxes within half a gutter; flag the rest |
| Headline sizes | Content-slide titles are 20 pt; 38 pt display titles (cover, divider, closing) are exempt | Flag only |
| Type styles | At most 8 distinct (size, weight) styles in the deck (the whole Atlas system), no per-slide limit beyond that | Flag |

Skipped on purpose: logo/Gemic-named shapes, slide-number/date/footer placeholders, pictures (for text rules), charts, shapes in the footer band (below 6.79 in), full-width backgrounds, and shapes inside filled cards for grid checks.

## Atlas rules this tool does NOT check

Headline position (0.96 in) and body top (1.96 in); content inside margins; text-box right edges and the 6-column width limit; spacing tiers (4 pt / 0.15 in / 0.35 in); bullet spacing; label styling and fill order; text colors (no extended hues, taupe only at 24 px+); italics/underline use; center alignment ban; image bleed and alignment rules; footer position; shadows, rounded corners and gradients; native-table and native-chart requirements; punctuation (typographic quotes, en dashes); ALL-CAPS limits; slide size.

A clean report therefore means "no issues in the rules above", not "fully on brand".
