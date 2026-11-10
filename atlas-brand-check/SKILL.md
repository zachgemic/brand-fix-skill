---
name: atlas-brand-check
description: Check and clean up a PowerPoint (.pptx) against the Gemic Atlas design system - fonts, text-box padding, colors, type sizes, line spacing, table cell padding, and 12/5-column grid alignment - and optionally auto-fix the mechanical issues. Use whenever the user uploads a deck and asks whether it is on brand, to brand-check, audit, lint, clean up, or "fix the formatting" of slides, mentions Atlas, the Gemic template, brand compliance, off-grid or off-palette elements, wrong fonts or spacing, or wants a designer's mechanical checking automated - even if they don't say "brand check". Produces an audit report and, only when asked, a corrected copy of the deck. Does not judge content or overall design quality.
---

# Atlas brand check

## Why this exists

Designers lose time to mechanical checking: is every font Inter, is padding zero, are boxes on the grid, are the colors in the palette. This skill runs `scripts/brand_fix.py`, which checks those things against the Atlas design system, so designers can spend their time on the slides themselves.

The design system is the source of truth, and the script's constants were copied from it (`references/atlas-rules.md` lists every rule, its value, and whether it is auto-fixed or only flagged). If a user disagrees with a flag, point them to the rule in the reference and suggest changing the design system first; don't bend the result to please the user, and don't edit the script's values without them confirming the design system changed.

The script never edits text content, data, or slide meaning. It only touches font names, text-box margins, fill colors, line spacing, table cell padding, and (for small drift) text-box position.

## Procedure

Work on a copy. The script writes its audit and fixed files into the current directory, so run it from the outputs folder, and never overwrite the user's original.

Dependencies: `pip install python-pptx --break-system-packages` if missing.

### 1. Prove the checker works (once per session, seconds)

```bash
python scripts/selftest.py
```

It must print `SELFTEST PASSED`. It builds a deck with compliant Atlas elements (which must never be flagged or moved) and seeded violations (which must be flagged or fixed). If it fails, stop and tell the user rather than reporting on their deck.

### 2. Analyze (default; changes nothing)

```bash
python scripts/brand_fix.py deck.pptx
```

This writes `deck_audit.md`. Read it. Don't paste the whole thing: it can run to hundreds of rows. Summarize by category, with counts and a few representative examples each:

- **Auto-fixable** (the script can correct these safely): non-Inter fonts or weight variants, text-box padding on unfilled shapes, pure black/white fills, line spacing, table cell padding, and small grid drift on text boxes.
- **Flagged for the designer** (a judgment call, never auto-changed): off-palette fills and strokes, font sizes that aren't Atlas type roles, shapes meaningfully off the grid, padding under a rule that isn't 0.15 in, headline size differences.

Lead with the answer: how many issues, how many the tool can fix, how many need the designer. Then the per-category summary. Name slide numbers and shape names so the designer can find each one.

### 3. Fix only when asked

If the user wants the fixes applied (or answers yes to your offer), run:

```bash
python scripts/brand_fix.py deck.pptx --fix --out deck_fixed.pptx
```

Tell the user what this does to the file, because it surprises people:

- It adds numbered flag markers on shapes needing review, plus a matching list in each slide's speaker notes (`--no-annotate` skips these).
- It inserts a **hidden "before" copy** of every slide it edited, directly before the edited slide, matching the designers' duplicate-and-hide workflow (`--no-before-after` skips these).
- When the designer has finished reviewing, `python scripts/brand_fix.py deck_fixed.pptx --clean-flags` removes all markers, notes, and before-slides.

Offer `--no-annotate --no-before-after` when the user just wants a clean corrected file.

### 4. Verify the fix

Re-run the analysis on the fixed file. The auto-fixable count should now be zero, and the remaining items should be exactly the flagged-for-designer ones. If anything auto-fixable remains, say so and investigate; don't report success.

If the deck has hidden before-slides and the user wants proof the text wasn't touched, the `slide-content-qa` skill (if installed) checks each hidden-original/edited pair for content changes.

### 5. Report

Give: the verdict in one line, the per-category summary, what was fixed (if anything), what still needs the designer, and the file(s) produced. Be precise about scope: this checks mechanical brand formatting against the rules in the reference. It does not check content, hierarchy, imagery, layout quality, or the Atlas rules the script doesn't cover (listed in the reference). So never say a deck is "fully on brand"; say it has no issues in the checked categories.

## Judgment notes

- The grid check accepts both 12-column edges and five-column track edges; a text box that sits on either is fine.
- 38 pt display text, 10 pt Bold labels, orange/teal/purple/yellow/pink 700-step label fills, taupe #847D77, and 0.10 in table padding are all correct Atlas values. If the report flags one of them, something is wrong with the check, so investigate before reporting it as a designer error.
- Shapes named logo/Gemic, slide-number, date and footer placeholders, charts, and pictures are skipped on purpose.
