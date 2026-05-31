"""Bizzdesign-tailored composition tools.

Each composition wraps a specific source slide from the bundled brand
template (`Bizzdesign_CorporateDeck_2026.pptx`). The model picks a
composition by name and hands over content as named fields. The server
clones the source slide's shape tree into a freshly-added slide on the
working presentation and text-replaces each tagged shape with the
supplied content.

This is the deliberate alternative to letting the model assemble shapes
from primitives. The model can't see what it builds, so we ship pre-built
compositions proven by the template's own example slides — the same
visual quality bar the brand designer set, with the model just supplying
text.
"""
from copy import deepcopy
from io import BytesIO
from typing import Any, Dict, List, Optional
from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations
from pptx import Presentation
from pptx.util import Emu, Inches


# ------------------------------------------------------------------
# Composition catalogue
# ------------------------------------------------------------------
# Each composition declares:
#   source_slide_index: 0-based index into the brand template (None = no clone, uses layout directly)
#   description: short blurb (returned by list_compositions)
#   use_when: when to pick this composition (returned to model)
#   fields: dict of field_name -> { "match": <prefix>, optional "required": bool }
#           For repeatable groups (e.g. details, pillars), the field value is a list of matchers.
#
# Matchers: we locate shapes by the leading text content in the template,
# because (a) it's stable across deepcopies, (b) shape names are reused
# across shapes in the template (e.g. "Text Placeholder 6" appears more
# than once on a single slide), and (c) it's the most semantically clear.

COMPOSITIONS: Dict[str, Dict[str, Any]] = {
    "cover": {
        "source_slide_index": 0,
        "description": "Opening cover slide: title + subtitle on the brand cover background.",
        "use_when": "First slide of every deck.",
        "fields": {
            "title": {"match": "Corporate Deck", "required": True},
            "subtitle": {"match": "2026", "required": True},
        },
    },
    "section_divider": {
        "source_slide_index": 4,
        "description": "Section break: small kicker text above a large title.",
        "use_when": "Between major sections of a deck (rare in 7-slide decks; useful in longer ones).",
        "fields": {
            "kicker": {"match": "What We Offer", "required": True},
            "title": {"match": "Transformation That Flows", "required": True},
        },
    },
    "value_props_4": {
        "source_slide_index": 8,
        "description": "Lead claim + 4 value-proposition pillars + subtitle. Right-side stack of 4 short pillars.",
        "use_when": "4 parallel differentiators or value propositions. Don't use for fewer than 4.",
        "fields": {
            "title": {"match": "Why", "required": True},
            "lead_claim": {"match": "Only true end-to-end", "required": True},
            "subtitle": {"match": "The combined resources", "required": True},
            "pillars": [
                {"match": "Proven Market Leadership"},
                {"match": "Trusted by 2,000+ Enterprises"},
                {"match": "Continuous AI-driven innovation"},
                {"match": "Global reach, local expertise"},
            ],
        },
    },
    "solution_detail": {
        "source_slide_index": 13,
        "description": "Category kicker + offering title + benefits paragraph + EXACTLY 3 detail items.",
        "use_when": "A specific named topic that has exactly 3 named sub-features AND a 'why this matters' summary. NOT for generic 3-item lists — for those use `add_bullets_slide`. NOT for high-level claims — for those use `add_value_props_slide`. Don't use this composition more than 3 times in a row.",
        "fields": {
            "category": {"match": "Transformation Planning", "required": True},
            "title": {"match": "Strategic Portfolio Management", "required": True},
            "benefits": {"match": "Align investments with strategic goals", "required": True},
            "details": [
                {"match": "Business Strategy Alignment"},
                {"match": "IT Investment Optimization"},
                {"match": "Strategic"},
            ],
        },
        # Template slide 14 carries 3 product-UI screenshots of Bizzdesign
        # Unify. Those screenshots are content-coupled — they sell the
        # specific product feature this template slide describes. On any
        # other deck topic ("Squad breakdown", "Architecture decisions",
        # etc.) they show up as visually-impressive but completely off-
        # topic dashboard images. Strip them on this composition.
        "strip_pictures": True,
        # Iter 66: template's 3 product UI screenshots filled y=3.14→5.60.
        # After strip_pictures, the detail textboxes (y=5.54) sit awkwardly
        # at the bottom of the slide with ~2.4in of empty whitespace above
        # them. Shift each detail role-tagged shape up by 2.4in (2.4 *
        # 914400 EMU/in = 2194560 EMU) so they sit just below the benefits
        # paragraph (which ends at y=3.00).
        "shift_role_y_emu": {
            "details": -2194560,
        },
    },
    "stat_cards": {
        "source_slide_index": 10,  # slide 11
        "description": "Subhead + intro line + 4 big-number stat cards (each: value, label, attribution).",
        "use_when": "When you have 3-4 NUMBERS that tell the story (revenue, customer counts, time savings, capability counts, team counts, dates). Numbers make decks land — pick this composition any time the content has stats. Each stat is one string: 'VALUE\\nLABEL\\nATTRIBUTION' (e.g. '60+\\nCapabilities\\nAcross 5 platform layers').",
        "fields": {
            "subhead": {"match": "Our Impact", "required": True},
            "intro": {"match": "See how customers", "required": True},
            "stats": [
                {"match": "£4.35M"},  # £4.35M
                {"match": "66%"},
                {"match": "$5.5M"},
                {"match": "15%"},
            ],
        },
        # Post-process tightening: the 40pt VALUE paragraph reserves a
        # ~44pt line-box even for single-digit glyphs like "6". This
        # registers as a visible gap between VALUE and LABEL. Compress
        # the line spacing on the first paragraph of each stat shape so
        # short values sit closer to their label.
        "tighten_first_paragraph_line_spacing": {
            "shape_prefixes": ["£4.35M", "66%", "$5.5M", "15%"],
            "spc_pct": 80000,  # 80%
        },
    },
    "capability_grid": {
        "source_slide_index": 12,  # slide 13
        "description": "Section subhead + cross-cutting label + intro + 3 column headings + 9 capability cards in a 3x3 grid (each card has a heading AND a short description).",
        "use_when": "When the content is a MAP of many capabilities organised into 3 columns, with multiple sub-items per column. Each card is 'Heading\\nOne-line description'. Up to 9 cards (3 per column).",
        # Iter 70: when n=3 or n=6 cards, the bottom row(s) of each
        # column are empty but the column-background rectangles still
        # extend to y=7.50in (slide bottom). User reported the result
        # looks "off, missing background images" because of the large
        # empty grey area at the bottom of each column. Shrink the 3
        # column backgrounds (Rectangle 3/5/6) to cover only the rows
        # that have content.
        "cap_grid_column_bg_shapes": ["Rectangle 3", "Rectangle 5", "Rectangle 6"],
        "fields": {
            "subhead": {"match": "Bizzdesign", "required": True},
            "cross_label": {"match": "Transformation Collaboration", "required": True},
            "cross_intro": {"match": "Align teams around", "required": True},
            "col_headings": [
                {"match": "Transformation Planning"},
                {"match": "Transformation Design"},
                {"match": "Transformation Governance"},
            ],
            # 9 paired cards: heading + description per card, ordered
            # column-by-column (top to bottom within each column).
            "cards": [
                {"heading_match": "Strategic Portfolio Management", "description_match": "Align investments with strategic goals"},
                {"heading_match": "Application Portfolio Management", "description_match": "Reduce IT costs"},
                {"heading_match": "Technology Portfolio Management", "description_match": "Stay ahead of obsolescence"},
                {"heading_match": "Enterprise Architecture Management", "description_match": "Strengthen governance"},
                {"heading_match": "Business Architecture Management", "description_match": "Turn insight into execution"},
                {"heading_match": "Solution Architecture Management", "description_match": "Accelerate design with business context"},
                {"heading_match": "Business Process Management", "description_match": "Increase process visibility"},
                {"heading_match": "Data Management", "description_match": "Create a common data language"},
                {"heading_match": "Governance, Risk & Compliance", "description_match": "Improve compliance visibility"},
            ],
        },
    },
    "bullets": {
        # Special: built directly from the Basic Text layout, not cloned
        # from a source slide. The Basic Text layout has placeholder idx 13
        # (subhead) and idx 14 (body bullets). The layout sets <a:buNone/>
        # on the body — we override per-paragraph with explicit buChar so
        # bullets actually render.
        "source_slide_index": None,
        "layout_name": "Basic Text",
        "description": "Title + small subhead + brand-styled bullet list.",
        "use_when": "Short list of items WITHOUT sub-detail (3-6 bullets, ≤ 15 words each). Use for asks / takeaways / quick lists / agenda items. Reach for this BEFORE solution_detail when the content is just a list, not a 'topic with 3 features'.",
        "fields": {
            "title": "layout_title",
            "subhead": "placeholder_idx_13",
            "bullets": "placeholder_idx_14",
        },
    },
    "process_steps": {
        "source_slide_index": 27,  # slide 28 (0-indexed)
        "description": "Subhead + title + intro paragraph + 4 sequential steps in a horizontal flow (with arrows between).",
        "use_when": "When the content is a SEQUENCE of 4 stages or steps (onboarding flow, project phases, journey stages). Each step is `Name\\nShort description`. Use for any chronological or workflow narrative. Don't use for parallel concepts — that's `value_cards` or `value_props_4`.",
        "fields": {
            "subhead": {"match": "Customer Onboarding", "required": True},
            "title_and_intro": {"match": "Hit the Ground Running", "required": True},
            "steps": [
                {"match": "Kick-Off"},
                {"match": "Set-Up"},
                {"match": "Support"},
                {"match": "Stay Connected"},
            ],
        },
    },
    "timeline": {
        "source_slide_index": 3,  # slide 4 (0-indexed) — "Our Growth Story" year-by-year
        "description": "Subhead + 4-12 chronological events on a horizontal timeline with year/date markers.",
        "use_when": "When the content is a chronological story (founding history, product launches, hiring sequence, contract milestones, customer journey by year). Each event is `Date Description` in one string — the year/date at the start anchors the marker. Events 4-12, each ≤ 12 words.",
        # Iter 71: when n<12 events, delete the Oval markers + dashed
        # connector lines for empty slots so only the dots that have
        # labels remain. User reported "There are dots with no label
        # shouldn't be there" on a 4-event delivery-phases roadmap.
        "timeline_clean_empty_markers": True,
        "fields": {
            "subhead": {"match": "Our Growth Story", "required": True},
            "events": [
                {"match": "2000"},
                {"match": "2004"},
                {"match": "2009"},
                {"match": "2014"},
                {"match": "2015"},
                {"match": "2019"},
                {"match": "2022"},
                {"match": "2023"},
                {"match": "2024"},
                {"match": "Jan 2025"},
                {"match": "Sept 2025"},
                {"match": "Apr 2026"},
            ],
        },
    },
    "split_benefits": {
        "source_slide_index": 25,  # slide 26 (0-indexed) — "Optimize and Reinforce your Operations"
        "description": "Split layout: large brand title (left), 4-6 short benefit lines over a stock background image (right).",
        "use_when": "When you want to land a benefit list visually — the right-hand stock background gives the slide weight. Best for 'here's what you get' / 'why this matters' moments mid-deck. The title is a verb-led claim about the section; benefits are 4-6 short outcome statements (5-12 words each).",
        "fields": {
            "title": {"match": "Optimize and Reinforce", "required": True},
            "benefits": {"match": "Achieve 360", "required": True},
        },
    },
    "diagram": {
        # Built directly from the Basic Text layout: title placeholder
        # at the top, body placeholder replaced with the rendered PNG.
        "source_slide_index": None,
        "layout_name": "Basic Text",
        "description": "Title + a server-rendered Mermaid diagram embedded as PNG. Optional caption.",
        "use_when": "When you need to show STRUCTURE, FLOW, or RELATIONSHIPS visually — architecture diagrams, sequence flows, org charts, decision trees, dependency graphs. Supply Mermaid source text; the server renders it and embeds the result. Don't use for static lists or bullets — those are bullets/solution_detail.",
        # Handled by add_diagram_slide directly, not by _apply_fields.
        "fields": {},
    },
    "chart": {
        "source_slide_index": None,
        "layout_name": "Basic Text",
        "description": "Title + a server-rendered chart (bar / line / pie / doughnut) embedded as PNG. Optional caption.",
        "use_when": "When you need to show QUANTITIES — revenue over time, market share split, growth comparisons, before/after metrics. Supply a Chart.js v4 config (chart_type + data); the server renders via QuickChart and embeds the result. For 4 standalone stat numbers WITHOUT axis context, prefer add_stat_cards_slide.",
        "fields": {},
    },
    "closing": {
        "source_slide_index": 36,  # slide 37 (0-indexed) — "Thank You" + CTA list
        "description": "Final 'Thank you' slide with title + 2–6 short CTA / contact lines.",
        "use_when": "ALWAYS the last slide of every deck. Carries the closing message (recommendation, ask, or thank-you) plus contact / next-step lines. This is the slide audiences look at while you wrap — make it count.",
        "fields": {
            "title": {"match": "Thank You", "required": True},
            "cta_lines": {"match": "Book a Demo", "required": True},
        },
        # Iter 57: template's cta_lines render at 10pt — too small for
        # board / leadership decks. Visual judge (iter 55 platform-rearch
        # slide 9) flagged "small font size makes bullets hard to read".
        # Bumping to 14pt — combined with the iter 57 cap tightening to
        # ≤ 50 chars per line, the larger text still fits the column.
        "set_run_sz_for_role": {
            "cta_lines": 1400,
        },
    },
    "value_cards": {
        "source_slide_index": 29,  # slide 30 (0-indexed)
        "description": "Subhead + title + 4 icon-topped cards, each with a heading and rich description.",
        "use_when": "When you want to showcase 4 parallel benefits / value props / pillars WITH rich descriptions (~12-25 words each). The 4 cards are visually distinct (each with an icon on top) — strong payoff slide. Differs from value_props_4 (short pillar text only) by allowing real descriptions per card. Differs from capability_grid (smaller cards, more of them) by being more visual / impactful.",
        "fields": {
            "subhead": {"match": "Why Partner with", "required": True},
            "title": {"match": "Accelerate customer impact", "required": True},
            "cards": [
                {"match": "Expand your footprint"},
                {"match": "Build Long-Term Sticky Relationships"},
                {"match": "Continuous Innovation Advantage"},
                {"match": "Successful"},  # "Successful | Implementations | End to End..."
            ],
        },
        # Card 1's heading is in a SEPARATE shape ("Footprint", TextBox 29) —
        # cards 2-4 have heading+description combined in one shape. Clearing
        # the standalone heading shape so card 1's new content (which goes
        # entirely into TextBox 21) doesn't conflict with the leftover.
        "clear_text_starting_with": ["Footprint"],
    },
}


# ------------------------------------------------------------------
# Locator + shape walking
# ------------------------------------------------------------------

def _shape_text(shape) -> str:
    """Concatenate all text runs in a shape's text frame (single string)."""
    if not shape.has_text_frame:
        return ""
    parts = []
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            if run.text:
                parts.append(run.text)
    return "".join(parts)


_UNICODE_ESCAPE_RE = __import__("re").compile(r"\\u([0-9A-Fa-f]{4})")


def _normalise_text(s: str) -> str:
    """Normalise common LLM escape artefacts in incoming text.

    - Two-char literal `\\n` / `\\t` → real newline / tab (models sometimes
      pass these escape sequences when they think they're producing
      JSON-escaped strings).
    - Literal `\\uXXXX` escape sequences → the corresponding Unicode
      codepoint. Models occasionally double-escape special chars like
      em-dash ("\\u2014") in tool args, which arrives at the server as
      a 6-char literal string instead of the intended single character.
      Decode them here so rendered slides show "—" not "\\u2014".
    - Multiple consecutive newlines collapse to a single newline (extra
      blank lines in stat cards / multi-tier shapes render as visible
      blank paragraphs and break the layout).
    - Leading/trailing whitespace stripped.
    """
    if not isinstance(s, str):
        return s
    s = s.replace("\\n", "\n").replace("\\t", "\t")
    s = _UNICODE_ESCAPE_RE.sub(lambda m: chr(int(m.group(1), 16)), s)
    # Collapse runs of \n\n+ (and \n \n etc.) to single \n. Trim per-line
    # whitespace so " \n " also collapses cleanly.
    parts = [p.strip() for p in s.split("\n")]
    # Drop empty parts produced by collapsing, but keep one between any
    # two non-empty parts (i.e. join with single \n).
    parts = [p for p in parts if p]
    return "\n".join(parts)


def _set_shape_text(shape, new_text: str) -> None:
    """Replace a shape's entire text content, preserving the styling of the
    FIRST run in the FIRST paragraph (font, size, colour, bold, etc.).

    python-pptx-friendly approach: keep the first paragraph + first run,
    overwrite its text, delete any subsequent runs / paragraphs.
    """
    if not shape.has_text_frame:
        return
    new_text = _normalise_text(new_text)
    tf = shape.text_frame
    # Keep first paragraph; delete the rest.
    paragraphs = tf.paragraphs
    if not paragraphs:
        tf.text = new_text
        return
    first_para = paragraphs[0]
    # Delete all paragraphs after the first by manipulating XML.
    for para in list(paragraphs[1:]):
        para._p.getparent().remove(para._p)
    # Within the first paragraph, keep only the first run.
    runs = first_para.runs
    if not runs:
        first_para.text = new_text
        return
    first_run = runs[0]
    for run in list(runs[1:]):
        run._r.getparent().remove(run._r)
    first_run.text = new_text


def _set_shape_multiline(shape, lines: List[str]) -> None:
    """Replace a shape's text with multiple lines (one paragraph each),
    preserving EACH ORIGINAL PARAGRAPH'S styling per-line.

    The template's stat card (and similar multi-tier compositions) have
    distinct font sizes per paragraph: huge VALUE on line 1, medium
    LABEL on line 2, small ATTRIBUTION on line 3. Cloning the first
    paragraph's styling onto every new line dumps the giant VALUE font
    onto the LABEL and causes overflow.

    Algorithm:
      - For i in range(min(len(lines), len(original_paragraphs))):
          replace the text of paragraph i, keep its run styling
      - If new content has MORE lines than original paragraphs, clone the
        LAST original paragraph's style for the overflow.
      - If new content has FEWER lines, delete the extra paragraphs.
    """
    if not shape.has_text_frame or not lines:
        _set_shape_text(shape, "")
        return
    lines = [_normalise_text(l) for l in lines]
    tf = shape.text_frame
    original_paras = list(tf.paragraphs)
    if not original_paras:
        tf.text = lines[0]
        if len(lines) > 1:
            for line in lines[1:]:
                p = tf.add_paragraph()
                p.text = line
        return

    a_ns = "http://schemas.openxmlformats.org/drawingml/2006/main"

    # Iter 54: detect source paragraphs that use `<a:br/>` to separate
    # multiple text runs at DIFFERENT styles (e.g. process_steps source
    # `TextBox 4` has heading-run + <a:br/> + description-run with
    # sz=1400 in a single paragraph). Treat each run-segment as its own
    # "line slot" and apply the user's lines preserving each segment's
    # per-run styling. Without this, the function would keep only the
    # FIRST run's style and apply it to every line, making the
    # description render at the heading's larger font and overflow the
    # step box (user complaint 2026-05-30 on slide 4 of the AI
    # capability deck).
    def _segment_count(p_el):
        """Count run-or-br segments in this paragraph that we can target.
        A run between brs counts as a segment; an empty <a:br/> counts
        as a separator, not a segment. Pure-run-no-br paragraphs return 1.
        """
        children = [c for c in p_el if c.tag.endswith("}r") or c.tag.endswith("}br")]
        if not children:
            return 0
        # Count contiguous run blocks separated by br.
        segments = 0
        in_run_block = False
        for c in children:
            if c.tag.endswith("}r"):
                if not in_run_block:
                    segments += 1
                    in_run_block = True
            elif c.tag.endswith("}br"):
                in_run_block = False
        return segments

    def _set_segments_preserving_styles(p_el, lines):
        """For a paragraph with N <a:br/>-separated run-segments, split
        the segments into separate <a:p> paragraphs and apply lines[i]
        to the i-th paragraph, preserving each segment's first-run rPr.

        Iter 55 change: previously this helper wrote back as a SINGLE
        paragraph with <a:br/> between segments. That produced tighter
        line-spacing than sibling shapes that legitimately used 2
        paragraphs (e.g. process_steps step 0 came out visibly
        compressed against steps 1–3, which the template authored as
        2 paragraphs). Splitting into separate paragraphs makes step 0
        match the others' line-spacing while still preserving each
        segment's rPr (so the description still renders at 14pt).

        Returns True if all lines were applied (one per segment slot),
        False otherwise (caller falls through to per-paragraph path).
        """
        from lxml import etree

        children = [c for c in p_el if c.tag.endswith("}r") or c.tag.endswith("}br") or c.tag.endswith("}endParaRPr")]
        # Group runs into segments, separated by brs.
        segments = []   # list of run-element-lists
        cur = []
        for c in children:
            if c.tag.endswith("}r"):
                cur.append(c)
            elif c.tag.endswith("}br"):
                if cur:
                    segments.append(cur)
                    cur = []
            elif c.tag.endswith("}endParaRPr"):
                if cur:
                    segments.append(cur)
                    cur = []
        if cur:
            segments.append(cur)
        if not segments:
            return False
        n_apply = min(len(lines), len(segments))
        if n_apply == 0:
            return False

        # First segment stays in p_el. Strip the br separators from p_el
        # and remove everything past the first segment's first run.
        first_run = segments[0][0]
        # Set the first segment's text.
        t = first_run.find(f"{{{a_ns}}}t")
        if t is None:
            t = etree.SubElement(first_run, f"{{{a_ns}}}t")
        t.text = lines[0]
        # Drop sibling runs after the first run, and drop brs.
        for c in list(p_el):
            if c is first_run:
                continue
            if c.tag.endswith("}r"):
                p_el.remove(c)
            elif c.tag.endswith("}br"):
                p_el.remove(c)

        # Build a new <a:p> for each remaining segment we need to apply.
        parent = p_el.getparent()
        insert_at = list(parent).index(p_el)
        for seg_idx in range(1, n_apply):
            seg_runs = segments[seg_idx]
            seg_first_run = seg_runs[0]
            # Construct a fresh <a:p> with optional pPr cloned from the
            # original paragraph (so alignment / margins carry across).
            new_p = etree.SubElement(parent, f"{{{a_ns}}}p")
            parent.remove(new_p)
            orig_ppr = p_el.find(f"{{{a_ns}}}pPr")
            if orig_ppr is not None:
                new_p.append(deepcopy(orig_ppr))
            # Clone the segment's first run (preserving its rPr) and
            # set its text to the user's line.
            new_run = deepcopy(seg_first_run)
            new_t = new_run.find(f"{{{a_ns}}}t")
            if new_t is None:
                new_t = etree.SubElement(new_run, f"{{{a_ns}}}t")
            new_t.text = lines[seg_idx]
            new_p.append(new_run)
            insert_at += 1
            parent.insert(insert_at, new_p)

        return n_apply == len(lines)

    def _set_para_text_preserving_style(p_el, new_text: str) -> None:
        """Set a paragraph's text to `new_text` while preserving the
        styling of its first run (font, size, colour, bold, etc.)."""
        runs = p_el.findall(f"{{{a_ns}}}r")
        if not runs:
            # No <a:r> children — just set the paragraph text directly.
            # Strip any existing inline text first.
            for child in list(p_el):
                if child.tag.endswith("}r") or child.tag.endswith("}br") or child.tag.endswith("}fld"):
                    p_el.remove(child)
            # python-pptx-style: add a single run with the text.
            from lxml import etree
            r = etree.SubElement(p_el, f"{{{a_ns}}}r")
            t = etree.SubElement(r, f"{{{a_ns}}}t")
            t.text = new_text
            return
        # Keep the first run (style preserved), drop the rest, update text.
        first_run = runs[0]
        for r in runs[1:]:
            p_el.remove(r)
        t = first_run.find(f"{{{a_ns}}}t")
        if t is None:
            from lxml import etree
            t = etree.SubElement(first_run, f"{{{a_ns}}}t")
        t.text = new_text

    # Iter 54: if the source has a SINGLE paragraph with multiple
    # <a:br/>-separated run-segments at different rPr styles (e.g.
    # process_steps' TextBox 4 has heading-run + br + 14pt description-run
    # in one paragraph), map our lines onto those segments preserving
    # each segment's style. Falls through to the normal per-paragraph
    # path when source has multiple paragraphs or single-segment paragraphs.
    if (
        len(original_paras) == 1
        and _segment_count(original_paras[0]._p) >= 2
        and _set_segments_preserving_styles(original_paras[0]._p, lines)
    ):
        return

    # Update existing paragraphs in place (preserves per-paragraph styling).
    for i in range(min(len(lines), len(original_paras))):
        _set_para_text_preserving_style(original_paras[i]._p, lines[i])

    if len(lines) > len(original_paras):
        # Clone the LAST original paragraph for any overflow lines.
        last_p = original_paras[-1]._p
        for extra_line in lines[len(original_paras):]:
            new_p = deepcopy(last_p)
            _set_para_text_preserving_style(new_p, extra_line)
            tf._txBody.append(new_p)
    elif len(lines) < len(original_paras):
        # Delete extra paragraphs from the txBody when the user supplies
        # fewer lines than the template had. Previously we cleared the
        # text (set to "") which preserved styling but left visible blank
        # lines at the bottom of the rendered shape — caught by the
        # composition-variety probe added in iter 16 (closing.cta_lines
        # had 2 trailing blanks). Deleting the `<a:p>` element entirely
        # removes the blank-line artifact.
        #
        # Safety: each composition call is a fresh slide — we never need
        # to "restore" the deleted paragraphs later. Multi-tier shapes
        # like stat_cards always receive all their paragraphs (the
        # composition tool's validator enforces VALUE + LABEL minimum)
        # so this won't accidentally collapse them.
        for extra_p in original_paras[len(lines):]:
            extra_p._p.getparent().remove(extra_p._p)


def _rename_shape(shape, role: str) -> None:
    """Rename the shape so downstream tooling (the assessment harness'
    LLM judge in particular) can identify which field role this shape
    plays without inferring from font size or position.

    The new name is prefixed with `role:` so it doesn't collide with
    legitimate template shape names. python-pptx exposes `shape.name`
    as a settable property that writes to the underlying `<p:cNvPr>`
    `name` attribute, which OOXML treats as a free-form label not
    visible in any rendered slide chrome.
    """
    try:
        shape.name = f"role:{role}"
    except Exception:
        # Some shapes (e.g. group containers without nvSpPr) reject the
        # rename — fall through silently rather than abort the build.
        pass


def _widen_title_full_width(slide, working_pres, side_margin_in: float = 0.5) -> None:
    """Resize the slide's title placeholder to span the full slide width.

    The Basic Text layout title placeholder is narrower than the slide
    by default (the template designed it for short titles). Real titles
    routinely run 8-12 words and wrap to two lines in that narrow box,
    eating vertical space and shoving body content downward. Stretching
    the title to full slide width (minus a small side margin) lets most
    titles fit on a single line without changing font size.

    Only applied to layout-built slides (bullets, diagram, chart). The
    cloned compositions inherit their title geometry from the template
    on purpose — those slides are visually tuned with the title sitting
    beside other shapes.
    """
    title_shape = slide.shapes.title
    if title_shape is None:
        return
    try:
        slide_w = working_pres.slide_width
        margin = Inches(side_margin_in)
        # Resolve current position from the placeholder's effective
        # frame (falls back through the layout/master if the slide
        # hasn't overridden it). We MUST set all four (left/top/width/
        # height) — touching one writes <a:off>/<a:ext> on the slide
        # which OVERRIDES the layout entirely; missing axes default to
        # zero and the title slides up out of the visible area.
        current_top = title_shape.top
        current_height = title_shape.height
        if current_top is None or current_height is None:
            # Layout-fallback defaults for the Basic Text layout: title
            # sits ~0.45in from the top and is ~1.0in tall.
            current_top = Inches(0.45) if current_top is None else current_top
            current_height = Inches(1.0) if current_height is None else current_height
        title_shape.left = margin
        title_shape.top = current_top
        title_shape.width = slide_w - 2 * margin
        title_shape.height = current_height
    except Exception:
        # Some title shapes refuse geometry edits (rare); fail soft so
        # the build continues with the template-default narrow title.
        pass


def _find_shape_by_match(shapes, match_prefix: str, original_texts=None, consumed=None):
    """Return the first shape whose ORIGINAL text starts with the given
    prefix (case-sensitive, stripped). None if not found.

    `original_texts` is a dict mapping id(shape) → snapshot of the shape's
    text before any mutation. This must be supplied when matching across
    multiple fields in one composition, because one field's new value
    can collide with another field's matcher prefix (e.g. cover title
    "2026 is a regime change..." starts with "2026" and would match the
    subtitle matcher after being applied). Falling back to the live text
    is the legacy path used for first-match-only callers.

    `consumed` is a set of id(shape) values already claimed by another
    field in this pass — those shapes are skipped.
    """
    target = match_prefix.strip()
    for shape in shapes:
        if consumed is not None and id(shape) in consumed:
            continue
        if original_texts is not None:
            text = original_texts.get(id(shape), "").strip()
        else:
            text = _shape_text(shape).strip()
        if text.startswith(target):
            return shape
    return None


def _clear_shape_text(shape) -> None:
    """Empty a shape's text frame (used when a repeatable field has fewer
    entries than the template provides).
    """
    if shape.has_text_frame:
        _set_shape_text(shape, "")


# ------------------------------------------------------------------
# Slide cloning
# ------------------------------------------------------------------

def _clone_pic_into(pic_element, src_slide, dst_slide):
    """Clone a `<p:pic>` element from `src_slide` into `dst_slide`,
    rewiring EVERY `r:embed` / `r:link` attribute so the new picture
    points at the same media part(s) via fresh relationships on the
    destination slide.

    Iter 12 added this helper but only rewired refs on `<a:blip>`
    elements. That worked for plain pictures but missed two nested
    extensions that the Bizzdesign template uses heavily:

    1. `<asvg:svgBlip r:embed="...">` — Office 2016+ SVG sibling
       reference, sitting inside `<a:blip>/<a:extLst>/<a:ext
       uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">/<asvg:svgBlip>`.
       Carries an extra rId pointing at the SVG version of the same
       graphic (e.g. brand-icon vector for crisp scaling).
    2. `<a14:imgLayer r:embed="...">` — Office 2010 picture-effects
       layer, sitting inside `<a:blip>/<a:extLst>/<a:ext
       uri="{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}">/<a14:imgLayer>`.
       Carries an rId pointing at the un-effected source image so the
       picture-effects (brightness, sharpening, crop) can reapply on
       open.

    Both extensions are deep-copied along with the rest of the picture
    XML, but their rIds reference the SOURCE slide's rel graph — not
    the destination's. Result: PowerPoint opens the file, can't resolve
    the embed, prompts "this file needs to be repaired", and on accept
    silently drops the picture. Same failure root cause as iter 12's
    original `_clone_slide_into` skipping pictures entirely.

    Fix: walk every descendant element of the cloned `<p:pic>` and for
    every element that has an `r:embed` or `r:link` attribute, rewire
    it via the same `dst_part.relate_to` dance the iter-12 code did
    for `<a:blip>`. Generic over namespace — works for `<a:blip>`,
    `<asvg:svgBlip>`, `<a14:imgLayer>`, and any future extension that
    follows the OOXML relationship-attribute convention.

    Driven by user feedback 2026-05-29: generated decks were prompting
    "this file needs to be repaired" on first open in Microsoft
    PowerPoint despite iter 41's save-time zip cleanup. Diagnosis on
    the user's example deck found dangling r:embed refs on slides 1,
    3, and 7 — exactly the picture-rich slides (cover, stat_cards,
    value_cards). Now caught by the per-slide `embed-refs-resolve`
    probe in assess.ts as a regression guardrail.
    """
    from copy import deepcopy
    R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    EMBED_ATTR = f"{{{R_NS}}}embed"
    LINK_ATTR = f"{{{R_NS}}}link"

    cloned = deepcopy(pic_element)
    _rewire_rids(cloned, src_slide.part, dst_slide.part)
    return cloned


def _rewire_rids(el, src_part, dst_part):
    """Walk a shape XML subtree and rewire every relationship-attribute
    (r:embed, r:link, r:id) so refs point at fresh relationships on the
    destination slide.

    Iter 43 introduced this for <p:pic> (image embeds + SVG sibling +
    Office picture-effects layer). Iter 53 extends it to ALL cloned
    shapes: <p:sp> shapes can carry <a:hlinkClick r:id="rIdN"> for
    hyperlinks, <a:hlinkMouseOver>, <p:custDataLst r:id="...">,
    chart-data refs, OLE embeds, audio/video media, etc. All of those
    use the same relationship-attribute convention; without rewiring,
    the cloned shape carries source-slide rIds that don't exist in the
    destination slide's rels graph, triggering Microsoft's repair
    prompt.

    Driven by user feedback 2026-05-30: closing slide kept asking for
    repair because the template's "Book a Demo / Partner Program /
    Careers" CTA buttons carry <a:hlinkClick r:id="rId2..rId5">
    pointing at external URLs — those rels never reached the
    destination slide because the clone path only rewired r:embed/r:link
    on <p:pic> elements.

    Returns the count of rIds rewired (for debugging / probe purposes).
    """
    R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    EMBED_ATTR = f"{{{R_NS}}}embed"
    LINK_ATTR = f"{{{R_NS}}}link"
    ID_ATTR = f"{{{R_NS}}}id"
    rid_attrs = (EMBED_ATTR, LINK_ATTR, ID_ATTR)
    count = 0
    for sub in el.iter():
        for attr_name in rid_attrs:
            old_rid = sub.get(attr_name)
            if old_rid is None:
                continue
            try:
                src_rel = src_part.rels[old_rid]
            except KeyError:
                # Source rel already missing — strip the attribute so
                # we don't write a dangling ref into the destination.
                # PowerPoint forgives a missing attribute; it does not
                # forgive a dangling rId.
                del sub.attrib[attr_name]
                continue
            # For external-target rels (hyperlinks point at URLs), the
            # rel has target_mode="External" and no target_part. Use
            # relate_to with the URL string instead of a part. python-pptx
            # supports this via add_relationship — fall back if relate_to
            # rejects an external target.
            try:
                new_rid = dst_part.relate_to(src_rel.target_part, src_rel.reltype)
            except (AttributeError, ValueError, TypeError):
                # External-target rel — copy directly via add_relationship.
                try:
                    target = src_rel.target_ref  # the URL string
                    new_rid = dst_part.rels.get_or_add_ext_rel(
                        src_rel.reltype, target
                    )
                except Exception:
                    # Last resort: strip the attribute rather than leave
                    # a dangling ref. Better to lose the hyperlink than
                    # to trigger the repair prompt.
                    del sub.attrib[attr_name]
                    continue
            sub.set(attr_name, new_rid)
            count += 1
    return count


def _clone_slide_into(working_pres, library_pres, source_index: int, strip_pictures: bool = False):
    """Append a new slide to `working_pres` that is a deep clone of
    `library_pres.slides[source_index]`. Returns the new Slide.

    The clone uses the same slide layout (matched by name from the
    library's layout). All shapes from the source slide are deep-copied
    into the new slide's shape tree.
    """
    src_slide = library_pres.slides[source_index]
    src_layout = src_slide.slide_layout
    src_layout_name = src_layout.name

    # Find the matching layout in the working presentation by name.
    target_layout = None
    for layout in working_pres.slide_layouts:
        if layout.name == src_layout_name:
            target_layout = layout
            break
    if target_layout is None:
        # Fallback: use the first layout. Shouldn't happen because both
        # presentations were loaded from the same template.
        target_layout = working_pres.slide_layouts[0]

    new_slide = working_pres.slides.add_slide(target_layout)

    # Strip the default placeholders that add_slide inserted — they would
    # clash with the placeholders coming from the cloned source.
    for ph in list(new_slide.placeholders):
        sp = ph._element
        sp.getparent().remove(sp)

    # Deep-copy each shape from source into the new slide's shape tree.
    # Skip {nv,}GrpSpPr (group props at the top of spTree).
    # `<p:pic>` elements are now cloned WITH their image relationship
    # rewired (see _clone_pic_into) — picture-bearing slides (the AI
    # capabilities slide, split-layout benefit slides 24-27, stat-cards
    # background image) now render correctly instead of as empty boxes.
    src_tree = src_slide.shapes._spTree
    new_tree = new_slide.shapes._spTree
    for child in list(src_tree):
        tag = child.tag
        if tag.endswith("}nvGrpSpPr") or tag.endswith("}grpSpPr"):
            continue
        if tag.endswith("}pic"):
            if strip_pictures:
                continue
            new_tree.append(_clone_pic_into(child, src_slide, new_slide))
            continue
        cloned = deepcopy(child)
        _replace_spautofit_with_normautofit(cloned)
        # iter 53: rewire r:id / r:embed / r:link on every cloned
        # non-pic shape. <p:pic> elements are handled separately by
        # _clone_pic_into; everything else (text shapes with hyperlinks,
        # group shapes containing pics, OLE objects) needs the same
        # treatment or hyperlinks point at missing rels and Microsoft
        # PowerPoint prompts to repair.
        _rewire_rids(cloned, src_slide.part, new_slide.part)
        new_tree.append(cloned)

    return new_slide


_RENDER_UA = "bizzdesign-pptx-mcp/0.15 (+contact a.dodd@bizzdesign.com)"


# Bizzdesign brand palette — derived from the corporate template's title
# accent (the bright blue used on cover/divider titles) and supporting
# tones. Used to theme rendered diagrams + charts so they look like they
# belong to the same deck instead of generic black-on-white kroki output.
_BIZZ_BRAND = {
    "primary": "#2E7CF6",       # bright blue (title accent)
    "primary_dark": "#1A4FB8",  # darker blue for borders / strokes
    "primary_light": "#D5E5FF", # tinted fill for diagram nodes
    "navy": "#0F1F4D",          # deep navy for text on light fills
    "accent": "#F47B7B",        # warm coral for contrast series
    "accent_light": "#FCD9D9",
    "neutral": "#5C6675",       # medium gray for axis labels
    "neutral_light": "#E6E9EE", # grid lines
}

# Series palette for charts with multiple datasets — keeps the brand
# blue as the lead colour and pairs it with the coral accent + supporting
# tones so a 2-5 series chart still reads coherently.
_CHART_SERIES_PALETTE = [
    "#2E7CF6",  # primary
    "#F47B7B",  # accent coral
    "#1A4FB8",  # primary dark
    "#7AA8FA",  # primary light variant
    "#D5905C",  # warm tan
]


def _theme_mermaid(mermaid_source: str) -> str:
    """Prepend a Bizzdesign brand init directive to the Mermaid source.

    Mermaid supports a `%%{init: {...} }%%` directive that sets theme
    variables before parsing. The 'base' theme honours every variable
    we set (the 'default' theme partially ignores overrides). If the
    caller already supplied an init directive (sophisticated callers
    might tune individual diagrams), leave it alone.
    """
    if mermaid_source.lstrip().startswith("%%{init"):
        return mermaid_source
    theme_vars = {
        "primaryColor": _BIZZ_BRAND["primary_light"],
        "primaryTextColor": _BIZZ_BRAND["navy"],
        "primaryBorderColor": _BIZZ_BRAND["primary"],
        "lineColor": _BIZZ_BRAND["primary_dark"],
        "secondaryColor": _BIZZ_BRAND["accent_light"],
        "tertiaryColor": "#FFFFFF",
        "background": "#FFFFFF",
        "mainBkg": _BIZZ_BRAND["primary_light"],
        "secondBkg": _BIZZ_BRAND["accent_light"],
        "fontFamily": "Helvetica, Arial, sans-serif",
        "fontSize": "16px",
        "nodeBorder": _BIZZ_BRAND["primary"],
        "clusterBkg": "#F5F8FF",
        "clusterBorder": _BIZZ_BRAND["primary_dark"],
        "edgeLabelBackground": "#FFFFFF",
        "actorBkg": _BIZZ_BRAND["primary_light"],
        "actorBorder": _BIZZ_BRAND["primary"],
        "actorTextColor": _BIZZ_BRAND["navy"],
        "labelBoxBkgColor": _BIZZ_BRAND["primary_light"],
        "labelBoxBorderColor": _BIZZ_BRAND["primary"],
        "noteBkgColor": "#FFF8E1",
        "noteBorderColor": _BIZZ_BRAND["neutral"],
    }
    import json as _json
    init_directive = f"%%{{init: {_json.dumps({'theme': 'base', 'themeVariables': theme_vars})} }}%%"
    return init_directive + "\n" + mermaid_source


def _fetch_diagram_png(mermaid_source: str, timeout: int = 30) -> bytes:
    """POST Mermaid source to kroki.io and return the rendered PNG bytes.

    Kroki is a free public diagram-rendering gateway. The Mermaid path
    runs the official mermaid renderer server-side and returns PNG.
    POC-grade dependency — fine for sandbox, swap for self-hosted
    kroki container if/when this matters for production SLAs.

    Kroki rejects the default Python urllib User-Agent — set a real one.
    Source is auto-themed with Bizzdesign brand colours before sending.
    """
    import urllib.request
    themed = _theme_mermaid(mermaid_source)
    req = urllib.request.Request(
        "https://kroki.io/mermaid/png",
        data=themed.encode("utf-8"),
        headers={
            "Content-Type": "text/plain",
            "Accept": "image/png",
            "User-Agent": _RENDER_UA,
        },
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return resp.read()


def _theme_chart(chart_config: Dict[str, Any]) -> Dict[str, Any]:
    """Apply Bizzdesign brand palette + chart-chrome styling to a Chart.js config.

    Mutates a copy of the supplied config so the caller's dict isn't
    altered. Auto-assigns dataset colours from the brand series palette
    if the caller didn't specify backgroundColor / borderColor; sets
    axis label / grid line styling so the chart looks branded without
    requiring the caller to know Chart.js styling syntax.
    """
    import copy as _copy
    cfg = _copy.deepcopy(chart_config)
    chart_type = cfg.get("type", "bar")
    data = cfg.setdefault("data", {})
    datasets = data.setdefault("datasets", [])

    for i, ds in enumerate(datasets):
        colour = _CHART_SERIES_PALETTE[i % len(_CHART_SERIES_PALETTE)]
        if chart_type in ("pie", "doughnut"):
            # Per-slice colours — one per label.
            if "backgroundColor" not in ds:
                ds["backgroundColor"] = [
                    _CHART_SERIES_PALETTE[j % len(_CHART_SERIES_PALETTE)]
                    for j in range(len(ds.get("data", [])))
                ]
            ds.setdefault("borderColor", "#FFFFFF")
            ds.setdefault("borderWidth", 2)
        else:
            ds.setdefault("backgroundColor", colour)
            ds.setdefault("borderColor", colour)
            ds.setdefault("borderWidth", 2)
            if chart_type == "line":
                ds.setdefault("tension", 0.3)
                ds.setdefault("pointRadius", 4)
                ds.setdefault("pointBackgroundColor", colour)
                ds.setdefault("fill", False)

    options = cfg.setdefault("options", {})
    options.setdefault("responsive", True)
    options.setdefault("maintainAspectRatio", False)
    plugins = options.setdefault("plugins", {})
    legend = plugins.setdefault("legend", {})
    legend.setdefault("position", "top")
    legend_labels = legend.setdefault("labels", {})
    legend_labels.setdefault("color", _BIZZ_BRAND["navy"])
    legend_labels.setdefault("font", {"family": "Helvetica, Arial, sans-serif", "size": 14})

    if chart_type not in ("pie", "doughnut"):
        scales = options.setdefault("scales", {})
        for axis_key in ("x", "y"):
            axis = scales.setdefault(axis_key, {})
            ticks = axis.setdefault("ticks", {})
            ticks.setdefault("color", _BIZZ_BRAND["neutral"])
            ticks.setdefault("font", {"family": "Helvetica, Arial, sans-serif", "size": 12})
            grid = axis.setdefault("grid", {})
            grid.setdefault("color", _BIZZ_BRAND["neutral_light"])
            grid.setdefault("borderColor", _BIZZ_BRAND["neutral_light"])

    return cfg


def _fetch_chart_png(chart_config: Dict[str, Any], width: int = 800, height: int = 500, timeout: int = 30) -> bytes:
    """POST a Chart.js config to quickchart.io and return the rendered PNG.

    Same POC-vs-prod tradeoff as kroki — fine here, swap for self-hosted
    quickchart later if needed. Config is auto-themed with the Bizzdesign
    brand palette before sending.
    """
    import urllib.request
    import json as _json
    themed = _theme_chart(chart_config)
    payload = {
        "chart": themed,
        "width": width,
        "height": height,
        "backgroundColor": "white",
        "format": "png",
        "version": "4",  # Chart.js v4
    }
    req = urllib.request.Request(
        "https://quickchart.io/chart",
        data=_json.dumps(payload).encode("utf-8"),
        headers={
            "Content-Type": "application/json",
            "Accept": "image/png",
            "User-Agent": _RENDER_UA,
        },
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return resp.read()


def _tag_slide_with_composition(slide, composition_name: str) -> None:
    """Set `<p:cSld name="composition:<name>">` on the slide so downstream
    tooling (assessment harnesses, debug viewers) can identify which
    composition produced each slide without inferring from shape names.

    The `name` attribute on `<p:cSld>` is part of the OOXML schema
    (ECMA-376 part 1, §19.3.1.16) and is not rendered anywhere visible to
    the viewer, so it's safe to use as a metadata channel.
    """
    P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
    csld = slide.element.find(f"{{{P_NS}}}cSld")
    if csld is not None:
        csld.set("name", f"composition:{composition_name}")


def _replace_spautofit_with_normautofit(element) -> None:
    """Walk an element's text-frame body props and swap any <a:spAutoFit/>
    (shape grows to fit text) for <a:normAutofit/> (text shrinks to fit
    shape). Brand layouts have fixed positions; auto-growing shapes
    overflow the design (e.g. stat-card text spilling below the white
    card rectangle).
    """
    from lxml import etree
    A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
    for bodyPr in element.iter(f"{{{A_NS}}}bodyPr"):
        sp_autofit = bodyPr.find(f"{{{A_NS}}}spAutoFit")
        if sp_autofit is not None:
            # Replace with normAutofit (text auto-shrinks on overflow).
            bodyPr.remove(sp_autofit)
            norm = etree.SubElement(bodyPr, f"{{{A_NS}}}normAutofit")
            norm.set("fontScale", "100000")
            norm.set("lnSpcReduction", "0")


# ------------------------------------------------------------------
# Field application
# ------------------------------------------------------------------

def _apply_fields(slide, fields_spec: Dict[str, Any], content: Dict[str, Any]) -> List[str]:
    """Apply text replacements to the slide's shapes per the composition's
    fields spec.

    `fields_spec` maps field_name → matcher dict OR list of matcher dicts (repeatable).
    Matcher dict can be either:
      - simple:  `{"match": "Prefix..."}`  — one shape per item
      - paired:  `{"heading_match": "Prefix...", "description_match": "Prefix..."}`  —
                 two shapes per item; user value is split at the first newline,
                 line 1 goes to the heading shape, lines 2+ go to the description shape.
    `content` maps field_name → string OR list of strings.

    Returns a list of warnings (missing shapes, etc).
    """
    warnings: List[str] = []
    shapes = list(slide.shapes)
    # Snapshot every shape's text BEFORE any mutation so the matcher
    # binds to the template's original content. Without this, applying
    # one field can change a shape's text in a way that makes it match
    # a different field's prefix on the next iteration — e.g. cover
    # title "2026 is a regime change..." starts with "2026" and was
    # being clobbered by the subtitle matcher that also looks for "2026".
    original_texts: Dict[int, str] = {id(s): _shape_text(s) for s in shapes}
    consumed: set = set()

    def _claim(shape):
        consumed.add(id(shape))

    def _apply_text(shape, value):
        if isinstance(value, list):
            _set_shape_multiline(shape, [str(v) for v in value])
        else:
            text = str(value) if not isinstance(value, str) else value
            if "\n" in text:
                _set_shape_multiline(shape, text.split("\n"))
            else:
                _set_shape_text(shape, text)

    def _apply_paired(shapes, matcher, value, role_base=None):
        # value is a string; split at the FIRST newline.
        text = _normalise_text(value) if isinstance(value, str) else str(value)
        head, _, body = text.partition("\n")
        head_shape = _find_shape_by_match(
            shapes, matcher["heading_match"], original_texts, consumed
        )
        desc_shape = _find_shape_by_match(
            shapes, matcher["description_match"], original_texts, consumed
        )
        if head_shape is None:
            warnings.append(
                f"Paired matcher: heading shape not found (prefix: {matcher['heading_match']!r})"
            )
        else:
            _set_shape_text(head_shape, head)
            _claim(head_shape)
            if role_base:
                _rename_shape(head_shape, f"{role_base}.heading")
        if desc_shape is None:
            warnings.append(
                f"Paired matcher: description shape not found (prefix: {matcher['description_match']!r})"
            )
        else:
            _set_shape_text(desc_shape, body)
            _claim(desc_shape)
            if role_base:
                _rename_shape(desc_shape, f"{role_base}.description")

    def _clear_paired(shapes, matcher):
        head_shape = _find_shape_by_match(
            shapes, matcher["heading_match"], original_texts, consumed
        )
        desc_shape = _find_shape_by_match(
            shapes, matcher["description_match"], original_texts, consumed
        )
        if head_shape is not None:
            _clear_shape_text(head_shape)
            _claim(head_shape)
        if desc_shape is not None:
            _clear_shape_text(desc_shape)
            _claim(desc_shape)

    for field_name, spec in fields_spec.items():
        value = content.get(field_name)
        if isinstance(spec, list):
            # Repeatable group.
            values = value or []
            if not isinstance(values, list):
                warnings.append(f"Field '{field_name}' expected a list, got {type(values).__name__}")
                values = []
            for i, matcher in enumerate(spec):
                is_paired = "heading_match" in matcher and "description_match" in matcher
                # Iter 56: None at index i means "leave this slot empty"
                # (used by capability_grid to remap 3/6 cards across cols).
                # Treat the same as "no value supplied" → clear the shape.
                if i < len(values) and values[i] is not None:
                    if is_paired:
                        _apply_paired(shapes, matcher, values[i], role_base=f"{field_name}[{i}]")
                    else:
                        shape = _find_shape_by_match(
                            shapes, matcher["match"], original_texts, consumed
                        )
                        if shape is None:
                            warnings.append(
                                f"Could not locate shape for {field_name}[{i}] (match prefix: {matcher['match']!r})"
                            )
                            continue
                        _apply_text(shape, values[i])
                        _claim(shape)
                        _rename_shape(shape, f"{field_name}[{i}]")
                else:
                    # User supplied fewer values than the template has slots
                    # (or explicit None) — blank the unused shape(s).
                    if is_paired:
                        _clear_paired(shapes, matcher)
                    else:
                        shape = _find_shape_by_match(
                            shapes, matcher["match"], original_texts, consumed
                        )
                        if shape is not None:
                            _clear_shape_text(shape)
                            _claim(shape)
        else:
            # Scalar field.
            if value is None:
                if spec.get("required", False):
                    warnings.append(f"Required field '{field_name}' missing from content")
                continue
            shape = _find_shape_by_match(
                shapes, spec["match"], original_texts, consumed
            )
            if shape is None:
                warnings.append(f"Could not locate shape for '{field_name}' (match prefix: {spec['match']!r})")
                continue
            _apply_text(shape, value)
            _claim(shape)
            _rename_shape(shape, field_name)

    return warnings


def _strip_empty_placeholders(slide) -> None:
    """Iter 52: drop any placeholder shape on the slide whose text frame
    is empty (no text content) AND whose name is NOT prefixed with
    `role:` (which means a composition tool deliberately populated it
    and an empty paragraph there is structural, not a stray prompt).

    PowerPoint renders an unfilled placeholder with the layout's
    default "Click to add text" prompt. That prompt is visible to the
    end user even though there's no real content. Stripping the empty
    placeholders is the only way to suppress the prompt from showing
    on the rendered slide.

    Layout-built compositions (chart, diagram, bullets) inherit all of
    the layout's placeholders when add_slide is called. Compositions
    only populate a subset; the rest sit empty until we remove them.
    """
    P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
    for shape in list(slide.placeholders):
        sp = shape._element
        name_el = sp.find(f".//{{{P_NS}}}cNvPr")
        name = (name_el.get("name") if name_el is not None else "") or ""
        # Skip role-tagged shapes — those were touched by the composition
        # and any empty paragraphs there are intentional layout artefacts.
        if name.startswith("role:"):
            continue
        # If the text frame has any non-empty paragraph, leave it.
        text_frame = shape.text_frame if shape.has_text_frame else None
        has_content = False
        if text_frame is not None:
            for p in text_frame.paragraphs:
                if (p.text or "").strip():
                    has_content = True
                    break
        if has_content:
            continue
        # Remove the placeholder from the slide's shape tree.
        sp.getparent().remove(sp)


def _clear_starting_with(slide, prefixes: List[str]) -> None:
    """Defensive cleanup: blank any shape whose text starts with one of
    the given prefixes. Used by compositions that leave behind original
    template text in unaddressed shapes.
    """
    shapes = list(slide.shapes)
    for prefix in prefixes:
        shape = _find_shape_by_match(shapes, prefix)
        if shape is not None:
            _clear_shape_text(shape)


def _tighten_first_paragraph_line_spacing(slide, spec: Dict[str, Any]) -> None:
    """Compress the line-height on the FIRST paragraph of shapes that
    were originally located by `spec["shape_prefixes"]`. Used to close
    the visible gap between large VALUE paragraphs and the smaller
    LABEL/ATTRIBUTION paragraphs that follow them in stat-card shapes.

    `spec` shape (composition metadata):
        {"shape_prefixes": [...], "spc_pct": 80000}  # 80% line spacing
    """
    from lxml import etree
    A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
    shapes = list(slide.shapes)
    pct_val = int(spec.get("spc_pct", 90000))
    # NOTE: by the time this runs, the shapes have NEW text (the user's
    # content) — they no longer start with the source-slide prefix.
    # Instead, find the shapes whose ORIGINAL prefix would have located
    # them, by walking ALL shapes and matching the new value at the same
    # position. Simpler: just walk all shapes the composition cares
    # about by re-running _find_shape_by_match, but since text changed,
    # we instead apply to ALL shapes that have a multi-paragraph text
    # frame where p0 is much larger than p1 (heuristic).
    for shape in shapes:
        if not shape.has_text_frame:
            continue
        tf = shape.text_frame
        paragraphs = list(tf.paragraphs)
        if len(paragraphs) < 2:
            continue
        # Get first-run font sizes of p0 and p1.
        p0_runs = paragraphs[0]._p.findall(f"{{{A_NS}}}r")
        p1_runs = paragraphs[1]._p.findall(f"{{{A_NS}}}r")
        if not p0_runs or not p1_runs:
            continue
        p0_rPr = p0_runs[0].find(f"{{{A_NS}}}rPr")
        p1_rPr = p1_runs[0].find(f"{{{A_NS}}}rPr")
        if p0_rPr is None or p1_rPr is None:
            continue
        p0_sz = p0_rPr.get("sz")
        p1_sz = p1_rPr.get("sz")
        if p0_sz is None or p1_sz is None:
            continue
        # Apply tightening when p0 font is at least 2× p1 font (heuristic
        # for "this is a big-value paragraph above small-label paragraphs").
        if int(p0_sz) < 2 * int(p1_sz):
            continue
        # Add/replace <a:lnSpc><a:spcPct val="80000"/></a:lnSpc> on p0's pPr.
        p_el = paragraphs[0]._p
        pPr = p_el.find(f"{{{A_NS}}}pPr")
        if pPr is None:
            pPr = etree.SubElement(p_el, f"{{{A_NS}}}pPr")
            p_el.remove(pPr)
            p_el.insert(0, pPr)
        # Remove existing lnSpc if any.
        for existing in pPr.findall(f"{{{A_NS}}}lnSpc"):
            pPr.remove(existing)
        # lnSpc must be the first child of pPr per OOXML schema.
        lnSpc = etree.Element(f"{{{A_NS}}}lnSpc")
        spcPct = etree.SubElement(lnSpc, f"{{{A_NS}}}spcPct")
        spcPct.set("val", str(pct_val))
        pPr.insert(0, lnSpc)


def _timeline_clean_empty_markers(slide) -> None:
    """Iter 71: delete the Oval marker dots + dashed Straight-Connector
    lines for timeline slots whose text label is empty.

    Template slide 4 layout (per slot):
      - TextBox at (label_x, label_y), with a year prefix
      - Oval (ellipse, 0.26in diameter) at (label_x + ~0.4, label_y - 1.85)
      - Straight Connector (only on alternating slots) at
        (label_x + ~0.55, label_y - 1.4)

    Identify slot dots+lines by horizontal proximity to a TextBox shape.
    Delete those whose nearest TextBox is empty.
    """
    A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
    P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
    EMU = 914400

    def shape_xy(shape) -> Optional[tuple]:
        sp = shape._element
        sp_pr = sp.find(f"{{{P_NS}}}spPr")
        if sp_pr is None:
            return None
        xfrm = sp_pr.find(f"{{{A_NS}}}xfrm")
        if xfrm is None:
            return None
        off = xfrm.find(f"{{{A_NS}}}off")
        if off is None:
            return None
        try:
            return (int(off.get("x", 0)) / EMU, int(off.get("y", 0)) / EMU)
        except Exception:
            return None

    def shape_geom(shape) -> str:
        sp = shape._element
        sp_pr = sp.find(f"{{{P_NS}}}spPr")
        if sp_pr is None:
            return ""
        geom = sp_pr.find(f"{{{A_NS}}}prstGeom")
        return geom.get("prst", "") if geom is not None else ""

    def shape_text(shape) -> str:
        if not getattr(shape, "has_text_frame", False):
            return ""
        try:
            return (shape.text_frame.text or "").strip()
        except Exception:
            return ""

    # Categorise shapes
    text_boxes = []  # (x, y, text)
    ovals = []       # python-pptx shape objects (ellipses)
    connectors = []  # python-pptx shape objects (lines / cxnSp)
    for shape in list(slide.shapes):
        xy = shape_xy(shape)
        if xy is None:
            continue
        x, y = xy
        geom = shape_geom(shape)
        # Only consider timeline-region shapes (y > 1.0 to skip title)
        if y < 1.0:
            continue
        if geom == "ellipse":
            ovals.append(shape)
        elif geom == "line":
            connectors.append(shape)
        elif shape_text(shape) and y > 2.0:
            # TextBox with content sits in the lower half of the slide
            text_boxes.append((x, y, shape))

    # For each ellipse, the matching label sits ~0.4in to the LEFT and
    # EITHER ~0.5in below (template's even slots, no connector) OR
    # ~1.85in below (odd slots, with connector). Try both expected
    # positions — keep the oval if a labelled TextBox is found at
    # either one within tolerance.
    for oval in list(ovals):
        ox, oy = shape_xy(oval)
        target_x = ox - 0.4
        matched = False
        for target_dy in (0.5, 1.85):
            target_y = oy + target_dy
            for tx, ty, _t in text_boxes:
                if abs(tx - target_x) < 0.5 and abs(ty - target_y) < 0.5:
                    matched = True
                    break
            if matched:
                break
        if not matched:
            sp = oval._element
            sp.getparent().remove(sp)

    # Now delete connectors that share an X position with a deleted
    # oval — i.e. connectors whose nearest oval no longer exists.
    # Reread ovals AFTER deletions:
    remaining_ovals = []
    for shape in list(slide.shapes):
        xy = shape_xy(shape)
        if xy is None:
            continue
        if shape_geom(shape) == "ellipse" and xy[1] >= 1.0:
            remaining_ovals.append((xy[0], xy[1], shape))
    for conn in list(connectors):
        cx, cy = shape_xy(conn)
        # A connector is paired with the oval ~0.15in to its left.
        target_x = cx - 0.15
        if any(abs(ox - target_x) < 0.3 for ox, _, _ in remaining_ovals):
            continue
        # No remaining oval at this x → delete the connector
        sp = conn._element
        sp.getparent().remove(sp)


def _resize_cap_grid_backgrounds(slide, shape_names: List[str], n_filled: int) -> None:
    """Iter 70: shrink the capability_grid column-background rectangles
    when only some rows are populated. Default template extends each
    background to y=7.50in (bottom of slide) under the assumption that
    all 3 rows are filled (n_filled=9). At n=6 (rows 1-2 only) shrink
    so the bottom matches the bottom card's bottom; at n=3 (row 1 only)
    shrink further.

    Card y positions in the template (after iter-56 remap):
      - row 1 cards: heading y=3.23in, description y=3.53in (ends ~3.93)
      - row 2 cards: heading y=4.42in, description y=4.85in (ends ~5.25)
      - row 3 cards: heading y=5.75in, description y=6.19in (ends ~6.59)
    Backgrounds start at y=2.52in (under column headings).
    """
    A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
    P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
    if n_filled <= 3:
        new_cy_in = 1.85  # ends at y=4.37in, just below row-1 cards
    elif n_filled <= 6:
        new_cy_in = 3.10  # ends at y=5.62in, just below row-2 cards
    else:
        return  # n=9, no resize needed
    new_cy_emu = int(new_cy_in * 914400)
    for shape in list(slide.shapes):
        if shape.name not in shape_names:
            continue
        sp = shape._element
        sp_pr = sp.find(f"{{{P_NS}}}spPr")
        if sp_pr is None:
            continue
        xfrm = sp_pr.find(f"{{{A_NS}}}xfrm")
        if xfrm is None:
            continue
        ext = xfrm.find(f"{{{A_NS}}}ext")
        if ext is None:
            continue
        try:
            ext.set("cy", str(new_cy_emu))
        except Exception:
            pass


def _shift_role_y_emu(slide, role_to_delta: Dict[str, int]) -> None:
    """Iter 66: shift the Y position of shapes by role. Negative delta
    moves up. Used by solution_detail to close the mid-slide whitespace
    gap left after `strip_pictures` removes the template's product UI
    screenshots.

    Matches against the role-prefix root (e.g. role `details` matches
    `role:details[0]`, `role:details[1]`, etc.).
    """
    A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
    P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
    for shape in list(slide.shapes):
        name = (shape.name or "")
        if not name.startswith("role:"):
            continue
        role = name[len("role:"):]
        root = role
        for sep in ("[", "."):
            idx = root.find(sep)
            if idx != -1:
                root = root[:idx]
        delta = role_to_delta.get(role) or role_to_delta.get(root)
        if delta is None:
            continue
        sp = shape._element
        sp_pr = sp.find(f"{{{P_NS}}}spPr")
        if sp_pr is None:
            continue
        xfrm = sp_pr.find(f"{{{A_NS}}}xfrm")
        if xfrm is None:
            continue
        off = xfrm.find(f"{{{A_NS}}}off")
        if off is None:
            continue
        try:
            y = int(off.get("y", 0))
            off.set("y", str(y + int(delta)))
        except Exception:
            pass


def _set_run_sz_for_role(slide, role_to_sz: Dict[str, int]) -> None:
    """Iter 57: post-clone hook to override the font size on every run
    inside shapes whose role matches a key in `role_to_sz`. Used when
    the template's source-slide font size is too small/large for the
    composition's intended use (e.g. closing.cta_lines is 10pt in the
    template — too small for a leadership audience).

    Shape role is determined by the `role:` prefix of its `<p:cNvPr name>`
    attribute (set by `_rename_shape` during field application).
    """
    A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
    for shape in list(slide.shapes):
        if not shape.has_text_frame:
            continue
        name = (shape.name or "")
        if not name.startswith("role:"):
            continue
        role = name[len("role:"):]
        # Allow either exact match ("cta_lines") or prefix-style for
        # paired fields ("cards[0].heading" → check "cards" base).
        sz = role_to_sz.get(role)
        if sz is None:
            # Try the field root before the first '['/'.'
            root = role
            for sep in ("[", "."):
                idx = root.find(sep)
                if idx != -1:
                    root = root[:idx]
            sz = role_to_sz.get(root)
        if sz is None:
            continue
        for p in shape.text_frame._txBody.findall(f"{{{A_NS}}}p"):
            for r in p.findall(f"{{{A_NS}}}r"):
                rpr = r.find(f"{{{A_NS}}}rPr")
                if rpr is None:
                    from lxml import etree as _et
                    rpr = _et.SubElement(r, f"{{{A_NS}}}rPr")
                    r.insert(0, rpr)
                rpr.set("sz", str(int(sz)))


# ------------------------------------------------------------------
# Layout-built compositions (no source slide clone)
# ------------------------------------------------------------------

def _apply_layout_fields(slide, content: Dict[str, Any]) -> List[str]:
    """Populate a layout-built slide (Basic Text) with title + subhead + bullets.

    Currently only used by the `bullets` composition. The Basic Text
    layout has the title placeholder, a subhead at idx 13, and a body
    placeholder at idx 14 (brand bullets baked in).
    """
    warnings: List[str] = []
    title_text = content.get("title")
    if title_text:
        if slide.shapes.title is not None:
            slide.shapes.title.text = _normalise_text(str(title_text))
            _rename_shape(slide.shapes.title, "title")
        else:
            warnings.append("No title placeholder on this layout")

    subhead_text = content.get("subhead")
    if subhead_text is not None:
        ph_13 = None
        for ph in slide.placeholders:
            if ph.placeholder_format.idx == 13:
                ph_13 = ph
                break
        if ph_13 is not None:
            ph_13.text_frame.text = str(subhead_text)
            _rename_shape(ph_13, "subhead")
        else:
            warnings.append("No placeholder idx 13 (subhead) on this layout")

    bullets = content.get("bullets")
    if bullets:
        if isinstance(bullets, str):
            bullets = [bullets]
        ph_14 = None
        for ph in slide.placeholders:
            if ph.placeholder_format.idx == 14:
                ph_14 = ph
                break
        if ph_14 is None:
            warnings.append("No placeholder idx 14 (body) on this layout")
        else:
            tf = ph_14.text_frame
            tf.text = str(bullets[0])
            for bullet in bullets[1:]:
                p = tf.add_paragraph()
                p.text = str(bullet)
            # The Basic Text layout sets <a:buNone/> on the body
            # placeholder — paragraphs inherit "no bullet" and render
            # as plain paragraphs. Override per-paragraph with explicit
            # <a:buChar char="•"/> so bullets actually appear.
            _force_bullet_glyphs(ph_14.text_frame)
            _rename_shape(ph_14, "bullets")
    return warnings


def _force_bullet_glyphs(text_frame, char: str = "•") -> None:
    """Add an explicit <a:buChar char="•"/> to every paragraph's <a:pPr>,
    overriding any inherited <a:buNone/>.
    """
    from lxml import etree
    A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
    for p in text_frame.paragraphs:
        p_el = p._p
        pPr = p_el.find(f"{{{A_NS}}}pPr")
        if pPr is None:
            pPr = etree.SubElement(p_el, f"{{{A_NS}}}pPr")
            # pPr must be the first child of <a:p>
            p_el.remove(pPr)
            p_el.insert(0, pPr)
        # Remove any inherited bullet-suppression and existing bullet
        # chars so our new one is authoritative.
        for tag in ("buNone", "buChar", "buAutoNum"):
            for el in pPr.findall(f"{{{A_NS}}}{tag}"):
                pPr.remove(el)
        buChar = etree.SubElement(pPr, f"{{{A_NS}}}buChar")
        buChar.set("char", char)


# ------------------------------------------------------------------
# Library cache
# ------------------------------------------------------------------
# We need a fresh Presentation instance per working deck (so the model's
# changes don't leak into the library). Reload from disk per session.

_LIBRARY_TEMPLATE_CACHE: Dict[str, str] = {}  # presentation_id → template path used


def _get_library(presentation_id: str, library_template_paths: Dict[str, str]):
    """Re-open the library template from disk to get a fresh copy."""
    path = library_template_paths.get(presentation_id)
    if path is None:
        return None
    return Presentation(path)


# ------------------------------------------------------------------
# Tool registration
# ------------------------------------------------------------------

def register_composition_tools(
    app: FastMCP,
    presentations: Dict,
    get_current_presentation_id,
    library_template_paths: Dict[str, str],
    set_current_presentation_id=None,
):
    """Register the composition tools with the FastMCP app.

    `library_template_paths` is a presentation_id → template path dict
    maintained by `create_presentation`. It lets us reload the library
    from disk per composition call.
    """

    def _build_composition(
        composition_name: str,
        content: Dict[str, Any],
        presentation_id: Optional[str],
    ) -> Dict[str, Any]:
        comp = COMPOSITIONS.get(composition_name)
        if comp is None:
            return {"error": f"Unknown composition: {composition_name!r}"}

        pres_id = presentation_id if presentation_id is not None else get_current_presentation_id()
        if pres_id is None or pres_id not in presentations:
            return {"error": "No presentation is currently loaded."}

        working = presentations[pres_id]
        library = _get_library(pres_id, library_template_paths)
        if library is None:
            return {
                "error": (
                    "No brand template library available for this presentation. "
                    "Compositions require a presentation created via the default-template "
                    "path (PPTX_DEFAULT_TEMPLATE) — `blank=True` decks can't use compositions."
                )
            }

        source_idx = comp["source_slide_index"]
        try:
            if source_idx is None:
                # Layout-built composition (currently just `bullets`).
                # Find the named layout in the working presentation and add
                # a fresh slide, then populate title + subhead + bullets via
                # the layout's placeholders directly.
                layout_name = comp.get("layout_name")
                layout = next(
                    (l for l in working.slide_layouts if l.name == layout_name),
                    None,
                )
                if layout is None:
                    return {"error": f"Layout {layout_name!r} not found in working presentation."}
                new_slide = working.slides.add_slide(layout)
                warnings = _apply_layout_fields(new_slide, content)
                _widen_title_full_width(new_slide, working)
                # iter 52: strip any placeholder we didn't populate. The
                # Basic Text layout has title + subhead + body slots; if
                # the composition only fills (say) the title shape, the
                # other placeholders inherit the layout's "Click to add
                # text" default text and render that prompt to the user.
                # User feedback (2026-05-30) flagged this on chart slides
                # where the layout's body and right-side placeholder both
                # showed "Click to add text" boxes flanking the chart.
                _strip_empty_placeholders(new_slide)
            else:
                new_slide = _clone_slide_into(
                    working,
                    library,
                    source_idx,
                    strip_pictures=comp.get("strip_pictures", False),
                )
                warnings = _apply_fields(new_slide, comp["fields"], content)
                # Optional defensive cleanup: blank any shapes whose text
                # starts with one of these prefixes. Used for compositions
                # whose source slides have leftover heading/intro shapes
                # the composition can't address by field.
                clear_prefixes = comp.get("clear_text_starting_with")
                if clear_prefixes:
                    _clear_starting_with(new_slide, clear_prefixes)
                tighten = comp.get("tighten_first_paragraph_line_spacing")
                if tighten:
                    _tighten_first_paragraph_line_spacing(new_slide, tighten)
                # Iter 57: per-composition font size override by role.
                set_run_sz_for_role = comp.get("set_run_sz_for_role")
                if set_run_sz_for_role:
                    _set_run_sz_for_role(new_slide, set_run_sz_for_role)
                # Iter 66: shift shape Y by role to close mid-slide
                # whitespace gaps left by strip_pictures (e.g. solution_detail).
                shift_role_y_emu = comp.get("shift_role_y_emu")
                if shift_role_y_emu:
                    _shift_role_y_emu(new_slide, shift_role_y_emu)
                # Iter 70: shrink capability_grid column backgrounds when
                # only rows 1-2 (n=6) or row 1 (n=3) are filled.
                cap_grid_bg_shapes = comp.get("cap_grid_column_bg_shapes")
                if cap_grid_bg_shapes:
                    n_filled = sum(
                        1 for c in (content.get("cards") or []) if c is not None
                    )
                    _resize_cap_grid_backgrounds(
                        new_slide, cap_grid_bg_shapes, n_filled
                    )
                # Iter 71: delete timeline Oval markers + connector lines
                # for empty (label-less) slots.
                if comp.get("timeline_clean_empty_markers"):
                    _timeline_clean_empty_markers(new_slide)
        except Exception as e:
            return {"error": f"Failed to build {composition_name}: {e}"}

        _tag_slide_with_composition(new_slide, composition_name)

        slide_index = len(working.slides) - 1
        result: Dict[str, Any] = {
            "composition": composition_name,
            "slide_index": slide_index,
            "message": f"Added '{composition_name}' as slide {slide_index}.",
        }
        if warnings:
            result["warnings"] = warnings
        return result

    @app.tool(
        annotations=ToolAnnotations(
            title="Validate Deck",
            readOnlyHint=True,
        ),
    )
    def validate_deck(
        headline_message: Optional[str] = None,
        presentation_id: Optional[str] = None,
    ) -> Dict:
        """Run deck-level deterministic quality probes on the working
        presentation. Returns pass/fail per probe with actionable
        rewrite hints for any failures.

        Call this AFTER building the deck and BEFORE
        `save_presentation_to_url`. If any probe fails, use
        `delete_slide` + the appropriate composition tool to fix the
        specific slide indices reported, then re-validate before saving.

        Probes run:
        - `deck-has-closing-slide`: last slide must be `composition:closing`
        - `composition-variety-min-4-distinct`: 7+ slide decks need ≥ 4 distinct
          compositions in the body (excluding cover + closing)
        - `composition-no-streak-over-2`: no composition repeats > 2 times in a row
        - `title-ladder`: if `headline_message` is supplied, cover must
          reference ≥ 40% of its key terms and body slides ≥ 60% collectively

        Args:
            headline_message: Optional — the brief's ONE-line headline message.
                When supplied, enables the title-ladder probe. Skip if you don't
                have a stable headline (some kickoff decks don't).
        """
        import re
        pres_id = presentation_id if presentation_id is not None else get_current_presentation_id()
        if pres_id is None or pres_id not in presentations:
            return {"error": "No presentation is currently loaded."}
        pres = presentations[pres_id]
        slides = list(pres.slides)
        P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"

        def slide_composition(slide):
            csld = slide.element.find(f"{{{P_NS}}}cSld")
            if csld is None:
                return None
            name = csld.get("name") or ""
            return name[len("composition:"):] if name.startswith("composition:") else None

        compositions = [slide_composition(s) for s in slides]
        probe_results = []

        # Deck-has-closing-slide.
        last_comp = compositions[-1] if compositions else None
        probe_results.append({
            "name": "deck-has-closing-slide",
            "passed": last_comp == "closing",
            "detail": None if last_comp == "closing" else (
                f"Last slide is `{last_comp}` (slide index {len(slides) - 1}). "
                f"Every deck must end with add_closing_slide."
            ),
            "fix_hint": None if last_comp == "closing" else (
                "Call add_closing_slide(title='Thank You' or a deck-specific "
                "wrap-up, cta_lines=[…]) — it'll append as the new last slide."
            ),
        })

        if len(slides) >= 7:
            body = compositions[1:-1]
            distinct = set(c for c in body if c)
            # For the fix_hint, find the most over-used composition AND
            # the slide indices that use it, so the model can pick which
            # to swap.
            from collections import Counter
            comp_counts = Counter(c for c in body if c)
            most_common = comp_counts.most_common(1)
            most_common_comp = most_common[0][0] if most_common else None
            most_common_indices = [
                i + 1 for i, c in enumerate(body) if c == most_common_comp
            ]
            # Suggest compositions NOT yet used in the deck.
            unused_compositions = sorted(
                {"bullets", "stat_cards", "value_cards", "split_benefits", "capability_grid", "value_props_4", "section_divider", "timeline", "process_steps"}
                - distinct
            )
            probe_results.append({
                "name": "composition-variety-min-4-distinct",
                "passed": len(distinct) >= 4,
                "detail": None if len(distinct) >= 4 else (
                    f"Body has only {len(distinct)} distinct compositions: "
                    f"{sorted(distinct)}. A {len(slides)}-slide deck needs ≥ 4."
                ),
                "fix_hint": None if len(distinct) >= 4 else (
                    f"`{most_common_comp}` is the most over-used (appears on slide indices "
                    f"{most_common_indices}). Pick ONE of those slides, delete_slide(idx), "
                    f"and re-add the same content using a different composition not yet in "
                    f"the deck: {unused_compositions}. Match the new composition to the "
                    f"slide's content shape (a list → bullets; numbers → stat_cards; "
                    f"4 short statements → value_props_4; outcomes → split_benefits; "
                    f"3 cols of capabilities → capability_grid)."
                ),
            })
            longest_streak = 1
            streak_comp = body[0] if body else None
            streak_start = 1
            current_start = 1
            current_streak = 1
            for i in range(1, len(body)):
                if body[i] == body[i - 1] and body[i] is not None:
                    current_streak += 1
                    if current_streak > longest_streak:
                        longest_streak = current_streak
                        streak_comp = body[i]
                        streak_start = current_start + 1  # +1 because body[0] is slide index 1
                else:
                    current_start = i + 1
                    current_streak = 1
            streak_end = streak_start + longest_streak - 1
            probe_results.append({
                "name": "composition-no-streak-over-2",
                "passed": longest_streak <= 2,
                "detail": None if longest_streak <= 2 else (
                    f"`{streak_comp}` repeats {longest_streak} times in a row "
                    f"(slides {streak_start} through {streak_end}). Vary the rhythm."
                ),
                "fix_hint": None if longest_streak <= 2 else (
                    f"Delete one of the streaked slides (e.g. delete_slide({streak_end - 1})) "
                    f"and re-add the same content with a different composition. "
                    f"If content is the same SHAPE, try the suggested swap in the "
                    f"list_compositions() catalogue."
                ),
            })

        # Title-ladder.
        if headline_message and len(slides) >= 3:
            stopwords = {
                'a', 'an', 'the', 'and', 'or', 'but', 'is', 'are', 'was', 'were',
                'be', 'been', 'being', 'to', 'of', 'in', 'on', 'at', 'for', 'with',
                'by', 'from', 'as', 'that', 'this', 'these', 'those', 'it', 'its',
                'we', 'our', 'us', 'you', 'your', 'their', 'they', 'them',
                'have', 'has', 'had', 'will', 'would', 'should', 'can', 'could',
                'may', 'might', 'must', 'do', 'does', 'did', 'not', 'no',
                'than', 'then', 'now', 'so', 'if', 'when', 'where', 'who',
                'what', 'how', 'why', 'one', 'two', 'three', 'four', 'five',
                'more', 'most', 'less', 'each', 'into', 'about',
            }
            def key_terms(text):
                return {
                    w for w in re.sub(r"[^a-z0-9\s%$£€]", " ", text.lower()).split()
                    if len(w) >= 4 and w not in stopwords
                }
            headline_terms = key_terms(headline_message)
            if len(headline_terms) >= 3:
                def slide_text(slide):
                    parts = []
                    for shape in slide.shapes:
                        if shape.has_text_frame:
                            parts.append(shape.text_frame.text)
                    return " ".join(parts)
                cover_terms = key_terms(slide_text(slides[0]))
                body_terms = set()
                for s in slides[1:-1]:
                    body_terms |= key_terms(slide_text(s))
                cover_cov = len(cover_terms & headline_terms) / max(1, len(headline_terms))
                body_cov = len(body_terms & headline_terms) / max(1, len(headline_terms))
                cover_ok = cover_cov >= 0.4
                body_ok = body_cov >= 0.6
                missing_cover = sorted(headline_terms - cover_terms)
                missing_body = sorted(headline_terms - body_terms)
                hint_parts = []
                if not cover_ok:
                    # Pick a handful of the missing terms to suggest verbatim.
                    suggest = missing_cover[: max(3, len(headline_terms) // 3)]
                    hint_parts.append(
                        f"call delete_slide(0) and then add_cover_slide again. "
                        f"The new title MUST include at least these words from the "
                        f"headline (verbatim): {suggest}. Keep within 7-16 words — "
                        f"pick a single sharp sentence that uses those words."
                    )
                if not body_ok:
                    hint_parts.append(
                        f"Add or rewrite a body slide title to include these "
                        f"missing headline terms: {missing_body}. Use add_bullets_slide "
                        f"or add_solution_detail_slide with a title that literally "
                        f"uses 2-3 of those words."
                    )
                probe_results.append({
                    "name": "title-ladder",
                    "passed": cover_ok and body_ok,
                    "detail": None if (cover_ok and body_ok) else (
                        (f"cover covers {int(cover_cov*100)}% of headline terms (need ≥40%) — missing: {missing_cover}. " if not cover_ok else "") +
                        (f"body covers {int(body_cov*100)}% (need ≥60%) — missing: {missing_body}." if not body_ok else "")
                    ).strip(),
                    "fix_hint": None if (cover_ok and body_ok) else " ".join(hint_parts),
                })

        failures = [p for p in probe_results if not p["passed"]]
        return {
            "passed": len(failures) == 0,
            "slide_count": len(slides),
            "probes": probe_results,
            "failure_count": len(failures),
            "next_step": (
                "Deck passed all deck-level probes. Safe to call save_presentation_to_url."
                if not failures else
                f"{len(failures)} probe(s) failed. Address each using the `fix_hint` field "
                "(delete_slide + replacement composition call), then call validate_deck "
                "again before save_presentation_to_url. Do NOT save while validation is failing."
            ),
        }

    @app.tool(
        annotations=ToolAnnotations(
            title="List Compositions",
            readOnlyHint=True,
        ),
    )
    def list_compositions() -> Dict:
        """List the high-level slide compositions available in the brand template.

        Each composition is backed by a known source slide from the bundled
        corporate template. Pick a composition by matching content shape to
        `use_when`. Compositions handle layout, placeholders, fonts, and
        styling for you — you just supply the content.
        """
        return {
            "compositions": {
                name: {
                    "description": meta["description"],
                    "use_when": meta["use_when"],
                    "fields": {
                        fname: ("(repeatable list of strings)" if isinstance(fspec, list)
                                else "(string)")
                        for fname, fspec in meta["fields"].items()
                    },
                }
                for name, meta in COMPOSITIONS.items()
            },
            "guidance": (
                "Always prefer these compositions over `add_slide` for branded output. "
                "Each composition clones a proven branded slide from the template — the "
                "model only supplies content. For repeatable fields (e.g. `details`, "
                "`pillars`), supply a list of strings; if shorter than the template "
                "expects, unused slots are blanked. Multi-line content can include `\\n`."
            ),
        }

    @app.tool(
        annotations=ToolAnnotations(title="Add Cover Slide"),
    )
    def add_cover_slide(
        title: str,
        subtitle: str,
        presentation_id: Optional[str] = None,
    ) -> Dict:
        """Add the deck's opening cover slide with brand-styled title and subtitle.

        Use this for slide 1 of every deck. The cover uses the dark
        Bizzdesign brand background with the large title centred and a
        subtitle line below.

        CRITICAL: The cover slide is NOT a name-tag for the deck. It is
        the opening claim. `title` must be a CLAIM the audience will
        debate — a verb / number / delta / tension — NOT a topic label.
        The deck's topic / audience / year / event goes in `subtitle`.

        Args:
            title: The deck's HEADLINE CLAIM (an assertion, not a topic).
                Drop the deck's headline message here, the ONE thing the
                audience should walk away believing. **Length: 7-16 words.**
                Shorter is a topic label; longer overflows the title shape
                and renders unreadably small. If your brief's
                headline_message exceeds 16 words, compress it to its
                sharpest single sentence — move the longer rationale to
                slide 2 (a solution_detail or bullets slide).
                If you find yourself writing a noun phrase like
                "2026 Economic Outlook" or "Company Objectives 2026" or
                "LLM Integration Plan" — STOP. That's the subtitle. The
                title is the claim that justifies why the audience should
                sit through the deck.

                Good titles (assertions):
                  ✓ "2026 is a regime change, not a continuation"
                  ✓ "Net retention climbs to 118% by Q4"
                  ✓ "Three integration points, one kill-switch"
                  ✓ "Earn the right to grow — three priorities, no fourth"

                Bad titles (topic labels — go in subtitle):
                  ✗ "2026 Economic Outlook"
                  ✗ "Company Objectives 2026"
                  ✗ "LLM Integration Plan"
                  ✗ "Q1 Strategy Review"

            subtitle: Audience / scope / year / event. This is where the
                deck's TOPIC name goes (the noun-phrase). Example forms:
                  • "Strategy & Finance Leaders · 2026 Outlook"
                  • "All-Hands · 2026 Company Objectives"
                  • "Architecture Review Board · Q1 2026"
        """
        # Hard validation — the model otherwise defaults to topic-label
        # cover titles ("2026 Economic Outlook", "Company Objectives 2026")
        # no matter how strongly the skill or this docstring push the
        # opposite. Refuse on shape violation so the model must retry
        # with an actual claim.
        import re
        title_text = (title or "").strip()
        violations: List[str] = []
        # Strip out non-word separators (·, |, &, /) before counting so
        # decorative joiners don't pad the word count.
        content_only = re.sub(r"[·|&/]+", " ", title_text)
        content_words = [w for w in content_only.split() if w]
        word_count = len(content_words)
        if word_count < 7:
            violations.append(
                f"title carries only {word_count} content words. Cover titles "
                f"must be ≥ 7 words to carry a claim — anything shorter is a "
                f"topic label. Got: {title!r}. Move it to `subtitle` and put "
                f"your brief's headline message in `title` instead."
            )
        if word_count > 12:
            violations.append(
                f"title is {word_count} words. Cover titles must be ≤ 12 words — "
                f"longer ones overflow the title shape and orphan words on a "
                f"second line. Got: {title!r}. Compress the claim to its "
                f"sharpest form (single sentence, one comma at most) and move "
                f"the longer-form rationale to slide 2."
            )
        if len(title_text) > 60:
            violations.append(
                f"title is {len(title_text)} characters. Cover titles must be "
                f"≤ 60 chars total — longer ones wrap and orphan words even "
                f"under the word cap. Got: {title!r}. Tighten the wording."
            )
        # An assertion has at least one verb-indicator OR a clause-break
        # punctuation mark. Topic labels typically have neither.
        VERB_INDICATORS = (
            r"\b(?:is|are|was|were|am|be|been|being|has|have|had|"
            r"will|won't|would|wouldn't|shall|should|shouldn't|"
            r"can|cannot|can't|could|couldn't|may|might|must|mustn't|"
            r"do|does|did|don't|doesn't|didn't|"
            r"requires?|demands?|needs?|wants?|becomes?|contains?|"
            r"climbs?|rises?|falls?|drops?|jumps?|grows?|shrinks?|"
            r"beats?|wins?|loses?|takes?|gives?|holds?|drives?|"
            r"ships?|delivers?|lands?|launches?|"
            r"changes?|shifts?|moves?|signals?|marks?|matters?|"
            r"earns?|costs?|saves?|protects?|enables?|preserves?|"
            r"bolts?|cuts?|raises?|builds?|breaks?|opens?|closes?|"
            r"defends?|fights?|wins?|loses?|stops?|starts?|"
            r"unlocks?|forces?|tightens?|loosens?|denies?|admits?)\b"
        )
        has_verb = re.search(VERB_INDICATORS, title_text, re.IGNORECASE) is not None
        has_clause_punct = any(p in title_text for p in (",", "—", "–", ":", "?", "!"))
        if not has_verb and not has_clause_punct:
            violations.append(
                f"title has no verb and no clause-break punctuation — looks "
                f"like a noun-phrase label, not a claim. Got: {title!r}. "
                f"Rewrite as an assertion (a sentence with a verb, or a "
                f"comma/em-dash claim) e.g. 'X is Y' / 'X drives Y' / "
                f"'Three Xs — one Y'."
            )
        if violations:
            return {
                "error": "Cover title is a topic label, not a claim. NO slide was added.",
                "violations": violations,
                "action": (
                    "Rewrite the title as an ASSERTION (a verb / number / delta / "
                    "tension — the ONE thing the audience should walk away "
                    "believing) and move the topic noun-phrase to `subtitle`. "
                    "Use the brief's headline_message verbatim if it fits."
                ),
            }
        return _build_composition("cover", {"title": title, "subtitle": subtitle}, presentation_id)

    @app.tool(
        annotations=ToolAnnotations(title="Add Section Divider"),
    )
    def add_section_divider(
        kicker: str,
        title: str,
        presentation_id: Optional[str] = None,
    ) -> Dict:
        """Add a section divider slide — a small kicker label above a large title.

        Use this between major sections of a deck. For 7-slide decks this is
        rare; for 12+ slide decks it gives the audience a visual breath.

        Args:
            kicker: Short label above the title (e.g. "What We Offer", "Where We Go Next").
            title: Section title (claim form).
        """
        return _build_composition("section_divider", {"kicker": kicker, "title": title}, presentation_id)

    @app.tool(
        annotations=ToolAnnotations(title="Add Timeline Slide"),
    )
    def add_timeline_slide(
        subhead: str,
        events: List[str],
        presentation_id: Optional[str] = None,
    ) -> Dict:
        """Add a horizontal timeline slide — subhead + 4-12 chronological
        events with year/date markers along a horizontal axis.

        Clones the brand template's "Our Growth Story" slide (year-by-year
        history with dots between events). Use for founding histories,
        product launches by year, hiring milestones, customer-journey
        timelines — anything where the SEQUENCE of dates carries the
        meaning.

        Don't use for:
        - 4 sequential PHASES with stages → use `add_process_steps_slide`.
        - Non-chronological lists → use `add_bullets_slide`.

        Args:
            subhead: Section label, ≤ 5 words (e.g. "Our Growth Story",
                "Product Milestones", "Q1 Rollout Plan").
            events: 4-12 chronological items. Each is ONE string starting
                with the date/year, then a short description:
                  "2000 Founded to design better ways of working"
                  "Q3 2024 Closed Series C at $200M valuation"
                  "Sept 2025 Launched the unified platform beta"
                Each event ≤ 12 words total (including the date prefix).
        """
        violations: List[str] = []
        if len((subhead or "").split()) > 5:
            violations.append(
                f"subhead is {len(subhead.split())} words; max is 5. Got: {subhead!r}"
            )
        if not isinstance(events, list) or len(events) < 4 or len(events) > 12:
            violations.append(
                f"events must be a list of 4-12 items; got {len(events) if isinstance(events, list) else type(events).__name__}"
            )
        else:
            for i, ev in enumerate(events):
                text = (ev or "").strip()
                words = text.split()
                wc = len(words)
                if wc < 3 or wc > 10:
                    violations.append(
                        f"events[{i}] is {wc} words; needs 3-10 (date prefix + "
                        f"short description). Got: {ev!r}"
                    )
                # Each event slot is narrow; a single word longer than 12
                # chars will break mid-word ("observability" → "observabilit y").
                long_words = [w for w in words if len(w) > 12]
                if long_words:
                    violations.append(
                        f"events[{i}] contains words > 12 chars: {long_words}. "
                        f"Long words break mid-word in the narrow event slot "
                        f"(e.g. 'observability' renders as 'observabilit y'). "
                        f"Pick shorter synonyms or split the event. Got: {ev!r}"
                    )
        if violations:
            return {
                "error": "timeline content doesn't fit the format. NO slide was added.",
                "violations": violations,
                "action": (
                    "Rewrite per the violations. Each event is one string "
                    "starting with the date/year ('2024 ...', 'Q3 2025 ...', "
                    "'Sept 2025 ...') followed by a 2-10 word description."
                ),
            }
        # Iter 69: redistribute N<12 events evenly across the 12 template
        # slots so the text labels span the full timeline width instead of
        # clumping on the left. User reported (2026-05-31): "milestones
        # should be evenly spread out, not clumped together on the bottom
        # left" on a 6-event roadmap deck.
        SLOT_COUNT = 12
        if 1 < len(events) < SLOT_COUNT:
            n = len(events)
            remapped: List[Any] = [None] * SLOT_COUNT
            # Spread events across [0, SLOT_COUNT-1] inclusive — first
            # always at slot 0, last always at slot SLOT_COUNT-1.
            for i in range(n):
                slot = round(i * (SLOT_COUNT - 1) / (n - 1))
                remapped[slot] = events[i]
            events = remapped
        return _build_composition(
            "timeline",
            {"subhead": subhead, "events": events},
            presentation_id,
        )

    @app.tool(
        annotations=ToolAnnotations(title="Add Split Benefits Slide"),
    )
    def add_split_benefits_slide(
        title: str,
        benefits: List[str],
        presentation_id: Optional[str] = None,
    ) -> Dict:
        """Add a split-layout benefits slide — large title on the left,
        4-6 short benefit lines over the brand stock background on the right.

        Clones the brand template's "Optimize and Reinforce your Operations"
        slide, which is the visual workhorse for any 'here's what you get'
        moment. The right-hand rebar/grid background image gives the slide
        weight and prevents the text-only flatness of bullets.

        Use this when:
        - You're listing 4-6 OUTCOME statements (not bullets of action items).
        - You want the slide to feel visually substantive without writing
          per-card descriptions.
        - The list is parallel — each benefit is the same shape of
          statement (verb-led OR noun-led, pick one).

        Don't use for:
        - Lists of actions / asks → use `add_bullets_slide`.
        - Items that need descriptions → use `add_value_cards_slide`.
        - Capabilities organised by category → use `add_capability_grid_slide`.

        Args:
            title: Section claim, 3-8 words. Pattern: "<Verb> <noun>"
                or "<Noun phrase>" (e.g. "Optimize operations across the
                enterprise", "Three guardrails on every LLM call").
            benefits: 4-6 short benefit / outcome lines, 5-12 words each.
                Pick parallel grammar: all verb-led OR all noun-led, not
                mixed. Examples:
                  ["Achieve 360° insight across your digital core",
                   "Reduce operational risk and technology costs",
                   "Reinforce governance across the enterprise",
                   "Accelerate transformation with automation and AI"]
        """
        violations: List[str] = []
        title_text = (title or "").strip()
        title_wc = len(title_text.split())
        if title_wc < 3 or title_wc > 8:
            violations.append(
                f"title is {title_wc} words; needs 3-8. Got: {title!r}"
            )
        # Template slide 26 has 6 benefit paragraphs each preceded by a
        # separate decorative-bar shape. Supplying 5 leaves one orphaned
        # bar visible (paragraph-delete clears the text but doesn't
        # touch the separate decoration shape). Restrict to 4 or 6 to
        # avoid the awkward trailing-bar artifact.
        if not isinstance(benefits, list) or len(benefits) not in (4, 6):
            count = len(benefits) if isinstance(benefits, list) else type(benefits).__name__
            violations.append(
                f"benefits must be a list of EXACTLY 4 or 6 items; got "
                f"{count}. The template has 6 slots each with a decorative "
                f"bar; supplying 5 leaves an orphaned bar visible. Pad to 6 "
                f"or trim to 4."
            )
        else:
            for i, line in enumerate(benefits):
                wc = len((line or "").split())
                if wc < 5 or wc > 12:
                    violations.append(
                        f"benefits[{i}] is {wc} words; needs 5-12. Got: {line!r}"
                    )
        if violations:
            return {
                "error": "split_benefits content doesn't fit the format. NO slide was added.",
                "violations": violations,
                "action": (
                    "Rewrite per the violations. title 3-8 words, benefits "
                    "4-6 lines of 5-12 words each, all sharing the same "
                    "grammatical pattern (verb-led OR noun-led, not mixed)."
                ),
            }
        return _build_composition(
            "split_benefits",
            {"title": title, "benefits": benefits},
            presentation_id,
        )

    def _add_image_slide(layout_name: str, title: str, png_bytes: bytes, caption: Optional[str], composition_name: str, presentation_id: Optional[str]) -> Dict:
        """Shared internal helper for diagram + chart compositions.

        Adds a fresh slide using the named layout, sets the title
        placeholder, clears the body placeholder, drops the PNG centered
        in the slide body area with proper aspect ratio, optionally adds
        a small caption below, and tags the slide with composition_name.
        """
        pres_id = presentation_id if presentation_id is not None else get_current_presentation_id()
        if pres_id is None or pres_id not in presentations:
            return {"error": "No presentation is currently loaded."}
        working = presentations[pres_id]
        layout = next((l for l in working.slide_layouts if l.name == layout_name), None)
        if layout is None:
            return {"error": f"Layout {layout_name!r} not found in working presentation."}

        new_slide = working.slides.add_slide(layout)

        # Set title via the layout's title placeholder and widen it to
        # the full slide width so 8-12 word titles fit on a single line
        # instead of wrapping into the body region.
        if new_slide.shapes.title is not None:
            new_slide.shapes.title.text = _normalise_text(title)
            _rename_shape(new_slide.shapes.title, "title")
            _widen_title_full_width(new_slide, working)

        # Iter 52 fix: REMOVE every non-title placeholder rather than
        # just clearing its text frame. Clearing leaves the placeholder
        # shape in place, and PowerPoint then renders the layout's
        # default "Click to add text" prompt for that empty shape. The
        # user-reported chart slide had two visible "Click to add text"
        # placeholders flanking the chart image (langfuse trace
        # 4be8f6cc, 2026-05-30). Removing the shapes outright is the
        # only way to suppress the prompt.
        for ph in list(new_slide.placeholders):
            if ph.placeholder_format.idx == 0:
                continue  # keep title
            sp = ph._element
            sp.getparent().remove(sp)

        # Drop the picture centered in the body region.
        # Standard 16:9 slide is 13.333in × 7.5in (12192000 × 6858000 EMU).
        # Reserve top ~1.9in for the title — the Basic Text layout's title
        # placeholder spans ~1.0in tall but sits ~0.5in from the top with
        # extra breathing room before the body, so anything under ~1.8in
        # collides with the title (seen on chart slides where wide PNGs
        # take maximum height and shove up against the title text).
        slide_w = working.slide_width
        slide_h = working.slide_height
        top_reserve = Inches(1.9)
        bottom_reserve = Inches(0.8) if caption else Inches(0.4)
        avail_w = slide_w - Inches(1.0)  # 0.5in margin each side
        avail_h = slide_h - top_reserve - bottom_reserve

        # Probe PNG dimensions to compute aspect ratio.
        try:
            from PIL import Image  # Pillow is already a dep
            with Image.open(BytesIO(png_bytes)) as im:
                img_w, img_h = im.size
        except Exception:
            img_w, img_h = 800, 500  # safe fallback

        img_aspect = img_w / max(img_h, 1)
        avail_aspect = avail_w / max(avail_h, 1)
        if img_aspect > avail_aspect:
            picture_w = avail_w
            picture_h = Emu(int(avail_w / img_aspect))
        else:
            picture_h = avail_h
            picture_w = Emu(int(avail_h * img_aspect))
        picture_left = Emu(int((slide_w - picture_w) / 2))
        picture_top = top_reserve + Emu(int((avail_h - picture_h) / 2))

        png_stream = BytesIO(png_bytes)
        pic = new_slide.shapes.add_picture(png_stream, picture_left, picture_top, picture_w, picture_h)
        _rename_shape(pic, "diagram_image" if composition_name == "diagram" else "chart_image")

        # Optional caption below the picture.
        if caption:
            cap_top = picture_top + picture_h + Inches(0.1)
            cap_left = Inches(0.5)
            cap_w = slide_w - Inches(1.0)
            cap_h = Inches(0.5)
            cap_box = new_slide.shapes.add_textbox(cap_left, cap_top, cap_w, cap_h)
            cap_box.text_frame.text = caption
            for para in cap_box.text_frame.paragraphs:
                para.alignment = 2  # CENTER
            _rename_shape(cap_box, "caption")

        _tag_slide_with_composition(new_slide, composition_name)

        return {
            "composition": composition_name,
            "slide_index": len(working.slides) - 1,
            "message": f"Added '{composition_name}' as slide {len(working.slides) - 1}.",
        }

    @app.tool(
        annotations=ToolAnnotations(title="Add Diagram Slide"),
    )
    def add_diagram_slide(
        title: str,
        mermaid: str,
        caption: Optional[str] = None,
        presentation_id: Optional[str] = None,
    ) -> Dict:
        """Add a slide with a server-rendered Mermaid diagram.

        Use this when the content is best shown as a diagram — architecture
        layouts, sequence flows, decision trees, org charts, dependency
        graphs. Supply the diagram in Mermaid source syntax; the server
        renders it via kroki.io and embeds the resulting PNG into a new
        slide using the Basic Text layout (title at top, diagram centered).

        Args:
            title: Slide title, ≤ 10 words. A claim about what the diagram
                shows, not a topic label. ✓ "Side-car sits between modeler
                and Bedrock" ✗ "Architecture Diagram".
            mermaid: Mermaid source text. Supported diagram types include
                flowchart, sequenceDiagram, classDiagram, stateDiagram,
                erDiagram, journey, gantt, pie, mindmap. Example:
                  "flowchart LR\\n  A[User] --> B[Modeler] --> C[Side-car]\\n  C --> D[Bedrock]"
                Keep it focused — 3-12 nodes is the sweet spot. Diagrams
                with 30+ nodes render unreadably small in a slide.
            caption: Optional one-line caption below the diagram (≤ 15
                words). Use to call out the headline reading of the
                diagram, e.g. "Three integration points, one kill-switch".

        Returns the slide_index of the new slide and a composition tag
        of `diagram` so validate_deck + the assess loop know how to
        score it.
        """
        violations: List[str] = []
        if not title or not title.strip():
            violations.append("title is empty")
        elif len(title.split()) > 10:
            violations.append(f"title is {len(title.split())} words; max is 10. Got: {title!r}")
        if not mermaid or not mermaid.strip():
            violations.append("mermaid source is empty")
        elif len(mermaid) > 4000:
            violations.append(f"mermaid source is {len(mermaid)} chars; max is 4000 (≈ a couple dozen nodes).")
        if caption is not None and len(caption.split()) > 15:
            violations.append(f"caption is {len(caption.split())} words; max is 15.")
        if violations:
            return {
                "error": "add_diagram_slide content doesn't fit. NO slide was added.",
                "violations": violations,
            }
        try:
            png_bytes = _fetch_diagram_png(mermaid)
        except Exception as e:
            return {
                "error": f"Diagram render failed via kroki.io: {e!r}",
                "action": (
                    "Check the mermaid source for syntax errors. Common "
                    "issues: missing 'flowchart LR' header, unbalanced "
                    "brackets, reserved words. Try rendering at "
                    "https://mermaid.live/ to validate first."
                ),
            }
        return _add_image_slide(
            layout_name="Basic Text",
            title=title,
            png_bytes=png_bytes,
            caption=caption,
            composition_name="diagram",
            presentation_id=presentation_id,
        )

    @app.tool(
        annotations=ToolAnnotations(title="Add Chart Slide"),
    )
    def add_chart_slide(
        title: str,
        chart_type: str,
        labels: List[str],
        datasets: List[Dict[str, Any]],
        caption: Optional[str] = None,
        y_axis_title: Optional[str] = None,
        x_axis_title: Optional[str] = None,
        presentation_id: Optional[str] = None,
    ) -> Dict:
        """Add a slide with a server-rendered chart (bar, line, pie, doughnut).

        Use this when the content is a quantitative comparison best shown
        as a chart — revenue over time, market share split, before/after
        metrics. The server posts to QuickChart (Chart.js v4) and embeds
        the resulting PNG.

        For 4 standalone stat numbers WITHOUT axis context (totals,
        percentages, counts) prefer add_stat_cards_slide instead.

        Args:
            title: Slide title, ≤ 10 words. A claim about what the chart
                shows. ✓ "ARR triples while burn stays flat" ✗ "ARR Chart".
            chart_type: One of "bar", "line", "pie", "doughnut".
            labels: x-axis labels (or category names for pie/doughnut),
                e.g. ["FY22", "FY23", "FY24", "FY25", "FY26"].
            datasets: One or more Chart.js dataset objects. Each is a dict
                with at least:
                  - label: legend entry (string, optional for single-series)
                  - data: list of numbers, same length as labels
                Example:
                  [{"label": "ARR ($M)", "data": [8, 18, 31, 42, 55]}]
                For pie/doughnut, use a single dataset; the values become
                the slice sizes.
            caption: Optional one-line caption (≤ 15 words).
            y_axis_title: REQUIRED for bar/line charts. Y-axis label
                including unit, e.g. "ARR ($M)", "GDP growth (% YoY)",
                "Latency p95 (ms)". Visual judge consistently flags
                un-labelled axes as a severe defect.
            x_axis_title: Optional X-axis label. Often unnecessary when
                labels are self-explanatory (years, quarters, regions).

        Returns the slide_index of the new slide and composition tag `chart`.
        """
        violations: List[str] = []
        if not title or not title.strip():
            violations.append("title is empty")
        elif len(title.split()) > 10:
            violations.append(f"title is {len(title.split())} words; max is 10.")
        valid_types = {"bar", "line", "pie", "doughnut"}
        if chart_type not in valid_types:
            violations.append(f"chart_type must be one of {sorted(valid_types)}; got {chart_type!r}")
        if not isinstance(labels, list) or len(labels) < 2 or len(labels) > 20:
            violations.append(f"labels must be a list of 2-20 strings; got {labels!r}")
        if not isinstance(datasets, list) or len(datasets) < 1 or len(datasets) > 5:
            violations.append(f"datasets must be a list of 1-5 dataset objects; got {datasets!r}")
        else:
            for i, ds in enumerate(datasets):
                if not isinstance(ds, dict):
                    violations.append(f"datasets[{i}] must be a dict")
                    continue
                data = ds.get("data")
                if not isinstance(data, list) or not data:
                    violations.append(f"datasets[{i}].data must be a non-empty list of numbers")
                elif isinstance(labels, list) and len(data) != len(labels):
                    violations.append(
                        f"datasets[{i}].data has {len(data)} values but labels has {len(labels)}. They must match."
                    )
        if caption is not None and len(caption.split()) > 15:
            violations.append(f"caption is {len(caption.split())} words; max is 15.")
        # Iter 64: require y_axis_title for bar/line charts — visual
        # judge flagged un-labelled axes as severe defect on multiple
        # iterations (econ-2026 slide 4/5 iter 55+62+63).
        if chart_type in ("bar", "line") and not (y_axis_title and y_axis_title.strip()):
            violations.append(
                f"y_axis_title is required for {chart_type} charts. "
                f"Provide the unit, e.g. 'ARR ($M)', 'GDP growth (% YoY)', "
                f"'Latency p95 (ms)'."
            )
        if violations:
            return {
                "error": "add_chart_slide content doesn't fit. NO slide was added.",
                "violations": violations,
            }

        chart_options: Dict[str, Any] = {
            "plugins": {
                "legend": {"display": len(datasets) > 1 or chart_type in ("pie", "doughnut")},
            },
            "responsive": False,
        }
        # Iter 64: add axis titles + data labels for bar/line charts.
        if chart_type in ("bar", "line"):
            scales: Dict[str, Any] = {}
            if y_axis_title and y_axis_title.strip():
                scales["y"] = {"title": {"display": True, "text": y_axis_title.strip()}}
            if x_axis_title and x_axis_title.strip():
                scales["x"] = {"title": {"display": True, "text": x_axis_title.strip()}}
            if scales:
                chart_options["scales"] = scales
            # Data labels on top of each point/bar — agent-visible numbers
            # the visual judge looks for.
            chart_options["plugins"]["datalabels"] = {
                "anchor": "end",
                "align": "top",
                "color": "#333",
                "font": {"weight": "bold", "size": 12},
            }

        # Iter 67: apply Bizzdesign brand palette to datasets that don't
        # already carry per-series colors. Improves visual cohesion with
        # the brand template.
        BRAND_PALETTE = ["#1659FF", "#00247E", "#53D576", "#FFD607", "#9259EF"]
        styled_datasets: List[Dict[str, Any]] = []
        for i, ds in enumerate(datasets):
            ds = dict(ds)  # avoid mutating caller's dict
            color = BRAND_PALETTE[i % len(BRAND_PALETTE)]
            if chart_type in ("pie", "doughnut"):
                if "backgroundColor" not in ds:
                    n = len(ds.get("data") or [])
                    ds["backgroundColor"] = [BRAND_PALETTE[k % len(BRAND_PALETTE)] for k in range(n)]
            elif chart_type == "line":
                ds.setdefault("borderColor", color)
                ds.setdefault("backgroundColor", color)
                ds.setdefault("borderWidth", 3)
                ds.setdefault("pointRadius", 4)
            else:  # bar
                ds.setdefault("backgroundColor", color)
                ds.setdefault("borderColor", color)
            styled_datasets.append(ds)

        chart_config = {
            "type": chart_type,
            "data": {
                "labels": labels,
                "datasets": styled_datasets,
            },
            "options": chart_options,
        }
        try:
            png_bytes = _fetch_chart_png(chart_config)
        except Exception as e:
            return {
                "error": f"Chart render failed via quickchart.io: {e!r}",
                "action": "Check the chart_type, labels, and datasets for typos. Each dataset's `data` array must be the same length as `labels`.",
            }
        return _add_image_slide(
            layout_name="Basic Text",
            title=title,
            png_bytes=png_bytes,
            caption=caption,
            composition_name="chart",
            presentation_id=presentation_id,
        )

    @app.tool(
        annotations=ToolAnnotations(title="Add Closing Slide"),
    )
    def add_closing_slide(
        title: str,
        cta_lines: List[str],
        presentation_id: Optional[str] = None,
    ) -> Dict:
        """Add the deck's closing slide — must be the FINAL slide of every deck.

        The closing slide carries the wrap-up message and the calls-to-action
        the audience should leave with (book a demo, contact us, visit the
        portal, etc.). Clones the brand template's "Thank You" slide.

        Args:
            title: Closing title. Defaults to "Thank You" if you want the
                stock wording. You can override with a deck-specific
                closing claim like "Earn the right to grow — start Monday"
                or "Three asks, one decision".
            cta_lines: 2 to 6 short call-to-action lines. Each is one line
                (8-20 words). Example for an external-facing deck:
                  ["Book a demo to see the platform in action.",
                   "Contact your account team to scope a pilot.",
                   "Visit the support portal for technical docs."]
                For an internal all-hands, use action-orientated lines:
                  ["Q1 reforecast happens 30 days from today.",
                   "Owner check-ins: Sarah, Marcus, Priya every Friday.",
                   "If you have a question, message #2026-priorities."]
        """
        violations: List[str] = []
        title_text = (title or "").strip()
        if not title_text:
            violations.append("title is empty")
        elif len(title_text.split()) > 6:
            # Tightened from 8 → 6 in iter 40 after the platform-rearch
            # closing slide rendered an 8-word title across THREE lines and
            # overflowed the title shape (the template title font is ~80pt
            # and the shape only fits ~5 words per line). 6-word cap gives
            # the model one wrap and still reads as a clean closer.
            violations.append(
                f"title is {len(title_text.split())} words; closing titles must "
                f"be ≤ 6 words to fit the closing layout's 80pt title font "
                f"without wrapping past two lines. Got: {title!r}."
            )
        elif len(title_text) > 40:
            violations.append(
                f"title is {len(title_text)} chars; closing titles must be ≤ 40 "
                f"chars total (tightened from 50 in iter 40 — long-word titles "
                f"like 'Kafka ships week 6 — #platform-rearch' overflow even "
                f"under the 8-word cap). Got: {title!r}."
            )
        if not isinstance(cta_lines, list) or len(cta_lines) < 2:
            violations.append(
                f"cta_lines must be a list of 2-6 strings; got {cta_lines!r}"
            )
        elif len(cta_lines) > 6:
            violations.append(
                f"cta_lines has {len(cta_lines)} items; max is 6."
            )
        else:
            for i, line in enumerate(cta_lines):
                line_text = (line or "").strip()
                wc = len(line_text.split())
                cc = len(line_text)
                if wc < 4 or wc > 12:
                    # Tightened max from 20 → 12 in iter 40. The closing
                    # template splits cta_lines into two narrow columns;
                    # anything > 12 words wraps to 3+ lines and the right
                    # edge truncates against the column boundary.
                    violations.append(
                        f"cta_lines[{i}] is {wc} words; needs 4-12 to fit the "
                        f"narrow CTA columns without right-edge truncation. "
                        f"Got: {line!r}"
                    )
                elif cc > 50:
                    # Iter 57: tightened from 70 → 50 chars. Combined with
                    # the 14pt cta_lines font (set_run_sz_for_role), 50
                    # chars fits the narrow column without wrapping;
                    # 70 chars wrapped to two lines on platform-rearch
                    # iter 55 slide 9.
                    violations.append(
                        f"cta_lines[{i}] is {cc} chars; closing CTAs must be "
                        f"≤ 50 chars to fit the column at 14pt without "
                        f"wrapping. Got: {line!r}"
                    )
        if violations:
            return {
                "error": "closing slide content doesn't fit the format. NO slide was added.",
                "violations": violations,
                "action": (
                    "Rewrite per the violations. title is the final claim "
                    "(≤ 6 words AND ≤ 40 chars). cta_lines are 2-6 short action/"
                    "contact lines (4-12 words AND ≤ 70 chars each)."
                ),
            }
        # Render cta_lines as newline-joined string so _set_shape_multiline
        # writes each as its own paragraph in the CTA shape.
        return _build_composition(
            "closing",
            {"title": title, "cta_lines": cta_lines},
            presentation_id,
        )

    @app.tool(
        annotations=ToolAnnotations(title="Add Value Props (4 Pillars)"),
    )
    def add_value_props_slide(
        title: str,
        lead_claim: str,
        subtitle: str,
        pillars: List[str],
        presentation_id: Optional[str] = None,
    ) -> Dict:
        """Add a 4-pillar value-props slide with a lead claim and subtitle.

        Right side: a stack of 4 short pillar statements (each ~5-8 words).
        Left side: the deck's lead claim with a smaller subtitle beneath.

        Use this when you have exactly 4 parallel differentiators or value
        propositions. Don't use for 3 or fewer — pick `solution_detail` or
        `bullets` instead.

        Args:
            title: Main title (e.g. "Why Bizzdesign?").
            lead_claim: One-line lead claim (e.g. "Only true end-to-end Enterprise Transformation offering").
            subtitle: Supporting subtitle.
            pillars: List of exactly 4 short pillar statements.
        """
        return _build_composition(
            "value_props_4",
            {"title": title, "lead_claim": lead_claim, "subtitle": subtitle, "pillars": pillars},
            presentation_id,
        )

    @app.tool(
        annotations=ToolAnnotations(title="Add Solution Detail"),
    )
    def add_solution_detail_slide(
        category: str,
        title: str,
        benefits: str,
        details: List[str],
        presentation_id: Optional[str] = None,
    ) -> Dict:
        """Add the workhorse "solution detail" composition.

        Use this for a substantive topic that needs 3 named sub-items
        each with a real description. The slide has:
        - A small kicker (`category`) at the top
        - A large product/area title beneath
        - A multi-line benefits paragraph (the "why this matters")
        - 3 detail items in a row at the bottom (each: bold heading + body description)

        IMPORTANT: this composition only EARNS its space when the
        details have substantive descriptions. A solution_detail slide
        with three bare headings ("Chat & Orchestration", "Skills",
        "Data Access") and nothing else looks empty — use
        `add_bullets_slide` for that shape instead. The composition
        validates and rejects calls where details lack descriptions.

        Args:
            category: Top-of-slide kicker (≤ 6 words). E.g. "Layer 1 · Foundation", "Now — Q2 2026".
            title: Main heading (≤ 10 words). The topic name.
            benefits: Substantive paragraph (15-60 words) explaining why
                this topic matters — what changes, what unlocks, what
                the reader takes away. Use real newlines for paragraph
                breaks if needed.
            details: EXACTLY 3 detail items. Each item is a string with
                a heading on line 1 and a description on lines 2+, e.g.
                "Chat & Orchestration\\nPersistent conversations, interruptable
                streams, and a routing layer that picks the right agent per
                turn — the runtime every assistant talks through."
                Each heading: ≤ 6 words. Each description: 12-30 words
                (substantive, specifics-laden). Headings without
                descriptions are REJECTED.
        """
        # Hard validation — reject on shape violations rather than
        # silently produce a thin slide. Matches the stat_cards pattern.
        violations: List[str] = []
        if len(category.split()) > 6:
            violations.append(
                f"category is {len(category.split())} words; max is 6. Got: {category!r}"
            )
        if len(title.split()) > 10:
            violations.append(
                f"title is {len(title.split())} words; max is 10. Got: {title!r}"
            )
        benefits_text = _normalise_text(benefits) if isinstance(benefits, str) else str(benefits)
        benefits_word_count = len(benefits_text.split())
        if benefits_word_count < 15:
            violations.append(
                f"benefits is only {benefits_word_count} words; needs 15-60 to "
                f"justify this composition. Got: {benefits!r}"
            )
        if benefits_word_count > 60:
            violations.append(
                f"benefits is {benefits_word_count} words; max is 60. Got: {benefits!r}"
            )
        if len(details) != 3:
            violations.append(
                f"details must have exactly 3 items; got {len(details)}."
            )
        for i, item in enumerate(details[:3]):
            normalised = _normalise_text(item) if isinstance(item, str) else str(item)
            lines = [l.strip() for l in normalised.split("\n") if l.strip()]
            if len(lines) < 2:
                violations.append(
                    f"details[{i}] needs a heading AND a description separated by "
                    f"newline. Got just: {item!r}"
                )
                continue
            heading = lines[0]
            description = " ".join(lines[1:])
            if len(heading.split()) > 6:
                violations.append(
                    f"details[{i}] heading is {len(heading.split())} words; max is 6. Got: {heading!r}"
                )
            description_word_count = len(description.split())
            if description_word_count < 8:
                violations.append(
                    f"details[{i}] description is only {description_word_count} words; "
                    f"needs ≥ 8 to be substantive. Got: {description!r}"
                )
            if description_word_count > 35:
                violations.append(
                    f"details[{i}] description is {description_word_count} words; max is 35. Got: {description!r}"
                )
        if violations:
            return {
                "error": "solution_detail content doesn't justify the composition. NO slide was added.",
                "violations": violations,
                "action": (
                    "Rewrite per the violations and retry. If you don't have "
                    "12-30 words of real description for each detail, this is the "
                    "wrong composition for your content — use `add_bullets_slide` "
                    "(headings only) or `add_capability_grid_slide` (more headings, "
                    "no per-item description) instead. solution_detail is for "
                    "topics that DESERVE a deep slide."
                ),
            }
        return _build_composition(
            "solution_detail",
            {"category": category, "title": title, "benefits": benefits, "details": details},
            presentation_id,
        )

    @app.tool(
        annotations=ToolAnnotations(title="Add Bullets Slide"),
    )
    def add_bullets_slide(
        title: str,
        subhead: str,
        bullets: List[str],
        presentation_id: Optional[str] = None,
    ) -> Dict:
        """Add a clean title + subhead + brand-styled bullet list slide.

        Reach for this BEFORE `add_solution_detail_slide` when the content
        is a simple LIST without sub-features per item. The brand styling
        on the bullets comes from the Basic Text layout's body
        placeholder — bullet glyphs, font, and indentation are all
        template-driven.

        Args:
            title: Slide title (claim form, e.g. "Three things we need from leadership").
            subhead: SHORT kicker line above the title — 1-5 words, ALL CAPS style
                (e.g. "NEXT STEPS", "PRIORITY 1", "HIRING PLAN"). NOT a one-line
                intro paragraph — that overflows the kicker shape and looks
                broken. If you need a full sentence to introduce the bullets,
                use solution_detail's `benefits` field instead.
            bullets: List of 3-6 short bullets (≤ 15 words each, parallel grammar).
        """
        violations: List[str] = []
        sub_wc = len((subhead or "").split())
        if sub_wc > 5:
            violations.append(
                f"subhead is {sub_wc} words; bullets.subhead must be 1-5 words "
                f"(kicker-style). Got: {subhead!r}. If you need a longer "
                f"intro line, use solution_detail with that text in `benefits`."
            )
        if violations:
            return {
                "error": "bullets.subhead too long. NO slide was added.",
                "violations": violations,
                "action": (
                    "Shorten subhead to 1-5 words (an ALL-CAPS kicker like "
                    "'NEXT STEPS' or 'HIRING PLAN'). For a full intro "
                    "sentence, switch to add_solution_detail_slide."
                ),
            }
        return _build_composition(
            "bullets",
            {"title": title, "subhead": subhead, "bullets": bullets},
            presentation_id,
        )

    @app.tool(
        annotations=ToolAnnotations(title="Add Stat Cards Slide"),
    )
    def add_stat_cards_slide(
        subhead: str,
        intro: str,
        stats: List[str],
        presentation_id: Optional[str] = None,
    ) -> Dict:
        """Add a slide with 4 big-number stat cards. Each card has a small
        physical footprint (~3in wide x 1in tall) — content MUST be tight.

        HARD CHARACTER LIMITS — content that exceeds these overflows the
        card boundary and looks broken. The composition returns a warning
        if any limit is exceeded; rewrite shorter and retry.

        Each stat is a single string with up to THREE lines separated by
        actual newlines:

            VALUE\\nLABEL\\nATTRIBUTION

        - VALUE:   1-6 chars. A number or short symbol. ("60+", "$5.5M", "100%", "Q2", "3x")
        - LABEL:   1-2 words, max 14 chars. The "what" of the number. ("Capabilities", "Ship Date", "Layers")
        - ATTRIBUTION: optional. max 32 chars. Brief context. ("Across 5 layers", "Achieved by UK insurer")

        GOOD:   "60+\\nCapabilities\\nAcross 5 layers"
        BAD:    "60+\\nCapabilities Mapped"           (LABEL 19 chars → wraps and overflows)
        BAD:    "6\\nCapability Layers"               (LABEL 17 chars → wraps and overflows)
        BAD:    "Six different\\nThings\\nFor reasons"  (VALUE should be a number/symbol, not words)

        IF YOU GET AN OVERFLOW ERROR: rewrite shorter and retry — do NOT
        call this tool again with different content; that creates a duplicate
        slide. The validation runs BEFORE any slide is built; an error
        means nothing was added to the deck.

        Args:
            subhead: Small kicker line at the top, max 5 words. ("By the numbers", "Today's scale", "Our impact").
            intro: One-line intro sentence above the cards, max 12 words.
            stats: List of 3 or 4 stats. Each follows the VALUE/LABEL/ATTRIBUTION
                contract above. Fewer than 4 supplied → unused slots are blanked.
        """
        # Validate field lengths BEFORE building the slide. Any limit
        # violation rejects the whole call — the model rewrites and
        # retries with compliant content. No slide is added to the deck
        # on rejection, so retries don't accumulate duplicates.
        violations: List[str] = []
        if len(subhead.split()) > 5:
            violations.append(
                f"subhead is {len(subhead.split())} words; max is 5. Got: {subhead!r}"
            )
        intro_text = (intro or "").strip()
        intro_wc = len(intro_text.split())
        if intro_wc > 12:
            violations.append(
                f"intro is {intro_wc} words; max is 12. Got: {intro!r}"
            )
        if intro_wc < 7:
            violations.append(
                f"intro is only {intro_wc} words; needs ≥ 7 to carry a claim "
                f"about what the stats SHOW. A short noun phrase like 'Four data "
                f"points...' is a topic label, not the slide's headline. Got: {intro!r}"
            )
        # Require clause-break punctuation in the intro. Topic-shaped
        # noun phrases ("Five data points that define the regime") fail
        # this check and force the model to write a multi-clause claim
        # ("Growth slows, inflation persists, rate cuts delayed").
        intro_has_clause_punct = any(p in intro_text for p in (",", "—", "–", ":", ";", "?", "!"))
        if intro_text and not intro_has_clause_punct:
            violations.append(
                f"intro has no clause-break punctuation (comma / em-dash / colon). "
                f"Stat-card intros must be a claim sentence, not a topic noun "
                f"phrase. Got: {intro!r}. Rewrite as a multi-clause statement "
                f"about what the stats reveal, e.g. 'Growth slows, inflation "
                f"persists, rate cuts delayed' or 'Capex bifurcates — AI surges, "
                f"non-tech freezes'."
            )
        for i, stat in enumerate(stats[:4]):
            normalised = _normalise_text(stat) if isinstance(stat, str) else str(stat)
            lines = normalised.split("\n")
            if len(lines) < 2:
                violations.append(
                    f"stats[{i}] has only {len(lines)} line(s); needs at least VALUE + LABEL "
                    f"separated by newline. Got: {stat!r}"
                )
                continue
            value, label = lines[0].strip(), lines[1].strip()
            attribution = lines[2].strip() if len(lines) > 2 else ""
            if len(value) > 6:
                violations.append(
                    f"stats[{i}] VALUE is {len(value)} chars (max 6). Got: {value!r}"
                )
            if len(label) > 14:
                violations.append(
                    f"stats[{i}] LABEL is {len(label)} chars (max 14). Got: {label!r}"
                )
            if attribution and len(attribution) > 32:
                violations.append(
                    f"stats[{i}] ATTRIBUTION is {len(attribution)} chars (max 32). Got: {attribution!r}"
                )
        if violations:
            return {
                "error": "Stat-card content exceeds character limits and would overflow visually. NO slide was added.",
                "violations": violations,
                "action": (
                    "Rewrite the offending stats with shorter wording, then call "
                    "add_stat_cards_slide again. Examples of fits: LABEL like "
                    "'Capabilities', 'Layers', 'Ship Date'; ATTRIBUTION like "
                    "'Across 5 layers', 'Q2 2026 ship'. If you can't compress a "
                    "stat without losing meaning, drop it from the stats list — "
                    "3 sharp stats land harder than 4 wrapped ones."
                ),
            }
        return _build_composition(
            "stat_cards",
            {"subhead": subhead, "intro": intro, "stats": stats},
            presentation_id,
        )

    @app.tool(
        annotations=ToolAnnotations(title="Delete Slide", destructiveHint=True),
    )
    def delete_slide(
        slide_index: int,
        presentation_id: Optional[str] = None,
    ) -> Dict:
        """Delete a single slide from the working presentation by index.

        Use this to clean up duplicate or overflow slides created during
        iteration — e.g. if you called a composition tool, got an
        overflow error, and an earlier attempt left an unwanted slide.
        After deletion, downstream slide indices shift down by 1.

        Args:
            slide_index: 0-based index of the slide to remove.
        """
        pres_id = presentation_id if presentation_id is not None else get_current_presentation_id()
        if pres_id is None or pres_id not in presentations:
            return {"error": "No presentation is currently loaded."}
        pres = presentations[pres_id]
        xml_slides = pres.slides._sldIdLst
        slide_entries = list(xml_slides)
        if slide_index < 0 or slide_index >= len(slide_entries):
            return {
                "error": (
                    f"slide_index {slide_index} out of range. Deck has "
                    f"{len(slide_entries)} slides (valid indices 0..{len(slide_entries) - 1})."
                )
            }
        slide_entry = slide_entries[slide_index]
        try:
            pres.part.drop_rel(slide_entry.rId)
            xml_slides.remove(slide_entry)
        except Exception as e:
            return {"error": f"Failed to delete slide {slide_index}: {e}"}
        return {
            "message": f"Deleted slide {slide_index}.",
            "remaining_slide_count": len(pres.slides),
        }

    @app.tool(
        annotations=ToolAnnotations(title="Add Capability Grid Slide"),
    )
    def add_capability_grid_slide(
        subhead: str,
        cross_label: str,
        cross_intro: str,
        col_headings: List[str],
        cards: List[str],
        presentation_id: Optional[str] = None,
    ) -> Dict:
        """Add a 3-column × 3-row grid of capability cards.

        Use this for "the platform at a glance" — many capabilities
        organised into 3 columns with up to 3 cards per column. Each
        card has a heading AND a short description.

        Args:
            subhead: Top-of-slide subhead, ≤ 6 words.
            cross_label: Cross-cutting label that spans the top center, ≤ 5 words.
            cross_intro: Intro paragraph framing the grid, 8-20 words.
            col_headings: EXACTLY 3 column headings, each ≤ 4 words.
            cards: EXACTLY 3, 6, or 9 cards. The server distributes cards
                ACROSS the three columns so each column shows roughly the
                same number of cards.
                  - n=3 → ONE card per column at row 1. Order: [col1,
                    col2, col3]. e.g. col_headings ["Search", "Canvas",
                    "Documentation"] + cards [search-card, canvas-card,
                    doc-card].
                  - n=6 → TWO cards per column. Order:
                    [col1-row1, col1-row2, col2-row1, col2-row2,
                    col3-row1, col3-row2].
                  - n=9 → full 3×3 grid. Order:
                    [col1-row1, col1-row2, col1-row3, col2-row1, ...,
                    col3-row3].
                If your content is 4 cards → use value_cards; 5 →
                bullets; 7-8 → split across two capability_grid slides.
                Each card is "Heading\\nOne-line description". Heading
                ≤ 4 words. Description 5-15 words.
        """
        violations: List[str] = []
        if len(subhead.split()) > 6:
            violations.append(f"subhead is {len(subhead.split())} words; max 6. Got: {subhead!r}")
        if len(cross_label.split()) > 5:
            violations.append(f"cross_label is {len(cross_label.split())} words; max 5. Got: {cross_label!r}")
        cross_intro_text = _normalise_text(cross_intro) if isinstance(cross_intro, str) else str(cross_intro)
        ci_words = len(cross_intro_text.split())
        if ci_words < 8 or ci_words > 20:
            violations.append(f"cross_intro is {ci_words} words; needs 8-20. Got: {cross_intro!r}")
        if len(col_headings) != 3:
            violations.append(f"col_headings must have exactly 3 items; got {len(col_headings)}.")
        for i, h in enumerate(col_headings[:3]):
            if len(h.split()) > 4:
                violations.append(f"col_headings[{i}] is {len(h.split())} words; max 4. Got: {h!r}")
        # The template arranges cards 3 per column × 3 columns. We need a
        # FULL column at minimum — supplying fewer than 3 cards visibly
        # empties the right-hand column(s). Acceptable counts are 3
        # (1 column populated), 6 (2 columns), 9 (full grid). 4 / 5 / 7 /
        # 8 leave half-empty columns and look broken — refuse those.
        if len(cards) not in (3, 6, 9):
            violations.append(
                f"cards must have 3, 6, or 9 items (one full column at a time); "
                f"got {len(cards)}. The template lays out cards 3-per-column × 3 "
                f"columns; partial columns render as awkward whitespace. Round "
                f"to 3/6/9, or use a different composition: 4 cards → "
                f"value_cards or value_props_4; 5 → bullets or split_benefits; "
                f"7-8 → bullets or two capability_grid slides."
            )
        for i, card in enumerate(cards[:9]):
            normalised = _normalise_text(card) if isinstance(card, str) else str(card)
            lines = [l.strip() for l in normalised.split("\n") if l.strip()]
            if len(lines) < 2:
                violations.append(
                    f"cards[{i}] needs 'Heading\\nDescription'. Got just: {card!r}"
                )
                continue
            heading = lines[0]
            description = " ".join(lines[1:])
            if len(heading.split()) > 4:
                violations.append(f"cards[{i}] heading is {len(heading.split())} words; max 4. Got: {heading!r}")
            d_words = len(description.split())
            if d_words < 5 or d_words > 15:
                violations.append(
                    f"cards[{i}] description is {d_words} words; needs 5-15 (small card slot). "
                    f"Got: {description!r}"
                )
        if violations:
            return {
                "error": "capability_grid content doesn't fit the grid format. NO slide was added.",
                "violations": violations,
                "action": (
                    "Rewrite per the violations and retry. The grid format constrains "
                    "each card to ≤4-word heading + 5-15 word description. If your content "
                    "needs more detail per item, use add_solution_detail_slide (3 items "
                    "with rich descriptions) or split across multiple capability_grid slides."
                ),
            }
        # Iter 56: remap N<9 cards across columns so they sit one-per-column
        # (top rows of each column) instead of stacking all in column 1.
        # The composition's `cards` matchers are ordered column-by-column
        # (slots 0/1/2 in col 1, 3/4/5 in col 2, 6/7/8 in col 3). Without
        # this remap, n=3 fills slots [0,1,2] = entire col 1, leaving cols
        # 2/3 visibly empty (llm-legacy iter 55 slide 5).
        if len(cards) == 3:
            # one card per column at the top row → slots [0, 3, 6]
            remapped = [None] * 9
            remapped[0] = cards[0]
            remapped[3] = cards[1]
            remapped[6] = cards[2]
            cards = remapped
        elif len(cards) == 6:
            # two cards per column → slots [0,1,3,4,6,7]
            remapped = [None] * 9
            remapped[0] = cards[0]
            remapped[1] = cards[1]
            remapped[3] = cards[2]
            remapped[4] = cards[3]
            remapped[6] = cards[4]
            remapped[7] = cards[5]
            cards = remapped
        return _build_composition(
            "capability_grid",
            {
                "subhead": subhead,
                "cross_label": cross_label,
                "cross_intro": cross_intro,
                "col_headings": col_headings,
                "cards": cards,
            },
            presentation_id,
        )

    @app.tool(
        annotations=ToolAnnotations(title="Add Process Steps Slide"),
    )
    def add_process_steps_slide(
        subhead: str,
        title: str,
        intro: str,
        steps: List[str],
        presentation_id: Optional[str] = None,
    ) -> Dict:
        """Add a 4-step horizontal process flow with arrows between steps.

        Use this for any SEQUENCE of 4 stages: onboarding, project
        phases, customer journey, workflow steps, rollout phases.
        Each step is a name + a short description.

        Args:
            subhead: Section label at the top, ≤ 4 words. ("Customer Onboarding", "Project Phases".)
            title: Main slide claim, ≤ 8 words.
            intro: One-line intro paragraph framing the steps, 8-18 words.
            steps: EXACTLY 4 items. Each is "Name\\nShort description".
                Name ≤ 3 words. Description 5-12 words (the step box is
                small — longer descriptions get cut off at the bottom).
        """
        violations: List[str] = []
        if len(subhead.split()) > 4:
            violations.append(f"subhead is {len(subhead.split())} words; max 4. Got: {subhead!r}")
        if len(title.split()) > 8:
            violations.append(f"title is {len(title.split())} words; max 8. Got: {title!r}")
        intro_text = _normalise_text(intro) if isinstance(intro, str) else str(intro)
        ic = len(intro_text.split())
        if ic < 8 or ic > 18:
            violations.append(f"intro is {ic} words; needs 8-18. Got: {intro!r}")
        if len(steps) != 4:
            violations.append(f"steps must have exactly 4 items; got {len(steps)}.")
        for i, step in enumerate(steps[:4]):
            normalised = _normalise_text(step) if isinstance(step, str) else str(step)
            lines = [l.strip() for l in normalised.split("\n") if l.strip()]
            if len(lines) < 2:
                violations.append(
                    f"steps[{i}] needs 'Name\\nDescription'. Got just: {step!r}"
                )
                continue
            name = lines[0]
            description = " ".join(lines[1:])
            if len(name.split()) > 3:
                violations.append(f"steps[{i}] name is {len(name.split())} words; max 3. Got: {name!r}")
            dw = len(description.split())
            if dw < 5 or dw > 12:
                violations.append(
                    f"steps[{i}] description is {dw} words; needs 5-12. "
                    f"Each step box is small and longer text overflows the rectangle "
                    f"(the 'validated' / 'reforecasted' get cut off). Got: {description!r}"
                )
        if violations:
            return {
                "error": "process_steps content doesn't fit the step format. NO slide was added.",
                "violations": violations,
                "action": (
                    "Rewrite per the violations and retry. process_steps is designed "
                    "for 4 sequential stages with tight names and short descriptions. "
                    "If your content isn't sequential, use add_value_cards_slide "
                    "(parallel cards) instead."
                ),
            }
        # Combine title + intro into a single newline-separated string for
        # the title_and_intro shape (which holds both as paragraphs).
        return _build_composition(
            "process_steps",
            {
                "subhead": subhead,
                "title_and_intro": f"{title}\n{intro}",
                "steps": steps,
            },
            presentation_id,
        )

    @app.tool(
        annotations=ToolAnnotations(title="Add Value Cards Slide"),
    )
    def add_value_cards_slide(
        subhead: str,
        title: str,
        cards: List[str],
        presentation_id: Optional[str] = None,
    ) -> Dict:
        """Add a slide with 4 icon-topped value cards (heading + description).

        High-impact composition for showcasing 4 parallel benefits with
        real descriptions. Each card has an icon (decorative — kept from
        the template), a heading (the value prop), and a 12-25 word
        description (the why).

        Use for:
        - "Why work with us" type slides
        - 4 differentiators / value props (richer than value_props_4)
        - 4 strategic pillars
        - 4 customer benefits

        DON'T use for: capability lists (use capability_grid), sequences
        (use process_steps), short pillar statements (use value_props_4).

        Args:
            subhead: Section label, ≤ 5 words. ("Why partner with us", "Our value props".)
            title: Main claim, ≤ 10 words. The headline message.
            cards: EXACTLY 4 items. Each is "Heading\\nDescription".
                Heading ≤ 5 words. Description 12-25 words.
        """
        violations: List[str] = []
        if len(subhead.split()) > 5:
            violations.append(f"subhead is {len(subhead.split())} words; max 5. Got: {subhead!r}")
        if len(title.split()) > 10:
            violations.append(f"title is {len(title.split())} words; max 10. Got: {title!r}")
        if len(cards) != 4:
            violations.append(f"cards must have exactly 4 items; got {len(cards)}.")
        for i, card in enumerate(cards[:4]):
            normalised = _normalise_text(card) if isinstance(card, str) else str(card)
            lines = [l.strip() for l in normalised.split("\n") if l.strip()]
            if len(lines) < 2:
                violations.append(f"cards[{i}] needs 'Heading\\nDescription'. Got just: {card!r}")
                continue
            heading = lines[0]
            description = " ".join(lines[1:])
            if len(heading.split()) > 5:
                violations.append(f"cards[{i}] heading is {len(heading.split())} words; max 5. Got: {heading!r}")
            dw = len(description.split())
            if dw < 12 or dw > 25:
                violations.append(
                    f"cards[{i}] description is {dw} words; needs 12-25. Got: {description!r}"
                )
        if violations:
            return {
                "error": "value_cards content doesn't fit the card format. NO slide was added.",
                "violations": violations,
                "action": (
                    "Rewrite per the violations and retry. value_cards needs 4 cards "
                    "with rich descriptions. If you have only short pillar statements, "
                    "use add_value_props_slide instead. If your content needs more "
                    "depth per card, use add_solution_detail_slide (3 cards, deeper)."
                ),
            }
        return _build_composition(
            "value_cards",
            {"subhead": subhead, "title": title, "cards": cards},
            presentation_id,
        )

    # ------------------------------------------------------------------
    # ATOMIC BUILD: iter 49 rewrite
    # ------------------------------------------------------------------
    # `build_deck` takes the WHOLE deck spec as JSON and builds it in one
    # shot. No partial state, no delete_slide loops, no validate-then-fix
    # cycle. Either the spec is valid (deck builds and saves cleanly) or
    # it's not (model gets back per-slide field errors and resubmits).
    #
    # Iter 44-48 confirmed the incremental build-then-fix architecture
    # forces the agent through too many decision points where it panics:
    # validate_deck fails → delete_slide loops corrupt order → rebuild
    # cycles burn step budget → fail to save. Iter 48 final: 1 of 5
    # briefs saved with rebuild-on-fail skill rule. Need a different shape.
    #
    # build_deck shape:
    #   build_deck(
    #       cover={"title": "...", "subtitle": "..."},
    #       slides=[
    #         {"type": "stat_cards", "subhead": "...", "intro": "...", "stats": [...]},
    #         {"type": "value_cards", "subhead": "...", "title": "...", "cards": [...]},
    #         ...one entry per body slide...
    #       ],
    #       closing={"title": "...", "cta_lines": [...]},
    #   )
    # → {"download_url": "...", "slide_count": N, "compositions_used": [...]}
    # OR
    # → {"errors": [{"slide_index": 3, "type": "stat_cards", "violations": [...]}], ...}

    @app.tool(
        annotations=ToolAnnotations(
            title="Build Deck (Atomic)",
            destructiveHint=False,
        ),
    )
    def build_deck(
        cover: Dict[str, Any],
        slides: List[Dict[str, Any]],
        closing: Dict[str, Any],
        headline_message: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Build the whole deck in one atomic call from a structured spec.

        Replaces the incremental add_*_slide flow. Pass the entire deck
        as JSON: cover (title + subtitle), an ordered list of body
        slides (each with a `type` field naming the composition), and
        closing (title + cta_lines). The server creates a fresh
        presentation from the brand template, builds every slide, and
        saves to S3 — returning a single download URL on success or a
        list of per-slide validation errors on failure.

        ALWAYS use this. Do NOT mix with create_presentation /
        add_*_slide / delete_slide / save_presentation_to_url — those
        are legacy and being phased out.

        Body slide types and their fields (omit any field not used):
        - "stat_cards":     subhead, intro, stats[3-4] (each "VALUE\\nLABEL\\nATTRIBUTION")
        - "value_cards":    subhead, title, cards[4] (each "Heading\\nDescription", 12-25 word descriptions)
        - "value_props_4":  title, lead_claim, subtitle, pillars[4] (each ≤ 10 words)
        - "capability_grid": subhead, cross_label, cross_intro, col_headings[3], cards[3|6|9]
        - "process_steps":  subhead, title, intro, steps[4] (each "Name\\nDescription")
        - "split_benefits": title (3-8 words), benefits[4|6] (each 5-12 words)
        - "timeline":       subhead, events[4-12] (each "Date description", 3-10 words)
        - "diagram":        title, mermaid, caption? (server renders + embeds)
        - "chart":          title, chart_type ("bar"/"line"/"pie"/"doughnut"), labels[], datasets[], caption?
        - "bullets":        title, subhead (1-5 words), bullets[3-6] (each ≤ 15 words)
        - "solution_detail": category, title, benefits (15-60 words), details[3] (each "Heading\\nDescription", 12-30 word descriptions)
        - "section_divider": kicker, title

        cover format: {"title": "<assertion 7-12 words ≤ 60 chars>", "subtitle": "<topic line>"}
        closing format: {"title": "<wrap-up ≤ 6 words ≤ 40 chars>", "cta_lines": ["...", "..."] (2-6 items, each ≤ 12 words ≤ 70 chars)}

        headline_message (RECOMMENDED): pass the brief's ONE-line
        headline message — the single thing the audience should walk
        away believing. When supplied, build_deck runs a title-ladder
        check: the cover title+subtitle MUST reference at least 40% of
        the headline's key terms (substantive words, length ≥ 4,
        stopwords excluded). This catches the "cover doesn't anchor
        the headline" failure mode upfront. If you omit it the check
        is skipped, but the assess-side probe will fire post-hoc.

        Spec-level rules enforced upfront (rejection blocks build):
        - Deck length 5-15 slides (cover + body + closing).
        - Visual richness ≥ 50% — at least half of body slides use
          value_cards, capability_grid, stat_cards, split_benefits,
          timeline, diagram, chart, or process_steps. Bullets and
          solution_detail are the workhorses but mass overuse reads
          as a wall of text.
        - has-chart-or-diagram on 8+ slide decks.
        - No 3+ text-only body slides in a row (bullets, solution_detail,
          section_divider, value_props_4 count as text-only).
        - title-ladder if headline_message is supplied (cover must
          reference ≥ 40% of headline key terms).

        On failure, the response includes a structured `errors` array. Each
        entry says exactly which slide (by index in `slides`), which fields
        violated their format, and how to fix. Rewrite those fields in the
        spec and call `build_deck` again — no state to clean up.
        """
        import io
        import os
        import uuid
        from datetime import datetime, timezone
        import utils as ppt_utils

        # Dispatch table: composition type → builder function from this module
        builders = {
            "cover": lambda s: add_cover_slide(s["title"], s["subtitle"]),
            "section_divider": lambda s: add_section_divider(s["kicker"], s["title"]),
            "bullets": lambda s: add_bullets_slide(s["title"], s["subhead"], s["bullets"]),
            "solution_detail": lambda s: add_solution_detail_slide(
                s["category"], s["title"], s["benefits"], s["details"],
            ),
            "stat_cards": lambda s: add_stat_cards_slide(
                s["subhead"], s["intro"], s["stats"],
            ),
            "value_props_4": lambda s: add_value_props_slide(
                s["title"], s["lead_claim"], s["subtitle"], s["pillars"],
            ),
            "value_cards": lambda s: add_value_cards_slide(
                s["subhead"], s["title"], s["cards"],
            ),
            "capability_grid": lambda s: add_capability_grid_slide(
                s["subhead"], s["cross_label"], s["cross_intro"],
                s["col_headings"], s["cards"],
            ),
            "process_steps": lambda s: add_process_steps_slide(
                s["subhead"], s["title"], s["intro"], s["steps"],
            ),
            "split_benefits": lambda s: add_split_benefits_slide(
                s["title"], s["benefits"],
            ),
            "timeline": lambda s: add_timeline_slide(s["subhead"], s["events"]),
            "diagram": lambda s: add_diagram_slide(
                s["title"], s["mermaid"], s.get("caption"),
            ),
            "chart": lambda s: add_chart_slide(
                s["title"], s["chart_type"], s["labels"],
                s["datasets"], s.get("caption"),
                y_axis_title=s.get("y_axis_title"),
                x_axis_title=s.get("x_axis_title"),
            ),
            "closing": lambda s: add_closing_slide(s["title"], s["cta_lines"]),
        }

        # iter 50: structural spec checks BEFORE building anything. These
        # catch the most common iter 49 sweep failures upfront so the model
        # gets a fast, structured error rather than discovering the problem
        # after a successful build via post-hoc assess.ts probes:
        #   - title-ladder (cover must reference ≥ 40% of headline terms)
        #   - visual-richness (≥ 50% of body slides must use a visual-rich composition)
        #   - has-chart-or-diagram (8+ slide decks need at least one)
        #   - no 3+ text-only slides in a row
        #   - deck length (3-15 slides total — cover + closing + body)
        VISUAL_RICH_TYPES = {
            "value_cards", "capability_grid", "stat_cards",
            "split_benefits", "timeline", "diagram", "chart", "process_steps",
        }
        TEXT_ONLY_TYPES = {
            "bullets", "solution_detail", "section_divider", "value_props_4",
        }
        # Iter 58: which composition types actually carry a <p:pic>
        # element after building. solution_detail strips its source
        # pictures (iter 31). process_steps / timeline / closing are
        # layout-built with text shapes only. bullets is layout-built.
        HAS_PICTURE_TYPES = {
            "cover", "stat_cards", "split_benefits",
            "value_cards", "chart", "diagram", "capability_grid",
        }
        spec_errors: List[Dict[str, Any]] = []

        # Length cap.
        total_slides = len(slides) + 2  # cover + body + closing
        if total_slides < 5:
            spec_errors.append({
                "rule": "deck-length",
                "detail": f"Deck has {total_slides} slides (cover + {len(slides)} body + closing). Need ≥ 5 total — add more body slides.",
            })
        elif total_slides > 15:
            spec_errors.append({
                "rule": "deck-length",
                "detail": f"Deck has {total_slides} slides. Cap is 15 — drop {total_slides - 15} body slides.",
            })

        # Visual richness over body slides.
        body_types = [s.get("type") for s in slides if s.get("type")]
        if body_types:
            visual_count = sum(1 for t in body_types if t in VISUAL_RICH_TYPES)
            pct = visual_count / len(body_types)
            if pct < 0.5:
                text_only_idx = [
                    i for i, t in enumerate(body_types) if t in TEXT_ONLY_TYPES
                ]
                spec_errors.append({
                    "rule": "visual-richness-min-50-pct",
                    "detail": (
                        f"Only {int(pct*100)}% of body slides use a visual-rich "
                        f"composition ({visual_count}/{len(body_types)}). "
                        f"Body slide indices {text_only_idx} are text-only "
                        f"({set(body_types[i] for i in text_only_idx)}). Convert "
                        f"at least one to value_cards / capability_grid / "
                        f"stat_cards / split_benefits / timeline / diagram / "
                        f"chart based on the content shape."
                    ),
                })

        # Iter 58 (REVERTED iter 61): image-density spec validator
        # caused the sandbox agent to panic ('template missing')
        # whenever the threshold wasn't met on the first 1-2 calls.
        # Probe still runs post-build in assess.ts; spec validation
        # left only the older softer rules. Will revisit when
        # sandbox-agent retry behaviour is more robust.

        # has-chart-or-diagram for 8+ slide decks.
        if total_slides >= 8:
            has_cd = any(t in ("chart", "diagram") for t in body_types)
            if not has_cd:
                spec_errors.append({
                    "rule": "has-chart-or-diagram",
                    "detail": (
                        f"Deck has {total_slides} slides but no chart or diagram. "
                        f"For numeric content (growth / breakdown / comparison) "
                        f"use a chart; for architecture / flow / sequence use a "
                        f"diagram. Pick one body slide and convert."
                    ),
                })

        # No 3+ consecutive text-only slides.
        if len(body_types) >= 3:
            longest_text_run = 0
            current_run = 0
            worst_start = -1
            run_start = -1
            for i, t in enumerate(body_types):
                if t in TEXT_ONLY_TYPES:
                    if current_run == 0:
                        run_start = i
                    current_run += 1
                    if current_run > longest_text_run:
                        longest_text_run = current_run
                        worst_start = run_start
                else:
                    current_run = 0
            if longest_text_run > 2:
                spec_errors.append({
                    "rule": "no-text-only-streak-over-2",
                    "detail": (
                        f"Text-only compositions run {longest_text_run} in a row "
                        f"at body slide indices {list(range(worst_start, worst_start + longest_text_run))}. "
                        f"Three in a row is a wall of text. Swap the middle one "
                        f"for a visual-rich composition."
                    ),
                })

        # Title-ladder: cover must reference ≥ 40% of headline_message key terms.
        if headline_message:
            import re as _re
            stopwords = {
                "a", "an", "the", "and", "or", "but", "is", "are", "was", "were",
                "be", "been", "being", "to", "of", "in", "on", "at", "for", "with",
                "by", "from", "as", "that", "this", "these", "those", "it", "its",
                "we", "our", "us", "you", "your", "their", "they", "them",
                "have", "has", "had", "will", "would", "should", "can", "could",
                "may", "might", "must", "do", "does", "did", "not", "no",
                "than", "then", "now", "so", "if", "when", "where", "who",
                "what", "how", "why", "one", "two", "three", "four", "five",
                "more", "most", "less", "each", "into", "about",
            }
            def _key_terms(text):
                return {
                    w for w in _re.sub(r"[^a-z0-9\s%$£€]", " ", (text or "").lower()).split()
                    if len(w) >= 4 and w not in stopwords
                }
            headline_terms = _key_terms(headline_message)
            if len(headline_terms) >= 3:
                cover_text = " ".join([
                    cover.get("title", ""),
                    cover.get("subtitle", ""),
                ])
                cover_terms = _key_terms(cover_text)
                hits = headline_terms & cover_terms
                cov = len(hits) / len(headline_terms)
                if cov < 0.4:
                    missing = sorted(headline_terms - cover_terms)[:6]
                    spec_errors.append({
                        "rule": "title-ladder-cover",
                        "detail": (
                            f"Cover title+subtitle reference {int(cov*100)}% "
                            f"of headline key terms (need ≥ 40%). Headline: "
                            f"'{headline_message}'. Add at least 2-3 of these "
                            f"verbatim into the cover title or subtitle: {missing}."
                        ),
                    })

                # Iter 59 (REVERTED iter 60c): body-coverage validator
                # disabled because forcing the agent to hit 60% headline-
                # term coverage across body slides caused it to abandon
                # build_deck after one failure. Probe still runs in
                # assess.ts post-build; will revisit when agent retry
                # behaviour is more robust.
                pass

        if spec_errors:
            # Iter 60: return only ONE spec_error per call. The sandbox
            # agent panics ("template missing") when shown multiple spec
            # failures at once; one-at-a-time turns each retry into a
            # focused fix rather than a triage exercise. Order: structural
            # (deck-length, has-chart-or-diagram), then density rules,
            # then title-ladder. The remaining rules surface on the next
            # call after the agent fixes this one.
            order = [
                "deck-length",
                "has-chart-or-diagram",
                "no-text-only-streak-over-2",
                "visual-richness-min-50-pct",
                "image-density-min-60-pct",
                "title-ladder-cover",
                "title-ladder-body",
            ]
            ranked = sorted(
                spec_errors,
                key=lambda e: order.index(e["rule"]) if e["rule"] in order else 99,
            )
            top = ranked[0]
            return {
                "ok": False,
                "spec_errors": [top],
                "action": (
                    f"Fix the '{top['rule']}' violation above and retry "
                    f"build_deck with the revised spec. This is a CONTENT "
                    f"problem in your spec — NOT a template availability "
                    f"problem, NOT a server problem. The brand template "
                    f"is loaded; build_deck is the only correct tool. "
                    f"DO NOT call create_presentation, add_*_slide, "
                    f"list_template_files, or any other tool. Up to 6 "
                    f"retries are expected; each retry should land "
                    f"meaningful spec changes."
                ),
            }

        # Step 1: create a fresh presentation from the brand template.
        default_template = os.environ.get("PPTX_DEFAULT_TEMPLATE", "").strip()
        if not default_template:
            return {
                "error": "PPTX_DEFAULT_TEMPLATE env var is not set on the MCP server.",
            }
        # Resolve the template path the same way create_presentation does.
        if os.path.exists(default_template):
            resolved_template_path = default_template
        else:
            template_name = os.path.basename(default_template)
            resolved_template_path = None
            from tools.presentation_tools import (
                _strip_template_slides,
            )
            # Walk through known search directories — same set the
            # create_presentation tool uses.
            search_paths = [
                "/app/templates",
                ".",
                "./templates",
                "./assets",
                "./resources",
            ]
            for d in search_paths:
                candidate = os.path.join(d, template_name)
                if os.path.exists(candidate):
                    resolved_template_path = candidate
                    break
            if resolved_template_path is None:
                return {
                    "error": f"Template not found: {default_template}",
                }
        # Import _strip_template_slides if not imported above.
        from tools.presentation_tools import _strip_template_slides, _purge_orphan_slide_parts, _get_s3_client

        try:
            pres = ppt_utils.create_presentation_from_template(resolved_template_path)
            _strip_template_slides(pres)
        except Exception as e:
            return {"error": f"Failed to load template: {e}"}

        # Step 2: register the new presentation in the shared state so
        # the inner add_*_slide functions (which depend on the
        # presentations dict + library_template_paths) can find it.
        # iter 47 state-reset: flush prior decks first.
        pres_id = f"build_deck_{uuid.uuid4().hex[:8]}"
        presentations.clear()
        library_template_paths.clear()
        presentations[pres_id] = pres
        library_template_paths[pres_id] = resolved_template_path
        # The inner add_*_slide functions resolve presentation_id from
        # get_current_presentation_id(); set it to our new id so they
        # operate on this deck. set_current_presentation_id is the
        # closure passed into register_composition_tools — required
        # for build_deck to function correctly.
        if set_current_presentation_id is None:
            return {
                "error": (
                    "Server misconfiguration: set_current_presentation_id "
                    "not passed to register_composition_tools. build_deck "
                    "cannot reset the current_presentation_id and would "
                    "leave inner add_*_slide calls stranded."
                ),
            }
        set_current_presentation_id(pres_id)

        # Step 3: build all slides in spec order. Cover → body → closing.
        # Collect errors per slide rather than short-circuiting; that way
        # the caller gets a complete picture of what to fix in one shot.
        errors: List[Dict[str, Any]] = []
        compositions_used: List[str] = []

        def call_builder(slide_index: int, slide_type: str, slide_spec: Dict[str, Any]):
            builder = builders.get(slide_type)
            if builder is None:
                errors.append({
                    "slide_index": slide_index,
                    "type": slide_type,
                    "error": f"Unknown composition type: {slide_type!r}",
                    "valid_types": sorted(builders.keys()),
                })
                return
            try:
                result = builder(slide_spec)
            except KeyError as e:
                errors.append({
                    "slide_index": slide_index,
                    "type": slide_type,
                    "error": f"Missing required field: {e}",
                    "spec": slide_spec,
                })
                return
            except Exception as e:
                errors.append({
                    "slide_index": slide_index,
                    "type": slide_type,
                    "error": f"Builder exception: {type(e).__name__}: {e}",
                    "spec": slide_spec,
                })
                return
            if isinstance(result, dict) and result.get("error"):
                errors.append({
                    "slide_index": slide_index,
                    "type": slide_type,
                    "error": result["error"],
                    "violations": result.get("violations", []),
                    "action": result.get("action"),
                    "spec": slide_spec,
                })
                return
            compositions_used.append(slide_type)

        # Cover (always slide 0)
        call_builder(0, "cover", cover)
        # Body slides
        for i, slide_spec in enumerate(slides):
            slide_type = slide_spec.get("type")
            if not slide_type:
                errors.append({
                    "slide_index": i + 1,
                    "error": "Slide spec missing required 'type' field",
                    "spec": slide_spec,
                })
                continue
            call_builder(i + 1, slide_type, slide_spec)
        # Closing (always last)
        call_builder(len(slides) + 1, "closing", closing)

        # Step 4: if any errors, bail without saving. The presentation
        # is discarded (will be flushed by the next build_deck call).
        if errors:
            return {
                "ok": False,
                "errors": errors,
                "compositions_built_so_far": compositions_used,
                "action": (
                    "Fix the listed field violations in the spec and call build_deck "
                    "again with the corrected JSON. The deck is rebuilt from scratch "
                    "each call — no partial state to clean up."
                ),
            }

        # Step 5: save to S3.
        bucket = os.environ.get("PPTX_OUTPUT_BUCKET")
        if not bucket:
            return {"error": "PPTX_OUTPUT_BUCKET env var is not set on the MCP server."}
        region = os.environ.get("AWS_REGION") or os.environ.get(
            "AWS_DEFAULT_REGION"
        ) or "us-east-1"
        try:
            buffer = io.BytesIO()
            presentations[pres_id].save(buffer)
            cleaned = _purge_orphan_slide_parts(buffer.getvalue())
            timestamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
            key = f"pptx/{timestamp}-{uuid.uuid4()}.pptx"
            client = _get_s3_client()
            client.put_object(
                Bucket=bucket,
                Key=key,
                Body=cleaned,
                ContentType=(
                    "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                ),
            )
            url = f"https://s3.{region}.amazonaws.com/{bucket}/{key}"
        except Exception as e:
            return {"error": f"Failed to save deck: {type(e).__name__}: {e}"}

        return {
            "ok": True,
            "download_url": url,
            "s3_key": key,
            "slide_count": len(slides) + 2,  # cover + body + closing
            "compositions_used": compositions_used,
        }
