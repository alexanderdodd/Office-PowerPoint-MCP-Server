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
from typing import Any, Dict, List, Optional
from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations
from pptx import Presentation


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
    "closing": {
        "source_slide_index": 36,  # slide 37 (0-indexed) — "Thank You" + CTA list
        "description": "Final 'Thank you' slide with title + 2–6 short CTA / contact lines.",
        "use_when": "ALWAYS the last slide of every deck. Carries the closing message (recommendation, ask, or thank-you) plus contact / next-step lines. This is the slide audiences look at while you wrap — make it count.",
        "fields": {
            "title": {"match": "Thank You", "required": True},
            "cta_lines": {"match": "Book a Demo", "required": True},
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


def _normalise_text(s: str) -> str:
    """Normalise common LLM escape artefacts in incoming text.

    - Two-char literal `\\n` / `\\t` → real newline / tab (models sometimes
      pass these escape sequences when they think they're producing
      JSON-escaped strings).
    - Multiple consecutive newlines collapse to a single newline (extra
      blank lines in stat cards / multi-tier shapes render as visible
      blank paragraphs and break the layout).
    - Leading/trailing whitespace stripped.
    """
    if not isinstance(s, str):
        return s
    s = s.replace("\\n", "\n").replace("\\t", "\t")
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
        # Don't delete extra paragraphs — that would collapse the visual
        # structure of multi-tier shapes. Clear their text instead, so
        # leftover template lorem-ipsum doesn't show but the
        # paragraph-level styling and spacing remain intact.
        for extra_p in original_paras[len(lines):]:
            _set_para_text_preserving_style(extra_p._p, "")


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
    rewiring its `r:embed` so the new picture points at the same media
    part via a fresh relationship on the destination slide.

    Previously `_clone_slide_into` skipped pictures because copying the
    XML alone leaves a dangling embed reference — the destination slide's
    rels don't contain the source's rId, so the picture renders as an
    empty placeholder. This helper preserves the visual (stock
    backgrounds, hero photos, decorative imagery) by adding the rel.
    """
    from copy import deepcopy
    R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
    EMBED_ATTR = f"{{{R_NS}}}embed"
    LINK_ATTR = f"{{{R_NS}}}link"

    cloned = deepcopy(pic_element)
    src_part = src_slide.part
    dst_part = dst_slide.part

    for blip in cloned.iter(f"{{{A_NS}}}blip"):
        for attr_name in (EMBED_ATTR, LINK_ATTR):
            old_rid = blip.get(attr_name)
            if old_rid is None:
                continue
            try:
                src_rel = src_part.rels[old_rid]
            except KeyError:
                continue
            new_rid = dst_part.relate_to(src_rel.target_part, src_rel.reltype)
            blip.set(attr_name, new_rid)
    return cloned


def _clone_slide_into(working_pres, library_pres, source_index: int):
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
            new_tree.append(_clone_pic_into(child, src_slide, new_slide))
            continue
        cloned = deepcopy(child)
        _replace_spautofit_with_normautofit(cloned)
        new_tree.append(cloned)

    return new_slide


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
                if i < len(values):
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
                    # User supplied fewer values than the template has slots —
                    # blank the unused shape(s).
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
            slide.shapes.title.text = str(title_text)
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
            else:
                new_slide = _clone_slide_into(working, library, source_idx)
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
        if word_count > 16:
            violations.append(
                f"title is {word_count} words. Cover titles must be ≤ 16 words — "
                f"longer ones overflow the title shape and the rendered text "
                f"shrinks to be unreadable. Got: {title!r}. Compress the claim "
                f"to its sharpest form (single sentence, one comma at most) and "
                f"move the longer-form rationale to slide 2."
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
        elif len(title_text.split()) > 10:
            violations.append(
                f"title is {len(title_text.split())} words; closing titles must be ≤ 10."
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
                wc = len((line or "").split())
                if wc < 4 or wc > 20:
                    violations.append(
                        f"cta_lines[{i}] is {wc} words; needs 4-20. Got: {line!r}"
                    )
        if violations:
            return {
                "error": "closing slide content doesn't fit the format. NO slide was added.",
                "violations": violations,
                "action": (
                    "Rewrite per the violations. title is the final claim (≤10 words). "
                    "cta_lines are 2-6 short action/contact lines (4-20 words each)."
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
            subhead: Small kicker line above the bullets (e.g. "NEXT STEPS" or "WHAT YOU CAN DO").
            bullets: List of 3-6 short bullets (≤ 15 words each, parallel grammar).
        """
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
        if len(intro.split()) > 12:
            violations.append(
                f"intro is {len(intro.split())} words; max is 12. Got: {intro!r}"
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
            cards: Up to 9 cards. Each card is a string formatted as
                "Heading\\nOne-line description". Heading ≤ 4 words.
                Description 5-15 words (fits the small card slot). Cards
                fill column 1 first (top to bottom), then column 2, then
                column 3. Unused slots are blanked.
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
        if not (1 <= len(cards) <= 9):
            violations.append(f"cards must have 1-9 items; got {len(cards)}.")
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
