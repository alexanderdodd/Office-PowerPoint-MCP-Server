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
    },
    "capability_grid": {
        "source_slide_index": 12,  # slide 13
        "description": "Section subhead + cross-cutting label + intro + 3 column headings + 9 capability cards in a 3x3 grid.",
        "use_when": "When the content is a MAP of many capabilities organised into 3 columns, with multiple sub-items per column. Perfect for 'the platform at a glance' or 'all the layers and what's in each'. Each card is JUST a heading (the brand description shapes on the source slide are auto-blanked).",
        "fields": {
            "subhead": {"match": "Bizzdesign", "required": True},
            "cross_label": {"match": "Transformation Collaboration", "required": True},
            "cross_intro": {"match": "Align teams around", "required": True},
            "col_headings": [
                {"match": "Transformation Planning"},
                {"match": "Transformation Design"},
                {"match": "Transformation Governance"},
            ],
            # 9 cards: 3 per column, ordered by column then by vertical position
            "cards": [
                {"match": "Strategic Portfolio Management"},
                {"match": "Application Portfolio Management"},
                {"match": "Technology Portfolio Management"},
                {"match": "Enterprise Architecture Management"},
                {"match": "Business Architecture Management"},
                {"match": "Solution Architecture Management"},
                {"match": "Business Process Management"},
                {"match": "Data Management"},
                {"match": "Governance, Risk & Compliance"},
            ],
        },
        # Auto-clear: the source slide has 9 description text shapes paired
        # with each card heading. When the user supplies new card content
        # we don't have descriptions for it, so blank the originals rather
        # than leaving Bizzdesign EA copy on the model's deck.
        "clear_text_starting_with": [
            "Align investments with strategic goals",
            "Reduce IT costs",
            "Stay ahead of obsolescence",
            "Strengthen governance",
            "Turn insight into execution",
            "Accelerate design with business context",
            "Increase process visibility",
            "Create a common data language",
            "Improve compliance visibility",
        ],
    },
    "bullets": {
        # Special: built directly from the Basic Text layout, not cloned
        # from a source slide. The Basic Text layout has placeholder idx 13
        # (subhead) and idx 14 (body bullets, brand-styled).
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

    Models sometimes pass `\\n` (literal backslash + n) instead of an
    actual newline, especially when content is shaped to look like a
    JSON-escaped string. Treat those as real line breaks so they don't
    render as visible `\\n` in the slide.
    """
    if not isinstance(s, str):
        return s
    # Two-character literal '\n' → real newline.
    return s.replace("\\n", "\n").replace("\\t", "\t")


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
        # Remove extra paragraphs the new content doesn't need.
        for extra_p in original_paras[len(lines):]:
            extra_p._p.getparent().remove(extra_p._p)


def _find_shape_by_match(shapes, match_prefix: str):
    """Return the first shape whose text starts with the given prefix
    (case-sensitive, stripped). None if not found.
    """
    target = match_prefix.strip()
    for shape in shapes:
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
    # Skip:
    #   - {nv,}GrpSpPr (group props at the top of spTree)
    #   - <p:pic> (pictures — their <a:blip r:embed="rIdN"/> refs point to
    #     relationships on the SOURCE slide that we don't replicate, so on
    #     the clone they resolve to "missing image" empty boxes. Branded
    #     text shapes carry the styling; decorative imagery is a Phase 2
    #     concern that needs proper rel + media copying.)
    src_tree = src_slide.shapes._spTree
    new_tree = new_slide.shapes._spTree
    for child in list(src_tree):
        tag = child.tag
        if tag.endswith("}nvGrpSpPr") or tag.endswith("}grpSpPr"):
            continue
        if tag.endswith("}pic"):
            continue
        new_tree.append(deepcopy(child))

    return new_slide


# ------------------------------------------------------------------
# Field application
# ------------------------------------------------------------------

def _apply_fields(slide, fields_spec: Dict[str, Any], content: Dict[str, Any]) -> List[str]:
    """Apply text replacements to the slide's shapes per the composition's
    fields spec.

    `fields_spec` maps field_name → matcher dict OR list of matcher dicts (repeatable).
    `content` maps field_name → string OR list of strings.

    Returns a list of warnings (missing shapes, etc).
    """
    warnings: List[str] = []
    shapes = list(slide.shapes)

    for field_name, spec in fields_spec.items():
        value = content.get(field_name)
        if isinstance(spec, list):
            # Repeatable group.
            values = value or []
            if not isinstance(values, list):
                warnings.append(f"Field '{field_name}' expected a list, got {type(values).__name__}")
                values = []
            for i, matcher in enumerate(spec):
                shape = _find_shape_by_match(shapes, matcher["match"])
                if shape is None:
                    warnings.append(f"Could not locate shape for {field_name}[{i}] (match prefix: {matcher['match']!r})")
                    continue
                if i < len(values):
                    new_text = values[i]
                    if isinstance(new_text, str):
                        # Detect newlines → multiline
                        if "\n" in new_text:
                            _set_shape_multiline(shape, new_text.split("\n"))
                        else:
                            _set_shape_text(shape, new_text)
                    else:
                        _set_shape_text(shape, str(new_text))
                else:
                    # User supplied fewer values than the template has slots —
                    # blank the unused shape.
                    _clear_shape_text(shape)
        else:
            # Scalar field.
            if value is None:
                if spec.get("required", False):
                    warnings.append(f"Required field '{field_name}' missing from content")
                continue
            shape = _find_shape_by_match(shapes, spec["match"])
            if shape is None:
                warnings.append(f"Could not locate shape for '{field_name}' (match prefix: {spec['match']!r})")
                continue
            if isinstance(value, list):
                _set_shape_multiline(shape, [str(v) for v in value])
            else:
                if "\n" in str(value):
                    _set_shape_multiline(shape, str(value).split("\n"))
                else:
                    _set_shape_text(shape, str(value))

    return warnings


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
    return warnings


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
        except Exception as e:
            return {"error": f"Failed to build {composition_name}: {e}"}

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

        Args:
            title: The deck's headline (claim form, e.g. "From feature to platform").
            subtitle: Subtitle text (audience, year, or short scope statement).
        """
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

        Use this for any substantial topic with 3 detail items. The slide
        has:
        - A small kicker (`category`) at the top
        - A large product/area title beneath
        - A multi-line benefits paragraph (the "why this matters")
        - 3 detail items in a row at the bottom (each: short heading + description)

        For a roadmap or platform deck, use this for EACH phase or EACH
        pillar — it's the brand's standard "topic detail" template.

        Args:
            category: Top-of-slide kicker (e.g. "Transformation Planning", "Now — Q2 2026").
            title: Main heading (the topic name, ~3-5 words).
            benefits: Multi-line benefits paragraph. Use `\\n` between lines.
            details: List of 3 detail items, each formatted as
                "Heading\\nDescription text spanning one or two lines". The first
                line becomes the bold heading; subsequent lines are the body.
        """
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
        # Validate field lengths and surface warnings the model can act on.
        validation_warnings: List[str] = []
        if len(subhead.split()) > 5:
            validation_warnings.append(
                f"subhead is {len(subhead.split())} words; max is 5. Shorter is better."
            )
        if len(intro.split()) > 12:
            validation_warnings.append(
                f"intro is {len(intro.split())} words; max is 12."
            )
        for i, stat in enumerate(stats[:4]):
            normalised = _normalise_text(stat) if isinstance(stat, str) else str(stat)
            lines = normalised.split("\n")
            if len(lines) < 2:
                validation_warnings.append(
                    f"stats[{i}] has only {len(lines)} line(s); needs at least VALUE + LABEL "
                    f"separated by newline. Got: {stat!r}"
                )
                continue
            value, label = lines[0].strip(), lines[1].strip()
            attribution = lines[2].strip() if len(lines) > 2 else ""
            if len(value) > 6:
                validation_warnings.append(
                    f"stats[{i}] VALUE is {len(value)} chars (max 6). Got: {value!r}"
                )
            if len(label) > 18:
                validation_warnings.append(
                    f"stats[{i}] LABEL is {len(label)} chars (max 18). Got: {label!r}"
                )
            if attribution and len(attribution) > 40:
                validation_warnings.append(
                    f"stats[{i}] ATTRIBUTION is {len(attribution)} chars (max 40). Got: {attribution!r}"
                )
        result = _build_composition(
            "stat_cards",
            {"subhead": subhead, "intro": intro, "stats": stats},
            presentation_id,
        )
        if validation_warnings:
            existing = result.get("warnings", [])
            result["warnings"] = existing + validation_warnings
            # Add a clear actionable note so the model knows to rewrite.
            result["action"] = (
                "Some fields exceed the stat-card character limits and will overflow visually. "
                "Rewrite the offending stats to fit the limits and call add_stat_cards_slide again."
            )
        return result

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

        Use this for "the whole platform at a glance" slides — when the
        content is a MAP of many capabilities organised into 3 categories
        with multiple sub-items per category. 9 cards total, distributed
        as 3 cards per column in the order you supply them.

        Args:
            subhead: Top-of-slide subhead (e.g. "Platform map", "All Capabilities").
            cross_label: Cross-cutting label that spans the top center
                (e.g. "Powered by Unify", "Across every surface").
            cross_intro: Intro paragraph framing the grid.
            col_headings: List of 3 column headings (e.g. ["Core Platform", "App-Wide Assistant", "Eval & Observability"]).
            cards: List of up to 9 cards. Each card is a string formatted
                as "HEADING\\nDescription text". Cards fill column 1 first
                (top to bottom), then column 2, then column 3. Unused
                slots are blanked.
        """
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
