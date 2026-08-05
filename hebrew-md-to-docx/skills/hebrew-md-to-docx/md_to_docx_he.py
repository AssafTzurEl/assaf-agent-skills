#!/usr/bin/env python3
r"""
Hebrew Markdown -> Word (.docx) converter, RTL-aware.

Usage:
    python md_to_docx_he.py INPUT.md [-o OUTPUT.docx] [--overrides overrides.json]
                            [--code-font Consolas]

Design principles
-----------------
* Use Word's default template for everything (Normal, Heading 1-N, List Number,
  List Bullet, Quote, Table Grid). The ONLY added style is a "Code" character
  style: identical to Normal but with a monospace font.
* Direction is set per paragraph with w:bidi (RTL paragraphs get NO explicit jc,
  so Word's default right alignment for RTL applies) and per run with w:rtl.
* Runs are split by script and tagged RTL/LTR. Neutral characters (spaces,
  parens, quotes, dashes, punctuation) fall back to the *paragraph's own base
  direction* unless both neighbours are the opposite direction: under an RTL
  base a neutral is LTR only between two LTR characters (so "format string"
  stays together while the parens/quotes/dashes around English stay on the
  Hebrew side), and under an LTR base the mirror applies (so the period ending a
  mostly-English sentence that closes with a Hebrew word stays at its end).
  Digits are treated as LTR so numbers like 1920 never reverse.
* "Word-joiner" symbols (+, #, *, ., \, /, !, ?, %) that touch a Latin letter/digit
  are kept on the Latin side so tech tokens (C++, C#, .NET, *.txt), paths/URLs
  (C:\Users\foo, /home/assaf, https://x/y), embedded English exclamations
  ("Hello, World!") and ellipses ("... . . .") stay intact; a joiner at an LTR/RTL
  boundary is anchored with an invisible LRM (the only inserted control character
  - zero-width, never a glyph). The dot is forward-only so a sentence period is
  never captured; the percent sign is backward-only (plus a directly-adjacent
  Latin letter ahead) so "80%" never becomes "%80" and "%d" holds together.
  Inline-code runs are LRM-anchored the same way when their first/last character
  is not strongly LTR, so `%d` in a Hebrew paragraph does not come out as "d%".
* Markdown links and bare URLs become clickable Word hyperlinks ("Hyperlink"
  character style). Lists share one base direction across all their items (like
  heading levels), and bullets are font-agnostic so they render in LTR and RTL.
* A standalone image (`![alt](src)` on its own line) is embedded centered and
  scaled to fit the page's content width, with the alt text as a centered
  caption (Word "Caption" style) below it. Paths resolve relative to the input
  file; a missing image becomes a visible "[missing image: ...]" placeholder.

Paragraph direction is chosen by the *majority of strong directional words*
(Hebrew vs Latin), after stripping URLs/e-mails. Counting words rather than
characters keeps a Hebrew paragraph RTL when its English is only a few long or
repeated proper nouns (e.g. "Visual Studio" several times), and dropping URLs
keeps a labelled link line ("הורדה: https://...") from being pulled LTR by the
URL. Near-even paragraphs are reported as "ambiguous" and can be resolved with an
overrides file keyed by paragraph index:
    { "118": "rtl", "169": "ltr" }

Headings are a further exception: a single heading's own text is not reliable
(e.g. "הצהרה (declaration)" has more Latin than Hebrew letters, but it is a
Hebrew heading with an English gloss). Instead, each heading *level* (H1, H2,
H3, ...) gets one direction, chosen by the aggregate Hebrew/Latin word counts
across every heading at that level in the document. All headings at that level
then use it, even ones whose own text would individually score the other way.
Per-heading overrides still win via the same --overrides file.
"""

import argparse
import json
import os
import re
import sys

import mistune
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import nsmap, qn
from docx.shared import Pt, RGBColor

# ---------------------------------------------------------------------------
# Direction detection
# ---------------------------------------------------------------------------

# NOTE: the ranges are written as \u escapes on purpose. Spelling them with
# literal characters is fragile: the Hebrew presentation form U+FB1D decomposes
# to U+05D9 U+05B4 under Unicode normalisation, which silently turned the class
# into the huge range U+05B4-U+FB4F - so arrows, dashes, CJK and geometric
# shapes (e.g. the "▸" of "Build ▸ Build Solution") all counted as Hebrew.
HEB_RE = re.compile("[\u0590-\u05FF\uFB1D-\uFB4F]")
LAT_RE = re.compile("[A-Za-z\u00C0-\u024F]")

# Word-level patterns (maximal runs of one script) for base-direction detection.
# Direction is decided by counting *words*, not characters, so a Hebrew sentence
# that embeds long or repeated English proper nouns (e.g. "Visual Studio" several
# times) is not dragged LTR just because those tokens are long.
HEB_WORD_RE = re.compile("[\u0590-\u05FF\uFB1D-\uFB4F]+")
LAT_WORD_RE = re.compile("[A-Za-z\u00C0-\u024F]+")

# URLs and e-mail addresses are data, not prose: their many Latin characters (and
# the "words" a URL splits into) must not sway a paragraph's base direction, so
# they are removed before counting. Without this a labelled link line such as
# "הורדה: https://visualstudio.microsoft.com/vs/community/" would be counted as
# mostly-Latin and wrongly set LTR.
URL_RE = re.compile(
    r"(?:https?://|ftp://|www\.)\S+|\b[\w.+-]+@[\w-]+\.[^\s,;]+", re.I)

AMBIGUOUS_BAND = 0.20  # |heb-lat|/total below this -> flag for human review


def counts(text):
    """(Hebrew words, Latin words) after stripping URLs/e-mails. Counting words
    rather than characters keeps a Hebrew paragraph RTL when its only Latin is a
    handful of (possibly long, possibly repeated) proper nouns; the Hebrew prose
    that carries the grammar normally outnumbers them. A genuine tie that comes
    out wrong is still fixable with --overrides."""
    clean = URL_RE.sub(" ", text)
    return len(HEB_WORD_RE.findall(clean)), len(LAT_WORD_RE.findall(clean))


def detect_direction(text, base="rtl"):
    """Return (direction, ambiguous) for a block of text."""
    h, l = counts(text)
    if h == 0 and l == 0:
        return base, False
    if l == 0:
        return "rtl", False
    if h == 0:
        return "ltr", False
    total = h + l
    direction = "rtl" if h >= l else "ltr"
    ambiguous = abs(h - l) / total < AMBIGUOUS_BAND
    return direction, ambiguous


def _direction_from_counts(h, l, base="rtl"):
    if h == 0 and l == 0:
        return base
    if l == 0:
        return "rtl"
    if h == 0:
        return "ltr"
    return "rtl" if h >= l else "ltr"


def collect_headings(tokens):
    """Recursively collect (level, text) for every heading token. Headings
    normally live at the top level, but this also walks into block_quote
    children in case a heading ever appears nested there."""
    headings = []

    def walk(toks):
        for tok in toks or []:
            t = tok.get("type")
            if t == "heading":
                level = tok["attrs"]["level"]
                text = inline_text(tok["children"])
                headings.append((level, text))
                continue
            children = tok.get("children")
            if children:
                walk(children)

    walk(tokens)
    return headings


def heading_level_directions(tokens):
    """One direction per heading level (H1, H2, ...), chosen by the *aggregate*
    Hebrew/Latin character counts across all headings at that level - not each
    heading's own text. This keeps a heading level consistently RTL (or LTR)
    even when one heading happens to contain more English glyphs than Hebrew
    ones, e.g. a Hebrew term followed by an English gloss in parentheses like
    "הצהרה (declaration)"."""
    totals = {}
    for level, text in collect_headings(tokens):
        h, l = counts(text)
        th, tl = totals.get(level, (0, 0))
        totals[level] = (th + h, tl + l)
    return {level: _direction_from_counts(h, l) for level, (h, l) in totals.items()}


# Symbols that are neutral in the Unicode bidi sense but, in tech writing, are
# glued to a Latin token and must stay with it rather than drift into the RTL
# flow - e.g. the "++"/"#" in C++/C#, the dot in .NET / Node.js, the star in a
# glob like *.txt, the slashes in a path/URL like C:\Users\foo, /home/assaf,
# https://x/y, or the "!"/"?" ending an embedded English phrase ("Hello, World!").
# They are promoted to LTR only when a run of them actually touches a Latin
# letter/digit (see split_runs).
JOIN_LTR = set("+#\\/*.!?%‰")

# The dot is special: it is also the sentence period. A period ending a Hebrew
# sentence ("... Visual Studio." or "יש 123.") must stay in the RTL flow, so a dot
# joins a Latin token only when a Latin letter/digit FOLLOWS it (as in .NET,
# file.txt, 3.14) - never on a trailing dot that a space or line end follows.
# ("!" and "?" are not forward-only: they legitimately end an English phrase, so
# they attach backwards to Latin, but only when directly adjacent - a Hebrew
# "נתקעתם?" has no adjacent Latin and stays RTL.)
JOIN_FORWARD_ONLY = set(".")

# The percent (and per-mille) sign is the mirror case: it always trails its
# number ("80%", "27%"), never leads it, so it joins a Latin token only when a
# Latin letter/digit PRECEDES it. Without this it is a plain neutral, and under
# an RTL base a "%" after a number resolves to the RTL flow and renders on the
# wrong side ("%80" instead of "80%"). A lone "%" with no adjacent digit/letter
# (rare, e.g. a Hebrew sentence about "סימן ה-%") keeps neutral treatment.
JOIN_BACKWARD_ONLY = set("%‰")

# "!" and "?" are the same kind of special case as the dot: they end an embedded
# English phrase ("Hello, World!") but they *also* end the surrounding Hebrew
# sentence ("אין לכם Windows? היחידה..." - the "?" belongs to the Hebrew question,
# not to the word "Windows"). They attach backwards to Latin only when something
# immediately follows them, which is what an embedded phrase looks like - a
# closing quote/bracket or more Latin ('"Hello, World!"', "(Really?)"). When the
# next character is a space, or the text ends there, the mark is terminating the
# Hebrew sentence and is left in the RTL flow.
JOIN_BACKWARD_TIGHT = set("!?")

# Left-to-right mark: an invisible (zero-width) strong-L character used to anchor
# a trailing/leading joiner symbol to the Latin side. Without it, a "++" ending a
# Latin token sits between the token (L) and a following RTL run and the bidi
# algorithm resolves it to the RTL base, flipping "C++" to "++C". The LRM gives
# the joiner a strong-L neighbour so it stays put. It is the only inserted control
# character and never renders as a glyph.
LRM = "‎"


def _char_dir(ch):
    if HEB_RE.match(ch):
        return "rtl"
    if LAT_RE.match(ch) or ch.isdigit():
        return "ltr"   # digits are LTR so numbers like 1920 never reverse
    return "neutral"


def anchor_ltr(text):
    """Bracket an LTR-only run with invisible LRMs when its first/last character
    is not strongly LTR, so the surrounding RTL flow cannot pull that edge
    character across. A run marked LTR still has its *neutral* characters
    resolved by the paragraph's base direction, so inline code like `%d` (the run
    starts with a neutral "%") comes out as "d%" inside a Hebrew paragraph
    without this; likewise a trailing "*" or "-". Only added where needed, and
    the LRM is zero-width so nothing visible changes."""
    if not text:
        return text
    if _char_dir(text[0]) != "ltr":
        text = LRM + text
    if _char_dir(text[-1]) != "ltr":
        text = text + LRM
    return text


def split_runs(text, base="rtl"):
    """Split text into (segment, 'rtl'|'ltr') runs, applying the Unicode bidi
    neutral rule under the paragraph's own base direction: a neutral char
    (space, paren, quote, dash, punctuation) that sits between two opposite-
    direction neighbours - or at the text edge - falls back to the base. Under
    an RTL base a neutral is LTR only when it lies between two LTR characters,
    so English phrases like "format string" stay together while the
    parentheses/quotes/dashes around them stay on the Hebrew side. Under an LTR
    base the rule mirrors: a neutral is RTL only between two RTL characters, so
    the period ending a mostly-English sentence that happens to close with a
    Hebrew word ("This one is 10% עברית.") stays at the sentence's end instead
    of jumping into the Hebrew flow.

    Exception: "word-joiner" symbols in JOIN_LTR (the "++" of C++, the "#" of C#,
    the dot of .NET/Node.js, the star of *.txt, the slashes of a path/URL) are
    promoted to LTR whenever a contiguous run of them reaches a Latin letter/digit,
    so the token stays intact instead of the trailing symbols flipping to the
    Hebrew side at a token boundary. Most joiners qualify on reaching Latin on
    either side; the dot (JOIN_FORWARD_ONLY) qualifies only when Latin FOLLOWS it,
    so a sentence period is never captured; the percent sign
    (JOIN_BACKWARD_ONLY) qualifies only when Latin/digit PRECEDES it, so "80%"
    stays "80%". A joiner that touches no Latin
    character (e.g. "+" alone between Hebrew words, or "/" in a Hebrew "כן/לא")
    keeps the normal neutral treatment. A joiner sitting at the edge of an LTR run
    (facing an RTL run or the run boundary) is additionally anchored with an
    invisible LRM so the bidi algorithm cannot pull it across the boundary.
    """
    if not text:
        return []
    n = len(text)
    base_cls = [_char_dir(ch) for ch in text]

    def _reaches_ltr(i, step):
        j = i + step
        while 0 <= j < n and text[j] in JOIN_LTR:
            j += step
        return 0 <= j < n and base_cls[j] == "ltr"

    cls = list(base_cls)
    for i, ch in enumerate(text):
        if ch in JOIN_LTR and cls[i] == "neutral":
            if ch in JOIN_FORWARD_ONLY:
                reachable = _reaches_ltr(i, 1)
            elif ch in JOIN_BACKWARD_TIGHT:
                reachable = _reaches_ltr(i, -1) and \
                    (i + 1 < n and text[i + 1] != " ")
            elif ch in JOIN_BACKWARD_ONLY:
                # Backward (the "80" of "80%"), or a directly-adjacent Latin
                # letter/digit ahead, which is the printf-specifier shape
                # ("%d", "%s", "%5.2f") written outside backticks.
                reachable = _reaches_ltr(i, -1) or \
                    (i + 1 < n and base_cls[i + 1] == "ltr")
            else:
                reachable = _reaches_ltr(i, 1) or _reaches_ltr(i, -1)
            if reachable:
                cls[i] = "ltr"

    # An ellipsis ("..." or ". . ." with spaces) next to a Latin phrase belongs to
    # it, not the RTL flow. The dots are forward-only above, so a trailing ". . ."
    # is not caught there; promote a run of >=2 dots when the nearest non-space
    # neighbour on either side is Latin/digit (spaces are skipped so
    # "window . . ." counts; a Hebrew "רגע . . ." has no Latin neighbour and stays).
    for mt in re.finditer(r"\.(?:[ ]*\.)+", text):
        a, b = mt.start(), mt.end()
        li = a - 1
        while li >= 0 and text[li] == " ":
            li -= 1
        ri = b
        while ri < n and text[ri] == " ":
            ri += 1
        if (li >= 0 and base_cls[li] == "ltr") or (ri < n and base_cls[ri] == "ltr"):
            for k in range(a, b):
                if text[k] == ".":
                    cls[k] = "ltr"

    atoms = []  # [segment, dir]
    cur, cur_cls = [], None
    for ch, c in zip(text, cls):
        if cur_cls is None or c == cur_cls:
            cur.append(ch)
            cur_cls = c
        else:
            atoms.append(["".join(cur), cur_cls])
            cur, cur_cls = [ch], c
    atoms.append(["".join(cur), cur_cls])

    m = len(atoms)
    for i, atom in enumerate(atoms):
        if atom[1] != "neutral":
            continue
        prev_dir = next((atoms[j][1] for j in range(i - 1, -1, -1)
                         if atoms[j][1] != "neutral"), None)
        next_dir = next((atoms[j][1] for j in range(i + 1, m)
                         if atoms[j][1] != "neutral"), None)
        if base == "ltr":
            atom[1] = "rtl" if (prev_dir == "rtl" and next_dir == "rtl") else "ltr"
        else:
            atom[1] = "ltr" if (prev_dir == "ltr" and next_dir == "ltr") else "rtl"

    merged = []
    for seg, d in atoms:
        if merged and merged[-1][1] == d:
            merged[-1][0] += seg
        else:
            merged.append([seg, d])

    # Anchor edge joiners. A merged LTR run only ever borders RTL runs or the
    # text boundary (same-direction runs were already merged), so a joiner symbol
    # at its start or end faces a pull toward the RTL flow. Bracket it with an LRM
    # so it stays on the Latin side (keeps "C++", "C:\path\", ".../y" intact).
    for run in merged:
        if run[1] != "ltr":
            continue
        if run[0] and run[0][-1] in JOIN_LTR:
            run[0] += LRM
        if run[0] and run[0][0] in JOIN_LTR:
            run[0] = LRM + run[0]
    return [(s, d) for s, d in merged]


# ---------------------------------------------------------------------------
# Low-level OOXML helpers
# ---------------------------------------------------------------------------

_PPR_ORDER = [
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
    "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
    "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
    "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
    "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
    "suppressOverlap", "jc", "textDirection", "textAlignment",
    "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
]


def _insert_pPr_child(pPr, element):
    tag = element.tag.split("}")[-1]
    idx = _PPR_ORDER.index(tag)
    later = set(_PPR_ORDER[idx + 1:])
    for child in pPr:
        if child.tag.split("}")[-1] in later:
            child.addprevious(element)
            return
    pPr.append(element)


def set_paragraph_direction(paragraph, rtl):
    """Set paragraph base direction via w:bidi only.

    No w:jc is written: an RTL paragraph then defaults to right alignment and an
    LTR paragraph to left alignment, matching the template's defaults.
    """
    pPr = paragraph._p.get_or_add_pPr()
    for b in pPr.findall(qn("w:bidi")):
        pPr.remove(b)
    bidi = OxmlElement("w:bidi")
    if not rtl:
        bidi.set(qn("w:val"), "0")
    _insert_pPr_child(pPr, bidi)


_RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "color", "rtl"]


def _ordered_rPr(children):
    rPr = OxmlElement("w:rPr")
    for tag, el in sorted(children, key=lambda c: _RPR_ORDER.index(c[0])):
        rPr.append(el)
    return rPr


def add_run(paragraph, text, *, bold=False, italic=False, rtl=False, style=None,
            container=None):
    """Add a run. A w:rtl element is written only for RTL (Hebrew) runs; LTR
    runs omit it, exactly as Word itself does (Latin/digits resolve LTR under
    the RTL paragraph base on their own). When `container` is given (e.g. a
    w:hyperlink element) the run is appended there instead of to the paragraph."""
    r = OxmlElement("w:r")
    children = []
    if style:
        rStyle = OxmlElement("w:rStyle")
        rStyle.set(qn("w:val"), style)
        children.append(("rStyle", rStyle))
    if bold:
        children.append(("b", OxmlElement("w:b")))
        children.append(("bCs", OxmlElement("w:bCs")))
    if italic:
        children.append(("i", OxmlElement("w:i")))
        children.append(("iCs", OxmlElement("w:iCs")))
    if rtl:
        children.append(("rtl", OxmlElement("w:rtl")))
    r.append(_ordered_rPr(children))
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    (paragraph._p if container is None else container).append(r)
    return r


def add_break(paragraph, container=None):
    r = OxmlElement("w:r")
    r.append(OxmlElement("w:br"))
    (paragraph._p if container is None else container).append(r)


# ---------------------------------------------------------------------------
# Inline rendering
# ---------------------------------------------------------------------------

def inline_text(tokens):
    out = []
    for tok in tokens or []:
        t = tok.get("type")
        if t == "text":
            out.append(tok.get("raw", ""))
        elif t == "codespan":
            out.append(tok.get("raw", ""))
        elif t in ("strong", "emphasis", "link", "strikethrough", "image"):
            out.append(inline_text(tok.get("children")))
        elif t in ("softbreak", "linebreak"):
            out.append(" ")
    return "".join(out)


def render_inline(paragraph, tokens, bold=False, italic=False, container=None,
                  style=None, base="rtl"):
    """`base` is the containing paragraph's base direction; it decides how
    neutral characters (spaces, punctuation, a trailing period) resolve - see
    split_runs."""
    for tok in tokens or []:
        t = tok.get("type")
        if t == "text":
            for seg, d in split_runs(tok.get("raw", ""), base=base):
                add_run(paragraph, seg, bold=bold, italic=italic,
                        rtl=(d == "rtl"), style=style, container=container)
        elif t == "strong":
            render_inline(paragraph, tok.get("children"), bold=True, italic=italic,
                          container=container, style=style, base=base)
        elif t == "emphasis":
            render_inline(paragraph, tok.get("children"), bold=bold, italic=True,
                          container=container, style=style, base=base)
        elif t == "strikethrough":
            render_inline(paragraph, tok.get("children"), bold=bold, italic=italic,
                          container=container, style=style, base=base)
        elif t == "codespan":
            # Monospace LTR run. Marking the run LTR is not enough on its own:
            # a neutral character at the run's edge (the "%" of `%d`, a leading
            # "." or "*", a trailing "-") is still resolved by the paragraph's
            # RTL base and would flip across the token, so anchor those edges.
            add_run(paragraph, anchor_ltr(tok.get("raw", "")), bold=bold,
                    italic=italic, rtl=False, style="Code", container=container)
        elif t == "link":
            render_link(paragraph, tok, bold, italic, container, base=base)
        elif t == "image":
            # Inline image (embedded mid-text): fall back to rendering its alt
            # text. Standalone images on their own line are handled as centered
            # image blocks in render_block instead.
            render_inline(paragraph, tok.get("children"), bold=bold, italic=italic,
                          container=container, style=style, base=base)
        elif t == "linebreak":
            add_break(paragraph, container=container)
        elif t == "softbreak":
            add_run(paragraph, " ", container=container)


def render_link(paragraph, tok, bold, italic, container=None, base="rtl"):
    """Render a Markdown link (including an auto-linked bare URL) as a real,
    clickable Word hyperlink: an external relationship + a w:hyperlink wrapping the
    link's text runs, styled with the built-in 'Hyperlink' character style. The
    visible text keeps normal bidi run-splitting, so a URL stays LTR and a Hebrew
    link label stays RTL."""
    url = (tok.get("attrs") or {}).get("url", "")
    if not url:
        render_inline(paragraph, tok.get("children"), bold=bold, italic=italic,
                      container=container, base=base)
        return
    r_id = paragraph.part.relate_to(url, RT.HYPERLINK, is_external=True)
    hyper = OxmlElement("w:hyperlink")
    hyper.set(qn("r:id"), r_id)
    render_inline(paragraph, tok.get("children"), bold=bold, italic=italic,
                  container=hyper, style="Hyperlink", base=base)
    (paragraph._p if container is None else container).append(hyper)


# ---------------------------------------------------------------------------
# Block rendering
# ---------------------------------------------------------------------------

class Converter:
    def __init__(self, doc):
        self.doc = doc
        self.para_index = 0
        self.ambiguous = []
        self.missing_images = []
        self.base_dir = "."
        self.max_img_width = None

    def direction_for(self, text, overrides, force=None):
        """Base direction for a paragraph. A per-index --overrides entry wins; then
        a `force` direction (e.g. one shared across a whole list); otherwise the
        per-text word-count heuristic, which flags near-even paragraphs as
        ambiguous. When overridden or forced the paragraph is not flagged."""
        idx = self.para_index
        ovr = overrides.get(str(idx))
        direction, ambiguous = detect_direction(text, base="rtl")
        if ovr in ("rtl", "ltr"):
            direction = ovr
        elif force in ("rtl", "ltr"):
            direction = force
        elif ambiguous:
            self.ambiguous.append((idx, direction, text[:60]))
        self.para_index += 1
        return direction

    def render_blocks(self, tokens, para_style=None):
        for tok in tokens or []:
            self.render_block(tok, para_style)

    def render_block(self, tok, para_style=None):
        t = tok.get("type")
        if t == "heading":
            self.render_heading(tok)
        elif t in ("paragraph", "block_text"):
            image = self._lone_image(tok)
            if image is not None:
                self.render_image_block(image)
            else:
                self.render_paragraph(tok, para_style)
        elif t == "block_code":
            self.render_code_block(tok)
        elif t == "block_quote":
            self.render_quote(tok)
        elif t == "list":
            self.render_list(tok)
        elif t == "table":
            self.render_table(tok)
        # thematic_break / blank_line: rely on default paragraph flow

    def render_heading(self, tok):
        level = tok["attrs"]["level"]
        idx = self.para_index
        ovr = self.overrides.get(str(idx))
        if ovr in ("rtl", "ltr"):
            direction = ovr
        else:
            direction = self.heading_directions.get(level, "rtl")
        self.para_index += 1
        p = self.doc.add_paragraph(style=f"Heading {min(level, 9)}")
        set_paragraph_direction(p, direction == "rtl")
        render_inline(p, tok["children"], base=direction)

    def render_paragraph(self, tok, para_style=None, force_dir=None):
        text = inline_text(tok["children"])
        direction = self.direction_for(text, self.overrides, force=force_dir)
        p = self.doc.add_paragraph(style=para_style)
        set_paragraph_direction(p, direction == "rtl")
        render_inline(p, tok["children"], base=direction)
        return p

    @staticmethod
    def _lone_image(tok):
        """Return the image token if this paragraph is a single standalone image
        (the `![alt](src)` on its own line case), else None. Whitespace-only text
        and soft/hard breaks around the image are ignored; any real text alongside
        the image makes it an inline image, not a block."""
        children = tok.get("children") or []
        images = [c for c in children if c.get("type") == "image"]
        if len(images) != 1:
            return None
        for c in children:
            ct = c.get("type")
            if ct == "image" or ct in ("softbreak", "linebreak"):
                continue
            if ct == "text" and not (c.get("raw") or "").strip():
                continue
            return None  # real content beside the image -> treat as inline
        return images[0]

    def render_image_block(self, image_tok):
        """Render a standalone image centered, scaled down to the page's content
        width if larger, with the alt text as a centered caption below it."""
        self.para_index += 1  # keep --overrides paragraph indices stable
        src = (image_tok.get("attrs") or {}).get("url", "")
        path = src
        if path and not os.path.isabs(path):
            path = os.path.normpath(os.path.join(self.base_dir, path))

        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if src and os.path.exists(path):
            try:
                pic = p.add_run().add_picture(path)
                if self.max_img_width and pic.width > self.max_img_width:
                    ratio = self.max_img_width / pic.width
                    pic.width = int(self.max_img_width)
                    pic.height = int(pic.height * ratio)
            except Exception as e:  # unreadable/corrupt image: don't abort
                add_run(p, f"[image error: {src}]", rtl=False)
                self.missing_images.append((src, str(e)))
        else:
            add_run(p, f"[missing image: {src}]", rtl=False)
            self.missing_images.append((src, "file not found"))

        cap_text = inline_text(image_tok.get("children"))
        if cap_text.strip():
            style = "Caption" if "Caption" in (s.name for s in self.doc.styles) else None
            cap = self.doc.add_paragraph(style=style)
            cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            direction, _ = detect_direction(cap_text, base="rtl")
            set_paragraph_direction(cap, direction == "rtl")
            render_inline(cap, image_tok.get("children"), base=direction)

    def render_code_block(self, tok):
        raw = tok.get("raw", "").rstrip("\n")
        lines = raw.split("\n")
        p = self.doc.add_paragraph()
        set_paragraph_direction(p, rtl=False)
        for i, line in enumerate(lines):
            add_run(p, line if line else " ", rtl=False, style="Code")
            if i != len(lines) - 1:
                add_break(p)
        self.para_index += 1

    def render_quote(self, tok):
        style = "Quote" if "Quote" in (s.name for s in self.doc.styles) else None
        self.render_blocks(tok.get("children"), para_style=style)

    def _new_num(self, ordered):
        """Allocate a fresh numId (same abstract list) so each list restarts at 1."""
        abs_id = _ABS_ORDERED if ordered else _ABS_BULLET
        nid = self.next_num_id
        self.next_num_id += 1
        num = OxmlElement("w:num")
        num.set(qn("w:numId"), str(nid))
        ai = OxmlElement("w:abstractNumId")
        ai.set(qn("w:val"), str(abs_id))
        num.append(ai)
        lo = OxmlElement("w:lvlOverride")
        lo.set(qn("w:ilvl"), "0")
        so = OxmlElement("w:startOverride")
        so.set(qn("w:val"), "1")
        lo.append(so)
        num.append(lo)
        self.numbering_el.append(num)
        return nid

    @staticmethod
    def _list_direction(tok):
        """One base direction for a whole list, from the aggregate word count of
        all its items (like heading levels). A Hebrew instruction list whose short
        items lead with a Hebrew label but are dominated per-item by a technical
        term - "מקש Ctrl+F5", "תפריט Debug ▸ Start Without Debugging" - would each
        score LTR alone; aggregated with their Hebrew siblings the list reads RTL,
        so the items align consistently. --overrides still fixes a genuine outlier."""
        h = l = 0
        stack = list(tok.get("children", []))
        while stack:
            node = stack.pop()
            for child in node.get("children", []) or []:
                ct = child.get("type")
                if ct in ("block_text", "paragraph"):
                    hh, ll = counts(inline_text(child.get("children")))
                    h += hh
                    l += ll
                elif ct in ("list", "list_item"):
                    stack.append(child)
        return _direction_from_counts(h, l)

    def render_list(self, tok):
        ordered = tok["attrs"].get("ordered", False)
        num_id = self._new_num(ordered)
        list_dir = self._list_direction(tok)
        for item in tok.get("children", []):
            first = True
            for child in item.get("children", []):
                if child.get("type") in ("block_text", "paragraph"):
                    p = self.render_paragraph(
                        child, para_style=list_paragraph_style(self.doc),
                        force_dir=list_dir)
                    if first:
                        apply_numbering(p, num_id)
                        first = False
                else:
                    self.render_block(child)

    def render_table(self, tok):
        head = None
        body_rows = []
        for child in tok.get("children", []):
            if child["type"] == "table_head":
                head = child["children"]
            elif child["type"] == "table_body":
                body_rows = child["children"]
        ncols = len(head) if head else (len(body_rows[0]["children"]) if body_rows else 0)
        if ncols == 0:
            return
        table = self.doc.add_table(rows=0, cols=ncols)
        set_table_borders(table)
        table._tbl.tblPr.append(OxmlElement("w:bidiVisual"))
        if head:
            row = table.add_row()
            for cell, cellfmt in zip(row.cells, head):
                self._fill_cell(cell, cellfmt, header=True)
        for r in body_rows:
            row = table.add_row()
            for cell, cellfmt in zip(row.cells, r["children"]):
                self._fill_cell(cell, cellfmt, header=False)

    def _fill_cell(self, cell, cellfmt, header):
        p = cell.paragraphs[0]
        for r in list(p._p.findall(qn("w:r"))):
            p._p.remove(r)
        text = inline_text(cellfmt.get("children"))
        direction, _ = detect_direction(text, base="rtl")
        set_paragraph_direction(p, direction == "rtl")
        render_inline(p, cellfmt.get("children"), bold=header, base=direction)


# ---------------------------------------------------------------------------
# Document setup
# ---------------------------------------------------------------------------

def ensure_code_style(doc, code_font):
    if "Code" in (s.name for s in doc.styles):
        return
    st = doc.styles.add_style("Code", WD_STYLE_TYPE.CHARACTER)
    st.base_style = doc.styles["Default Paragraph Font"] if \
        "Default Paragraph Font" in (s.name for s in doc.styles) else None
    st.font.name = code_font
    rPr = st.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:cs"), code_font)
    rFonts.set(qn("w:hAnsi"), code_font)
    rFonts.set(qn("w:ascii"), code_font)


def ensure_caption_style(doc):
    """Provide a 'Caption' paragraph style for image captions. If the template
    already defines one (built-in or custom), leave it alone; otherwise create a
    Word-like Caption: smaller, italic, subdued color, based on Normal."""
    if "Caption" in (s.name for s in doc.styles):
        return
    st = doc.styles.add_style("Caption", WD_STYLE_TYPE.PARAGRAPH)
    if "Normal" in (s.name for s in doc.styles):
        st.base_style = doc.styles["Normal"]
    st.font.italic = True
    st.font.size = Pt(9)
    st.font.color.rgb = RGBColor(0x44, 0x44, 0x44)


def ensure_hyperlink_style(doc):
    """Provide a 'Hyperlink' character style (blue + underlined) for links, unless
    the template already defines one. Word's built-in name is 'Hyperlink'."""
    if "Hyperlink" in (s.name for s in doc.styles):
        return
    st = doc.styles.add_style("Hyperlink", WD_STYLE_TYPE.CHARACTER)
    if "Default Paragraph Font" in (s.name for s in doc.styles):
        st.base_style = doc.styles["Default Paragraph Font"]
    st.font.color.rgb = RGBColor(0x05, 0x63, 0xC1)
    st.font.underline = True


def set_table_borders(table, color="auto", sz="4"):
    """Single borders on all sides + inside grid, without depending on the
    'Table Grid' style (which may be absent from a custom template)."""
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), sz)
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    tblPr.append(borders)


def list_paragraph_style(doc):
    return "List Paragraph" if "List Paragraph" in (s.name for s in doc.styles) else None


NUMID_ORDERED = 1101
NUMID_BULLET = 1102
_ABS_ORDERED = 1101
_ABS_BULLET = 1102

_NUMBERING_XML = f"""<w:numbering xmlns:w="{nsmap['w']}">
  <w:abstractNum w:abstractNumId="{_ABS_ORDERED}">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:start="720" w:hanging="360"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="{_ABS_BULLET}">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:start="720" w:hanging="360"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="{NUMID_ORDERED}"><w:abstractNumId w:val="{_ABS_ORDERED}"/></w:num>
  <w:num w:numId="{NUMID_BULLET}"><w:abstractNumId w:val="{_ABS_BULLET}"/></w:num>
</w:numbering>"""


def ensure_numbering(doc):
    """Attach a numbering part with one decimal and one bullet definition, so
    lists work regardless of whether the base template defines list styles."""
    from docx.opc.constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT
    from docx.opc.packuri import PackURI
    from docx.oxml import parse_xml
    from docx.parts.numbering import NumberingPart
    try:
        doc.part.numbering_part  # already present
        return
    except NotImplementedError:
        pass
    part = NumberingPart(PackURI("/word/numbering.xml"), CT.WML_NUMBERING,
                         parse_xml(_NUMBERING_XML), doc.part.package)
    doc.part.relate_to(part, RT.NUMBERING)


def apply_numbering(paragraph, num_id, ilvl=0):
    pPr = paragraph._p.get_or_add_pPr()
    numPr = OxmlElement("w:numPr")
    il = OxmlElement("w:ilvl"); il.set(qn("w:val"), str(ilvl)); numPr.append(il)
    ni = OxmlElement("w:numId"); ni.set(qn("w:val"), str(num_id)); numPr.append(ni)
    pPr.insert(0, numPr)


def strip_leading_empty(doc):
    """Drop blank paragraphs a template may start with, so content begins clean."""
    body = doc.element.body
    for child in list(body):
        tag = child.tag.split("}")[-1]
        if tag == "p":
            if child.find(qn("w:r")) is None and not (child.text or "").strip():
                body.remove(child)
            else:
                break
        else:
            break


def convert(md_path, out_path, overrides, code_font, template=None):
    with open(md_path, encoding="utf-8") as f:
        text = f.read()
    md = mistune.create_markdown(
        renderer=None,
        plugins=["table", "strikethrough", "url"],
    )
    tokens = md(text)
    # Base the document on a template if given (or a bundled template.docx next
    # to this script); otherwise use python-docx's stock default. The converter
    # is template-agnostic: it supplies its own numbering and table borders and
    # only relies on Normal / Heading N / Quote, which any Word template has.
    if template is None:
        bundled = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                               "template.docx")
        if os.path.exists(bundled):
            template = bundled
    doc = Document(template) if template else Document()
    ensure_code_style(doc, code_font)
    ensure_caption_style(doc)
    ensure_hyperlink_style(doc)
    ensure_numbering(doc)
    conv = Converter(doc)
    conv.overrides = overrides
    conv.heading_directions = heading_level_directions(tokens)
    conv.numbering_el = doc.part.numbering_part.element
    conv.next_num_id = 2000
    conv.base_dir = os.path.dirname(os.path.abspath(md_path))
    section = doc.sections[0]
    conv.max_img_width = (section.page_width - section.left_margin
                          - section.right_margin)
    conv.render_blocks(tokens)
    strip_leading_empty(doc)
    doc.save(out_path)
    return conv


def main():
    ap = argparse.ArgumentParser(description="Hebrew Markdown -> DOCX (RTL-aware)")
    ap.add_argument("input")
    ap.add_argument("-o", "--output")
    ap.add_argument("--overrides", help="JSON map of paragraph index -> rtl/ltr")
    ap.add_argument("--code-font", default="Consolas")
    ap.add_argument("--template", help="base .docx to inherit styles from "
                    "(e.g. a blank doc saved from your Word default template)")
    args = ap.parse_args()

    out = args.output or re.sub(r"\.md$", "", args.input) + ".docx"
    overrides = {}
    if args.overrides:
        with open(args.overrides, encoding="utf-8") as f:
            overrides = json.load(f)

    conv = convert(args.input, out, overrides, args.code_font, args.template)
    print(f"Wrote {out}")
    if conv.missing_images:
        print(f"\n{len(conv.missing_images)} image(s) could not be embedded "
              f"(a visible placeholder was inserted instead):", file=sys.stderr)
        for src, reason in conv.missing_images:
            print(f"  {src!r}: {reason}", file=sys.stderr)
    if conv.ambiguous:
        print(f"\n{len(conv.ambiguous)} ambiguous paragraph(s) "
              f"(near-even Hebrew/Latin counts). Resolve via --overrides "
              f"(index -> rtl/ltr):", file=sys.stderr)
        for idx, d, snippet in conv.ambiguous:
            print(f"  [{idx}] used={d}: {snippet!r}", file=sys.stderr)


if __name__ == "__main__":
    main()
