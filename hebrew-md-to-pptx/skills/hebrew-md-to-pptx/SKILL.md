---
name: hebrew-md-to-pptx
description: >-
  Convert a Markdown file written in the user's slide-authoring format (with
  Hebrew or mixed Hebrew/English) into a basic, RTL-aware PowerPoint .pptx
  skeleton. Use whenever the user wants to turn slide-Markdown (blocks separated
  by `---`, content/notes separated by `===`, section/slide labels and presenter
  instructions in [brackets]) into an editable deck, or complains that pasting
  such Markdown into PowerPoint breaks right-to-left alignment, flips the
  punctuation around English, or flags every Hebrew word as a typo. Produces real
  title/body placeholders, native PowerPoint sections, speaker notes, tables,
  embedded images, and per-paragraph text direction. The output is intentionally
  unstyled, ready to be themed afterwards by hand or with the pptx skill.
metadata:
  version: 1.0.0
  author: Assaf Tzur-El
---

# Hebrew slide-Markdown → PPTX

Converts a Markdown deck written in the slide-authoring format into a clean,
RTL-aware `.pptx` skeleton. This is the slide counterpart of the
`hebrew-md-to-docx` skill, and the inverse of `pptx-to-slide-md` (so a deck can
round-trip: PPTX → slide-Markdown → PPTX).

## Why this skill exists

Pasting slide-Markdown into PowerPoint forces the user to right-align every
Hebrew paragraph by hand, and still leaves bugs: punctuation and brackets around
English (`ו-CTO.`, `[^]`) drift to the wrong side, and every Hebrew word is
underlined as a spelling mistake because the runs carry no Hebrew language tag.
This skill fixes all of that at the OOXML level and lays out one slide per block
with the title, body, and speaker notes in the right places.

## How to use

The converter is `md_to_pptx_he.py` (Python 3; needs `python-pptx`, `mistune`,
and `Pillow`):

```bash
pip install python-pptx mistune Pillow --break-system-packages   # if needed
python md_to_pptx_he.py INPUT.md -o OUTPUT.pptx
```

Options:

- `--template DECK.pptx` – build on a template so the output inherits a specific
  theme, fonts, and slide layouts (see below). Without it, python-pptx's plain
  default theme is used.
- `--overrides overrides.json` – pin the base direction of ambiguous paragraphs
  (`{ "3": "rtl", "12": "ltr" }`, keyed by the index printed to stderr).
- `--code-font Consolas` – monospace font for code spans/blocks. Default
  `Consolas`.

Slide size: without a template the deck is **16:9 widescreen** (13.33"×7.5"), and
the title/body placeholders are widened to fill it. With `--template`, the
template's own slide size and layouts are kept as-is.

### When running this skill, first ask the user:

> Do you want to provide a custom .pptx template (a deck with your theme, fonts,
> and layouts) so the output matches your style? Otherwise I'll use the plain
> default theme, since the output is meant to be styled afterwards anyway.

## Input format

The document is a sequence of slide blocks. `---` on its own line separates
slides; `===` separates a slide's body from its speaker notes.

```markdown
# [Section 1: Section Name (~time)]   <- starts a native PowerPoint section
## [Slide 1: Brief Description]        <- scaffolding label, dropped from the slide

# Slide Title                          <- the real slide title
Body markdown: bullets, numbered lists, tables, **bold**, *italic*, `code`, [links](url).
[presenter instruction]                <- moved to the speaker notes
[image: path/to/pic.png — "alt"]       <- embedded onto the slide
[^]                                    <- a click marker; also moved to notes

===

**Speaker Notes:**
Notes in markdown. Single line breaks are preserved.

---

## [Slide 2: ...]
# Next Slide Title
...
```

Key rules:

- **Section labels** `# [Section N: Name]` start a native PowerPoint section. The
  label `Section N:` is stripped; only the name (everything after the colon,
  e.g. `Section Name (~time)`) is used. Every following slide belongs to that
  section until the next section label.
- **Slide labels** `## [Slide N: ...]` are scaffolding for the author and are
  dropped. A lone `[Slide N]` line (as produced by `pptx-to-slide-md`) is also
  recognised as a slide label and dropped, so an extracted deck round-trips.
- **Slide title** is the first `# Heading` inside the body. `# (untitled)` (and
  `# (Slide N — untitled)`) is treated as an empty title.
- **Presenter instructions** in `[brackets]` inside the body — animations, demos,
  `[^]` click markers, etc. — are removed from the slide and appended to the
  speaker notes under a "Presenter instructions:" heading. (`[text](url)` is a
  link, not an instruction, and is kept.)
- **Images** written as `[image: path — "alt"]` are embedded onto the slide
  (path resolved relative to the input .md). The alt text is ignored.
- `[hidden slide]` markers are dropped (the converter does not re-hide slides).

## Direction handling (the RTL core)

Unlike Word, PowerPoint sets text direction per **paragraph**, and there is no
per-run RTL flag. So:

- Each paragraph's base direction is chosen by the **majority of strong
  directional characters** (Hebrew vs Latin), and written as `a:pPr/@rtl` plus a
  matching `@algn` (right for RTL, left for LTR).
- Within a paragraph, text is split into runs by script, and each run is tagged
  with a **language** (`a:rPr/@lang` = `he-IL` or `en-US`). This does two jobs:
  the spell-checker uses the Hebrew dictionary for Hebrew runs (no more
  everything-is-a-typo), and PowerPoint resolves bidi per run so neutral
  punctuation (dash, dot, brackets) around English joins the Hebrew side instead
  of drifting LTR. Digits stay LTR so numbers like `1920` never reverse.
- Near-even paragraphs are printed to stderr as **ambiguous**; resolve them with
  `--overrides` keyed by the printed index.
- Tables whose header is Hebrew-dominant get right-to-left column order
  (`a:tbl` `@rtl`).

## What gets converted

| Markdown | PowerPoint |
| --- | --- |
| `# Title` (first heading in body) | slide title placeholder |
| paragraphs, `##`+ headings | body placeholder paragraphs (headings bold) |
| `-` / `1.` lists | bulleted / auto-numbered paragraphs with hanging indent |
| `**bold**`, `*italic*`, `~~strike~~` | run formatting |
| `` `code` `` / ```` ``` ```` blocks | monospace runs (LTR) |
| `[text](url)` | hyperlinked runs |
| tables | a table shape (RTL column order when Hebrew) |
| `[image: ...]` | embedded picture |
| body `[brackets]`, `[^]` | appended to speaker notes |
| speaker notes | notes pane; single line breaks preserved |

## Speaker notes

Notes keep the author's line structure: a single newline becomes a real line
break (Markdown would otherwise collapse it to a space). Bulleted/numbered notes
get a hanging indent so the marker is spaced from the text even though the notes
text frame defines no list styles.

## Limitations

- The output is a plain, unstyled skeleton (default theme). Styling — colors,
  fonts, layouts — is a deliberate later step.
- Table and image positioning is a rough estimate (text height can't be known in
  advance), so a slide that mixes a lot of body text with a table/image may need
  manual repositioning.
- Animations are not created; their description is only recorded in the notes.
- Hidden slides come through visible (the marker is dropped).
- Every content slide uses the "Title and Content" layout; pick other layouts by
  hand afterwards.
