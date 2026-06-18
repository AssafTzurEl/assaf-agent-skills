---
name: pptx-to-slide-md
description: >-
  Extract a PowerPoint (.pptx) file into the user's plain-text slide authoring
  format: a Markdown title + body per slide, presenter instructions in
  [brackets], a `===` separator, then speaker notes, with `---` between slides.
  Use this whenever the user uploads or points to a .pptx and wants the slides
  "as text", "in my slide format", reverse-engineered back into editable
  source, or wants the bullets / notes / animations / images pulled out of an
  existing deck. Trigger even if they just say "turn this deck into text" or
  "extract the slides from this pptx" — that is what this skill is for.
metadata:
  version: 1.0.0
  author: Assaf Tzur-El
---

# PPTX → slide-Markdown

Convert an existing `.pptx` back into the plain-text format the user writes
slides in, so a finished deck becomes editable source again.

## Target format

One block per slide. `---` separates slides; `===` separates slide content from
speaker notes. Presenter instructions (images, animations, demos) live in
`[brackets]` inside the content area. If the deck has sections, each section
begins with a `[Section N: Section Name]` label before the first slide in that
section.

```
[Section 1: Introduction]

[Slide 1]

# Slide Title

Slide content in markdown: bullets, numbered lists, tables, **bold**, *italic*, ~~strikethrough~~, [links](url).

[image: assets/slide3_img1.png — "alt text"]

[animation — 2 click steps:
  click 1: entrance of "First point"
  click 2: entrance of "Second point"]

===

**Speaker Notes:**

Notes in markdown.

---

[Slide 2]

# Next Slide Title

...

===

**Speaker Notes:**

*(none)*
```

Each slide is preceded by a `[Slide N]` label (N is the slide's position in the
deck, so a subset extraction keeps the original numbers). `---` separates slides
and `===` separates content from notes, as before.

Blank lines between elements ensure Markdown viewers render each block as a
separate paragraph rather than collapsing consecutive lines into soft-wrapped text.

The user also uses `[^]` inside speaker notes to mark a click. See the note on
animations below for why extraction puts animation detail in the content area
rather than fabricating `[^]` positions in the notes.

## How to run it

The work is done by `scripts/extract.py` (uses `python-pptx`, already available).

```bash
python scripts/extract.py <input.pptx> -o <output.md>
```

Options:
- `-o, --output` — output path (default: `<input-stem>.md` in the cwd).
- `--slides` — subset to extract, e.g. `--slides 3-20` or `--slides 1,4,7-9`
  (1-based; default: all). Numbers in `[Slide N]` stay absolute.
- `--images-dir DIR` — where to write extracted images (default:
  `<output-stem>_images/` next to the output file).
- `--no-images` — don't write image files; still emit `[image: filename]` inline.

After running, read the output file and make it available to the user. Surface
the generated `.md` (and the images folder) using whatever file-sharing
mechanism your runtime offers — a file preview or attachment if one exists,
otherwise point the user to the output path. Don't paste a huge deck inline if
it's long — write the file and show a representative excerpt.

## What gets extracted

- **Title** → `# Heading` (from the slide's title placeholder).
- **Body text** → bullets by indent level (`-`, nested with two spaces),
  numbered lists for auto-numbered paragraphs, plain paragraphs when a paragraph
  has no bullet. Run formatting becomes `**bold**`, `*italic*`, `***both***`,
  `~~strikethrough~~`, and `[text](url)` for hyperlinks. Adjacent runs with the
  same formatting are merged; soft line breaks within a paragraph become spaces.
- **Tables** → Markdown tables (pipes inside cells are escaped).
- **Images** → extracted to the images folder and referenced as
  `[image: <path> — "<alt text>"]` (alt text included when the shape has it).
- **Charts** → `[chart: <type> — "<title>"]` (not the underlying data).
- **SmartArt / diagrams** → `[SmartArt/diagram — text: ...]` when text is
  reachable, otherwise a note that the text isn't extractable.
- **Speaker notes** → Markdown under `**Speaker Notes:**` (`*(none)*` if empty).
- **Animations** → best-effort, see below.
- **Hidden slides** → flagged with `[hidden slide]`.
- **Sections** → `[Section N: Name]` emitted before the first slide of each
  section (only when the deck actually uses sections).
- Grouped shapes are walked recursively; shapes keep document order, with the
  title hoisted to the top.

## Animations — what is and isn't recoverable

PowerPoint stores animations as a time-node tree (`<p:timing>`) that
`python-pptx` does not expose, so the script parses the raw slide XML. Be honest
with the user about the boundaries:

- **Reliable:** the number of click-advance steps, and which shape/paragraph
  each step reveals (the script maps the animated paragraph range back to its
  text). Build-by-paragraph reveals (one bullet per click) come through.
- **Reliable:** the effect *category* — entrance, exit, emphasis, motion path.
- **Not reported:** the specific effect *name* (Fade vs. Wipe vs. Fly-In). The
  preset-ID → name mapping is version-dependent and easy to get wrong, so the
  script deliberately reports only the category rather than guessing.
- **By design:** animation detail goes in a `[animation — …]` block in the
  content area (matching the user's convention that presenter instructions live
  in brackets). The script does **not** insert `[^]` click markers into the
  notes, because where each click belongs in the authored notes prose is not
  stored in the file and can't be reconstructed faithfully. If the user wants
  `[^]` markers in the notes, offer to add them using the click count as a guide.
- If timing XML is present but can't be parsed, the slide gets
  `[animation present — could not parse (…)]` instead of failing.

## Build sequences across slides (always surface these)

Decks often split one idea across several consecutive slides: a list that grows
one bullet per slide, or a fixed background with different things overlaid. The
script flags runs of consecutive slides that share a title and prints them to
stderr under `BUILD SEQUENCES DETECTED`, classifying each as either an additive
build or a shared base with changing overlays.

Always extract these as separate slides (one block each, as normal) — do not
merge them automatically. But after extraction, **tell the user which sequences
were detected and offer to compact each into a single slide** with the build
expressed as progressive-disclosure instructions and `[^]` click markers in the
notes. Only compact if the user agrees. This applies to additive lists,
overlay-style sequences, and restyled finales (e.g. a slide that repeats the
previous one with text struck through).

## Optional follow-ups

- **Describe images:** if your runtime can view images, open the extracted image
  files and, if the user wants, replace `[image: file]` with a short description.
- **Verify a tricky deck:** if the output looks off (e.g. a text box that should
  be a heading came through as a bullet, or vice-versa), the slide XML can be
  inspected directly; bullet detection relies on explicit `buNone`/`buChar`/
  `buAutoNum` formatting and falls back to "body placeholders bullet, everything
  else is plain", which is a heuristic, not a guarantee.

## Edge cases worth knowing

- Embedded media, OLE objects, and equations are not extracted (only noted if
  they carry text).
- Merged table cells repeat their content per the underlying model.
- A slide with no title placeholder is emitted as `# (Slide N — untitled)`.