# The slide-Markdown format

A plain-text format for authoring slide decks in Markdown. It is the shared
contract between two skills in this collection:

- [`hebrew-md-to-pptx`](../hebrew-md-to-pptx) **consumes** this format and builds a `.pptx`.
- [`pptx-to-slide-md`](../pptx-to-slide-md) **produces** this format by extracting an existing `.pptx`.

Because they are inverses, a deck can round-trip: `PPTX → slide-Markdown → PPTX`.

The format is ordinary Markdown plus a few conventions that carry slide
structure (where slides start, what's a title, what's a speaker note, what's a
presenter instruction). It is not a generic "Markdown dump" of a deck - it is a
specific, editable authoring shape.

## At a glance

| Marker | Meaning |
| --- | --- |
| `---` (on its own line) | Separates one slide from the next |
| `===` (on its own line) | Within a slide, separates the body from the speaker notes |
| `# Heading` | The slide title (the first `#` heading in the body) |
| `[Section N: Name]` | Starts a native presentation section; applies to the slides that follow |
| `[Slide N]` | A slide label/position marker; scaffolding only |
| `[ ... ]` (other brackets) | A presenter instruction - moved to the notes, not shown on the slide |
| `[image: path — "alt"]` | An image to embed on the slide |
| `[^]` | A "click" / build-step marker |

Blank lines between blocks are significant for readability: they make Markdown
viewers render each block as its own paragraph instead of soft-wrapping them
together.

## Structure of one slide

```markdown
[Slide 1]

# Slide Title

Body content as Markdown: bullets, numbered lists, tables, **bold**, *italic*,
~~strikethrough~~, `code`, and [links](https://example.com).

- a bullet
- another bullet
  - a nested bullet

[image: assets/diagram.png — "architecture overview"]

[animation — 2 click steps:
  click 1: entrance of "First point"
  click 2: entrance of "Second point"]

===

**Speaker Notes:**

What the presenter says here. Single line breaks are preserved. This [^] is a mouse click reminder at the right time.

---
```

Reading that top to bottom:

1. `[Slide 1]` labels the slide by position. It is scaffolding - when
   `hebrew-md-to-pptx` builds the deck, this label is dropped.
2. `# Slide Title` is the first heading and becomes the slide's title
   placeholder. Use `# (untitled)` (or `# (Slide N — untitled)`) for a slide
   with no title.
3. Everything between the title and `===` is the slide body.
4. `[brackets]` that are not links, images, or section/slide labels are
   presenter instructions (animations, demos, build cues). They are removed
   from the visible slide and appended to the speaker notes.
5. `===` ends the body and begins the speaker notes.
6. `**Speaker Notes:**` introduces the notes. Use `*(none)*` when there are no
   notes.
7. `[^]` hints the presenter when to initiate an event (mouse click, usually for animation). As a rule of thumb, the number of `[^]`s in the speaker notes should match the number of clicks in the slide body's `[animation]` labels.
8. `---` closes the slide; the next slide begins after it.

## Sections

If a deck is organized into sections, start each one with a section label on its
own line before the first slide of that section:

```markdown
[Section 1: Introduction]

[Slide 1]

# Welcome

...
```

`hebrew-md-to-pptx` turns `[Section N: Name]` into a real presentation section.
Only the name after the colon is used (`Introduction`), and the section applies
to every following slide until the next section label. When authoring by hand
you may also see the longer scaffolding form `## [Slide N: short description]`;
the `Slide N:` part is dropped on build.

## Body content

The body is regular Markdown:

- **Bullets** with `-`, nested by indenting two spaces.
- **Numbered lists** with `1.`, `2.`, ...
- **Tables** with pipes (`|`).
- **Inline formatting**: `**bold**`, `*italic*`, `***both***`,
  `~~strikethrough~~`, `` `code` ``, and `[text](url)` links.

## Images

Embed an image with:

```markdown
[image: path/to/picture.png — "alt text"]
```

The path is resolved relative to the Markdown file. The separator before the
alt text is an en dash (`–`), and the alt text is quoted. When
`pptx-to-slide-md` extracts a deck it writes images to an images folder and
emits this same line; when `hebrew-md-to-pptx` builds a deck it embeds the
picture (the alt text is currently ignored on build).

## Presenter instructions and clicks

Anything in `[brackets]` that isn't a link, an image, or a section/slide label
is treated as a presenter instruction: animation descriptions, demo cues,
reminders. These are moved off the visible slide into the speaker notes under a
"Presenter instructions" heading. A lone `[^]` marks a click / build step and is
handled the same way.

Animations are descriptive only - the format records *what* happens (number of
click steps, which paragraph each step reveals, and the effect category:
entrance / exit / emphasis / motion), but building a `.pptx` does not recreate
live animations; the description is preserved in the notes.

## A fuller example

```markdown
[Section 1: Overview]

[Slide 1]

# What we'll cover

- The problem
- Our approach
- Results

===

**Speaker Notes:**

Keep this short - it's just a roadmap.

---

[Slide 2]

# Results

| Metric | Before | After |
| --- | --- | --- |
| Latency | 800ms | 120ms |

[animation: entrance of chart]
[image: assets/latency.png — "latency chart"]

===

**Speaker Notes:**

Discuss the table, then [^] walk through the table.

---
```

## Round-trip notes

The two skills are inverses, but a round-trip is not byte-for-byte identical,
because some information lives in PowerPoint structures the text format does not
try to reconstruct:

- `[Slide N]` labels are dropped when building and re-added (by position) when
  extracting.
- Live animations are described in the notes, not recreated as timing data.
- Styling (theme, fonts, colors, layouts) is intentionally out of scope -
  `hebrew-md-to-pptx` produces an unstyled skeleton meant to be themed
  afterwards.

For the authoritative, skill-specific behavior, see each skill's `SKILL.md`:
[`hebrew-md-to-pptx`](../hebrew-md-to-pptx/skills/hebrew-md-to-pptx/SKILL.md) and
[`pptx-to-slide-md`](../pptx-to-slide-md/skills/pptx-to-slide-md/SKILL.md).
