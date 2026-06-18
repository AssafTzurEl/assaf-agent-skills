---
name: hebrew-md-to-docx
description: >-
  Convert a Markdown file that contains Hebrew (usually mixed Hebrew/English) into
  a properly formatted, RTL-aware Word .docx. Use whenever the user asks to turn a
  Hebrew or mixed Hebrew/English .md file into a Word document, or complains that
  pasting Markdown into Word/Google Docs breaks right-to-left alignment, flips
  inline code, or mangles numbered lists. Handles headings, bold/italic, inline
  code, fenced code blocks, tables, blockquotes, and ordered/unordered lists, and
  sets per-paragraph text direction automatically.
metadata:
  version: 1.0.0
  author: Assaf Tzur-El
---

# Hebrew Markdown → DOCX

Converts Hebrew (or mixed Hebrew/English) Markdown to a clean, RTL-aware `.docx`.

## Why this skill exists

Pasting Markdown into Word/Google Docs forces the user to manually right-align
every paragraph, left-align the English ones, and still leaves bugs: inline code
like `` `a = b` `` reorders to `b = a`, numbered markers (`1. `) render LTR and
detach from their Hebrew item, parentheses/quotes around English drift to the
wrong side, and a Hebrew paragraph that starts with an English word gets aligned
the wrong way. This skill fixes all of that at the OOXML level.

## How to use

The converter is `md_to_docx_he.py` (Python 3; needs `python-docx` and `mistune`):

```
pip install python-docx mistune --break-system-packages   # if not already present
python md_to_docx_he.py INPUT.md -o OUTPUT.docx
```

Options:

- `--template BLANK.docx` – build on a template so the output matches a specific
  set of Word styles/fonts (see below).
- `--code-font Consolas` – monospace font for the `Code` style. Default `Consolas`.
- `--overrides overrides.json` – manual direction fixes for ambiguous paragraphs.

## Templates (matching the user's Word styles)

A `.docx` always carries its own styles, so the output's look (fonts, heading
styles, spacing) comes from the template it is built on — not from the user's
`Normal.dotm` at open time. Three levels:

1. **Bundled default** – if `template.docx` sits next to the script, it is used
   automatically. The bundled one here matches the user's Word defaults
   (Aptos + Arial for Hebrew).
2. **Per-run override** – pass `--template path.docx` to use a different one.
3. **Stock** – with no template available, python-docx's standard default
   (Calibri) is used.

A template should be an (essentially) empty document saved from the desired Word
template. The converter is template-agnostic: it supplies its own list numbering
and table borders, and only relies on `Normal`, `Heading 1..9`, and `Quote`,
which every Word template has — so any template works without missing-style errors.

### When running this skill, first ask the user:

> Do you want to provide a custom .docx template (a blank doc saved from your Word
> default template) so the output matches your styles/fonts? Otherwise I'll use
> the bundled default.

If they provide one, pass it via `--template`.

## Direction logic (heuristic first, AI for ties)

Each paragraph's base direction is chosen by the **majority of strong directional
characters** (Hebrew vs Latin), not by the first character — so "Claude הוא כלי
בינה מלאכותית" is correctly RTL. When the counts are near-even (within 20%) the
script picks a direction but prints the paragraph to stderr as **ambiguous**:

```
2 ambiguous paragraph(s) ... Resolve via --overrides (index -> rtl/ltr):
  [118] used=rtl: 'scanf אוגרת הקלדות: ה-buffer'
```

Read those paragraphs, decide which language is primary, and pass an overrides
file keyed by the printed index:

```json
{ "118": "rtl", "169": "ltr" }
```

then re-run with `--overrides overrides.json`. Most documents need none.

## How direction is represented (matches Word's own output)

- Paragraph direction via `w:bidi`; **no `w:jc`** — an RTL paragraph then defaults
  to right alignment, an LTR one to left, exactly as Word writes it.
- Only Hebrew (RTL) runs get `w:rtl`; English/number/code runs omit it.
- Neutral characters (spaces, parens, quotes, dashes, punctuation) are LTR only
  when between two LTR characters (so "format string" stays together); otherwise
  they join the Hebrew flow (so the parens/quotes/dashes around English stay on
  the Hebrew side). Digits are LTR so numbers like 1920 never reverse.
- No Unicode control characters are inserted; nothing visible is added.

## What gets converted

| Markdown | Word output |
| --- | --- |
| `#`/`##`/`###` headings | `Heading 1..9`, direction per content |
| **bold**, *italic* | bold / italic runs |
| `` `inline code` `` | `Code` character style (Normal + monospace), kept LTR |
| ```` ``` ```` fenced blocks | LTR monospace lines (`Code` style) |
| tables | RTL table (`bidiVisual`) with its own borders, bold header row |
| `>` blockquotes | built-in `Quote` style |
| `1.` / `-` lists | self-contained numbering (each list restarts at 1), RTL-aware |

## Limitations

- Nested / multi-level lists render at a single level.
- Images and raw inline HTML are skipped.
- Tuned for Hebrew-dominant content; primarily-English docs still convert but the
  base direction is RTL.
