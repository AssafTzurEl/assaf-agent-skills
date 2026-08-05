---
name: hebrew-md-to-docx
description: >-
  Convert a Markdown file that contains Hebrew (usually mixed Hebrew/English) into
  a properly formatted, RTL-aware Word .docx. Use whenever the user asks to turn a
  Hebrew or mixed Hebrew/English .md file into a Word document, or complains that
  pasting Markdown into Word/Google Docs breaks right-to-left alignment, flips
  inline code, or mangles numbered lists. Handles headings, bold/italic, inline
  code, fenced code blocks, tables, blockquotes, ordered/unordered lists, clickable
  hyperlinks, and centered images with captions, and sets per-paragraph text
  direction automatically.
metadata:
  version: 1.1.1
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

### Running this skill

Just run the converter with no `--template` flag — the bundled default is used
automatically, no need to ask the user first. Only pass `--template path.docx`
if the user explicitly asks for a different template (e.g. "use my company's
Word template" or provides their own .docx to match).

## Direction logic (heuristic first, AI for ties)

Each paragraph's base direction is chosen by the **majority of strong directional
words** (Hebrew vs Latin), not by the first character — so "Claude הוא כלי בינה
מלאכותית" is correctly RTL. When the counts are near-even (within 20%) the script
picks a direction but prints the paragraph to stderr as **ambiguous**:

**Words, not characters.** Counting *words* (maximal runs of one script) instead
of characters keeps a Hebrew paragraph RTL even when its English is a few long or
repeated proper nouns. For example "אנחנו נשתמש ב-Visual Studio Community –
המהדורה החינמית והמלאה של Visual Studio…" has slightly more Latin *letters* than
Hebrew (long, repeated "Visual Studio"), which would wrongly flip it LTR by
character count; by word count the Hebrew prose clearly wins, so it stays RTL.

**URLs and e-mails are excluded.** A URL or e-mail address is data, not prose, and
its characters must not sway direction. They are stripped before counting, so a
labelled link line like "**הורדה**: https://visualstudio.microsoft.com/vs/…" is
driven by the Hebrew label and stays RTL instead of being dragged LTR by the URL.

**Headings are a further exception.** A single heading's own text is not a
reliable signal — e.g. `### הצהרה (declaration)` is a Hebrew heading with an
English gloss. So each heading *level* (H1, H2, H3, ...) gets one direction,
chosen by the aggregate Hebrew/Latin **word** count across every heading at that
level in the document. All headings at that level use it, even ones whose own text
would individually score the other way. This assumes each level is predominantly
one language; if a document genuinely mixes fully-Hebrew and fully-English
headings at the same level, use `--overrides` to fix the minority ones by
paragraph index.

**Lists share one direction, like heading levels.** All items in a list get a
single base direction from the aggregate word count of the whole list, so short
items that individually score LTR — a Hebrew label followed by a technical term,
e.g. "מקש Ctrl+F5" or "תפריט Debug ▸ Start Without Debugging" — align RTL with
their Hebrew siblings instead of coming out ragged. A genuinely mixed list can be
corrected per item with `--overrides`.

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
- Only the actual Hebrew blocks (`U+0590–U+05FF`, `U+FB1D–U+FB4F`) count as RTL.
  Symbols such as `▸`, `→`, `–` and `…` are neutral and take their direction from
  their neighbours, so a menu path like `Build ▸ Build Solution` reads as one
  left-to-right unit. (The script spells those ranges as `\uXXXX` escapes on
  purpose – written as literal characters, `U+FB1D` decomposes under Unicode
  normalisation and silently widens the class to `U+05B4–U+FB4F`, which swallows
  arrows, dashes and CJK.)
- Only Hebrew (RTL) runs get `w:rtl`; English/number/code runs omit it.
- Neutral characters (spaces, parens, quotes, dashes, punctuation) fall back to
  the **paragraph's own base direction** unless both neighbours are the opposite
  direction. Under an RTL paragraph a neutral is LTR only when it lies between
  two LTR characters (so "format string" stays together, while the
  parens/quotes/dashes around English stay on the Hebrew side). Under an LTR
  paragraph the rule mirrors: a neutral is RTL only between two RTL characters,
  so the period ending a mostly-English sentence that happens to close with a
  Hebrew word – "This one, however, is 10% עברית." – stays at the sentence's end
  instead of jumping into the Hebrew flow. Digits are LTR so numbers like 1920
  never reverse.
- "Word-joiner" symbols (`+`, `#`, `*`, `.`, `\`, `/`, `!`, `?`, `%`) that touch a
  Latin letter/digit are kept on the Latin side, so tech tokens and embedded
  English phrases stay intact: `C++`/`C#` don't flip to `++C`/`#C`; `.NET`,
  `Node.js`, `*.txt`, `3.14` hold together; paths/URLs like `C:\Users\foo`,
  `/home/assaf`, `https://x/y` read left-to-right as one unit; and the `!`/`?` of
  an embedded English exclamation stays put (`"Hello, World!"` keeps its `!`). A
  joiner sitting at the boundary between the Latin run and the surrounding Hebrew
  is anchored with an invisible LRM (`U+200E`) so the bidi algorithm can't drag it
  across. A joiner that touches no Latin character (a lone `+` between Hebrew
  words, `/` in a Hebrew `כן/לא`, the `?` in a Hebrew `נתקעתם?`) is left alone.
  The dot is a special case: because it is also the sentence period, it joins a
  token only when a Latin letter/digit *follows* it (as in `.NET`, `file.txt`), so
  a period ending a Hebrew sentence — even after an English word or number, like
  `Visual Studio.` or `123.` — correctly stays in the RTL flow.
  The percent sign is the mirror case: it always trails its number and never
  leads it, so it joins only when a Latin letter/digit *precedes* it. That keeps
  `80%` / `27%` intact instead of rendering as `%80` / `%27`, and it puts the
  digits and the sign in one LTR run so they also share one font. A lone `%` with
  no adjacent digit or letter keeps the normal neutral treatment; `‰` behaves the
  same way. It also joins *forward* to a directly-adjacent Latin letter/digit, so
  a printf specifier written outside backticks (`%d`, `%s`, `%5.2f`) holds
  together too.
  `!` and `?` need the same care, because they end an embedded English phrase
  *and* the surrounding Hebrew sentence. They attach to the Latin token only when
  something immediately follows them – a closing quote or bracket, or more Latin
  (`"Hello, World!"`, `(Really?)`). When the next character is a space or the text
  ends there, the mark is terminating the Hebrew sentence and stays in the RTL
  flow, so "אין לכם Windows? היחידה עדיין בשבילכם" puts the `?` on the Hebrew side
  of "Windows", not the English one.
- An ellipsis (`...` or a spaced `. . .`) next to an English phrase stays with it
  (`"…this window . . ."` keeps the dots at the end); a Hebrew-side ellipsis is
  left in the RTL flow.
- Inline code gets the same anchoring. Marking a code run LTR is not enough on
  its own: a *neutral* character at the run's edge is still resolved by the
  paragraph's base direction, so `` `%d` `` inside a Hebrew paragraph came out as
  `d%`. A code run whose first or last character is not strongly LTR is bracketed
  with LRMs, which also protects a leading `` `.NET` `` / `` `*.txt` `` and a
  trailing `` `--flag-` ``.
- The LRM anchor is the only inserted control character — zero-width, it
  never renders as a glyph; nothing else visible is added.

## Links

A Markdown link `[text](url)` — and a bare URL that the parser auto-links — becomes
a real, clickable Word hyperlink (an external relationship plus the built-in
`Hyperlink` character style, blue and underlined, created if the template lacks
it). The visible text keeps normal bidi splitting, so a URL stays LTR while a
Hebrew link label stays RTL. A labelled line like `**הורדה**: https://…` is driven
RTL by the Hebrew label (URLs are excluded from direction counting, see above)
while the URL itself is left-to-right and clickable.

## Images

A standalone image (`![alt text](path/to/img.png)` on its own line) is embedded
as a centered picture. If it is wider than the page's content area (page width
minus margins) it is scaled down proportionally to fit; smaller images keep their
native size. The **alt text becomes a centered caption** directly below the image,
using Word's `Caption` paragraph style (a smaller italic style, created
automatically if the template lacks one). Caption direction (RTL/LTR) follows the
same Hebrew/Latin majority rule as body text.

Image `src` paths are resolved relative to the input `.md` file's folder, so a
reference like `Assets/setup.png` works as long as the assets sit beside the
Markdown. If a file can't be found or read, a visible `[missing image: …]`
placeholder is inserted and the problem is reported to stderr instead of aborting.

## What gets converted

| Markdown | Word output |
| --- | --- |
| `#`/`##`/`###` headings | `Heading 1..9`, direction per content |
| **bold**, *italic* | bold / italic runs |
| `` `inline code` `` | `Code` character style (Normal + monospace), kept LTR, edges LRM-anchored |
| ```` ``` ```` fenced blocks | LTR monospace lines (`Code` style) |
| tables | RTL table (`bidiVisual`) with its own borders, bold header row |
| `>` blockquotes | built-in `Quote` style |
| `1.` / `-` lists | self-contained numbering (each list restarts at 1), one shared direction per list, bullets render in both directions |
| `[text](url)` / bare URLs | clickable Word hyperlink (`Hyperlink` style, blue + underlined) |
| `![alt](src)` images | centered picture (scaled to fit page width), alt text as a centered `Caption` below it |

## Limitations

- Nested / multi-level lists render at a single level.
- Raw inline HTML is skipped.
- Only standalone images (a `![alt](src)` on its own line) become centered image
  blocks with captions. An image sitting inline within a text paragraph falls back
  to rendering just its alt text. Image paths are resolved relative to the input
  `.md` file's folder; a missing/unreadable image is replaced with a visible
  `[missing image: …]` placeholder and reported to stderr.
- Tuned for Hebrew-dominant content; primarily-English docs still convert but the
  base direction is RTL.
