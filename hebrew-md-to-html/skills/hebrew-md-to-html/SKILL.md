---
name: hebrew-md-to-html
description: >-
  Turn Hebrew (usually mixed Hebrew/English) Markdown into a paste-ready, RTL-aware HTML page with a
  Copy button, so the text can be pasted into Outlook, Gmail, Teams, Word, or any other rich-text
  editor with correct right-to-left alignment. Use whenever the user wants to send, paste, or copy
  Hebrew text written in the chat (an email, a message, a post, a reply) into another app, or
  complains that copying Hebrew from the chat breaks right-to-left alignment, shows Markdown
  symbols, or scrambles English words, numbers, and punctuation inside Hebrew sentences. For a Word
  file use hebrew-md-to-docx instead.
metadata:
  version: 1.0.0
  author: Assaf Tzur-El
---

# Hebrew Markdown → paste-ready HTML

Converts Hebrew (or mixed Hebrew/English) Markdown into a single self-contained HTML page. The page shows the content on a white sheet with a **Copy** button. Clicking it puts the content on the clipboard twice: as HTML, which rich-text editors such as Outlook, Gmail, Teams, and Word paste formatted and correctly aligned, and as plain text, which plain-text editors such as Notepad paste without Markdown symbols.

## Why this skill exists

Copying Hebrew from a chat into an email client gives the Markdown source, or formatted text with every paragraph aligned left. English terms, numbers, and punctuation inside Hebrew sentences drift to the wrong side, and numbered lists lose their direction. Going through Word (hebrew-md-to-docx) works but takes several steps. This skill produces the formatted, correctly directed text in one step.

## How to run it

The converter is `scripts/md_to_html_he.py` (Python 3; needs `mistune` and `beautifulsoup4`):

```bash
pip install mistune beautifulsoup4 --break-system-packages   # if not already present
python scripts/md_to_html_he.py INPUT.md -o OUTPUT.html --title "Page title"
```

Options:

- `-o, --output` – output path (default: `<input-stem>.html` next to the input).
- `--title` – the page title (default: the first Markdown heading, else the file name).

## Workflow

1. Write the content as Markdown to a file. For an email, put the subject on the first line as `**נושא:** …` so the user can copy it into the subject field separately, and put each signature line in its own paragraph (separated by blank lines).
2. Run the converter.
3. Deliver the page so the user can open it and click Copy:
   - If your runtime can publish a hosted HTML page (for example, a Claude artifact), publish the HTML file there. This is the smoothest path: the user opens the page and clicks Copy.
   - Otherwise, send the `.html` file. The user opens it in a browser and clicks Copy.
4. Tell the user in one line to click the Copy button (or select the text and press Ctrl+C) and paste into the target app with **Keep Source Formatting** or **Use Destination Styles** (in Word and Outlook: press Ctrl right after pasting, or right-click, to open the paste options). A plain paste applies the destination's paragraph direction and styles, which turns English paragraphs right-to-left and drops the formatting. After copying, the page shows the same hint in green next to the Copy button.

Don't paste the HTML source into the chat: the user needs the rendered page, not the markup.

## Direction rules

- Each paragraph, heading, table cell and quote gets `dir="rtl"` or `dir="ltr"` by the majority of Hebrew vs. English *words* in it; ties go right-to-left. Code, URLs and e-mail addresses are not counted, so a Hebrew label followed by a long URL stays right-to-left, and a line that is only a URL is left-to-right.
- All items in one list share a direction: right-to-left when at least half of the items contain Hebrew. Short items that are mostly a technical term (`Ctrl+F5`) stay aligned with their Hebrew siblings.
- Inside a right-to-left block, every run of English is wrapped in `<span dir="ltr">`, so it cannot be reordered. Tech tokens stay whole (`C++`, `.NET`, `file.txt`, `%d`, paths). Neighbouring English words are merged into one run when nothing Hebrew separates them, so a menu path such as `Debug ▸ Start Without Debugging` or a list of English names keeps its order. A sentence-ending period, colon or comma after an English word stays on the Hebrew side.
- Inside a left-to-right block, every Hebrew run is wrapped in `<span dir="rtl" lang="he">`, so a Hebrew word inside an English sentence stays in place.
- Inline code and fenced code blocks are left-to-right and monospace.
- Direction is written both as a `dir` attribute and as CSS (`direction`, `unicode-bidi:embed`), the way Word writes it in its own HTML.

## Email-safe markup

Outlook and Word turn some standard HTML into something else when pasted, so the converter avoids it:

- **Strikethrough** is a `<span>` with `text-decoration:line-through`, never `<del>` or `<s>`: Word and Outlook read `<del>` as a tracked deletion, and Outlook drops the text.
- **Fenced code blocks** are one shaded paragraph with line breaks, not `<pre>`, whose shading Outlook drops. Line breaks become `<br>` and spaces become non-breaking, so indentation survives.
- **Inline code** is a `<span>` with a monospace font and `background-color`, not `<code>`.
- **Quotes** are one-cell tables with a side border, not `<blockquote>`, whose border Outlook drops.
- All formatting is written as inline styles, because email clients drop page stylesheets.

## What gets converted

Headings, paragraphs, **bold**, *italic*, ~~strikethrough~~, inline code, fenced code blocks, ordered and unordered lists, tables, blockquotes, links, and bare URLs (which become clickable links).

## Limitations

- The direction rules are heuristics. A paragraph that mixes Hebrew and English almost evenly may get the wrong base direction; rephrasing it so one language clearly dominates fixes it.
- How each target app treats pasted HTML varies. The markup targets Outlook and Word; check any other target app once with a short sample.
- Images are not embedded. Use hebrew-md-to-docx for documents with images.
