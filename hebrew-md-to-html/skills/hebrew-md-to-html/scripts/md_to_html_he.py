#!/usr/bin/env python3
"""Convert Hebrew (or mixed Hebrew/English) Markdown into a paste-ready, RTL-aware HTML page.

The page shows the content in a white "sheet" with a Copy button. The button puts the sheet on
the clipboard as both HTML and plain text, so a rich-text editor (Outlook, Word, Teams, Gmail)
pastes it formatted and correctly aligned, while a plain-text editor gets clean text.

Direction rules:
- Each block (paragraph, heading, table cell, quote) gets dir="rtl" or dir="ltr" by the majority of
  Hebrew vs. Latin *words* in it; ties go RTL. Code, URLs, and e-mail addresses are not counted.
- All items of one list share one direction: RTL when at least half of the items contain Hebrew,
  so short items that are mostly a technical term still align with their Hebrew siblings.
- Inside an RTL block, every run of English (words, numbers attached to letters, tech tokens like
  C++, .NET, file.txt, paths) is wrapped in <span dir="ltr"> so it cannot be reordered. Sentence
  punctuation at the end of a run stays on the Hebrew side.
- Inside an LTR block, every Hebrew run is wrapped in <span dir="rtl" lang="he">.
- Inline code is LTR; fenced code blocks are LTR, monospace.
- Direction is written both as a dir attribute and as CSS (direction, unicode-bidi), as Word does in
  its own HTML export, and the markup avoids tags that Outlook and Word mistranslate: <del> (read as
  a tracked deletion), <pre> and <blockquote> (shading and borders dropped).
- Paste with "Keep Source Formatting" or "Use Destination Styles": a plain paste in Word and Outlook
  applies the destination's paragraph direction and styles, which turns English paragraphs
  right-to-left. After copying, the page shows this hint in green next to the Copy button.
"""
import argparse
import html
import re
import sys
from pathlib import Path

import mistune
from bs4 import BeautifulSoup, NavigableString

HEB_CHARS = re.compile(r'[\u0590-\u05FF\uFB1D-\uFB4F]')   # spelled as escapes on purpose: a literal U+FB1D decomposes under normalisation
LATIN_CHARS = re.compile(r'[A-Za-z]')
URLISH = re.compile(r'(https?://\S+|www\.\S+|\S+@\S+\.\S+)')
# A Latin token: optional leading dot or percent (".NET", "%d"), a Latin letter/digit, then letters,
# digits and word-joining symbols. Neighbouring tokens are merged by latin_spans().
LATIN_RUN = re.compile(
    r'[.%]?[A-Za-z0-9][A-Za-z0-9.+#/\\_:%\-]*')
TRAILING_PUNCT = re.compile(r'[.:,;\-]+$')

BLOCK_TAGS = ['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'th', 'td', 'blockquote']
FONT = 'Aptos, Calibri, Arial, sans-serif'
MONO = "Consolas, 'Courier New', monospace"


def text_of(tag):
    """Visible text of a tag, without inline/fenced code (code never decides direction)."""
    return ' '.join(t for t in tag.find_all(string=True) if t.find_parent(['code', 'pre']) is None)


def direction(text):
    """Majority of whitespace-separated words by script; URLs and e-mail addresses are not counted."""
    heb = lat = 0
    for word in URLISH.sub(' ', text).split():
        h, l = len(HEB_CHARS.findall(word)), len(LATIN_CHARS.findall(word))
        if h > l:
            heb += 1
        elif l > h:
            lat += 1
    if heb == lat == 0:  # e.g. a line that is only a URL
        return 'ltr' if LATIN_CHARS.search(text) else 'rtl'
    return 'ltr' if lat > heb else 'rtl'


def list_direction(lst):
    """A list is RTL when at least half of its items contain Hebrew."""
    items = lst.find_all('li', recursive=False)
    hebrew_items = sum(1 for li in items if HEB_CHARS.search(text_of(li)))
    return 'rtl' if items and hebrew_items * 2 >= len(items) else direction(text_of(lst))


def latin_spans(text):
    """(start, end) of each English run, merged across gaps that hold no Hebrew.

    Merging mirrors the Unicode bidi rule that a neutral between two LTR characters is LTR, so a
    menu path like "Debug > Start Without Debugging" or "Visual Studio, Rider" stays one unit.
    """
    spans = [[m.start(), m.end()] for m in LATIN_RUN.finditer(text)]
    merged = []
    for st, en in spans:
        if merged and not HEB_CHARS.search(text[merged[-1][1]:st]) and '\n' not in text[merged[-1][1]:st]:
            merged[-1][1] = en
        else:
            merged.append([st, en])
    return merged


def wrap_latin(soup, node):
    """Wrap each English run of a text node in <span dir="ltr">."""
    text = str(node)
    pieces, last = [], 0
    for start, end in latin_spans(text):
        tok = text[start:end]
        trail = TRAILING_PUNCT.search(tok)
        if trail:
            end -= len(trail.group(0))
            tok = tok[:-len(trail.group(0))]
        if not tok or not LATIN_CHARS.search(tok):
            continue  # plain numbers render correctly on their own
        pieces.append(NavigableString(text[last:start]))
        span = soup.new_tag('span', dir='ltr')
        span.string = tok
        pieces.append(span)
        last = end
    if not pieces:
        return
    pieces.append(NavigableString(text[last:]))
    for p in pieces:
        node.insert_before(p)
    node.extract()


HEB_RUN = re.compile(HEB_CHARS.pattern + r"+(?:[\s\-'\"׳״]+" + HEB_CHARS.pattern + r'+)*')


def wrap_hebrew(soup, node):
    """Wrap each Hebrew run of a text node inside an LTR block in <span dir="rtl" lang="he">."""
    text = str(node)
    pieces, last = [], 0
    for m in HEB_RUN.finditer(text):
        pieces.append(NavigableString(text[last:m.start()]))
        span = soup.new_tag('span', dir='rtl', lang='he')
        span.string = m.group(0)
        pieces.append(span)
        last = m.end()
    if not pieces:
        return
    pieces.append(NavigableString(text[last:]))
    for p in pieces:
        node.insert_before(p)
    node.extract()


def bidi(d):
    """Direction as CSS too: Word and Outlook read this form, as in Word's own HTML export."""
    return f"direction:{d};unicode-bidi:embed;text-align:{'right' if d == 'rtl' else 'left'}"


def style(tag, extra=''):
    tag['style'] = f"{bidi(tag.get('dir', 'rtl'))};{extra}"


def code_block_html(pre):
    """A fenced code block as one shaded paragraph with line breaks, as Word writes shading."""
    lines = pre.get_text().rstrip('\n').split('\n')
    body = '<br>'.join(html.escape(ln).replace(' ', '&nbsp;') or '&nbsp;' for ln in lines)
    return (f'<p dir="ltr" style="{bidi("ltr")};margin:0 0 8pt 0;padding:6pt 8pt;'
            f'background:#F2F2F2;font-family:{MONO};font-size:10pt">{body}</p>')


def convert(md_text):
    markdown = mistune.create_markdown(plugins=['strikethrough', 'table', 'url'])
    soup = BeautifulSoup(markdown(md_text), 'html.parser')

    # 1. Directions. Code never decides direction and is always LTR.
    for code in soup.find_all('code'):
        code['dir'] = 'ltr'
    for lst in soup.find_all(['ol', 'ul']):
        d = list_direction(lst)
        lst['dir'] = d
        pad = 'padding-right:24pt;padding-left:0' if d == 'rtl' else 'padding-left:24pt;padding-right:0'
        lst['style'] = f'direction:{d};margin:0 0 8pt 0;{pad}'
        for li in lst.find_all('li', recursive=False):
            li['dir'] = d
            style(li, 'margin:0 0 4pt 0')
    for blk in soup.find_all(BLOCK_TAGS):
        if blk.name == 'p' and blk.find_parent('li') is not None:
            blk['dir'] = blk.find_parent('li')['dir']
        else:
            blk['dir'] = direction(text_of(blk))
        if blk.name.startswith('h'):
            style(blk, 'margin:12pt 0 6pt 0')
        elif blk.name in ('th', 'td'):
            style(blk, 'padding:4pt 6pt;border:1px solid #bbb')
        else:
            style(blk, 'margin:0 0 8pt 0')
    for table in soup.find_all('table'):
        table['dir'] = direction(text_of(table))
        table['style'] = f"direction:{table['dir']};border-collapse:collapse;margin:0 0 8pt 0"

    # 2. Keep the other script's runs in place: English inside RTL blocks, Hebrew inside LTR ones.
    for t in list(soup.find_all(string=True)):
        if not t.strip() or t.find_parent(['code', 'pre']) is not None:
            continue
        holder = t.find_parent(attrs={'dir': True})
        if holder is None or holder.name == 'span':
            continue
        if holder.get('dir') == 'rtl' and LATIN_CHARS.search(t):
            a = t.find_parent('a')
            if a is not None and not HEB_CHARS.search(a.get_text()):
                a['dir'] = 'ltr'
                continue
            wrap_latin(soup, t)
        elif holder.get('dir') == 'ltr' and HEB_CHARS.search(t):
            wrap_hebrew(soup, t)

    # 3. Email-safe markup. Outlook and Word read <del>/<ins> as tracked changes, drop a <pre>'s
    #    shading and a <blockquote>'s border, so use plain spans, a shaded paragraph for code, and a
    #    one-cell table for quotes instead.
    for tag in soup.find_all(['del', 's', 'strike']):
        tag.name = 'span'
        tag.attrs = {'style': 'text-decoration:line-through'}
    for tag in soup.find_all('ins'):
        tag.name = 'span'
        tag.attrs = {'style': 'text-decoration:underline'}
    for pre in soup.find_all('pre'):
        pre.replace_with(BeautifulSoup(code_block_html(pre), 'html.parser'))
    for code in soup.find_all('code'):
        code.name = 'span'
        code.attrs = {'dir': 'ltr',
                      'style': f'font-family:{MONO};font-size:10pt;background-color:#f2f2f2'}
    for quote in soup.find_all('blockquote'):
        d = quote['dir']
        side = 'right' if d == 'rtl' else 'left'
        table = soup.new_tag('table', dir=d, cellpadding='0', cellspacing='0',
                             style=f'direction:{d};border-collapse:collapse;margin:0 0 8pt 0')
        tr = soup.new_tag('tr')
        td = soup.new_tag('td', dir=d, style=f'{bidi(d)};border-{side}:3px solid #ccc;'
                                             f'padding-{side}:8pt;color:#444')
        for child in list(quote.contents):
            td.append(child.extract())
        tr.append(td)
        table.append(tr)
        quote.replace_with(table)
    return str(soup)


PAGE = '''<!DOCTYPE html>
<html lang="he"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title}</title>
<style>
 body {{ margin:0; background:#f4f4f2; color:#222; font-family:{font}; }}
 .bar {{ max-width:780px; margin:24px auto 8px; padding:0 16px; display:flex; gap:12px;
        align-items:center; direction:rtl; font-size:14px; }}
 button {{ font:inherit; padding:8px 16px; border:1px solid #888; border-radius:6px;
          background:#fff; cursor:pointer; }}
 #status {{ color:#555; }}
 #status.ok {{ color:#1a7f37; font-weight:bold; }}
 #sheet {{ max-width:780px; margin:0 auto 40px; padding:28px 32px; background:#fff;
          border:1px solid #ddd; border-radius:8px; box-sizing:border-box; }}
</style></head><body>
<div class="bar"><button id="copy" type="button">העתקה / Copy</button>
<span id="status">אפשר גם לסמן את הטקסט ולהעתיק עם Ctrl+C.</span></div>
<div id="sheet" style="font-family:{font};font-size:11pt;line-height:1.4">
{body}
</div>
<script>
const st = document.getElementById('status');
const copied = () => {{
  st.textContent = 'הועתק. בהדבקה ב-Word או ב-Outlook יש לבחור Keep Source Formatting או Use Destination Styles.';
  st.className = 'ok';
}};
// Also catches a manual Ctrl+C, so the paste hint shows whichever way the user copies.
document.addEventListener('copy', copied);
document.getElementById('copy').addEventListener('click', async () => {{
  const el = document.getElementById('sheet');
  try {{
    await navigator.clipboard.write([new ClipboardItem({{
      'text/html': new Blob([el.outerHTML], {{ type: 'text/html' }}),
      'text/plain': new Blob([el.innerText], {{ type: 'text/plain' }})
    }})]);
    copied();
  }} catch (e) {{
    try {{
      const r = document.createRange(); r.selectNodeContents(el);
      const s = window.getSelection(); s.removeAllRanges(); s.addRange(r);
      if (!document.execCommand('copy')) st.textContent = 'הטקסט מסומן – יש ללחוץ Ctrl+C.';
    }} catch (e2) {{
      st.textContent = 'יש לסמן את הטקסט ולהעתיק עם Ctrl+C.';
    }}
  }}
}});
</script>
</body></html>
'''


def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument('input', help='Markdown file')
    ap.add_argument('-o', '--output', help='output .html (default: <input-stem>.html next to input)')
    ap.add_argument('--title', help='page title (default: first heading, else the file name)')
    args = ap.parse_args()

    src = Path(args.input)
    md_text = src.read_text(encoding='utf-8')
    body = convert(md_text)
    title = args.title
    if not title:
        m = re.search(r'^#{1,6}\s+(.+)$', md_text, re.M)
        title = m.group(1).strip() if m else src.stem
    out = Path(args.output) if args.output else src.with_suffix('.html')
    out.write_text(PAGE.format(title=html.escape(title), font=FONT, body=body), encoding='utf-8')
    print(f'Wrote {out}')


if __name__ == '__main__':
    sys.exit(main())
