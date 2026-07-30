# Markdown ⇄ PDF Studio

A light-themed desktop app that converts in **both directions**:

- **Markdown → PDF** — clean, print-ready A4 PDFs with MS-Word-style margins,
  so the output reads like a properly laid-out Word document.
- **PDF → Markdown** — lean, LLM-friendly markdown recovered from any
  text-based PDF, with images extracted alongside. Convert a PDF once, then
  paste/attach the `.md` — it costs a fraction of the tokens of re-parsing
  the PDF every time.

Switch direction with the toggle at the top of the window; the options adapt
to whichever way you're converting.

**Drag and drop** — drop a `.md` or `.pdf` anywhere on the window instead of
browsing for it. The file picks the direction: drop a PDF and the app flips to
PDF → Markdown, drop markdown and it flips back (and out of *Paste text* mode
if you were in it). The output path follows the dropped file. Drop several and
the first usable one wins; anything that isn't markdown or PDF is refused
without disturbing the current selection. Windows only — elsewhere **Browse**
is still there.

## Markdown → PDF

- **Headings** — styled and scaled relative to your chosen base size, with
  GitHub-style rules under H1/H2, and orphan control. Each level gets its own
  ink, darkest at H1 and lifting as it descends, so the hierarchy reads even
  where two levels are nearly the same size. Override the ramp with
  `--heading-colors`, `ConvertOptions.heading_colors`, or front matter:

  ```markdown
  ---
  title: My guide
  heading_colors: #0B1220, #16223A, #26364F
  h4_color: teal
  ---
  ```

  Front matter is stripped from the output either way. Colours may be `#rgb`,
  `#rrggbb` or a named colour; anything else is ignored rather than guessed at.
- **PDF bookmarks** — the full heading hierarchy appears in the PDF viewer's
  outline panel.
- **Links** — external links are clickable; internal `[text](#anchor)` links
  jump to the matching heading (GitHub-style anchors).
- **Images** — `![alt](path.png)` embeds the image, sized to fit the page
  (tall images are capped to the frame, never overflow). Relative paths
  resolve next to the `.md` file; `https://` images are downloaded;
  HTML `<img src width height>` tags are honoured, including inline and in
  table cells. Alt text becomes a small caption, which stays on the same page
  as its image.
- **SVG** — drawn as **vector**, not rasterised, so diagrams stay sharp at any
  zoom and cost far fewer bytes than a PNG. Scaled to the column width by
  default, aspect ratio preserved. Done with `svglib`, which is pure Python
  and so survives the PyInstaller build (cairosvg and rsvg would drag in
  native libraries). Two ways to write one — pick with **Diagrams (SVG)** in
  the app, or `--no-inline-svg` on the CLI:

  | | Inline `<svg>` *(default)* | Linked `.svg` files |
  |---|---|---|
  | How | `<svg …>…</svg>` in the `.md` | `![alt](diagrams/x.svg)` |
  | Needs | the whole element in the file | the `.md` to travel with its image folder |
  | Caption | a `<title>` first inside `<svg>` | the alt text |
  | Preview | VS Code yes, **GitHub strips it** | anywhere, GitHub included |
  | Good for | one self-contained document | reuse, and repo-hosted docs |

  Blank lines *inside* an inline `<svg>` are fine — the engine rewrites the
  element so CommonMark keeps it in one piece. An SVG shown inside a fenced
  code block stays a code sample. Explicit `width`/`height` on an HTML `<img>`
  tag are honoured for the linked form.

  *Limit:* an SVG **inline in a sentence or a table cell** falls back to an
  `[alt]` label — reportlab's inline `<img>` markup takes rasters only, and
  rasterising would need exactly the native backend this avoids. Put the SVG
  on a line of its own to get the vector figure.
- **Symbols & emoji** — ✅ ❌ ⚠️ → and friends render through a symbol
  fallback font (Segoe UI Symbol on Windows) instead of showing as boxes;
  anything even that font lacks degrades to a lookalike (✅→✓).
- **Code** — fenced blocks get a light background, border, syntax colouring
  (Pygments), and long lines wrap. Inline code gets a grey chip.
- **Tables** — full-width, shaded header row, header repeats across pages.
- **Page breaks** — `\pagebreak`, `\newpage`, or `<!-- pagebreak -->` on its
  own line.
- **Pagination** — tables and code blocks up to about half a page move whole
  to the next page rather than split; bigger ones do split, but never leaving
  fewer than two rows (or two lines) on either side of the break, and a table
  header never ends up stranded. A sentence that introduces a code block or
  table travels with it. Paragraphs keep two lines on each side of a break.
- **Inline colour** — `<font color="…">` and `<span style="color: …">` are
  passed through; unknown colours and unbalanced tags are dropped rather than
  risking a failed build. Other inline HTML is still ignored.
- **Optional extras** — clickable table of contents, footer page numbers.
- Font family, base size 8–16 pt, Word margin presets. Defaults are **Aptos**
  at 11 pt with **Narrow (1.27 cm)** margins.

  The dropdown lists only families actually present on the machine. They are
  found three ways, in order:

  1. **By filename**, across the Windows font folder, per-user installs, and
     Office's own private font folder.
  2. **By display name in the Windows font registry** — so families whose
     files are named unpredictably still resolve.
  3. **By reading the name table of Office "cloud fonts"**. Microsoft 365
     downloads fonts on demand into
     `%LOCALAPPDATA%\Microsoft\FontCache\<n>\CloudFonts\<Family>\`, where the
     filenames are opaque numeric ids (`30153066857.ttf`) and nothing is
     registered with Windows. **Aptos lives here** on a typical Microsoft 365
     install — it is in Word but invisible to both lookups above, so each
     file's embedded family/style is read instead.

  If a family genuinely isn't installed, it is simply absent from the list and
  the engine falls back: Aptos → Calibri → Segoe UI → Arial.

## Getting an AI to write markdown for this app

Paste the block below into Claude (or any assistant) along with what you want
written. It describes only what this converter actually supports, so the
output needs no cleanup before conversion.

````markdown
Write the document as CommonMark markdown for a Markdown→PDF converter
(reportlab + markdown-it). Stick to what it renders:

SUPPORTED
- `#`…`######` headings. Each level gets its own colour automatically — don't
  fake hierarchy with bold text.
- Paragraphs, `**bold**`, `*italic*`, `~~strike~~`, `` `inline code` ``.
- Bullet and numbered lists, including nested ones.
- Tables with a header row (`| a | b |` / `|---|---|`). Header repeats if the
  table spans pages.
- Fenced code blocks. ALWAYS tag the language (```python, ```powershell,
  ```json) — that drives syntax colouring.
- Blockquotes with `>`.
- Links: `[text](https://…)` and internal `[text](#slug)`, where the slug is
  the heading lowercased with spaces as hyphens and punctuation removed.
- Images: `![alt](path.png)`. Alt text becomes a caption.
- Diagrams: write SVG straight into the document —
  `<svg xmlns="http://www.w3.org/2000/svg" width="900" height="400">…</svg>`
  with the opening tag starting its own line and a blank line before it.
  Put a `<title>` first inside `<svg>` to caption it. Blank lines inside the
  element are fine. Design for a ~900 px wide viewBox; it is scaled to the
  text column and stays vector.
- A line containing only `\pagebreak` forces a new page. Use it to start each
  major section on a fresh page.
- Inline colour, if needed: `<font color="#B45309">text</font>`. Hex or a
  standard colour name only.
- Emoji and symbols (✅ ❌ ⚠ →) are fine; they fall back to lookalikes.
- `<br>` for a hard line break inside a paragraph.
- Optional front matter, first thing in the file:
  ---
  title: Document title
  heading_colors: #0B1220, #16223A, #26364F
  ---

AVOID
- Any other raw HTML (`<div>`, `<details>`, HTML tables) — dropped silently.
- Footnotes, definition lists, task lists (`- [ ]`), math/LaTeX, Mermaid or
  other diagram-code fences, admonition blocks (`:::note`), and HTML comments
  used as directives. A Mermaid fence renders as a code listing, not a
  diagram — use inline SVG instead.
- Nested tables.
- SVG mid-sentence or inside a table cell: it degrades to a text label. Put
  diagrams on a line of their own. (Raster images inline are fine.)

LAYOUT THAT PAGINATES WELL
- Introduce every code block or table with a short sentence ending in `:` —
  it is kept on the same page as the block that follows.
- Prefer code blocks under ~30 lines; longer ones split across pages.
- Keep table cells short; wide tables get cramped at A4 width.
````

## PDF → Markdown

Built on `pymupdf4llm` (pinned to 0.0.27 — later 1.x builds switch to a
different parser that OCRs pages and drops body text). Recovers:

- **Headings** — inferred from font sizes, mapped to `#`/`##`/`###`.
- **Tables** — reconstructed as markdown tables.
- **Code blocks** — monospace text becomes fenced blocks (extraction
  artifacts like a blank line after every code line are cleaned up).
- **Images** — saved as PNGs into a `<name>_images/` folder next to the
  `.md`, referenced with clean relative paths (`![](guide_images/p3-1.png)`),
  so the markdown previews correctly in VS Code/GitHub and stays small for
  LLM use. Flat decorative rectangles are pruned automatically.
- **Page markers** — optional `<!-- page N -->` comments so you (or an LLM)
  can still cite page numbers. These deliberately do *not* trigger page
  breaks if you convert the markdown back to PDF.
- **Header/footer stripping** — bare page numbers at page edges are removed.
- **Linked contents** — optionally prepend a `## Contents` section with
  GitHub-style anchor links; dot-leader lines from the PDF's own printed
  TOC are compressed to compact bullets either way.

Known limits: hyperlink annotations aren't recovered (link text comes
through as plain text), visually wrapped code lines stay wrapped, and
scanned/image-only PDFs need OCR, which isn't included.

## Run from source

```powershell
py -m pip install -r requirements.txt
py md2pdf_app.py
```

## Build the .exe

```powershell
py -m pip install -r requirements.txt pyinstaller
py -m PyInstaller --noconfirm Markdown2PDF.spec
```

The finished exe lands in `dist\Markdown2PDF.exe`. Notes:

- Build from the spec, not a bare command line — it carries the bundling
  fixes below, and a build without them fails only at runtime.
- `collect_all('customtkinter')` is required — CustomTkinter ships JSON theme
  files PyInstaller misses otherwise (the exe opens then instantly closes
  without it).
- `collect_all('svglib')` plus explicit `cssselect2` / `tinycss2` hidden
  imports keep the SVG path alive. svglib only builds its CSS parser when an
  SVG actually has a `<style>` block, so PyInstaller cannot see those imports
  and SVGs would silently stop rendering in the exe while working from source.
- PyMuPDF is picked up by PyInstaller's own hook for the PDF → Markdown
  direction. Expect a noticeably bigger exe (~50–60 MB) — that's PyMuPDF.
- The spec bundles `assets/` for the window/taskbar icon.
- First launch of a `--onefile` build takes a few seconds while it unpacks.
  Drop `--onefile` for a faster-starting folder build.
- If SmartScreen complains about an unsigned exe: *More info → Run anyway*.

## Command-line use (no GUI)

Both engines double as CLIs, handy for batch conversion:

```powershell
py md2pdf_engine.py guide.md -o guide.pdf --font Aptos --size 11 --toc
py md2pdf_engine.py guide.md --heading-colors "#0B1220,#16223A,#26364F"
py pdf2md_engine.py guide.pdf -o guide.md
py pdf2md_engine.py guide.pdf --no-images --no-page-markers   # text only
py pdf2md_engine.py drawing.pdf --vector-graphics             # keep vector art
```

## Files

| File | Purpose |
|---|---|
| `md2pdf_app.py` | The GUI (CustomTkinter, light theme, direction toggle) |
| `md2pdf_engine.py` | Markdown → PDF engine (ReportLab + markdown-it + Pygments) |
| `pdf2md_engine.py` | PDF → Markdown engine (pymupdf4llm + cleanup passes) |
| `assets/icon.ico` | App/exe icon (multi-resolution) |
| `make_icon.py` | Regenerates the icon if you want to tweak it |
| `requirements.txt` | Python dependencies |
