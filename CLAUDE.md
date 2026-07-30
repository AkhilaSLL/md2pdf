# CLAUDE.md

Markdown ⇄ PDF Studio — a two-way Markdown/PDF converter shipped as a single
Windows exe. This directory is the repository root.

## Layout

```
md2pdf/
  md2pdf_app.py           CustomTkinter GUI, both directions
  md2pdf_engine.py        Markdown → PDF (reportlab + markdown-it-py + pygments)
  pdf2md_engine.py        PDF → Markdown (pymupdf4llm)
  Markdown2PDF.spec       PyInstaller spec (onefile, windowed)
  make_icon.py            regenerates assets/icon.ico
  assets/icon.ico         app/exe icon
  requirements.txt
  build/, dist/           PyInstaller output — generated, git-ignored
```

## Commands

Run everything from the repo root. `py` is the Windows launcher (Python 3.14).

```powershell
py -m pip install -r requirements.txt
py md2pdf_app.py                                   # GUI
py md2pdf_engine.py guide.md -o out.pdf --toc      # engine CLI
py pdf2md_engine.py out.pdf -o out.md
py -m PyInstaller --noconfirm Markdown2PDF.spec    # -> dist\Markdown2PDF.exe
```

Both engines are importable modules *and* CLIs (`if __name__ == "__main__"`),
so the fastest way to test a rendering change is the engine CLI — no GUI.

There is no test suite. Verify rendering changes by converting a document that
exercises the features and inspecting the PDF — PyMuPDF comes in via
`pymupdf4llm`, so pages can be rasterised to PNG and read back. A good
stress document needs `\pagebreak` markers, SVGs (inline and linked), several
tables and some long fenced blocks. The one this was built against lives
outside the repo at `..\MD_To_PDF_Project\cr1140-qt-guide.md`; it is not
committed because it contains an internal IP, an SSH password and device
serial numbers.

## md2pdf_engine.py — how it fits together

`convert_text()` is the entry point: markdown-it parses to a `SyntaxTreeNode`,
`_Renderer` walks it into a reportlab *story*, `GuideDocTemplate` builds the PDF.

- `_Renderer._blocks()` dispatches on node type to `_heading` / `_paragraph` /
  `_code_block` / `_list` / `_blockquote` / `_table` / `_image_block`;
  `_inline()` turns inline nodes into reportlab `Paragraph` markup.
- `_build_styles()` is the single place styling lives. Everything scales off
  `opts.base_size` (`self.S`) — sizes are multiples of it, never hard pt values.
  Heading sizes come from the `H_SCALE` dict, colours from the palette
  constants at the top of the file.
- `GuideDocTemplate.afterFlowable()` reads `flowable._bookmark_key` (stamped by
  `_heading`) to emit PDF outline entries and TOC notifications. Headings are
  matched by *style name* via `self.heading_levels` (`MDHeading1`…`MDHeading6`),
  so renaming a heading style silently kills bookmarks.
- Anchors are two-pass: `prescan_slugs()` walks all headings first so that
  `[text](#anchor)` links rendered before their target still resolve.
- `_inline()` is an allowlist — inline HTML it does not recognise is dropped
  rather than passed through, because reportlab's mini-XML parser raises on
  unknown tags and kills the whole build.
- `normalise_inline_svg()` runs on the raw markdown *before* parsing: a raw
  `<svg>` block would otherwise be cut in half by CommonMark's rule that an
  HTML block ends at the first blank line. It only triggers on a line that
  *opens* the element and skips fenced code — a table row or sentence merely
  mentioning `<svg>` must not start a diagram, and an SVG shown as a code
  sample must stay one. It always runs; `opts.inline_svg` decides only
  whether the resulting block is drawn.

### Conventions that matter here

- **Everything is pure Python.** The exe is built with PyInstaller and must not
  need native libraries. This rules out cairosvg/rsvg (use `svglib`), and any
  image backend beyond Pillow.
- **Never let a bad input abort the build.** Missing/unreadable images,
  unknown lexers, absent fonts and unparseable SVGs all degrade to a caption or
  a fallback, never an exception.
- **Fonts are resolved, not assumed.** `resolve_body_font()` walks a fallback
  chain and ends at built-in Type1. Glyphs the chosen font lacks are routed
  through a symbol fallback font by `markup_symbols()` — use it (or
  `plain_symbols()` outside Paragraph markup) for any author-supplied text, or
  emoji render as boxes.
- **Finding a font is three lookups, not one** (`_register_ttf_family`):
  filename candidates, then the Windows font registry by display name, then
  `_cloud_font_map()`. That last one matters — Microsoft 365 keeps on-demand
  fonts in `%LOCALAPPDATA%\Microsoft\FontCache\<n>\CloudFonts\<Family>\` under
  numeric filenames and never registers them with Windows, so they are
  invisible to the first two. **Aptos is one of these**: present in Word,
  undiscoverable by filename or registry. Identifying them means parsing each
  file's name table with `TTFontFile`. Don't "simplify" that away — if a font
  the user can see in Word goes missing from the dropdown, this is why.
- **Comment density is low and explanatory** — short `# ----- section -----`
  banners, and comments that say *why*, not *what*.

## md2pdf_app.py — drag and drop

`enable_file_drop()` is the whole mechanism. Tk has no drag-and-drop, and the
usual answer (`tkinterdnd2`) ships a native tkdnd Tcl extension — banned by the
pure-Python rule above. So it goes through the Win32 shell instead:
`DragAcceptFiles` on the toplevel plus a `SetWindowLongPtrW` subclass of its
window procedure that handles `WM_DROPFILES`, all via `ctypes`. Nothing extra
to bundle, and `enable_file_drop()` returns False rather than raising anywhere
it can't be set up, so non-Windows keeps working with Browse alone.

Two things there are load-bearing.

**The callback must not touch Tk — not even `after()`.** Tk's mainloop
releases the GIL while it waits in `Tcl_DoOneEvent`, and a real drop is
*posted*, so it is dispatched from inside that wait and reaches the callback
with no thread state. Plain Python is fine (ctypes reacquires the GIL on
entry), but re-entering Tcl from there kills the interpreter on the first
drop: `Fatal Python error: PyEval_RestoreThread: … the GIL is released`. So
the callback only does `self._drops.put(paths)`, and `_poll()` — already
ticking every 120 ms in proper Tk context — drains it.

The trap is that this is invisible to the obvious test: a synthetic
`SendMessage` from Python arrives with the GIL held and passes happily. Only a
**posted** message dispatched from inside a running `mainloop` reproduces it:
`PostMessageW(hwnd, WM_DROPFILES, hdrop, 0)` from an `after()` timer, with the
app in `mainloop()` — building the `HDROP` by hand in global memory.

**The boxed `WNDPROC` is kept alive on `window._drop_proc`.** Windows holds
the only other reference, so without it the callback is collected and the next
message takes the process down. Nothing inside the callback may raise either —
it runs where an exception has nowhere to go.

`ctypes.wintypes` is imported *inside* the function; importing it at module
level raises on non-Windows.

## pdf2md_engine.py

`pymupdf4llm` is pinned to **0.0.27** in requirements.txt on purpose — 1.x
switches parsers, OCRs pages and drops body text. Do not bump it.

## Packaging

`Markdown2PDF.spec` uses `collect_all('customtkinter')` (its JSON themes are
missed otherwise and the exe exits instantly on launch). Any new dependency
with data files or lazily-imported submodules needs its own `collect_all` /
`hiddenimports` entry there, and a frozen-build smoke test — the source run
proves nothing about the exe.
