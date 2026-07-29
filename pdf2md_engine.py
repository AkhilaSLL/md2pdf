"""
pdf2md_engine.py — PDF → Markdown conversion for Markdown ⇄ PDF Studio.

Built on pymupdf4llm, which analyses font sizes to recover headings,
detects tables and (monospace) code blocks, and can extract embedded
images. This module adds:

  * images saved to a sibling  <name>_images/  folder and re-referenced
    with clean *relative* paths, so the markdown renders anywhere and
    stays lightweight for LLM consumption (no base64 blobs);
  * optional  <!-- page N -->  markers so page provenance survives;
  * an optional linked "Contents" section (GitHub-style anchors);
  * whitespace tidying to keep the token count down.

Usage (CLI):
    python pdf2md_engine.py input.pdf [-o output.md] [--no-images]
        [--no-tables] [--no-code] [--no-page-markers] [--toc] [--dpi N]
"""

from __future__ import annotations

import argparse
import os
import re
import sys
from dataclasses import dataclass


@dataclass
class Pdf2MdOptions:
    extract_images: bool = True     # save images + reference them
    image_dpi: int = 150            # render dpi for non-raster graphics
    detect_tables: bool = True      # markdown table reconstruction
    detect_code: bool = True        # monospace text -> fenced blocks
    page_markers: bool = True       # <!-- page N --> between pages
    linked_toc: bool = False        # prepend a linked Contents section
    skip_margins: bool = True       # ignore header/footer bands (page nos.)
    vector_graphics: bool = False   # also rasterise vector drawings


_IMG_NAME = re.compile(r"-(\d+)-(\d+)\.(png|jpg|jpeg)$", re.I)


def _rename_images(imgdir_abs: str) -> dict[str, str]:
    """pymupdf4llm names files '<input>.pdf-<page>-<idx>.png'; rename to
    short 'p<page>-<idx>.png' and return old->new basename map."""
    mapping: dict[str, str] = {}
    try:
        files = sorted(os.listdir(imgdir_abs))
    except OSError:
        return mapping
    for f in files:
        m = _IMG_NAME.search(f)
        if not m:
            continue
        new = f"p{int(m.group(1)) + 1}-{int(m.group(2)) + 1}.{m.group(3).lower()}"
        if new == f or new in mapping.values():
            continue
        try:
            os.rename(os.path.join(imgdir_abs, f),
                      os.path.join(imgdir_abs, new))
            mapping[f] = new
        except OSError:
            pass
    return mapping


# --------------------------------------------------------------- helpers ---

def _safe_stem(path: str) -> str:
    stem = os.path.splitext(os.path.basename(path))[0]
    return re.sub(r"[^\w\-]+", "_", stem).strip("_") or "document"


def _slugify(text: str, taken: dict[str, int]) -> str:
    """GitHub-style anchor slugs (mirrors md2pdf_engine)."""
    slug = text.strip().lower()
    slug = re.sub(r"[^\w\s-]", "", slug, flags=re.UNICODE)
    slug = re.sub(r"[\s]+", "-", slug).strip("-") or "section"
    n = taken.get(slug, 0)
    taken[slug] = n + 1
    return slug if n == 0 else f"{slug}-{n}"


_IMG_REF = re.compile(r"(!\[[^\]]*\]\()([^)]+)(\))")


def _prune_blank_images(imgdir_abs: str) -> set[str]:
    """Delete extracted 'images' that are just flat background fills
    (rasterised code-block chips, rules, decorations). Returns the
    basenames removed so their refs can be dropped too."""
    removed: set[str] = set()
    try:
        from PIL import Image, ImageStat
    except ImportError:
        return removed
    try:
        files = os.listdir(imgdir_abs)
    except OSError:
        return removed
    for f in files:
        p = os.path.join(imgdir_abs, f)
        try:
            with Image.open(p) as im:
                w, h = im.size
                std = ImageStat.Stat(im.convert("L")).stddev[0]
        except Exception:
            continue
        if min(w, h) < 40 or std < 8.0:      # sliver, or essentially flat
            try:
                os.unlink(p)
                removed.add(f)
            except OSError:
                pass
    return removed


_FENCE = re.compile(r"^```.*?$\n(.*?)^```[ \t]*$", re.M | re.S)


def _deblank_fences(md: str) -> str:
    """pymupdf4llm inserts a blank line after nearly every code line;
    when that pattern dominates a fence, drop the blanks."""
    def fix(m):
        body = m.group(1)
        lines = body.split("\n")
        nonblank = [l for l in lines if l.strip()]
        blanks = len(lines) - len(nonblank)
        if nonblank and blanks >= len(nonblank) - 1 and blanks > 1:
            body = "\n".join(nonblank) + "\n"
        else:
            body = re.sub(r"\n{2,}", "\n\n", body).strip("\n") + "\n"
        head = m.group(0).split("\n", 1)[0]
        return f"{head}\n{body}```"
    return _FENCE.sub(fix, md)


def _polish(md: str) -> str:
    """Whole-document cleanups: un-bold headings, compress dot-leader
    lines from the PDF's own printed contents page."""
    md = re.sub(r"^#{1,6} .*$",
                lambda m: m.group(0).replace("**", ""), md, flags=re.M)
    md = re.sub(r"^(.{2,120}?)\s*(?:\.\s*){4,}(\d{1,4})\s*$",
                r"- \1 · p\2", md, flags=re.M)
    return md


def _relink_images(md: str, imgdir_abs: str, imgdir_rel: str,
                   renames: dict[str, str], removed: set[str]) -> str:
    """Rewrite image refs written by pymupdf4llm (absolute or cwd-relative,
    slashes either way) to clean paths relative to the .md file; drop refs
    whose image was pruned."""
    try:
        present = set(os.listdir(imgdir_abs))
    except OSError:
        present = set()

    def sub(m):
        target = m.group(2).strip().strip("<>")
        name = os.path.basename(target.replace("\\", "/"))
        name = renames.get(name, name)
        if name in removed:
            return ""
        if name in present:
            rel = f"{imgdir_rel}/{name}".replace(" ", "%20")
            return f"{m.group(1)}{rel}{m.group(3)}"
        return m.group(0)

    return _IMG_REF.sub(sub, md)


def _tidy(md: str) -> str:
    """Trim trailing spaces and collapse 3+ blank lines (token diet)."""
    md = re.sub(r"[ \t]+$", "", md, flags=re.M)
    md = re.sub(r"\n{4,}", "\n\n\n", md)
    return md.strip() + "\n"


def _strip_page_number(text: str, pageno: int) -> str:
    """Remove a bare footer/header page number ('7', 'Page 7', '- 7 -')
    sitting alone at the very start or end of a page's text."""
    pat = re.compile(rf"^\s*(?:page\s+)?-?\s*{pageno}\s*-?\s*$", re.I)
    lines = text.split("\n")
    while lines and not lines[-1].strip():
        lines.pop()
    if lines and pat.match(lines[-1]):
        lines.pop()
    i = 0
    while i < len(lines) and not lines[i].strip():
        i += 1
    if i < len(lines) and pat.match(lines[i]):
        del lines[i]
    return "\n".join(lines)


_HEADING = re.compile(r"^(#{1,3})\s+(.+?)\s*#*\s*$", re.M)


def _build_toc(md: str) -> str | None:
    """A linked Contents section from H1–H3 headings, if there are
    enough of them to be worth it."""
    taken: dict[str, int] = {}
    rows = []
    for m in _HEADING.finditer(md):
        level = len(m.group(1))
        title = re.sub(r"[*_`]", "", m.group(2)).strip()
        if not title:
            continue
        slug = _slugify(title, taken)
        rows.append("  " * (level - 1) + f"- [{title}](#{slug})")
    if len(rows) < 3:
        return None
    return "## Contents\n\n" + "\n".join(rows) + "\n\n---\n"


# ----------------------------------------------------------- conversion ---

def convert(pdf_path: str, md_path: str | None = None,
            options: Pdf2MdOptions | None = None) -> dict:
    """Convert a PDF to Markdown. Returns stats:
    {"md_path", "pages", "images", "image_dir"(or None), "chars"}."""
    import pymupdf
    import pymupdf4llm

    opts = options or Pdf2MdOptions()
    if md_path is None:
        md_path = os.path.splitext(pdf_path)[0] + ".md"
    out_dir = os.path.dirname(os.path.abspath(md_path)) or "."
    os.makedirs(out_dir, exist_ok=True)

    imgdir_rel = _safe_stem(md_path) + "_images"
    imgdir_abs = os.path.join(out_dir, imgdir_rel)

    kwargs: dict = dict(
        page_chunks=True,
        show_progress=False,
        table_strategy="lines_strict" if opts.detect_tables else None,
        ignore_code=not opts.detect_code,
        ignore_graphics=not opts.vector_graphics,
        margins=(0, 36, 0, 36) if opts.skip_margins else (0, 0, 0, 0),
    )
    if opts.extract_images:
        os.makedirs(imgdir_abs, exist_ok=True)
        kwargs.update(
            write_images=True,
            image_path=imgdir_abs,
            image_format="png",
            dpi=opts.image_dpi,
            image_size_limit=0.02,   # skip specks < 2% of page area
        )

    doc = pymupdf.open(pdf_path)
    try:
        chunks = pymupdf4llm.to_markdown(doc, **kwargs)
    finally:
        doc.close()

    renames = _rename_images(imgdir_abs) if opts.extract_images else {}
    removed = _prune_blank_images(imgdir_abs) if opts.extract_images else set()

    pages = []
    for i, ch in enumerate(chunks):
        text = (ch.get("text") if isinstance(ch, dict) else str(ch)) or ""
        text = text.strip("\n")
        if opts.skip_margins:
            text = _strip_page_number(text, i + 1)
        if opts.extract_images:
            text = _relink_images(text, imgdir_abs, imgdir_rel,
                                  renames, removed)
        if opts.page_markers:
            text = f"<!-- page {i + 1} -->\n\n{text}"
        pages.append(text)

    md = "\n\n".join(p for p in pages if p.strip())
    md = _deblank_fences(md)
    md = _polish(md)
    md = _tidy(md)

    if opts.linked_toc:
        toc = _build_toc(md)
        if toc:
            md = toc + "\n" + md

    with open(md_path, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(md)

    # image accounting (and drop the folder if nothing landed in it)
    n_images = 0
    image_dir = None
    if opts.extract_images and os.path.isdir(imgdir_abs):
        files = [f for f in os.listdir(imgdir_abs)
                 if os.path.isfile(os.path.join(imgdir_abs, f))]
        n_images = len(files)
        if n_images:
            image_dir = imgdir_abs
        else:
            try:
                os.rmdir(imgdir_abs)
            except OSError:
                pass

    return {"md_path": md_path, "pages": len(pages),
            "images": n_images, "image_dir": image_dir, "chars": len(md)}


# ------------------------------------------------------------------- CLI ---

def main(argv=None):
    ap = argparse.ArgumentParser(description="Convert PDF to Markdown.")
    ap.add_argument("pdf")
    ap.add_argument("-o", "--output", default=None, help="output .md path")
    ap.add_argument("--no-images", action="store_true")
    ap.add_argument("--no-tables", action="store_true")
    ap.add_argument("--no-code", action="store_true")
    ap.add_argument("--no-page-markers", action="store_true")
    ap.add_argument("--toc", action="store_true",
                    help="prepend a linked Contents section")
    ap.add_argument("--dpi", type=int, default=150)
    ap.add_argument("--keep-margins", action="store_true",
                    help="do not strip header/footer bands")
    ap.add_argument("--vector-graphics", action="store_true",
                    help="also rasterise vector drawings as images")
    a = ap.parse_args(argv)

    opts = Pdf2MdOptions(
        extract_images=not a.no_images,
        detect_tables=not a.no_tables,
        detect_code=not a.no_code,
        page_markers=not a.no_page_markers,
        linked_toc=a.toc,
        image_dpi=a.dpi,
        skip_margins=not a.keep_margins,
        vector_graphics=a.vector_graphics,
    )
    stats = convert(a.pdf, a.output, opts)
    msg = f"wrote {stats['md_path']}  ({stats['pages']} pages"
    if stats["images"]:
        msg += f", {stats['images']} images -> {os.path.basename(stats['image_dir'])}/"
    print(msg + ")")


if __name__ == "__main__":
    main()
