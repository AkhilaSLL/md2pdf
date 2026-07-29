"""
md2pdf_engine.py — Markdown → PDF conversion engine.

Produces A4 PDFs with MS-Word-like "Normal" layout (2.54 cm margins),
PDF outline bookmarks from headings, clickable external + internal links,
syntax-highlighted code blocks, tables, blockquotes, images, lists,
optional table of contents and page numbers.

Pure-Python stack (reportlab / markdown-it-py / pygments) so it packages
cleanly into a single .exe with PyInstaller.

Can also be used from the command line:
    python md2pdf_engine.py guide.md -o guide.pdf --font Aptos --size 11 --toc
"""

from __future__ import annotations

import html
import io
import os
import re
import sys
import tempfile
import urllib.parse
import urllib.request
from dataclasses import dataclass, field

from markdown_it import MarkdownIt
from markdown_it.tree import SyntaxTreeNode

from pygments import lex
from pygments.lexers import TextLexer, get_lexer_by_name
from pygments.styles import get_style_by_name
from pygments.util import ClassNotFound

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.pdfmetrics import registerFontFamily, stringWidth
from reportlab.pdfbase.ttfonts import TTFont, TTFontFile
from reportlab.graphics import renderPDF
from reportlab.platypus import (
    BaseDocTemplate,
    Flowable,
    Frame,
    HRFlowable,
    Image,
    PageBreak,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    XPreformatted,
)
from reportlab.platypus.tableofcontents import TableOfContents

# svglib turns an SVG into a reportlab Drawing, which goes straight into the
# story and stays vector in the PDF. Pure Python (unlike cairosvg/rsvg), so it
# packages with PyInstaller — but keep it optional so a stripped install still
# converts everything else.
try:
    from svglib.svglib import svg2rlg
except Exception:                                # pragma: no cover
    svg2rlg = None

APP_NAME = "Markdown ⇄ PDF Studio"

# ---------------------------------------------------------------- palette ---

INK = "#1F2937"          # body text
INK_DARK = "#111827"     # darkest heading / TOC ink
INK_SOFT = "#475569"     # blockquote text
GRAY_FOOT = "#9CA3AF"    # footer / captions
LINK = "#2563EB"
RULE = "#E5E7EB"
CODE_BG = colors.HexColor("#F6F8FA")
CODE_BORDER = colors.HexColor("#E2E8F0")
CODE_INLINE_BG = "#EFF1F4"
QUOTE_BAR = colors.HexColor("#6366F1")
QUOTE_BG = colors.HexColor("#F8FAFC")
TBL_HEAD_BG = colors.HexColor("#F3F4F6")
TBL_LINE = colors.HexColor("#E5E7EB")

# Per-level heading ink: a blue-leaning slate ramp, darkest at H1 and lifting
# as it descends, so the hierarchy reads at a glance even where two levels sit
# at nearly the same size. Every level stays above 4.5:1 on white.
# Override per document via ConvertOptions.heading_colors or front matter.
HEADING_COLORS = {
    1: "#0B1220",
    2: "#16223A",
    3: "#26364F",
    4: "#3A4A63",
    5: "#4A5B75",
    6: "#5A6B85",
}

# ------------------------------------------------------------ page presets --

# Word page-setup presets, in mm: (left, right, top, bottom)
MARGIN_PRESETS = {
    "Normal (2.54 cm)":   (25.4, 25.4, 25.4, 25.4),
    "Moderate (1.91 cm)": (19.1, 19.1, 25.4, 25.4),
    "Narrow (1.27 cm)":   (12.7, 12.7, 12.7, 12.7),
}

DEFAULT_FONT = "Aptos"
DEFAULT_MARGINS = "Narrow (1.27 cm)"

# ------------------------------------------------------------------- fonts --

def _font_dirs():
    dirs = []
    if os.name == "nt":
        windir = os.environ.get("WINDIR", r"C:\Windows")
        dirs.append(os.path.join(windir, "Fonts"))
        local = os.environ.get("LOCALAPPDATA")
        if local:
            # per-user installs ("Install for me only")
            dirs.append(os.path.join(local, "Microsoft", "Windows", "Fonts"))
        # Office keeps its own fonts outside the Windows font folder
        for pf in (os.environ.get("PROGRAMFILES"),
                   os.environ.get("PROGRAMFILES(X86)")):
            if pf:
                dirs.append(os.path.join(pf, "Microsoft Office", "root",
                                         "vfs", "Fonts", "private"))
    dirs += [
        "/usr/share/fonts/truetype/dejavu",
        "/usr/share/fonts/truetype/msttcorefonts",
        "/usr/share/fonts/truetype/liberation",
        "/usr/share/fonts/truetype/ubuntu",
        "/Library/Fonts",
        "/System/Library/Fonts/Supplemental",
        os.path.expanduser("~/.fonts"),
    ]
    return [d for d in dirs if os.path.isdir(d)]


_registry_fonts: dict[str, str] | None = None


def _registry_font_map() -> dict[str, str]:
    """Installed font display name (lowercased) -> file path, from Windows.

    Filenames differ between releases of the same font — Aptos alone ships as
    Aptos.ttf or Aptos-Regular.ttf depending on the build — so guessing them
    is unreliable. The registry records what is actually installed, which
    makes any family findable by the name a user would recognise.
    """
    global _registry_fonts
    if _registry_fonts is not None:
        return _registry_fonts
    _registry_fonts = {}
    if os.name != "nt":
        return _registry_fonts
    try:
        import winreg
    except ImportError:                       # pragma: no cover
        return _registry_fonts
    key = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
    for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        try:
            with winreg.OpenKey(root, key) as k:
                for i in range(winreg.QueryInfoKey(k)[1]):
                    try:
                        name, value, _ = winreg.EnumValue(k, i)
                    except OSError:
                        continue
                    if not isinstance(value, str) or not value:
                        continue
                    # "Aptos Bold (TrueType)" -> "aptos bold"
                    name = re.sub(r"\s*\((?:TrueType|OpenType)\)\s*$", "",
                                  name, flags=re.I).strip().lower()
                    if name and name not in _registry_fonts:
                        _registry_fonts[name] = value
        except OSError:
            continue
    return _registry_fonts


_cloud_fonts: dict[str, str] | None = None


def _cloud_font_map() -> dict[str, str]:
    """Office "cloud font" display name (lowercased) -> file path.

    Microsoft 365 downloads fonts on demand — Aptos, its own default, among
    them — into %LOCALAPPDATA%\\Microsoft\\FontCache\\<n>\\CloudFonts\\<Family>\\.
    These are invisible to every other lookup: the filenames are opaque
    numeric ids and the fonts are never added to the Windows font registry,
    so the only way to identify one is to read its internal name table.
    """
    global _cloud_fonts
    if _cloud_fonts is not None:
        return _cloud_fonts
    _cloud_fonts = {}
    local = os.environ.get("LOCALAPPDATA")
    if os.name != "nt" or not local:
        return _cloud_fonts
    base = os.path.join(local, "Microsoft", "FontCache")
    roots = []
    try:
        for entry in os.listdir(base):
            cf = os.path.join(base, entry, "CloudFonts")
            if os.path.isdir(cf):
                roots.append(cf)
    except OSError:
        return _cloud_fonts

    def text(value) -> str:
        if isinstance(value, bytes):
            value = value.decode("utf-8", "ignore")
        return (value or "").strip()

    for root in roots:
        for dirpath, _dirs, files in os.walk(root):
            for fname in files:
                if not fname.lower().endswith((".ttf", ".otf")):
                    continue
                path = os.path.join(dirpath, fname)
                try:
                    info = TTFontFile(path)
                except Exception:
                    continue
                family = text(getattr(info, "familyName", ""))
                style = text(getattr(info, "styleName", ""))
                keys = [text(getattr(info, "fullName", "")),
                        f"{family} {style}".strip()]
                if style.lower() in ("", "regular"):
                    keys.append(family)
                for key in keys:
                    key = key.lower()
                    if key:
                        _cloud_fonts.setdefault(key, path)
    return _cloud_fonts


def _find_installed_font(display_name: str) -> str | None:
    """Path for an installed font, looked up by the name a user would know.

    Covers fonts whose filenames cannot be guessed: registry-registered ones
    first, then Office's on-demand cloud fonts.
    """
    name = display_name.strip().lower()
    value = _registry_font_map().get(name)
    if value:
        if os.path.isabs(value):
            if os.path.isfile(value):
                return value
        else:
            found = _find_font_file(value)
            if found:
                return found
    return _cloud_font_map().get(name)


# family -> (regular, bold, italic, bold-italic); an entry may be "file.ttf",
# ("file.ttc", subfont_index), or a list of filenames to try in turn.
# Missing variants fall back; lookup is case-insensitive.
_TTF_FAMILIES = {
    # Microsoft 365's default family (Office 2024+). Shipped with Office
    # rather than Windows, so it is simply absent on machines without it and
    # resolve_body_font() falls down the chain below.
    "Aptos":              (["Aptos.ttf", "Aptos-Regular.ttf"],
                           ["Aptos-Bold.ttf", "AptosBold.ttf"],
                           ["Aptos-Italic.ttf", "AptosItalic.ttf"],
                           ["Aptos-BoldItalic.ttf", "AptosBoldItalic.ttf"]),
    "Aptos Display":      (["Aptos-Display.ttf", "AptosDisplay.ttf"],
                           ["Aptos-Display-Bold.ttf", "AptosDisplay-Bold.ttf"],
                           ["Aptos-Display-Italic.ttf", "AptosDisplay-Italic.ttf"],
                           ["Aptos-Display-BoldItalic.ttf", "AptosDisplay-BoldItalic.ttf"]),
    "Aptos Narrow":       (["Aptos-Narrow.ttf", "AptosNarrow.ttf"],
                           ["Aptos-Narrow-Bold.ttf", "AptosNarrow-Bold.ttf"],
                           ["Aptos-Narrow-Italic.ttf", "AptosNarrow-Italic.ttf"],
                           ["Aptos-Narrow-BoldItalic.ttf", "AptosNarrow-BoldItalic.ttf"]),
    "Aptos Serif":        (["Aptos-Serif.ttf", "AptosSerif.ttf"],
                           ["Aptos-Serif-Bold.ttf", "AptosSerif-Bold.ttf"],
                           ["Aptos-Serif-Italic.ttf", "AptosSerif-Italic.ttf"],
                           ["Aptos-Serif-BoldItalic.ttf", "AptosSerif-BoldItalic.ttf"]),
    # Ubuntu ships as Ubuntu-Regular.ttf on most systems and Ubuntu-R.ttf in
    # the original Canonical package
    "Ubuntu":             (["Ubuntu-Regular.ttf", "Ubuntu-R.ttf", "Ubuntu.ttf"],
                           ["Ubuntu-Bold.ttf", "Ubuntu-B.ttf"],
                           ["Ubuntu-Italic.ttf", "Ubuntu-RI.ttf"],
                           ["Ubuntu-BoldItalic.ttf", "Ubuntu-BI.ttf"]),
    "Ubuntu Light":       (["Ubuntu-Light.ttf", "Ubuntu-L.ttf"],
                           ["Ubuntu-Medium.ttf", "Ubuntu-M.ttf"],
                           ["Ubuntu-LightItalic.ttf", "Ubuntu-LI.ttf"],
                           ["Ubuntu-MediumItalic.ttf", "Ubuntu-MI.ttf"]),
    "Ubuntu Condensed":   (["UbuntuCondensed-Regular.ttf", "Ubuntu-C.ttf"],
                           None, None, None),
    "Ubuntu Mono":        (["UbuntuMono-Regular.ttf", "UbuntuMono-R.ttf"],
                           ["UbuntuMono-Bold.ttf", "UbuntuMono-B.ttf"],
                           ["UbuntuMono-Italic.ttf", "UbuntuMono-RI.ttf"],
                           ["UbuntuMono-BoldItalic.ttf", "UbuntuMono-BI.ttf"]),
    "Calibri":            ("calibri.ttf", "calibrib.ttf", "calibrii.ttf", "calibriz.ttf"),
    "Cambria":            (("cambria.ttc", 0), "cambriab.ttf", "cambriai.ttf", "cambriaz.ttf"),
    "Georgia":            ("georgia.ttf", "georgiab.ttf", "georgiai.ttf", "georgiaz.ttf"),
    "Segoe UI":           ("segoeui.ttf", "segoeuib.ttf", "segoeuii.ttf", "segoeuiz.ttf"),
    "Arial":              ("arial.ttf", "arialbd.ttf", "ariali.ttf", "arialbi.ttf"),
    "Verdana":            ("verdana.ttf", "verdanab.ttf", "verdanai.ttf", "verdanaz.ttf"),
    "Tahoma":             ("tahoma.ttf", "tahomabd.ttf", None, None),
    "Times New Roman":    ("times.ttf", "timesbd.ttf", "timesi.ttf", "timesbi.ttf"),
    "Palatino Linotype":  ("pala.ttf", "palab.ttf", "palai.ttf", "palabi.ttf"),
    "Garamond":           ("gara.ttf", "garabd.ttf", "garait.ttf", None),
    "Book Antiqua":       ("bkant.ttf", "antquab.ttf", "antquai.ttf", "antquabi.ttf"),
    # Non-Windows fallbacks (useful when running unpackaged on Linux/macOS)
    "DejaVu Sans":        ("DejaVuSans.ttf", "DejaVuSans-Bold.ttf",
                           "DejaVuSans-Oblique.ttf", "DejaVuSans-BoldOblique.ttf"),
    "DejaVu Serif":       ("DejaVuSerif.ttf", "DejaVuSerif-Bold.ttf",
                           "DejaVuSerif-Italic.ttf", "DejaVuSerif-BoldItalic.ttf"),
    "Liberation Sans":    ("LiberationSans-Regular.ttf", "LiberationSans-Bold.ttf",
                           "LiberationSans-Italic.ttf", "LiberationSans-BoldItalic.ttf"),
}

_CODE_FAMILIES = {
    "Consolas":         ("consola.ttf", "consolab.ttf", "consolai.ttf", "consolaz.ttf"),
    "Aptos Mono":       (["Aptos-Mono.ttf", "AptosMono.ttf"],
                         ["Aptos-Mono-Bold.ttf", "AptosMono-Bold.ttf"], None, None),
    "Ubuntu Mono":      (["UbuntuMono-Regular.ttf", "UbuntuMono-R.ttf"],
                         ["UbuntuMono-Bold.ttf", "UbuntuMono-B.ttf"],
                         ["UbuntuMono-Italic.ttf", "UbuntuMono-RI.ttf"],
                         ["UbuntuMono-BoldItalic.ttf", "UbuntuMono-BI.ttf"]),
    "Cascadia Mono":    ("CascadiaMono.ttf", None, None, None),
    "DejaVu Sans Mono": ("DejaVuSansMono.ttf", "DejaVuSansMono-Bold.ttf",
                         "DejaVuSansMono-Oblique.ttf", "DejaVuSansMono-BoldOblique.ttf"),
}

# Built-in Type1 families (always available, but limited symbol coverage)
_BUILTIN_FAMILIES = {
    "Helvetica (built-in)": ("Helvetica", "Helvetica-Bold", "Helvetica-Oblique", "Helvetica-BoldOblique"),
    "Times (built-in)":     ("Times-Roman", "Times-Bold", "Times-Italic", "Times-BoldItalic"),
}

_registered_families: dict[str, str | None] = {}


def _find_font_file(name):
    for d in _font_dirs():
        p = os.path.join(d, name)
        if os.path.isfile(p):
            return p
        # try case-insensitive match
        try:
            for f in os.listdir(d):
                if f.lower() == name.lower():
                    return os.path.join(d, f)
        except OSError:
            pass
    return None


def _register_ttf_family(family: str) -> str | None:
    """Register a TTF family with reportlab. Returns the fontName to use
    on ParagraphStyle (with <b>/<i> mapped), or None if unavailable."""
    if family in _registered_families:
        return _registered_families[family]

    spec = _TTF_FAMILIES.get(family) or _CODE_FAMILIES.get(family)
    if spec is None:
        _registered_families[family] = None
        return None

    def load(entry, suffix, style):
        subfont = 0
        names: list = []
        if isinstance(entry, tuple):
            names, subfont = [entry[0]], entry[1]
        elif entry is not None:
            # a list means "these filenames, first one that exists" — vendors
            # are not consistent about hyphenating variant names
            names = entry if isinstance(entry, list) else [entry]
        path = next((p for p in (_find_font_file(n) for n in names) if p), None)
        if not path:
            # last resort: ask what is actually installed, by display name
            path = _find_installed_font(f"{family} {style}".strip())
        if not path:
            return None
        psname = family.replace(" ", "") + suffix
        try:
            pdfmetrics.registerFont(TTFont(psname, path, subfontIndex=subfont))
            return psname
        except Exception:
            return None

    normal = load(spec[0], "", "")
    if not normal:
        _registered_families[family] = None
        return None
    bold = load(spec[1], "-Bold", "Bold") or normal
    italic = load(spec[2], "-Italic", "Italic") or normal
    bolditalic = load(spec[3], "-BoldItalic", "Bold Italic") or bold
    registerFontFamily(normal, normal=normal, bold=bold,
                       italic=italic, boldItalic=bolditalic)
    _registered_families[family] = normal
    return normal


def resolve_body_font(family: str) -> str:
    """Return a reportlab fontName for the requested family (best effort)."""
    if family in _BUILTIN_FAMILIES:
        return _BUILTIN_FAMILIES[family][0]
    name = _register_ttf_family(family)
    if name:
        return name
    # fall back down a sensible chain (Aptos ships with Office, not Windows)
    for fb in ("Aptos", "Calibri", "Segoe UI", "Arial", "DejaVu Sans",
               "Liberation Sans"):
        name = _register_ttf_family(fb)
        if name:
            return name
    return "Helvetica"


def resolve_code_font() -> str:
    for fam in ("Consolas", "Cascadia Mono", "Aptos Mono", "Ubuntu Mono",
                "DejaVu Sans Mono"):
        name = _register_ttf_family(fam)
        if name:
            return name
    return "Courier"


def list_font_families() -> list[str]:
    """Families actually available on this machine, for the UI dropdown."""
    out = []
    for fam in _TTF_FAMILIES:
        if _register_ttf_family(fam):
            out.append(fam)
    out += list(_BUILTIN_FAMILIES)
    return out


def _is_builtin(font_name: str) -> bool:
    return font_name in ("Helvetica", "Times-Roman", "Courier")


# Built-in Type1 fonts miss common symbols; degrade them gracefully
# (only used when no symbol-fallback font could be found at all).
_BUILTIN_CHAR_FALLBACK = {
    "\u2192": "->", "\u2190": "<-", "\u2191": "^", "\u2193": "v",
    "\u2713": "[x]", "\u2714": "[x]", "\u2717": "[ ]",
    "\u2026": "...", "\u2212": "-", "\u00d7": "x",
    "\u201c": '"', "\u201d": '"', "\u2018": "'", "\u2019": "'",
    "\u2265": ">=", "\u2264": "<=", "\u2248": "~",
}

# ---- symbol / emoji handling ------------------------------------------------
# Body fonts like Calibri have no glyphs for ✅ ❌ ⚠ etc., so those chars
# render as boxes. We route them through a symbol-capable fallback font
# (Segoe UI Symbol ships with Windows and covers nearly all of them as
# monochrome glyphs), and degrade to lookalikes when even that is missing.

_SYMBOL_FALLBACK_FILES = [
    "seguisym.ttf",                    # Segoe UI Symbol (Windows)
    "seguiemj.ttf",                    # Segoe UI Emoji (Windows)
    "NotoSansSymbols2-Regular.ttf",
    "NotoSansSymbols-Regular.ttf",
    "Symbola.ttf",
    "DejaVuSans.ttf",
]
_symbol_fallback: str | None | bool = False    # False = not probed yet


def _get_symbol_fallback() -> str | None:
    global _symbol_fallback
    if _symbol_fallback is False:
        _symbol_fallback = None
        for fname in _SYMBOL_FALLBACK_FILES:
            path = _find_font_file(fname)
            if not path:
                continue
            try:
                pdfmetrics.registerFont(TTFont("SymbolFB", path))
                _symbol_fallback = "SymbolFB"
                break
            except Exception:
                continue
    return _symbol_fallback


_coverage_cache: dict[str, dict] = {}


def _covers(font_name: str, cp: int) -> bool:
    """True if the registered font has a real glyph for this codepoint."""
    cov = _coverage_cache.get(font_name)
    if cov is None:
        try:
            f = pdfmetrics.getFont(font_name)
            cov = getattr(getattr(f, "face", None), "charToGlyph", None) or {}
        except Exception:
            cov = {}
        _coverage_cache[font_name] = cov
    if not cov:                        # built-in Type1: Latin-1-ish only
        return cp < 0x100
    return cov.get(cp, 0) != 0


# Lookalikes used when neither the body font nor the fallback has the glyph.
_SYMBOL_SUBSTITUTES = {
    0x2705: "\u2713", 0x2611: "\u2713", 0x2714: "\u2713",     # ✅ ☑ ✔ -> ✓
    0x274C: "\u2717", 0x2716: "\u2717", 0x274E: "\u2717",     # ❌ ✖ ❎ -> ✗
    0x1F6AB: "\u2717",                                        # 🚫
    0x26A0: "(!)", 0x2757: "!", 0x2755: "!",                  # ⚠ ❗ ❕
    0x2753: "?", 0x2754: "?",                                 # ❓ ❔
    0x2B50: "*", 0x1F4A1: "(i)", 0x2139: "(i)",               # ⭐ 💡 ℹ
    0x1F534: "\u25CF", 0x1F7E0: "\u25CF", 0x1F7E1: "\u25CF",  # 🔴 🟠 🟡
    0x1F7E2: "\u25CF", 0x1F535: "\u25CF", 0x1F7E3: "\u25CF",  # 🟢 🔵 🟣
    0x26AB: "\u25CF", 0x26AA: "\u25CB",                       # ⚫ ⚪
}

_SKIP_CPS = {0xFE0E, 0xFE0F, 0x200D}   # variation selectors, ZWJ


def markup_symbols(text: str, font_name: str) -> str:
    """XML-escape text, routing chars the given font can't draw through the
    symbol-fallback font (or substituting a supported lookalike)."""
    fb = _get_symbol_fallback()
    runs: list[list] = []              # [use_fallback, [chars]]

    def put(use_fb: bool, s: str):
        if runs and runs[-1][0] == use_fb:
            runs[-1][1].append(s)
        else:
            runs.append([use_fb, [s]])

    for ch in text:
        cp = ord(ch)
        if cp in _SKIP_CPS:
            continue
        if cp < 0x2000 or _covers(font_name, cp):
            put(False, ch)
            continue
        if fb and _covers(fb, cp):
            put(True, ch)
            continue
        rep = _SYMBOL_SUBSTITUTES.get(cp)
        if rep:
            if all(_covers(font_name, ord(c)) for c in rep):
                put(False, rep)
                continue
            if fb and all(_covers(fb, ord(c)) for c in rep):
                put(True, rep)
                continue
        put(False, ch)                 # last resort: let the font try

    out = []
    for use_fb, chars in runs:
        seg = _esc("".join(chars))
        out.append(f'<font face="{fb}">{seg}</font>' if use_fb else seg)
    return "".join(out)


def plain_symbols(text: str, font_name: str) -> str:
    """Markup-free variant for text drawn outside Paragraph markup
    (e.g. table-of-contents entries): substitute lookalikes only."""
    out = []
    for ch in text:
        cp = ord(ch)
        if cp in _SKIP_CPS:
            continue
        if cp < 0x2000 or _covers(font_name, cp):
            out.append(ch)
            continue
        rep = _SYMBOL_SUBSTITUTES.get(cp)
        if rep and all(_covers(font_name, ord(c)) for c in rep):
            out.append(rep)
        else:
            out.append(ch)
    return "".join(out)

# ----------------------------------------------------------------- options --

@dataclass
class ConvertOptions:
    font_family: str = DEFAULT_FONT
    base_size: float = 11.0          # paragraph pt; everything scales off this
    margins: str = DEFAULT_MARGINS
    page_numbers: bool = True
    toc: bool = False
    highlight_code: bool = True
    # True  -> render <svg>…</svg> written straight into the markdown
    # False -> only linked files, ![](diagrams/x.svg); raw <svg> is ignored
    inline_svg: bool = True
    title: str | None = None         # PDF metadata title (defaults to first H1)
    # level -> colour, merged over HEADING_COLORS; front matter merges on top
    heading_colors: dict[int, str] = field(default_factory=dict)


# heading size multipliers relative to base_size
H_SCALE = {1: 1.90, 2: 1.50, 3: 1.26, 4: 1.10, 5: 1.00, 6: 0.94}


# ------------------------------------------------------------ front matter --

_FRONT_MATTER = re.compile(r"\A﻿?[ \t]*---[ \t]*\r?\n(.*?)\r?\n[ \t]*---[ \t]*"
                           r"(?:\r?\n|\Z)", re.S)


def split_front_matter(md_text: str) -> tuple[dict[str, str], str]:
    """Peel a leading `--- ... ---` block off the document.

    Deliberately not YAML — one `key: value` per line is all this needs, and a
    real parser would mean another dependency in the exe. Unrecognised keys are
    simply ignored; the body is returned with the block removed so it never
    renders as stray text.
    """
    m = _FRONT_MATTER.match(md_text)
    if not m:
        return {}, md_text
    meta: dict[str, str] = {}
    for line in m.group(1).splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        key, sep, val = line.partition(":")
        if not sep:
            continue
        meta[key.strip().lower()] = val.strip().strip("'\"")
    return meta, md_text[m.end():]


def heading_colors_from_meta(meta: dict[str, str]) -> dict[int, str]:
    """Read `heading_colors: a, b, c…` (H1 first) and/or `h3_color: …`."""
    out: dict[int, str] = {}
    ramp = meta.get("heading_colors") or meta.get("heading_colours")
    if ramp:
        items = [p for p in re.split(r"[,\s]+", ramp.strip(" []")) if p]
        for i, raw in enumerate(items[:6], start=1):
            col = _safe_color(raw)
            if col:
                out[i] = col
    for lvl in range(1, 7):
        for key in (f"h{lvl}_color", f"h{lvl}_colour", f"heading{lvl}_color"):
            col = _safe_color(meta.get(key))
            if col:
                out[lvl] = col
                break
    return out


# ------------------------------------------------------------------ helpers --

def _esc(text: str) -> str:
    return html.escape(text, quote=True)


def _safe_color(value: str | None) -> str | None:
    """A colour literal safe to drop into reportlab Paragraph markup, or None.

    Accepts #rgb/#rrggbb(aa) and reportlab's own named colours; anything else
    (CSS functions, var(), garbage) is rejected rather than passed through,
    because an unparseable colour aborts the whole build.
    """
    if not value:
        return None
    v = value.strip().rstrip(";").strip()
    if re.fullmatch(r"#(?:[0-9a-fA-F]{3,4}|[0-9a-fA-F]{6}|[0-9a-fA-F]{8})", v):
        return v
    if re.fullmatch(r"[A-Za-z]+", v) and v.lower() in colors.getAllNamedColors():
        return v.lower()
    return None


def _color(value: str):
    """Colour literal -> reportlab colour. Handles the named colours
    _safe_color() lets through, which HexColor alone would choke on."""
    return colors.toColor(value)


_SVG_EXTS = (".svg", ".svgz")


def _looks_like_svg(path: str) -> bool:
    """Extension first, then sniff — remote images arrive with whatever
    suffix the URL happened to carry (often none)."""
    if path.lower().endswith(_SVG_EXTS):
        return True
    try:
        with open(path, "rb") as fh:
            head = fh.read(400)
    except OSError:
        return False
    if head[:2] == b"\x1f\x8b":          # gzipped .svgz served without extension
        return True
    return b"<svg" in head.lower()


def load_svg(src):
    """SVG -> reportlab Drawing (vector), or None if it can't be read.

    ``src`` is a file path or any file-like object, so inline markup never
    touches disk. Dimensions come back in points, svglib having already
    applied the 96 dpi px -> pt conversion.
    """
    if svg2rlg is None:
        return None
    try:
        d = svg2rlg(src)
    except Exception:
        return None
    if d is None or not float(getattr(d, "width", 0) or 0) or \
            not float(getattr(d, "height", 0) or 0):
        return None
    return d


# An <svg> element written straight into the markdown. Non-greedy so two
# diagrams in one HTML block stay separate.
_INLINE_SVG = re.compile(r"<svg\b[^>]*>.*?</svg\s*>", re.I | re.S)
# a <title> immediately inside <svg> is the element's accessible name, and
# makes a good caption — the same role alt text plays for a linked image
_SVG_TITLE = re.compile(r"<svg\b[^>]*>\s*<title[^>]*>(.*?)</title\s*>", re.I | re.S)


def normalise_inline_svg(text: str) -> str:
    """Make hand-written <svg> blocks survive CommonMark's HTML-block rules.

    A raw HTML block ends at the first blank line and cannot interrupt a
    paragraph, so an SVG with blank lines in it arrives cut in half and the
    tail renders as stray text. Squeeze the blank lines out of each <svg>
    element, put the opening tag on a line of its own, and make sure a blank
    line precedes it. Fenced code is left exactly as written — an SVG shown
    as a code sample must stay a code sample.
    """
    if not re.search(r"</svg\s*>", text, re.I):
        return text
    out: list[str] = []
    buf: list[str] | None = None
    fence: str | None = None
    for line in text.splitlines(keepends=True):
        if buf is not None:
            buf.append(line)
            if re.search(r"</svg\s*>", line, re.I):
                out.extend(_tidy_svg_lines(buf))
                buf = None
            continue
        stripped = line.lstrip()
        if fence:
            if stripped.startswith(fence):
                fence = None
            out.append(line)
            continue
        m = re.match(r"(`{3,}|~{3,})", stripped)
        if m:
            fence = m.group(1)
            out.append(line)
            continue
        # only a line that *opens* the element, not any line mentioning it —
        # a table row or sentence containing "<svg>" would otherwise swallow
        # everything up to the next real diagram
        if re.match(r"<svg\b", stripped, re.I):
            buf = [line]
            if re.search(r"</svg\s*>", line, re.I):
                out.extend(_tidy_svg_lines(buf))
                buf = None
            continue
        out.append(line)
    if buf:                      # never closed — leave it exactly as written
        out.extend(buf)
    return "".join(out)


def _tidy_svg_lines(lines: list[str]) -> list[str]:
    block = "".join(lines)
    # opening tag alone on its line first, then squeeze blank lines — doing it
    # the other way round leaves a blank line right after the tag, which ends
    # the HTML block before the element even starts
    block = re.sub(r"(<svg\b[^>]*?>)[ \t]*", r"\1\n", block, count=1)
    block = re.sub(r"\n[ \t]*(?=\n)", "", block)
    if not block.endswith("\n"):
        block += "\n"
    return ["\n", block, "\n"]   # blank line each side so it stands alone


def _slugify(text: str, taken: dict[str, int]) -> str:
    slug = text.strip().lower()
    slug = re.sub(r"[^\w\s-]", "", slug, flags=re.UNICODE)
    slug = re.sub(r"[\s]+", "-", slug).strip("-") or "section"
    n = taken.get(slug, 0)
    taken[slug] = n + 1
    return slug if n == 0 else f"{slug}-{n}"


def _plain_text(node) -> str:
    """Plain text of an inline subtree (for slugs / outline titles / widths)."""
    parts = []
    for n in getattr(node, "children", []) or []:
        if n.type in ("text", "code_inline"):
            parts.append(n.content)
        elif n.type in ("softbreak", "hardbreak"):
            parts.append(" ")
        elif n.children:
            parts.append(_plain_text(n))
    return "".join(parts)


class SvgFigure(Flowable):
    """A vector SVG in the story, kept sharp and small — no rasterising.

    The Drawing is held rather than subclassed: reportlab validates Drawing
    attributes against its _attrMap, so the spacing and keepWithNext flags
    platypus wants to set cannot live on it.
    """

    def __init__(self, drawing, width: float, height: float):
        Flowable.__init__(self)
        self.drawing = drawing
        self.width = width
        self.height = height

    def wrap(self, availWidth, availHeight):
        return self.width, self.height

    def draw(self):
        renderPDF.draw(self.drawing, self.canv, 0, 0)


class RuledHeading(Paragraph):
    """Heading paragraph that draws a full-width rule underneath (GitHub style)."""

    def __init__(self, text, style, rule_color=None):
        Paragraph.__init__(self, text, style)
        self._rule_color = rule_color

    def draw(self):
        Paragraph.draw(self)
        if self._rule_color is not None:
            c = self.canv
            c.saveState()
            c.setStrokeColor(self._rule_color)
            c.setLineWidth(0.9)
            c.line(0, -4, self.width, -4)
            c.restoreState()


# ---- pagination guards ------------------------------------------------------
# Platypus splits a flowable wherever the page happens to run out, which
# strands single table rows and lone code lines. Both classes below refuse the
# bad split instead of wrapping the block in KeepTogether: a KeepTogether is a
# _ContainerSpace, which doctemplate.handle_keepWithNext() will not absorb, so
# wrapping would silently break the keepWithNext link from the heading or
# lead-in paragraph above. Refusing a split (returning []) just moves the block
# to the next page, and both guards only do that when it can actually fit one.

# Moving a block whole leaves behind as much blank page as the block is tall,
# so only blocks up to this fraction of the frame are worth keeping together —
# past it, splitting cleanly beats a half-empty page.
KEEP_WHOLE_FRAC = 0.55


class GuardedTable(Table):
    """Table that will not split into a stranded fragment.

    A short table is moved whole rather than split; a longer one splits, but
    never so that fewer than ``min_rows`` body rows land on either side of the
    break. ``repeatRows`` still repeats the header on continuation pages.
    """

    min_rows = 2

    def __init__(self, *args, frame_height: float = 0.0, **kw):
        Table.__init__(self, *args, **kw)
        self.frame_height = frame_height
        self.keep_max = frame_height * KEEP_WHOLE_FRAC

    @staticmethod
    def _n_repeat(tbl) -> int:
        rr = tbl.repeatRows
        return rr if isinstance(rr, int) else len(rr)

    def _body_rows(self, tbl) -> int:
        return getattr(tbl, "_nrows", 0) - self._n_repeat(tbl)

    def split(self, availWidth, availHeight):
        total = sum(getattr(self, "_rowHeights", None) or [0])
        # deferring is only ever safe when a fresh page has more room than we
        # have here and the table actually fits one — otherwise we would loop
        room_next_page = availHeight < self.frame_height - 1
        can_defer = room_next_page and 0 < total <= self.frame_height
        if room_next_page and 0 < total <= self.keep_max:
            return []

        parts = Table.split(self, availWidth, availHeight)
        if len(parts) != 2:
            return parts
        heights = getattr(self, "_rowHeights", None) or []
        for _ in range(6):
            head, tail = self._body_rows(parts[0]), self._body_rows(parts[1])
            if head >= self.min_rows and tail >= self.min_rows:
                return parts
            if head < self.min_rows:
                # not even a sensible opening fragment fits — take it all over
                return [] if can_defer else parts
            # pull the last head row down so the continuation is not a lone row
            idx = self._n_repeat(self) + head - 1
            if not 0 <= idx < len(heights):
                return parts
            availHeight -= heights[idx]
            retry = Table.split(self, availWidth, availHeight)
            if len(retry) != 2:
                return [] if can_defer else parts
            parts = retry
        return parts


class GuardedCode(XPreformatted):
    """Fenced block that stays whole unless it is big enough to be worth
    splitting. When it does split, the style's widow/orphan settings keep at
    least two lines on each side of the break."""

    def __init__(self, *args, frame_height: float = 0.0, **kw):
        XPreformatted.__init__(self, *args, **kw)
        self.frame_height = frame_height
        self.keep_max = frame_height * KEEP_WHOLE_FRAC

    def split(self, availWidth, availHeight):
        height = getattr(self, "height", 0) or 0
        if 0 < height <= self.keep_max and availHeight < self.frame_height - 1:
            return []
        return XPreformatted.split(self, availWidth, availHeight)


class GuideDocTemplate(BaseDocTemplate):
    """A4 doc with footer page numbers, PDF outline + TOC notifications."""

    def __init__(self, filename, opts: ConvertOptions, heading_levels: dict[str, int], **kw):
        lm, rm, tm, bm = (v * mm for v in MARGIN_PRESETS.get(
            opts.margins, MARGIN_PRESETS[DEFAULT_MARGINS]))
        super().__init__(
            filename, pagesize=A4,
            leftMargin=lm, rightMargin=rm, topMargin=tm, bottomMargin=bm,
            title=opts.title or "", creator=APP_NAME, **kw)
        self._opts = opts
        self._heading_levels = heading_levels
        self._footer_font = resolve_body_font(opts.font_family)
        self._footer_size = max(7.0, opts.base_size * 0.78)
        self._outline_last = -1
        frame = Frame(lm, bm, A4[0] - lm - rm, A4[1] - tm - bm, id="body")
        self.addPageTemplates([PageTemplate(id="main", frames=[frame],
                                            onPage=self._draw_footer)])

    def _draw_footer(self, canv, doc):
        if not self._opts.page_numbers:
            return
        canv.saveState()
        canv.setFont(self._footer_font, self._footer_size)
        canv.setFillColor(colors.HexColor(GRAY_FOOT))
        canv.drawCentredString(A4[0] / 2.0, doc.bottomMargin * 0.45,
                               str(canv.getPageNumber()))
        canv.restoreState()

    def afterFlowable(self, flowable):
        if not isinstance(flowable, Paragraph):
            return
        level = self._heading_levels.get(flowable.style.name)
        if level is None:
            return
        key = getattr(flowable, "_bookmark_key", None)
        if not key:
            return
        text = flowable.getPlainText()
        self.canv.bookmarkPage(key)
        lvl = min(level - 1, self._outline_last + 1)
        try:
            self.canv.addOutlineEntry(text, key, level=lvl, closed=False)
            self._outline_last = lvl
        except Exception:
            pass
        # feed the (optional) table of contents; clamp depth to 3 levels
        toc_text = _esc(plain_symbols(text, self._footer_font))
        self.notify("TOCEntry", (min(level - 1, 2), toc_text, self.page, key))


# ------------------------------------------------------------- the renderer --

class _Renderer:
    def __init__(self, opts: ConvertOptions, base_dir: str, avail_width: float):
        self.opts = opts
        self.base_dir = base_dir
        self.avail = avail_width
        lm, rm, tm, bm = MARGIN_PRESETS.get(opts.margins,
                                            MARGIN_PRESETS[DEFAULT_MARGINS])
        self.avail_h = A4[1] - (tm + bm) * mm     # frame height for image caps
        self._tmp_files: list[str] = []          # downloaded remote images
        self.S = float(opts.base_size)
        self.body_font = resolve_body_font(opts.font_family)
        self.code_font = resolve_code_font()
        self.builtin_body = _is_builtin(self.body_font)
        self.heading_ink = dict(HEADING_COLORS)
        for lvl, col in (opts.heading_colors or {}).items():
            safe = _safe_color(col)
            if safe and lvl in self.heading_ink:
                self.heading_ink[lvl] = safe
        self.styles = self._build_styles()
        self.heading_levels = {f"MDHeading{i}": i for i in range(1, 7)}
        self._slug_registry: list[str] = []   # slugs in heading order (pre-pass)
        self._slug_set: set[str] = set()
        self._slug_cursor = 0
        try:
            self._pyg_style = get_style_by_name("friendly")
        except ClassNotFound:
            self._pyg_style = get_style_by_name("default")

    # ---- styles ----

    def _build_styles(self):
        S = self.S
        f = self.body_font
        st = {}
        # allowWidows=0 (with reportlab's allowOrphans=0 default) keeps at
        # least two lines of a paragraph on each side of a page break.
        st["body"] = ParagraphStyle(
            "MDBody", fontName=f, fontSize=S, leading=S * 1.45,
            spaceAfter=S * 0.65, textColor=colors.HexColor(INK),
            splitLongWords=1, allowWidows=0, allowOrphans=0)
        for lvl, scale in H_SCALE.items():
            size = S * scale
            st[f"h{lvl}"] = ParagraphStyle(
                f"MDHeading{lvl}", fontName=f, fontSize=size,
                leading=size * 1.22,
                spaceBefore=(S * 1.7 if lvl <= 2 else S * 1.25),
                spaceAfter=(S * 0.85 if lvl <= 2 else S * 0.5),
                textColor=_color(self.heading_ink[lvl]),
                keepWithNext=1, splitLongWords=1)
            st[f"h{lvl}"].fontName = f  # bold applied via <b> wrap at render
        st["code"] = ParagraphStyle(
            "MDCode", fontName=self.code_font, fontSize=S * 0.86,
            leading=S * 0.86 * 1.42, textColor=colors.HexColor("#24292F"),
            backColor=CODE_BG, borderColor=CODE_BORDER, borderWidth=0.75,
            borderPadding=7, borderRadius=4,
            spaceBefore=S * 0.85, spaceAfter=S * 1.0,
            allowWidows=0, allowOrphans=0)
        st["quote"] = ParagraphStyle(
            "MDQuote", parent=st["body"], textColor=colors.HexColor(INK_SOFT),
            spaceAfter=S * 0.35)
        st["tbl"] = ParagraphStyle(
            "MDTableCell", fontName=f, fontSize=S * 0.90,
            leading=S * 0.90 * 1.35, textColor=colors.HexColor(INK),
            splitLongWords=1)
        st["tblh"] = ParagraphStyle(
            "MDTableHead", parent=st["tbl"], textColor=colors.HexColor(INK_DARK))
        st["caption"] = ParagraphStyle(
            "MDCaption", fontName=f, fontSize=S * 0.82,
            leading=S * 0.82 * 1.3, alignment=1,
            textColor=colors.HexColor(GRAY_FOOT), spaceAfter=S * 0.8)
        return st

    # ---- entry points ----

    def prescan_slugs(self, root):
        taken: dict[str, int] = {}
        def walk(node):
            for n in node.children:
                if n.type == "heading":
                    inline = next((c for c in n.children if c.type == "inline"), None)
                    slug = _slugify(_plain_text(inline) if inline else "", taken)
                    self._slug_registry.append(slug)
                    self._slug_set.add(slug)
                elif n.children:
                    walk(n)
        walk(root)

    def render_document(self, root) -> list:
        return self._blocks(root.children)

    # ---- block level ----

    # blocks a lead-in paragraph should stay glued to across a page break
    _LEAD_IN_TARGETS = ("fence", "code_block", "table", "blockquote")

    def _blocks(self, nodes, ctx=None) -> list:
        chunks = [(n, self._block(n, ctx)) for n in nodes]
        self._bind_lead_ins(chunks)
        flow = []
        for _, fls in chunks:
            flow.extend(fls)
        return flow

    def _block(self, n, ctx=None) -> list:
        t = n.type
        if t == "heading":
            return self._heading(n)
        if t == "paragraph":
            return self._paragraph(n, ctx)
        if t in ("fence", "code_block"):
            return self._code_block(n)
        if t in ("bullet_list", "ordered_list"):
            return self._list(n, depth=0) + [Spacer(1, self.S * 0.45)]
        if t == "blockquote":
            return self._blockquote(n)
        if t == "table":
            return self._table(n)
        if t == "hr":
            return [HRFlowable(width="100%", thickness=0.9,
                               color=colors.HexColor(RULE),
                               spaceBefore=self.S * 0.9,
                               spaceAfter=self.S * 1.0)]
        if t == "html_block":
            content = n.content
            if re.search(r"page[-_ ]?break", content, re.I):
                return [PageBreak()]
            out = []
            if self.opts.inline_svg:
                for m in _INLINE_SVG.finditer(content):
                    out.extend(self._inline_svg_block(m.group(0)))
                content = _INLINE_SVG.sub("", content)   # don't rescan it
            for tag in re.findall(r"<img\b[^>]*>", content, re.I):
                src, alt, w, h = self._parse_img_tag(tag)
                if src:
                    out.extend(self._image_block(src, alt, w, h))
            return out
        return []                      # anything else is ignored

    def _bind_lead_ins(self, chunks) -> None:
        """Glue "…, as follows:" paragraphs to the block they introduce.

        Headings already carry keepWithNext; the sentence that hands off to a
        code block or table is just as bad to strand at a page bottom. Only
        blocks the guards would keep whole anyway qualify — gluing a lead-in to
        a block that is going to split regardless just moves the hole.
        """
        for i, (node, fls) in enumerate(chunks):
            if node.type != "paragraph" or not fls or i + 1 >= len(chunks):
                continue
            if not isinstance(fls[-1], Paragraph):
                continue              # image-only paragraph or a page break
            nxt_node, nxt_fls = chunks[i + 1]
            if nxt_node.type not in self._LEAD_IN_TARGETS or not nxt_fls:
                continue
            if getattr(nxt_fls[0], "_md_height", 0) > self.avail_h * KEEP_WHOLE_FRAC:
                continue
            fls[-1].keepWithNext = 1

    def _heading(self, node) -> list:
        level = int(node.tag[1]) if node.tag and node.tag[1:].isdigit() else 2
        inline = next((c for c in node.children if c.type == "inline"), None)
        markup = self._inline(inline.children if inline else [],
                              self.styles[f"h{level}"].fontSize)
        slug = (self._slug_registry[self._slug_cursor]
                if self._slug_cursor < len(self._slug_registry) else f"h-{id(node)}")
        self._slug_cursor += 1
        style = self.styles[f"h{level}"]
        text = f"<b>{markup}</b>"
        if level <= 2:
            para = RuledHeading(text, style, rule_color=colors.HexColor(RULE))
        else:
            para = Paragraph(text, style)
        para._bookmark_key = slug
        return [para]

    def _paragraph(self, node, ctx=None) -> list:
        inline = next((c for c in node.children if c.type == "inline"), None)
        if inline is None:
            return []
        kids = inline.children or []
        # explicit page break: a paragraph that is just \pagebreak / \newpage
        plain = _plain_text(inline).strip()
        if re.fullmatch(r"\\?(pagebreak|newpage)", plain, re.I):
            return [PageBreak()]
        # image-only paragraph -> block image
        solid = [k for k in kids if not (k.type == "text" and not k.content.strip())]
        if solid and all(k.type == "image" for k in solid):
            out = []
            for k in solid:
                out.extend(self._image(k))
            return out
        style = ctx or self.styles["body"]
        markup = self._inline(kids, style.fontSize)
        if not markup.strip():
            return []
        return [Paragraph(markup, style)]

    def _image(self, node) -> list:
        src = node.attrs.get("src", "")
        alt = _plain_text(node) or node.attrs.get("alt", "")
        return self._image_block(src, alt)

    def cleanup(self):
        for p in self._tmp_files:
            try:
                os.unlink(p)
            except OSError:
                pass
        self._tmp_files.clear()

    def _resolve_image(self, src: str) -> str | None:
        """Return a local file path for the image source, downloading
        http(s) sources to a temp file. None if unavailable."""
        src = (src or "").strip()
        if not src:
            return None
        if re.match(r"^https?://", src, re.I):
            try:
                req = urllib.request.Request(
                    src, headers={"User-Agent": f"{APP_NAME}/1.0"})
                with urllib.request.urlopen(req, timeout=8) as resp:
                    data = resp.read(25 * 1024 * 1024)
                suffix = os.path.splitext(
                    urllib.parse.urlparse(src).path)[1] or ".img"
                tf = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
                tf.write(data)
                tf.close()
                self._tmp_files.append(tf.name)
                return tf.name
            except Exception:
                return None
        path = urllib.parse.unquote(src)
        if not os.path.isabs(path):
            path = os.path.normpath(os.path.join(self.base_dir, path))
        return path if os.path.isfile(path) else None

    @staticmethod
    def _px_to_pt(v: float) -> float:
        return v * 72.0 / 96.0

    def _fit(self, nat_w: float, nat_h: float, w_px, h_px,
             max_w: float, max_h: float) -> float:
        """Scale factor for natural size (pt) honouring an explicit
        width/height in px, capped to the box and never upscaling."""
        if w_px and not h_px:
            h_px = w_px * nat_h / nat_w
        elif h_px and not w_px:
            w_px = h_px * nat_w / nat_h
        want = 1.0
        if w_px and h_px:
            # both given and out of ratio: letterbox rather than distort
            want = min(self._px_to_pt(w_px) / nat_w, self._px_to_pt(h_px) / nat_h)
        return min(want, max_w / nat_w, max_h / nat_h)

    def _svg_block(self, src, w_px=None, h_px=None):
        """SVG (path or file-like) -> a figure that stays vector in the PDF."""
        d = load_svg(src)
        if d is None:
            return None
        nat_w, nat_h = float(d.width), float(d.height)
        scale = self._fit(nat_w, nat_h, w_px, h_px, self.avail, self.avail_h * 0.82)
        d.scale(scale, scale)
        d.width, d.height = nat_w * scale, nat_h * scale
        return SvgFigure(d, d.width, d.height)

    def _inline_svg_block(self, markup: str) -> list:
        """An <svg> element written straight into the markdown."""
        fig = self._svg_block(io.BytesIO(markup.encode("utf-8")))
        if fig is None:
            return [Paragraph("[unreadable inline SVG]", self.styles["caption"])]
        fig.spaceBefore = self.S * 0.4
        fig.spaceAfter = self.S * 0.3
        m = _SVG_TITLE.search(markup)
        caption = html.unescape(m.group(1)).strip() if m else ""
        if not caption:
            fig.spaceAfter = self.S * 0.8
            return [fig]
        fig.keepWithNext = 1
        return [fig, Paragraph(markup_symbols(caption, self.body_font),
                               self.styles["caption"])]

    def _image_block(self, src: str, alt: str,
                     w_px: float | None = None, h_px: float | None = None) -> list:
        path = self._resolve_image(src)
        if path is None:
            if re.match(r"^https?://", src or "", re.I):
                cap = (f'<a href="{_esc(src)}" color="{LINK}">'
                       f"<u>{_esc(alt or src)}</u></a>")
                return [Paragraph(f"[image] {cap}", self.styles["caption"])]
            return [Paragraph(f"[missing image: {_esc(src)}]",
                              self.styles["caption"])]

        art = self._svg_block(path, w_px, h_px) if _looks_like_svg(path) else None
        if art is None:
            try:
                iw, ih = ImageReader(path).getSize()
            except Exception:
                return [Paragraph(f"[unreadable image: {_esc(src)}]",
                                  self.styles["caption"])]
            scale = self._fit(self._px_to_pt(iw), self._px_to_pt(ih), w_px, h_px,
                              self.avail, self.avail_h * 0.82)
            art = Image(path, width=self._px_to_pt(iw) * scale,
                        height=self._px_to_pt(ih) * scale)
            art.hAlign = "CENTER"

        art.spaceBefore = self.S * 0.4
        art.spaceAfter = self.S * 0.3
        if alt:
            # the caption is meaningless on its own page
            art.keepWithNext = 1
            return [art, Paragraph(markup_symbols(alt, self.body_font),
                                   self.styles["caption"])]
        art.spaceAfter = self.S * 0.8
        return [art]

    def _inline_img(self, src: str, alt: str, in_table: bool,
                    w_px: float | None = None, h_px: float | None = None) -> str:
        """An <img> markup fragment for images mixed into a text line."""
        path = self._resolve_image(src)
        if path is None:
            if re.match(r"^https?://", src or "", re.I):
                return (f'<a href="{_esc(src)}" color="{LINK}">'
                        f"<u>[{_esc(alt or 'image')}]</u></a>")
            return f"<i>[{_esc(alt or 'image')}]</i>"
        if _looks_like_svg(path):
            # reportlab's <img> markup only takes rasters, and rasterising an
            # SVG would need a native backend. On its own line an SVG becomes a
            # vector figure instead — see _image_block.
            return f"<i>[{_esc(alt or 'diagram')}]</i>"
        try:
            iw, ih = ImageReader(path).getSize()
        except Exception:
            return f"<i>[{_esc(alt or 'image')}]</i>"
        max_w = self.avail * (0.30 if in_table else 0.86)
        nat_w, nat_h = self._px_to_pt(iw), self._px_to_pt(ih)
        scale = self._fit(nat_w, nat_h, w_px, h_px, max_w, self.avail_h * 0.5)
        w, h = nat_w * scale, nat_h * scale
        return (f'<img src="{_esc(path)}" width="{w:.1f}" height="{h:.1f}" '
                f'valign="middle"/>')

    @staticmethod
    def _parse_img_tag(tag: str):
        """(src, alt, w_px, h_px) out of an HTML <img …> tag."""
        def attr(name):
            m = re.search(rf'{name}\s*=\s*("([^"]*)"|\'([^\']*)\'|(\S+))',
                          tag, re.I)
            return (m.group(2) or m.group(3) or m.group(4)) if m else None
        def px(v):
            if not v:
                return None
            m = re.match(r"([\d.]+)", v)
            return float(m.group(1)) if m else None
        return attr("src"), attr("alt") or "", px(attr("width")), px(attr("height"))

    def _list(self, node, depth: int) -> list:
        ordered = node.type == "ordered_list"
        start = int(node.attrs.get("start", 1)) if ordered else 1
        indent_step = self.S * 1.55
        left = indent_step * (depth + 1)
        item_style = ParagraphStyle(
            f"MDList{depth}", parent=self.styles["body"],
            leftIndent=left, bulletIndent=left - indent_step * 0.85,
            spaceAfter=self.S * 0.32)
        cont_style = ParagraphStyle(
            f"MDListCont{depth}", parent=item_style, spaceAfter=self.S * 0.4)
        flow = []
        n_item = start
        for item in node.children:
            if item.type != "list_item":
                continue
            bullet = f"{n_item}." if ordered else "\u2022"
            n_item += 1
            first_done = False
            for child in item.children:
                if child.type == "paragraph":
                    inline = next((c for c in child.children if c.type == "inline"), None)
                    markup = self._inline(inline.children if inline else [],
                                          item_style.fontSize)
                    if not first_done:
                        flow.append(Paragraph(markup, item_style, bulletText=bullet))
                        first_done = True
                    else:
                        flow.append(Paragraph(markup, cont_style))
                elif child.type in ("bullet_list", "ordered_list"):
                    if not first_done:
                        flow.append(Paragraph("", item_style, bulletText=bullet))
                        first_done = True
                    flow.extend(self._list(child, depth + 1))
                else:
                    if not first_done:
                        flow.append(Paragraph("", item_style, bulletText=bullet))
                        first_done = True
                    sub = self._blocks([child], ctx=cont_style)
                    for fl in sub:
                        # nudge nested blocks over to the list indent
                        if hasattr(fl, "style") and hasattr(fl.style, "leftIndent"):
                            pass
                    flow.extend(sub)
        return flow

    def _blockquote(self, node) -> list:
        inner = self._blocks(node.children, ctx=self.styles["quote"])
        if not inner:
            return []
        rows = [[fl] for fl in inner]
        tbl = Table(rows, colWidths=[self.avail - 4])
        tbl.setStyle(TableStyle([
            ("LINEBEFORE", (0, 0), (0, -1), 2.4, QUOTE_BAR),
            ("BACKGROUND", (0, 0), (-1, -1), QUOTE_BG),
            ("LEFTPADDING", (0, 0), (-1, -1), 14),
            ("RIGHTPADDING", (0, 0), (-1, -1), 10),
            ("TOPPADDING", (0, 0), (-1, -1), 1),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
            ("TOPPADDING", (0, 0), (-1, 0), 8),
            ("BOTTOMPADDING", (0, -1), (-1, -1), 8),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]))
        tbl.hAlign = "LEFT"
        tbl.spaceBefore = self.S * 0.3
        tbl.spaceAfter = self.S * 0.7
        tbl._md_height = self._measure(tbl)
        return [tbl]

    # ---- tables ----

    def _table(self, node) -> list:
        header_rows: list[list] = []
        body_rows: list[list] = []
        plain_rows: list[list[str]] = []

        def cells_of(tr, style):
            cells, plains = [], []
            for c in tr.children:
                inline = next((k for k in c.children if k.type == "inline"), None)
                kids = inline.children if inline else []
                markup = self._inline(kids, style.fontSize, in_table=True)
                if style is self.styles["tblh"]:
                    markup = f"<b>{markup}</b>"
                cells.append(Paragraph(markup, style))
                plains.append(_plain_text(inline) if inline else "")
            return cells, plains

        for part in node.children:
            if part.type == "thead":
                for tr in part.children:
                    cells, plains = cells_of(tr, self.styles["tblh"])
                    header_rows.append(cells)
                    plain_rows.append(plains)
            elif part.type == "tbody":
                for tr in part.children:
                    cells, plains = cells_of(tr, self.styles["tbl"])
                    body_rows.append(cells)
                    plain_rows.append(plains)

        data = header_rows + body_rows
        if not data:
            return []
        ncols = max(len(r) for r in data)
        for r in data:
            while len(r) < ncols:
                r.append(Paragraph("", self.styles["tbl"]))

        # proportional column widths from content, clamped, scaled to fill
        fnt, fsz = self.body_font, self.styles["tbl"].fontSize
        weights = []
        for col in range(ncols):
            wmax = 0.0
            for row in plain_rows:
                if col < len(row):
                    for chunk in row[col].split():
                        wmax = max(wmax, stringWidth(chunk, fnt, fsz))
                    wmax = max(wmax, min(stringWidth(row[col], fnt, fsz),
                                         self.avail * 0.55))
            weights.append(max(wmax + 18, self.S * 4.5))
        total = sum(weights)
        colw = [w / total * self.avail for w in weights]
        # ensure no column is starved below its longest word
        minw = self.S * 4.0
        colw = [max(w, minw) for w in colw]
        scale = self.avail / sum(colw)
        colw = [w * scale for w in colw]

        tbl = GuardedTable(data, colWidths=colw, repeatRows=len(header_rows),
                           splitByRow=1, frame_height=self.avail_h)
        style = [
            ("GRID", (0, 0), (-1, -1), 0.5, TBL_LINE),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 7),
            ("RIGHTPADDING", (0, 0), (-1, -1), 7),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]
        if header_rows:
            style.append(("BACKGROUND", (0, 0), (-1, len(header_rows) - 1), TBL_HEAD_BG))
        tbl.setStyle(TableStyle(style))
        tbl.hAlign = "LEFT"
        # spacing lives on the flowable, not in Spacers: a Spacer in front
        # would absorb the keepWithNext link from the heading or lead-in above
        tbl.spaceBefore = self.S * 0.35
        tbl.spaceAfter = self.S * 0.9
        tbl._md_height = self._measure(tbl)
        return [tbl]

    def _measure(self, flowable) -> float:
        """Laid-out height, for deciding what is small enough to keep with the
        text above it. If it can't be measured, call it a full page so nothing
        gets glued to a block that might turn out to be huge."""
        try:
            return float(flowable.wrap(self.avail, self.avail_h)[1])
        except Exception:
            return self.avail_h

    # ---- code blocks ----

    def _code_block(self, node) -> list:
        code = node.content.replace("\t", "    ").rstrip("\n")
        if not code.strip():
            return []
        lang = (node.info or "").strip().split()[0] if node.info else ""
        lines = self._highlight_lines(code, lang)
        style = self.styles["code"]
        wrap_width = self.avail - 2 * style.borderPadding - 2
        lines = self._wrap_code(lines, style.fontName, style.fontSize, wrap_width)
        markup_lines = []
        for runs in lines:
            segs = []
            for (color, bold, italic), text in runs:
                s = markup_symbols(text, style.fontName)
                if color:
                    s = f'<font color="{color}">{s}</font>'
                if bold:
                    s = f"<b>{s}</b>"
                if italic:
                    s = f"<i>{s}</i>"
                segs.append(s)
            markup_lines.append("".join(segs))
        pre = GuardedCode("\n".join(markup_lines), style,
                          frame_height=self.avail_h)
        pre._md_height = self._measure(pre)
        return [pre]

    def _highlight_lines(self, code: str, lang: str):
        """-> list of lines; each line is a list of ((color,bold,italic), text) runs."""
        default_fmt = (None, False, False)
        if self.opts.highlight_code and lang:
            try:
                lexer = get_lexer_by_name(lang)
            except ClassNotFound:
                lexer = TextLexer()
        else:
            lexer = TextLexer()
        lines: list[list] = [[]]
        if isinstance(lexer, TextLexer):
            for i, ln in enumerate(code.split("\n")):
                if i:
                    lines.append([])
                if ln:
                    lines[-1].append((default_fmt, ln))
            return lines
        for tok, val in lex(code, lexer):
            try:
                sd = self._pyg_style.style_for_token(tok)
                fmt = (("#" + sd["color"]) if sd["color"] else None,
                       bool(sd["bold"]), bool(sd["italic"]))
            except Exception:
                fmt = default_fmt
            for i, part in enumerate(val.split("\n")):
                if i:
                    lines.append([])
                if part:
                    lines[-1].append((fmt, part))
        # pygments usually appends a trailing newline
        if lines and not lines[-1]:
            lines.pop()
        return lines

    @staticmethod
    def _wrap_code(lines, font, size, max_width):
        """Wrap run-lines to the column width (monospace char counting)."""
        charw = stringWidth("M", font, size) or size * 0.6
        max_cols = max(20, int(max_width / charw))
        out = []
        for runs in lines:
            cur, cols = [], 0
            for fmt, text in runs:
                while text:
                    space = max_cols - cols
                    if space <= 0:
                        out.append(cur)
                        cur, cols = [], 0
                        space = max_cols
                    if len(text) <= space:
                        cur.append((fmt, text))
                        cols += len(text)
                        text = ""
                    else:
                        head = text[:space]
                        brk = head.rfind(" ")
                        cut = brk + 1 if brk >= int(space * 0.55) else space
                        cur.append((fmt, text[:cut]))
                        text = text[cut:]
                        out.append(cur)
                        cur, cols = [], 0
            out.append(cur)
        return out

    # ---- inline level ----

    def _inline(self, nodes, ctx_size: float, in_table: bool = False) -> str:
        parts = []
        colour_stack: list[bool] = []   # open <font>/<span> we actually emitted
        for n in nodes:
            t = n.type
            if t == "text":
                parts.append(self._text(n.content))
            elif t == "softbreak":
                parts.append(" ")
            elif t == "hardbreak":
                parts.append("<br/>")
            elif t == "strong":
                parts.append(f"<b>{self._inline(n.children, ctx_size, in_table)}</b>")
            elif t == "em":
                parts.append(f"<i>{self._inline(n.children, ctx_size, in_table)}</i>")
            elif t == "s":
                parts.append(f"<strike>{self._inline(n.children, ctx_size, in_table)}</strike>")
            elif t == "code_inline":
                size = round(ctx_size * 0.88, 1)
                parts.append(
                    f'<font face="{self.code_font}" size="{size}" '
                    f'backColor="{CODE_INLINE_BG}" color="#0F172A">'
                    f"{markup_symbols(n.content, self.code_font)}</font>")
            elif t == "link":
                href = n.attrs.get("href", "")
                inner = self._inline(n.children, ctx_size, in_table)
                if href.startswith("#"):
                    target = href[1:]
                    if target in self._slug_set:
                        parts.append(f'<a href="#{_esc(target)}" color="{LINK}">'
                                     f"<u>{inner}</u></a>")
                    else:
                        parts.append(f'<font color="{LINK}">{inner}</font>')
                elif href:
                    parts.append(f'<a href="{_esc(href)}" color="{LINK}">'
                                 f"<u>{inner}</u></a>")
                else:
                    parts.append(inner)
            elif t == "image":
                alt = _plain_text(n) or n.attrs.get("alt", "")
                parts.append(self._inline_img(n.attrs.get("src", ""), alt,
                                              in_table))
            elif t == "html_inline":
                raw = n.content.strip()
                low = raw.lower()
                if low in ("<br>", "<br/>", "<br />"):
                    parts.append("<br/>")
                elif low.startswith("<img"):
                    src, alt, w, h = self._parse_img_tag(raw)
                    if src:
                        parts.append(self._inline_img(src, alt, in_table, w, h))
                elif re.fullmatch(r"</?(b|strong|i|em|u|sub|sup)>", low):
                    tag = re.sub(r"strong", "b", re.sub(r"em", "i", low))
                    parts.append(tag.replace("sup", "super"))
                elif low.startswith("<font") or low.startswith("<span"):
                    # reportlab's Paragraph markup does colour natively; only
                    # the colour is carried over, and only if it parses
                    col = _safe_color(self._tag_color(raw))
                    colour_stack.append(bool(col))
                    if col:
                        parts.append(f'<font color="{col}">')
                elif low in ("</font>", "</span>"):
                    # never emit an unbalanced close — it would abort the build
                    if colour_stack and colour_stack.pop():
                        parts.append("</font>")
                # anything else is dropped
            elif n.children:
                parts.append(self._inline(n.children, ctx_size, in_table))
        parts.extend("</font>" for open_ in colour_stack if open_)
        return "".join(parts)

    @staticmethod
    def _tag_color(tag: str) -> str | None:
        """Colour out of `<font color=…>` or `<span style="color: …">`."""
        m = re.search(r'\bcolor\s*=\s*("([^"]*)"|\'([^\']*)\'|([^\s>]+))',
                      tag, re.I)
        if m:
            return m.group(2) or m.group(3) or m.group(4)
        m = re.search(r'style\s*=\s*("[^"]*"|\'[^\']*\')', tag, re.I)
        if m:
            m2 = re.search(r"(?<!-)\bcolor\s*:\s*([^;\"']+)", m.group(1), re.I)
            if m2:
                return m2.group(1)
        return None

    def _text(self, s: str) -> str:
        if self.builtin_body and _get_symbol_fallback() is None:
            for ch, rep in _BUILTIN_CHAR_FALLBACK.items():
                s = s.replace(ch, rep)
        return markup_symbols(s, self.body_font)


# ------------------------------------------------------------- public API ---

def _make_parser() -> MarkdownIt:
    try:
        md = MarkdownIt("commonmark", {"html": True, "linkify": True})
        md.enable(["table", "strikethrough", "linkify"])
    except Exception:
        md = MarkdownIt("commonmark", {"html": True})
        md.enable(["table", "strikethrough"])
    return md


def convert_text(md_text: str, pdf_path: str, options: ConvertOptions | None = None,
                 base_dir: str | None = None) -> str:
    """Convert markdown text to a PDF at pdf_path. Returns pdf_path."""
    opts = options or ConvertOptions()
    base_dir = base_dir or os.getcwd()

    # front matter is metadata, not content: strip it, then let it override
    # the caller's options (the document knows its own house style)
    meta, md_text = split_front_matter(md_text)
    if meta:
        opts.heading_colors = {**(opts.heading_colors or {}),
                               **heading_colors_from_meta(meta)}
        if not opts.title and meta.get("title"):
            opts.title = meta["title"]

    # always: this only makes a raw <svg> parse as one HTML block. Rendering
    # it is what opts.inline_svg decides — and an SVG left whole is one block
    # to ignore, rather than loose tags leaking their text into the page.
    md_text = normalise_inline_svg(md_text)

    tokens = _make_parser().parse(md_text)
    root = SyntaxTreeNode(tokens)

    lm, rm, tm, bm = MARGIN_PRESETS.get(opts.margins,
                                        MARGIN_PRESETS[DEFAULT_MARGINS])
    avail = A4[0] - (lm + rm) * mm

    r = _Renderer(opts, base_dir, avail)
    r.prescan_slugs(root)
    story = r.render_document(root)
    if not story:
        story = [Spacer(1, 1)]

    # PDF metadata title: explicit > first H1 > filename
    if not opts.title:
        m = re.search(r"^\s*#\s+(.+?)\s*$", md_text, re.M)
        opts.title = (m.group(1) if m else
                      os.path.splitext(os.path.basename(pdf_path))[0])

    if opts.toc:
        toc = TableOfContents()
        S = opts.base_size
        f = r.body_font
        # contents entries echo the heading ramp, so the outline reads the
        # same way the sections do
        ink = r.heading_ink
        toc.levelStyles = [
            ParagraphStyle("TOC1", fontName=f, fontSize=S, leading=S * 1.6,
                           leftIndent=0, textColor=_color(ink[1])),
            ParagraphStyle("TOC2", fontName=f, fontSize=S * 0.95,
                           leading=S * 1.5, leftIndent=S * 1.4,
                           textColor=_color(ink[2])),
            ParagraphStyle("TOC3", fontName=f, fontSize=S * 0.9,
                           leading=S * 1.45, leftIndent=S * 2.8,
                           textColor=_color(ink[3])),
        ]
        toc_heading = ParagraphStyle(
            "TOCHead", fontName=f, fontSize=S * H_SCALE[1],
            leading=S * H_SCALE[1] * 1.2, spaceAfter=S,
            textColor=_color(ink[1]))
        story = [Paragraph("<b>Contents</b>", toc_heading), toc,
                 PageBreak()] + story

    doc = GuideDocTemplate(pdf_path, opts, r.heading_levels)
    try:
        if opts.toc:
            doc.multiBuild(story)
        else:
            doc.build(story)
    finally:
        r.cleanup()
    return pdf_path


def convert_file(md_path: str, pdf_path: str | None = None,
                 options: ConvertOptions | None = None) -> str:
    with open(md_path, "r", encoding="utf-8-sig") as fh:
        text = fh.read()
    if not pdf_path:
        pdf_path = os.path.splitext(md_path)[0] + ".pdf"
    return convert_text(text, pdf_path, options,
                        base_dir=os.path.dirname(os.path.abspath(md_path)))


# ------------------------------------------------------------------ CLI -----

if __name__ == "__main__":
    import argparse

    ap = argparse.ArgumentParser(description="Markdown -> PDF (A4, Word-like layout)")
    ap.add_argument("input", help="input .md file")
    ap.add_argument("-o", "--output", help="output .pdf path")
    ap.add_argument("--font", default=DEFAULT_FONT, help="body font family")
    ap.add_argument("--size", type=float, default=11.0, help="base text size (pt)")
    ap.add_argument("--margins", default=DEFAULT_MARGINS,
                    choices=list(MARGIN_PRESETS))
    ap.add_argument("--toc", action="store_true", help="include table of contents")
    ap.add_argument("--no-page-numbers", action="store_true")
    ap.add_argument("--no-highlight", action="store_true")
    ap.add_argument("--no-inline-svg", action="store_true",
                    help="ignore <svg> written into the markdown; render only "
                         "linked files, ![](diagrams/x.svg)")
    ap.add_argument("--heading-colors", metavar="H1,H2,…",
                    help="comma-separated heading colours, H1 first "
                         f"(default {','.join(HEADING_COLORS.values())})")
    args = ap.parse_args()

    o = ConvertOptions(font_family=args.font, base_size=args.size,
                       margins=args.margins, toc=args.toc,
                       page_numbers=not args.no_page_numbers,
                       highlight_code=not args.no_highlight,
                       inline_svg=not args.no_inline_svg,
                       heading_colors=heading_colors_from_meta(
                           {"heading_colors": args.heading_colors or ""}))
    out = convert_file(args.input, args.output, o)
    print("wrote", out)
