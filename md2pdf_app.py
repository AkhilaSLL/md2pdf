"""
md2pdf_app.py — Markdown ⇄ PDF Studio (GUI).

Light, modern desktop front-end for md2pdf_engine (Markdown → PDF) and
pdf2md_engine (PDF → Markdown). Packages into a single Windows .exe with
PyInstaller — see README.md for the exact build command.
"""

from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import traceback
from tkinter import filedialog, messagebox

import customtkinter as ctk

import md2pdf_engine as engine
import pdf2md_engine

# ------------------------------------------------------------------ theme ---

BG = "#F4F5F7"
CARD = "#FFFFFF"
CARD_BORDER = "#E5E7EB"
INK = "#111827"
SUBTEXT = "#6B7280"
ACCENT = "#4F46E5"
ACCENT_HOVER = "#4338CA"
ACCENT_SOFT = "#EEF2FF"
FIELD = "#F9FAFB"
OK = "#16A34A"
ERR = "#DC2626"

MD2PDF = "Markdown → PDF"
PDF2MD = "PDF → Markdown"

SVG_INLINE = "Inline  <svg>"
SVG_LINKED = "Linked .svg files"
SVG_NOTE = {
    SVG_INLINE: ("Needs: the whole <svg>…</svg> in the .md (blank lines inside "
                 "it are fine). One self-contained file — but GitHub strips "
                 "<svg> from previews. A <title> becomes the caption."),
    SVG_LINKED: ("Needs: the .md to travel with its image folder; relative "
                 "paths resolve next to the .md. Previews anywhere, GitHub "
                 "included, and a diagram can be reused. Alt text is the "
                 "caption."),
}

ctk.set_appearance_mode("light")


def resource_path(rel: str) -> str:
    """Locate bundled assets both in source and in a PyInstaller onefile exe."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


class App(ctk.CTk):
    def __init__(self):
        super().__init__(fg_color=BG)
        self.title("Markdown ⇄ PDF Studio")
        # tall enough that the result buttons appearing after a conversion
        # don't squash the cards above them (the Markdown -> PDF side needs
        # ~905 px once the SVG note and the result row are both showing)
        self.geometry("760x910")
        self.minsize(680, 720)
        try:
            self.iconbitmap(resource_path(os.path.join("assets", "icon.ico")))
        except Exception:
            pass

        self._queue: queue.Queue = queue.Queue()
        self._last_out: str | None = None
        self._fonts = engine.list_font_families() or ["Helvetica (built-in)"]
        # Aptos ships with Office, so fall back to whatever this machine has
        default_font = next((f for f in (engine.DEFAULT_FONT, "Calibri",
                                         "Segoe UI") if f in self._fonts),
                            self._fonts[0])

        wrap = ctk.CTkFrame(self, fg_color="transparent")
        wrap.pack(fill="both", expand=True, padx=22, pady=(18, 14))

        # ---------------------------------------------------------- header --
        head = ctk.CTkFrame(wrap, fg_color="transparent")
        head.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(head, text="Markdown  ⇄  PDF", text_color=INK,
                     font=ctk.CTkFont(size=24, weight="bold")).pack(anchor="w")
        self.subtitle = ctk.CTkLabel(
            head, text="", text_color=SUBTEXT, font=ctk.CTkFont(size=13))
        self.subtitle.pack(anchor="w", pady=(2, 0))

        # ------------------------------------------------------- direction --
        self.direction = ctk.CTkSegmentedButton(
            wrap, values=[MD2PDF, PDF2MD], command=self._switch_direction,
            height=38, fg_color=CARD, selected_color=ACCENT,
            selected_hover_color=ACCENT_HOVER,
            unselected_color=CARD, unselected_hover_color=ACCENT_SOFT,
            text_color=INK, border_width=1,
            font=ctk.CTkFont(size=14, weight="bold"))
        self.direction.set(MD2PDF)
        self.direction.pack(fill="x", pady=(0, 12))

        # ---------------------------------------------------------- source --
        src = self._card(wrap, "Source")
        self.src_mode = ctk.CTkSegmentedButton(
            src, values=["Markdown file", "Paste text"],
            command=self._switch_source,
            fg_color=FIELD, selected_color=ACCENT,
            selected_hover_color=ACCENT_HOVER,
            unselected_color=FIELD, unselected_hover_color="#ECEEF2",
            text_color=INK)
        self.src_mode.set("Markdown file")
        self.src_mode.pack(anchor="w", padx=16, pady=(4, 10))

        self.file_row = ctk.CTkFrame(src, fg_color="transparent")
        self.file_row.pack(fill="x", padx=16, pady=(0, 14))
        self.in_entry = self._entry(self.file_row, "Choose a .md file…")
        self.in_entry.pack(side="left", fill="x", expand=True)
        ctk.CTkButton(self.file_row, text="Browse", width=92,
                      command=self._browse_input,
                      fg_color=ACCENT_SOFT, hover_color="#E0E7FF",
                      text_color=ACCENT,
                      font=ctk.CTkFont(size=13, weight="bold")
                      ).pack(side="left", padx=(10, 0))

        self.paste_box = ctk.CTkTextbox(
            src, height=140, fg_color=FIELD, text_color=INK,
            border_color=CARD_BORDER, border_width=1, corner_radius=8,
            font=ctk.CTkFont(family="Consolas", size=12))
        # (packed on demand by _switch_source)

        # --------------------------------------------------------- options --
        self.opt_card = self._card(wrap, "Options")

        # -- Markdown → PDF options -------------------------------------
        self.opts_md = ctk.CTkFrame(self.opt_card, fg_color="transparent")
        grid = ctk.CTkFrame(self.opts_md, fg_color="transparent")
        grid.pack(fill="x")
        grid.grid_columnconfigure((0, 1), weight=1, uniform="col")

        self._field_label(grid, "Font family").grid(row=0, column=0, sticky="w")
        self.font_menu = ctk.CTkOptionMenu(
            grid, values=self._fonts, width=240,
            fg_color=FIELD, button_color=ACCENT, button_hover_color=ACCENT_HOVER,
            text_color=INK, dropdown_fg_color=CARD, dropdown_text_color=INK,
            dropdown_hover_color=ACCENT_SOFT)
        self.font_menu.set(default_font)
        self.font_menu.grid(row=1, column=0, sticky="ew", padx=(0, 10), pady=(4, 10))

        self._field_label(grid, "Margins").grid(row=0, column=1, sticky="w")
        self.margin_menu = ctk.CTkOptionMenu(
            grid, values=list(engine.MARGIN_PRESETS), width=240,
            fg_color=FIELD, button_color=ACCENT, button_hover_color=ACCENT_HOVER,
            text_color=INK, dropdown_fg_color=CARD, dropdown_text_color=INK,
            dropdown_hover_color=ACCENT_SOFT)
        self.margin_menu.set(engine.DEFAULT_MARGINS)
        self.margin_menu.grid(row=1, column=1, sticky="ew", pady=(4, 10))

        self.size_label = self._field_label(grid, "")
        self.size_label.grid(row=2, column=0, columnspan=2, sticky="w")
        self.size_slider = ctk.CTkSlider(
            grid, from_=8, to=16, number_of_steps=16,
            command=self._size_changed,
            progress_color=ACCENT, button_color=ACCENT,
            button_hover_color=ACCENT_HOVER, fg_color="#E5E7EB")
        self.size_slider.set(11)
        self.size_slider.grid(row=3, column=0, columnspan=2, sticky="ew",
                              pady=(6, 12))
        self._size_changed(11)

        md_checks = ctk.CTkFrame(self.opts_md, fg_color="transparent")
        md_checks.pack(fill="x")
        self.toc_var = ctk.BooleanVar(value=True)
        self.pn_var = ctk.BooleanVar(value=True)
        self.hl_var = ctk.BooleanVar(value=True)
        for i, (label, var) in enumerate((
                ("Table of contents", self.toc_var),
                ("Page numbers", self.pn_var),
                ("Colour in code blocks", self.hl_var))):
            self._check(md_checks, label, var).grid(
                row=i // 2, column=i % 2, sticky="w", padx=(0, 24), pady=4)

        # -- where diagrams come from ------------------------------------
        svg_box = ctk.CTkFrame(self.opts_md, fg_color="transparent")
        svg_box.pack(fill="x", pady=(8, 0))
        svg_head = ctk.CTkFrame(svg_box, fg_color="transparent")
        svg_head.pack(fill="x")
        self._field_label(svg_head, "Diagrams (SVG)").pack(side="left",
                                                           padx=(0, 14))
        self.svg_mode = ctk.CTkSegmentedButton(
            svg_head, values=[SVG_INLINE, SVG_LINKED], height=26,
            command=self._switch_svg_mode,
            fg_color=FIELD, selected_color=ACCENT,
            selected_hover_color=ACCENT_HOVER,
            unselected_color=FIELD, unselected_hover_color="#ECEEF2",
            text_color=INK, font=ctk.CTkFont(size=12))
        self.svg_mode.set(SVG_INLINE)
        self.svg_mode.pack(side="left")
        self.svg_note = ctk.CTkLabel(
            svg_box, text="", text_color=SUBTEXT,
            font=ctk.CTkFont(size=12), wraplength=670, justify="left")
        self.svg_note.pack(anchor="w", pady=(4, 0))
        self._switch_svg_mode(SVG_INLINE)

        # -- PDF → Markdown options -------------------------------------
        self.opts_pdf = ctk.CTkFrame(self.opt_card, fg_color="transparent")
        self.img_var = ctk.BooleanVar(value=True)
        self.tbl_var = ctk.BooleanVar(value=True)
        self.code_var = ctk.BooleanVar(value=True)
        self.pm_var = ctk.BooleanVar(value=True)
        self.hf_var = ctk.BooleanVar(value=True)
        self.mtoc_var = ctk.BooleanVar(value=False)
        for i, (label, var) in enumerate((
                ("Extract images to a folder", self.img_var),
                ("Detect tables", self.tbl_var),
                ("Detect code blocks", self.code_var),
                ("Page markers  (<!-- page N -->)", self.pm_var),
                ("Skip headers / footers", self.hf_var),
                ("Linked contents section", self.mtoc_var))):
            self._check(self.opts_pdf, label, var).grid(
                row=i // 2, column=i % 2, sticky="w", padx=(0, 24), pady=4)
        ctk.CTkLabel(
            self.opts_pdf,
            text="Images are saved next to the .md as  <name>_images/  and "
                 "referenced with relative paths — markdown stays light "
                 "enough to paste straight into an LLM.",
            text_color=SUBTEXT, font=ctk.CTkFont(size=12),
            wraplength=640, justify="left",
        ).grid(row=3, column=0, columnspan=2, sticky="w", pady=(8, 0))

        # shared "open when done"
        shared = ctk.CTkFrame(self.opt_card, fg_color="transparent")
        shared.pack(side="bottom", fill="x", padx=16, pady=(0, 14))
        self.open_var = ctk.BooleanVar(value=True)
        self._check(shared, "Open result when done", self.open_var
                    ).pack(anchor="w", pady=(6, 0))

        self.opts_md.pack(fill="x", padx=16, pady=(2, 0))   # default mode

        # ---------------------------------------------------------- output --
        out = self._card(wrap, "Output")
        out_row = ctk.CTkFrame(out, fg_color="transparent")
        out_row.pack(fill="x", padx=16, pady=(4, 14))
        self.out_entry = self._entry(out_row, "Where to save the PDF…")
        self.out_entry.pack(side="left", fill="x", expand=True)
        ctk.CTkButton(out_row, text="Save as", width=92,
                      command=self._browse_output,
                      fg_color=ACCENT_SOFT, hover_color="#E0E7FF",
                      text_color=ACCENT,
                      font=ctk.CTkFont(size=13, weight="bold")
                      ).pack(side="left", padx=(10, 0))

        # --------------------------------------------------------- actions --
        actions = ctk.CTkFrame(wrap, fg_color="transparent")
        actions.pack(fill="x", pady=(4, 0))
        self.convert_btn = ctk.CTkButton(
            actions, text="Convert to PDF", height=44, corner_radius=10,
            fg_color=ACCENT, hover_color=ACCENT_HOVER,
            font=ctk.CTkFont(size=15, weight="bold"),
            command=self._convert)
        self.convert_btn.pack(fill="x")

        self.progress = ctk.CTkProgressBar(actions, mode="indeterminate",
                                           progress_color=ACCENT,
                                           fg_color="#E5E7EB", height=6)
        self.status = ctk.CTkLabel(actions, text="", text_color=SUBTEXT,
                                   font=ctk.CTkFont(size=13), anchor="w",
                                   justify="left", wraplength=660)
        self.status.pack(fill="x", pady=(10, 0))

        self.result_row = ctk.CTkFrame(actions, fg_color="transparent")
        ctk.CTkButton(self.result_row, text="Open result", width=110,
                      command=self._open_result, fg_color=CARD,
                      hover_color=ACCENT_SOFT, text_color=ACCENT,
                      border_color=ACCENT, border_width=1,
                      font=ctk.CTkFont(size=13, weight="bold")
                      ).pack(side="left", padx=(0, 10))
        ctk.CTkButton(self.result_row, text="Show in folder", width=120,
                      command=self._open_folder, fg_color=CARD,
                      hover_color=ACCENT_SOFT, text_color=ACCENT,
                      border_color=ACCENT, border_width=1,
                      font=ctk.CTkFont(size=13, weight="bold")
                      ).pack(side="left")

        self._apply_direction(MD2PDF)
        self._paint_segments(self.direction)
        self._paint_segments(self.src_mode)
        self.after(120, self._poll)

    # ------------------------------------------------------------ widgets --

    def _card(self, parent, title):
        card = ctk.CTkFrame(parent, fg_color=CARD, corner_radius=12,
                            border_color=CARD_BORDER, border_width=1)
        card.pack(fill="x", pady=(0, 12))
        ctk.CTkLabel(card, text=title.upper(), text_color=SUBTEXT,
                     font=ctk.CTkFont(size=11, weight="bold")
                     ).pack(anchor="w", padx=16, pady=(12, 6))
        return card

    def _entry(self, parent, placeholder):
        return ctk.CTkEntry(parent, placeholder_text=placeholder, height=36,
                            fg_color=FIELD, text_color=INK,
                            placeholder_text_color="#9CA3AF",
                            border_color=CARD_BORDER, border_width=1,
                            corner_radius=8, font=ctk.CTkFont(size=13))

    def _field_label(self, parent, text):
        return ctk.CTkLabel(parent, text=text, text_color=INK,
                            font=ctk.CTkFont(size=13, weight="bold"))

    @staticmethod
    def _paint_segments(seg):
        """White label on the selected segment, dark on the rest.

        CTkSegmentedButton has a single text_color for every segment, and dark
        ink on the accent fill reads poorly — so recolour the underlying
        buttons directly after each change.
        """
        buttons = getattr(seg, "_buttons_dict", None)
        if not buttons:
            return
        current = seg.get()
        for value, button in buttons.items():
            button.configure(text_color=("#FFFFFF" if value == current else INK))

    def _check(self, parent, label, var):
        return ctk.CTkCheckBox(parent, text=label, variable=var,
                               fg_color=ACCENT, hover_color=ACCENT_HOVER,
                               checkmark_color="white", text_color=INK,
                               border_color="#C7CBD4", corner_radius=6,
                               font=ctk.CTkFont(size=13))

    # ------------------------------------------------------------ handlers --

    def _is_md2pdf(self) -> bool:
        return self.direction.get() == MD2PDF

    def _switch_direction(self, mode):
        self._paint_segments(self.direction)
        self._apply_direction(mode)

    def _apply_direction(self, mode):
        md = mode == MD2PDF
        # subtitle
        self.subtitle.configure(text=(
            "A4 · Word-style margins · bookmarks, links, images and code preserved"
            if md else
            "Clean, LLM-friendly markdown · headings, tables, code and images recovered"))
        # source card
        if md:
            self.src_mode.pack(anchor="w", padx=16, pady=(4, 10))
            self._switch_source(self.src_mode.get())
            self.in_entry.configure(placeholder_text="Choose a .md file…")
        else:
            self.src_mode.pack_forget()
            self.paste_box.pack_forget()
            self.file_row.pack(fill="x", padx=16, pady=(0, 14))
            self.in_entry.configure(placeholder_text="Choose a .pdf file…")
        # options card
        if md:
            self.opts_pdf.pack_forget()
            self.opts_md.pack(fill="x", padx=16, pady=(2, 0))
        else:
            self.opts_md.pack_forget()
            self.opts_pdf.pack(fill="x", padx=16, pady=(2, 0))
        # output + button
        self.out_entry.configure(placeholder_text=(
            "Where to save the PDF…" if md else "Where to save the Markdown…"))
        self.convert_btn.configure(
            text="Convert to PDF" if md else "Convert to Markdown")
        # keep paths in step when a file is already chosen
        in_path = self.in_entry.get().strip()
        want, other = (".md", ".pdf") if md else (".pdf", ".md")
        if in_path and in_path.lower().endswith(other):
            self.in_entry.delete(0, "end")
        in_path = self.in_entry.get().strip()
        self.out_entry.delete(0, "end")
        if in_path:
            self.out_entry.insert(0, os.path.splitext(in_path)[0] +
                                  (".pdf" if md else ".md"))
        self._set_status("", SUBTEXT)
        self.result_row.pack_forget()

    def _switch_svg_mode(self, mode):
        self._paint_segments(self.svg_mode)
        self.svg_note.configure(text=SVG_NOTE.get(mode, ""))

    def _switch_source(self, mode):
        self._paint_segments(self.src_mode)
        if not self._is_md2pdf():
            return
        if mode == "Paste text":
            self.file_row.pack_forget()
            self.paste_box.pack(fill="x", padx=16, pady=(0, 14))
        else:
            self.paste_box.pack_forget()
            self.file_row.pack(fill="x", padx=16, pady=(0, 14))

    def _size_changed(self, value):
        s = int(round(float(value)))
        h1 = round(s * engine.H_SCALE[1])
        self.size_label.configure(
            text=f"Base text size — {s} pt   (headings scale with it: H1 ≈ {h1} pt)")

    def _browse_input(self):
        if self._is_md2pdf():
            path = filedialog.askopenfilename(
                title="Choose a Markdown file",
                filetypes=[("Markdown", "*.md *.markdown *.mdown *.txt"),
                           ("All files", "*.*")])
            new_ext = ".pdf"
        else:
            path = filedialog.askopenfilename(
                title="Choose a PDF file",
                filetypes=[("PDF", "*.pdf"), ("All files", "*.*")])
            new_ext = ".md"
        if path:
            self.in_entry.delete(0, "end")
            self.in_entry.insert(0, path)
            # keep the output path in step with the chosen file
            self.out_entry.delete(0, "end")
            self.out_entry.insert(0, os.path.splitext(path)[0] + new_ext)

    def _browse_output(self):
        md = self._is_md2pdf()
        ext, types = ((".pdf", [("PDF", "*.pdf")]) if md
                      else (".md", [("Markdown", "*.md")]))
        initial = self.out_entry.get().strip()
        path = filedialog.asksaveasfilename(
            title=f"Save {'PDF' if md else 'Markdown'} as",
            defaultextension=ext,
            initialfile=os.path.basename(initial) if initial else f"document{ext}",
            initialdir=os.path.dirname(initial) if initial else None,
            filetypes=types)
        if path:
            self.out_entry.delete(0, "end")
            self.out_entry.insert(0, path)

    # ------------------------------------------------------------- convert --

    def _convert(self):
        if self._is_md2pdf():
            self._convert_md2pdf()
        else:
            self._convert_pdf2md()

    def _convert_md2pdf(self):
        try:
            mode = self.src_mode.get()
            if mode == "Markdown file":
                md_path = self.in_entry.get().strip()
                if not md_path or not os.path.isfile(md_path):
                    raise ValueError("Choose a Markdown file first.")
                with open(md_path, "r", encoding="utf-8-sig") as fh:
                    text = fh.read()
                base_dir = os.path.dirname(os.path.abspath(md_path))
                default_out = os.path.splitext(md_path)[0] + ".pdf"
            else:
                text = self.paste_box.get("1.0", "end").strip()
                if not text:
                    raise ValueError("Paste some Markdown first.")
                base_dir = os.path.expanduser("~")
                default_out = os.path.join(
                    os.path.expanduser("~"), "Documents", "markdown.pdf")

            out_path = self.out_entry.get().strip() or default_out
            if not out_path.lower().endswith(".pdf"):
                out_path += ".pdf"
            os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
            self.out_entry.delete(0, "end")
            self.out_entry.insert(0, out_path)
        except ValueError as e:
            self._set_status(str(e), ERR)
            return

        opts = engine.ConvertOptions(
            font_family=self.font_menu.get(),
            base_size=float(int(round(self.size_slider.get()))),
            margins=self.margin_menu.get(),
            page_numbers=self.pn_var.get(),
            toc=self.toc_var.get(),
            highlight_code=self.hl_var.get(),
            inline_svg=self.svg_mode.get() == SVG_INLINE)

        def worker():
            try:
                engine.convert_text(text, out_path, opts, base_dir=base_dir)
                self._queue.put(("ok", out_path, ""))
            except Exception:
                self._queue.put(("err", traceback.format_exc(), ""))

        self._start_work("Building PDF…", worker)

    def _convert_pdf2md(self):
        try:
            pdf_path = self.in_entry.get().strip()
            if not pdf_path or not os.path.isfile(pdf_path):
                raise ValueError("Choose a PDF file first.")
            out_path = (self.out_entry.get().strip()
                        or os.path.splitext(pdf_path)[0] + ".md")
            if not out_path.lower().endswith((".md", ".markdown", ".txt")):
                out_path += ".md"
            os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
            self.out_entry.delete(0, "end")
            self.out_entry.insert(0, out_path)
        except ValueError as e:
            self._set_status(str(e), ERR)
            return

        opts = pdf2md_engine.Pdf2MdOptions(
            extract_images=self.img_var.get(),
            detect_tables=self.tbl_var.get(),
            detect_code=self.code_var.get(),
            page_markers=self.pm_var.get(),
            skip_margins=self.hf_var.get(),
            linked_toc=self.mtoc_var.get())

        def worker():
            try:
                stats = pdf2md_engine.convert(pdf_path, out_path, opts)
                extra = f"{stats['pages']} pages"
                if stats["images"]:
                    extra += (f", {stats['images']} images → "
                              f"{os.path.basename(stats['image_dir'])}/")
                self._queue.put(("ok", stats["md_path"], extra))
            except Exception:
                self._queue.put(("err", traceback.format_exc(), ""))

        self._start_work("Reading PDF…", worker)

    def _start_work(self, message, worker):
        self.convert_btn.configure(state="disabled", text="Converting…")
        self.result_row.pack_forget()
        self.progress.pack(fill="x", pady=(12, 0))
        self.progress.start()
        self._set_status(message, SUBTEXT)
        threading.Thread(target=worker, daemon=True).start()

    def _poll(self):
        try:
            kind, payload, extra = self._queue.get_nowait()
        except queue.Empty:
            self.after(120, self._poll)
            return
        self.progress.stop()
        self.progress.pack_forget()
        self.convert_btn.configure(
            state="normal",
            text="Convert to PDF" if self._is_md2pdf() else "Convert to Markdown")
        if kind == "ok":
            self._last_out = payload
            note = f"   ({extra})" if extra else ""
            self._set_status(f"Saved  ✓   {payload}{note}", OK)
            self.result_row.pack(fill="x", pady=(10, 0))
            if self.open_var.get():
                self._open_result()
        else:
            short = payload.strip().splitlines()[-1]
            self._set_status(f"Conversion failed: {short}", ERR)
            messagebox.showerror("Conversion failed", payload[-1500:])
        self.after(120, self._poll)

    def _set_status(self, text, color):
        self.status.configure(text=text, text_color=color)

    def _open_result(self):
        if self._last_out and os.path.isfile(self._last_out):
            self._open_path(self._last_out)

    def _open_folder(self):
        if self._last_out:
            if os.name == "nt":
                subprocess.Popen(["explorer", "/select,",
                                  os.path.normpath(self._last_out)])
            else:
                self._open_path(os.path.dirname(self._last_out))

    @staticmethod
    def _open_path(path):
        try:
            if os.name == "nt":
                os.startfile(path)  # noqa
            elif sys.platform == "darwin":
                subprocess.Popen(["open", path])
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception:
            pass


if __name__ == "__main__":
    App().mainloop()
