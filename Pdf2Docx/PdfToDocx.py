import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import os

# ── Try importing pdf2docx ──────────────────────────────────────────────────
try:
    from pdf2docx import Converter
    PDF2DOCX_AVAILABLE = True
except ImportError:
    PDF2DOCX_AVAILABLE = False

# ── Palette ─────────────────────────────────────────────────────────────────
BG       = "#0D0F14"
PANEL    = "#13161E"
BORDER   = "#1E2330"
ACCENT   = "#4F8EF7"
ACCENT2  = "#7B5CF0"
SUCCESS  = "#34D399"
ERROR    = "#F87171"
TEXT     = "#E8EAF0"
SUBTEXT  = "#6B7280"
FONT_HDR = ("Georgia", 22, "bold")
FONT_SUB = ("Courier New", 9)
FONT_LBL = ("Courier New", 10, "bold")
FONT_BTN = ("Courier New", 10, "bold")
FONT_LOG = ("Courier New", 9)

# ── Rounded-rect helper (Canvas) ────────────────────────────────────────────
def rounded_rect(canvas, x1, y1, x2, y2, r, **kw):
    points = [
        x1+r, y1,  x2-r, y1,
        x2,   y1,  x2,   y1+r,
        x2,   y2-r, x2,  y2,
        x2-r, y2,  x1+r, y2,
        x1,   y2,  x1,   y2-r,
        x1,   y1+r, x1,  y1,
        x1+r, y1,
    ]
    return canvas.create_polygon(points, smooth=True, **kw)

# ── Drop-zone widget ─────────────────────────────────────────────────────────
class DropZone(tk.Canvas):
    def __init__(self, parent, on_file_cb, **kw):
        super().__init__(parent, bg=PANEL, highlightthickness=0,
                         relief="flat", **kw)
        self._cb = on_file_cb
        self._hovered = False
        self._draw()
        self.bind("<Configure>", lambda e: self._draw())
        self.bind("<Button-1>", self._click)
        self.bind("<Enter>", self._enter)
        self.bind("<Leave>", self._leave)

    def _draw(self, highlight=False):
        self.delete("all")
        w, h = self.winfo_width() or 460, self.winfo_height() or 160
        color = ACCENT if highlight else BORDER
        # dashed border via segments
        dash = (8, 5)
        self.create_rectangle(2, 2, w-2, h-2,
                               outline=color, width=2,
                               dash=dash)
        # icon
        mid = w // 2
        cy  = h // 2 - 18
        # PDF glyph
        self.create_text(mid, cy, text="⬆", font=("Segoe UI Symbol", 28),
                          fill=ACCENT if highlight else SUBTEXT)
        self.create_text(mid, cy + 44,
                          text="Click to upload PDF here",
                          font=FONT_LBL, fill=ACCENT if highlight else SUBTEXT)

    def _enter(self, _):
        self._draw(highlight=True)

    def _leave(self, _):
        self._draw(highlight=False)

    def _click(self, _):
        path = filedialog.askopenfilename(
            title="Select PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if path:
            self._cb(path)

    def mark_loaded(self, filename):
        self.delete("all")
        w, h = self.winfo_width() or 460, self.winfo_height() or 160
        self.create_rectangle(2, 2, w-2, h-2,
                               outline=SUCCESS, width=2, dash=(8, 5))
        mid = w // 2
        self.create_text(mid, h//2 - 14, text="✔", font=("Segoe UI Symbol", 26),
                          fill=SUCCESS)
        self.create_text(mid, h//2 + 18, text=filename,
                          font=FONT_LBL, fill=SUCCESS)

# ── Animated progress bar ────────────────────────────────────────────────────
class ProgressBar(tk.Canvas):
    def __init__(self, parent, **kw):
        super().__init__(parent, bg=BG, highlightthickness=0, height=6, **kw)
        self._progress = 0.0
        self._animating = False
        self._anim_pos  = 0.0
        self.bind("<Configure>", lambda e: self._redraw())

    def _redraw(self):
        self.delete("all")
        w = self.winfo_width() or 460
        # track
        self.create_rectangle(0, 0, w, 6, fill=BORDER, outline="")
        if self._animating:
            bw = w * 0.35
            x1 = self._anim_pos * w - bw / 2
            x2 = x1 + bw
            x1 = max(0, min(x1, w))
            x2 = max(0, min(x2, w))
            self.create_rectangle(x1, 0, x2, 6, fill=ACCENT, outline="")
        else:
            fill_w = int(w * self._progress)
            if fill_w > 0:
                self.create_rectangle(0, 0, fill_w, 6, fill=SUCCESS, outline="")

    def start_indeterminate(self):
        self._animating = True
        self._anim_pos  = 0.0
        self._tick()

    def _tick(self):
        if not self._animating:
            return
        self._anim_pos = (self._anim_pos + 0.012) % 1.35
        self._redraw()
        self.after(16, self._tick)

    def stop(self, value=1.0):
        self._animating = False
        self._progress  = value
        self._redraw()

    def reset(self):
        self._animating = False
        self._progress  = 0.0
        self._redraw()

# ── Log box ──────────────────────────────────────────────────────────────────
class LogBox(tk.Frame):
    def __init__(self, parent, **kw):
        super().__init__(parent, bg=PANEL, **kw)
        self.text = tk.Text(
            self, bg=PANEL, fg=SUBTEXT, font=FONT_LOG,
            relief="flat", bd=0, highlightthickness=0,
            state="disabled", wrap="word", height=6,
            insertbackground=TEXT, selectbackground=ACCENT,
        )
        sb = tk.Scrollbar(self, command=self.text.yview,
                          bg=BORDER, troughcolor=PANEL,
                          relief="flat", bd=0, width=6)
        self.text.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.text.pack(side="left", fill="both", expand=True)

    def log(self, msg, color=None):
        self.text.configure(state="normal")
        tag = None
        if color:
            tag = f"c{color.replace('#','')}"
            self.text.tag_configure(tag, foreground=color)
        self.text.insert("end", f"▸  {msg}\n", tag or ())
        self.text.see("end")
        self.text.configure(state="disabled")

    def clear(self):
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self.text.configure(state="disabled")

# ── Main Application ─────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF → DOCX Converter")
        self.geometry("520x680")
        self.resizable(False, False)
        self.configure(bg=BG)

        self._pdf_path  = None
        self._docx_path = None
        self._converting = False

        self._build_ui()

        if not PDF2DOCX_AVAILABLE:
            self._log.log("pdf2docx not installed. Run: pip install pdf2docx", ERROR)

    # ── UI layout ────────────────────────────────────────────────────────────
    def _build_ui(self):
        pad = dict(padx=28)

        # ── header ──────────────────────────────────────────────────────────
        hdr = tk.Frame(self, bg=BG)
        hdr.pack(fill="x", pady=(30, 4), **pad)

        tk.Label(hdr, text="PDF", font=("Georgia", 22, "bold"),
                 fg=ACCENT, bg=BG).pack(side="left")
        tk.Label(hdr, text=" → ", font=("Georgia", 22),
                 fg=SUBTEXT, bg=BG).pack(side="left")
        tk.Label(hdr, text="DOCX", font=("Georgia", 22, "bold"),
                 fg=ACCENT2, bg=BG).pack(side="left")

        tk.Label(self, text="Convert PDF documents to editable Word files",
                 font=FONT_SUB, fg=SUBTEXT, bg=BG).pack(anchor="w", **pad)

        self._sep(8)

        # ── drop zone ───────────────────────────────────────────────────────
        tk.Label(self, text="INPUT  FILE", font=FONT_LBL,
                 fg=SUBTEXT, bg=BG).pack(anchor="w", **pad)
        self._drop = DropZone(self, self._on_pdf_selected,
                               width=464, height=140)
        self._drop.pack(**pad, pady=(6, 0))

        self._sep(16)

        # ── output path ─────────────────────────────────────────────────────
        tk.Label(self, text="OUTPUT  FILE", font=FONT_LBL,
                 fg=SUBTEXT, bg=BG).pack(anchor="w", **pad)

        out_row = tk.Frame(self, bg=BG)
        out_row.pack(fill="x", **pad, pady=(6, 0))

        self._out_var = tk.StringVar(value="No output path selected")
        out_entry = tk.Entry(
            out_row, textvariable=self._out_var,
            bg=PANEL, fg=TEXT, font=FONT_LOG,
            relief="flat", bd=0, highlightthickness=1,
            highlightbackground=BORDER, highlightcolor=ACCENT,
            insertbackground=TEXT,
        )
        out_entry.pack(side="left", fill="x", expand=True,
                       ipady=8, ipadx=8)

        self._browse_btn = self._small_btn(out_row, "Browse", self._browse_out)
        self._browse_btn.pack(side="left", padx=(8, 0))

        self._sep(20)

        # ── convert button ───────────────────────────────────────────────────
        self._conv_btn = self._big_btn("Convert", self._start_convert)
        self._conv_btn.pack(**pad, fill="x")

        self._sep(12)

        # ── progress bar ─────────────────────────────────────────────────────
        self._prog = ProgressBar(self, width=464)
        self._prog.pack(**pad)

        self._sep(12)

        # ── status label ─────────────────────────────────────────────────────
        self._status_var = tk.StringVar(value="Ready")
        tk.Label(self, textvariable=self._status_var,
                 font=FONT_SUB, fg=SUBTEXT, bg=BG).pack(anchor="w", **pad)

        self._sep(10)

        # ── log box ───────────────────────────────────────────────────────────
        tk.Label(self, text="LOG", font=FONT_LBL,
                 fg=SUBTEXT, bg=BG).pack(anchor="w", **pad)
        self._log = LogBox(self, width=464)
        self._log.pack(**pad, pady=(6, 0), fill="x")

        # ── footer ────────────────────────────────────────────────────────────
        tk.Label(self, text="Powered by pdf2docx",
                 font=FONT_SUB, fg=BORDER, bg=BG).pack(pady=(16, 8))

    # ── helpers ───────────────────────────────────────────────────────────────
    def _sep(self, h):
        tk.Frame(self, bg=BG, height=h).pack()

    def _big_btn(self, text, cmd):
        c = tk.Canvas(self, bg=BG, highlightthickness=0, height=48)
        def _draw(hover=False):
            c.delete("all")
            w = c.winfo_width() or 464
            color = ACCENT2 if hover else ACCENT
            rounded_rect(c, 0, 0, w, 48, 10, fill=color, outline="")
            c.create_text(w//2, 24, text=text, font=FONT_BTN, fill=BG)
        c.bind("<Configure>", lambda e: _draw())
        c.bind("<Enter>",     lambda e: _draw(True))
        c.bind("<Leave>",     lambda e: _draw(False))
        c.bind("<Button-1>",  lambda e: cmd())
        return c

    def _small_btn(self, parent, text, cmd):
        btn = tk.Label(
            parent, text=text, font=FONT_BTN,
            bg=BORDER, fg=TEXT, cursor="hand2",
            padx=12, pady=6,
        )
        btn.bind("<Button-1>", lambda e: cmd())
        btn.bind("<Enter>", lambda e: btn.configure(bg=ACCENT, fg=BG))
        btn.bind("<Leave>", lambda e: btn.configure(bg=BORDER, fg=TEXT))
        return btn

    # ── logic ─────────────────────────────────────────────────────────────────
    def _on_pdf_selected(self, path):
        self._pdf_path = path
        name = os.path.basename(path)
        self._drop.mark_loaded(name)

        # auto-suggest output
        base = os.path.splitext(path)[0]
        suggested = base + ".docx"
        self._docx_path = suggested
        self._out_var.set(suggested)
        self._log.clear()
        self._log.log(f"Selected: {name}", ACCENT)
        self._status_var.set("PDF loaded — ready to convert")
        self._prog.reset()

    def _browse_out(self):
        initial = self._docx_path or os.path.expanduser("~")
        path = filedialog.asksaveasfilename(
            title="Save DOCX as…",
            defaultextension=".docx",
            filetypes=[("Word Document", "*.docx")],
            initialfile=os.path.basename(initial) if self._docx_path else "output.docx",
        )
        if path:
            self._docx_path = path
            self._out_var.set(path)

    def _start_convert(self):
        if self._converting:
            return
        if not self._pdf_path:
            messagebox.showwarning("No PDF", "Please select a PDF file first.")
            return
        if not self._docx_path:
            messagebox.showwarning("No output", "Please choose an output file path.")
            return
        if not PDF2DOCX_AVAILABLE:
            messagebox.showerror("Missing library",
                                 "pdf2docx is not installed.\nRun: pip install pdf2docx")
            return
        self._converting = True
        self._log.clear()
        self._prog.start_indeterminate()
        self._status_var.set("Converting…")
        threading.Thread(target=self._convert_thread, daemon=True).start()

    def _convert_thread(self):
        try:
            self._log_main(f"Input : {os.path.basename(self._pdf_path)}", ACCENT)
            self._log_main(f"Output: {os.path.basename(self._docx_path)}", ACCENT)
            self._log_main("Starting conversion…", TEXT)

            cv = Converter(self._pdf_path)
            cv.convert(self._docx_path)
            cv.close()

            self._log_main("Conversion complete!", SUCCESS)
            self.after(0, self._on_success)
        except Exception as exc:
            self._log_main(f"Error: {exc}", ERROR)
            self.after(0, self._on_error)

    def _log_main(self, msg, color=None):
        self.after(0, lambda: self._log.log(msg, color))

    def _on_success(self):
        self._converting = False
        self._prog.stop(1.0)
        self._status_var.set("Done! ✔  File saved successfully.")
        messagebox.showinfo("Success", f"Saved to:\n{self._docx_path}")

    def _on_error(self):
        self._converting = False
        self._prog.stop(0.0)
        self._status_var.set("Conversion failed. See log for details.")


# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = App()
    app.mainloop()
