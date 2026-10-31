# -*- coding: utf-8 -*-
"""
wordcount_gui.py - a one-window desktop app around docx_wordcount.py.

Pick a single .docx to count it, or a folder to count every .docx one level
inside it (not recursive). Results show in a selectable, copyable table on
top; below it, every manual line break / optional hyphen found is listed with
its surrounding text; the bottom pane auto-fills with a ready-to-paste quote
block using each file's "字元(含空白)" count, once counting finishes.

Reuses docx_wordcount.py's counting and mark-finding logic unchanged; this
file is presentation only, no counting logic of its own.
"""
import os
import re
import sys
import threading
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, filedialog, messagebox

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)) + r"\..")
from docx_wordcount import count_docx, find_hidden_marks, mark_glyph  # noqa: E402

COLUMNS = (
    ("file", "檔案", 260),
    ("words", "字數", 80),
    ("chars", "字元(不計空白)", 110),
    ("chars_ws", "字元(含空白)", 100),
    ("fareast", "中文字元", 80),
    ("paragraphs", "段落", 60),
)

DEFAULT_RATE = "16"


def list_docx_one_level(folder):
    """*.docx directly inside folder, non-recursive, skipping Word lock files."""
    try:
        names = sorted(os.listdir(folder))
    except OSError as exc:
        raise RuntimeError("無法讀取資料夾：%s" % exc)
    out = []
    for name in names:
        if not name.lower().endswith((".docx", ".dotx")):
            continue
        if name.startswith("~$"):
            continue
        full = os.path.join(folder, name)
        if os.path.isfile(full):
            out.append(full)
    return out


def parse_rate(text):
    """A positive float from a rate field's text, or None if empty/invalid."""
    try:
        value = float(text.strip())
    except (ValueError, AttributeError):
        return None
    return value if value > 0 else None


def build_quote_text(entries, rate_value, rate_display):
    """entries: [(label, chars_with_spaces), ...] in display order.

    rate_value: positive number, price per 1000 characters, used to compute.
    rate_display: the exact string to show in the formula (matches what the
    user typed, so "16" stays "16" instead of becoming "16.0").

    Returns the exact copy-paste block the project quotes scripts with, e.g.:

        字數：
        健身教練：8250
        女大學生：8101

        總字數： 8250+8101=16351

        總報價：16351*16/1000=262

    Returns "" when there is nothing to quote.
    """
    if not entries:
        return ""
    lines = ["字數："]
    for label, n in entries:
        lines.append("%s：%d" % (label, n))
    lines.append("")
    total = sum(n for _, n in entries)
    expr = "+".join(str(n) for _, n in entries)
    lines.append("總字數： %s=%d " % (expr, total))
    lines.append("")
    quote = int(total * rate_value / 1000 + 0.5)  # 四捨五入
    lines.append("總報價：%d*%s/1000=%d" % (total, rate_display, quote))
    return "\n".join(lines)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("劇本字數統計")
        self.geometry("900x880")
        self.minsize(700, 640)

        self._rows = []          # (path, result_dict_or_None, error_or_None)
        self._buttons = []       # every ttk.Button, so busy-lock can find them regardless of nesting

        self._build_ui()

    # ------------------------------------------------------------------ UI

    def _build_ui(self):
        pad = dict(padx=10, pady=8)

        top = ttk.Frame(self)
        top.pack(fill="x", **pad)

        self._add_button(top, "選擇資料夾…", self._pick_folder).pack(side="left")
        self._add_button(top, "選擇檔案…", self._pick_file).pack(side="left", padx=(8, 0))

        self.status_var = tk.StringVar(value="請選擇一個 .docx 檔案，或一個資料夾（會算資料夾底下第一層的所有 .docx）。")
        ttk.Label(top, textvariable=self.status_var, foreground="#555").pack(side="left", padx=(16, 0))

        total = ttk.Frame(self)
        total.pack(fill="x", **pad)

        ttk.Label(total, text="總字數").pack(side="left")
        self.total_var = tk.StringVar(value="—")
        ttk.Label(total, textvariable=self.total_var).pack(side="left", padx=(12, 12))
        self._add_button(total, "複製總字數", self._copy_total).pack(side="left")
        self._add_button(total, "複製整份表格", self._copy_table).pack(side="left", padx=(8, 0))

        ttk.Label(total, text="單價（每千字）").pack(side="left", padx=(16, 4))
        self.rate_var = tk.StringVar(value=DEFAULT_RATE)
        vcmd = (self.register(self._validate_rate_keystroke), "%P")
        rate_entry = ttk.Entry(total, textvariable=self.rate_var, width=6,
                                justify="right", validate="key", validatecommand=vcmd)
        rate_entry.pack(side="left")
        rate_entry.bind("<FocusOut>", self._normalize_rate)
        self.rate_var.trace_add("write", self._on_rate_changed)

        # Top half / bottom half, with a draggable divider between them.
        paned = ttk.PanedWindow(self, orient="vertical")
        paned.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        table_pane = ttk.Frame(paned)
        paned.add(table_pane, weight=3)

        cols = [c[0] for c in COLUMNS]
        self.tree = ttk.Treeview(table_pane, columns=cols, show="headings", selectmode="extended")
        for key, label, width in COLUMNS:
            anchor = "w" if key == "file" else "e"
            self.tree.heading(key, text=label)
            self.tree.column(key, width=width, anchor=anchor, stretch=(key == "file"))
        vsb = ttk.Scrollbar(table_pane, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        # "same font as everything else, just bold" - derive from the actual
        # default rather than naming a family/size, so it never drifts from
        # whatever plain system font the rest of the UI is using.
        bold_font = tkfont.Font(**tkfont.nametofont("TkDefaultFont").actual())
        bold_font.configure(weight="bold")
        self._bold_font = bold_font
        self.tree.tag_configure("error", foreground="#B00020")
        self.tree.tag_configure("total", font=bold_font)
        self.tree.bind("<Control-c>", self._copy_selection)

        marks_pane = ttk.Frame(paned)
        paned.add(marks_pane, weight=1)

        self.marks_var = tk.StringVar(value="手動換行／選擇性連字號")
        ttk.Label(marks_pane, textvariable=self.marks_var).pack(fill="x", pady=(8, 4))

        mbody = ttk.Frame(marks_pane)
        mbody.pack(fill="both", expand=True)
        self.mark_tree = ttk.Treeview(mbody, columns=("file", "kind", "para", "context"),
                                      show="headings", selectmode="extended", height=4)
        for key, label, width, stretch in (("file", "檔案", 200, False),
                                           ("kind", "類型", 100, False),
                                           ("para", "段落", 60, False),
                                           ("context", "前後文", 400, True)):
            self.mark_tree.heading(key, text=label)
            self.mark_tree.column(key, width=width, stretch=stretch,
                                  anchor="e" if key == "para" else "w")
        msb = ttk.Scrollbar(mbody, orient="vertical", command=self.mark_tree.yview)
        self.mark_tree.configure(yscrollcommand=msb.set)
        self.mark_tree.pack(side="left", fill="both", expand=True)
        msb.pack(side="right", fill="y")
        self.mark_tree.bind("<Control-c>", self._copy_selection)

        quote_pane = ttk.Frame(paned)
        paned.add(quote_pane, weight=2)

        qhead = ttk.Frame(quote_pane)
        qhead.pack(fill="x", pady=(0, 4))
        ttk.Label(qhead, text="報價文字").pack(side="left")
        self._add_button(qhead, "複製", self._copy_quote).pack(side="right")

        qbody = ttk.Frame(quote_pane)
        qbody.pack(fill="both", expand=True)
        self.quote_text = tk.Text(qbody, wrap="word", state="disabled")
        qsb = ttk.Scrollbar(qbody, orient="vertical", command=self.quote_text.yview)
        self.quote_text.configure(yscrollcommand=qsb.set)
        self.quote_text.pack(side="left", fill="both", expand=True)
        qsb.pack(side="right", fill="y")

    def _add_button(self, parent, text, command):
        b = ttk.Button(parent, text=text, command=command)
        self._buttons.append(b)
        return b

    # ------------------------------------------------------------- actions

    def _pick_folder(self):
        folder = filedialog.askdirectory(title="選擇資料夾（會算第一層的所有 .docx）")
        if not folder:
            return
        try:
            files = list_docx_one_level(folder)
        except RuntimeError as exc:
            messagebox.showerror("錯誤", str(exc))
            return
        if not files:
            self.status_var.set("這個資料夾第一層沒有找到 .docx 檔案：%s" % folder)
            self._render([], [])
            self._render_marks([])
            self._refresh_quote()
            return
        self.status_var.set("資料夾：%s" % folder)
        self._run(files)

    def _pick_file(self):
        path = filedialog.askopenfilename(
            title="選擇一個 .docx 檔案",
            filetypes=[("Word 文件", "*.docx *.dotx"), ("所有檔案", "*.*")],
        )
        if not path:
            return
        self.status_var.set("檔案：%s" % path)
        self._run([path])

    def _run(self, files):
        self._set_busy(True)
        threading.Thread(target=self._count_worker, args=(files,), daemon=True).start()

    def _count_worker(self, files):
        rows = []
        marks = []               # (path, kind, paragraph_no, before, after)
        for f in files:
            try:
                r = count_docx(f)
                rows.append((f, r, None))
            except Exception as exc:  # noqa: BLE001 - surface any failure per file
                rows.append((f, None, str(exc)))
                continue
            try:
                marks.extend((f,) + m for m in find_hidden_marks(f))
            except Exception as exc:  # noqa: BLE001 - keep the count, but don't pass as "none found"
                marks.append((f, None, None, "掃描失敗：" + str(exc), ""))
        self.after(0, lambda: self._finish(rows, marks))

    def _finish(self, rows, marks):
        ok = [r for _, r, err in rows if err is None for r in [r]]
        bad = sum(1 for _, _, err in rows if err is not None)
        self._render(rows, ok)
        self._render_marks(marks)
        self._refresh_quote()
        self._set_busy(False)
        msg = "完成，共 %d 個檔案" % len(rows)
        if bad:
            msg += "，其中 %d 個讀取失敗" % bad
        self.status_var.set(msg)

    def _set_busy(self, busy):
        state = "disabled" if busy else "normal"
        for b in self._buttons:
            b.configure(state=state)
        if busy:
            self._status_before_busy = self.status_var.get()
            self.status_var.set(self.status_var.get() + "　計算中…")

    # -------------------------------------------------------------- render

    def _render(self, rows, ok_results):
        self._rows = rows
        self.tree.delete(*self.tree.get_children())

        total = dict(words=0, chars=0, chars_ws=0, fareast=0, paragraphs=0)
        for path, r, err in rows:
            name = os.path.basename(path)
            if err:
                self.tree.insert("", "end", values=(name, "錯誤：" + err, "", "", "", ""),
                                  tags=("error",))
                continue
            for k in total:
                total[k] += r[k]
            self.tree.insert("", "end", values=(
                name, format(r["words"], ","), format(r["chars"], ","),
                format(r["chars_ws"], ","), format(r["fareast"], ","), r["paragraphs"],
            ))

        if len(ok_results) > 1:
            self.tree.insert("", "end", values=(
                "總計 %d 個檔案" % len(ok_results), format(total["words"], ","),
                format(total["chars"], ","), format(total["chars_ws"], ","),
                format(total["fareast"], ","), total["paragraphs"],
            ), tags=("total",))

        self.total_var.set(format(total["words"], ","))
        self._total_raw = total["words"]

    def _render_marks(self, marks):
        self.mark_tree.delete(*self.mark_tree.get_children())
        found = 0
        for path, kind, para, before, after in marks:
            name = os.path.basename(path)
            if kind is None:
                self.mark_tree.insert("", "end", values=(name, "", "", before), tags=("error",))
                continue
            found += 1
            # Word's own editing-mark glyphs: ↵ / ¬ for the mark, ¶ where the paragraph ends
            context = "%s【%s】%s" % (before, mark_glyph(kind), after if after else "¶")
            self.mark_tree.insert("", "end", values=(name, kind, para, context))
        self.mark_tree.tag_configure("error", foreground="#B00020")
        self.marks_var.set("手動換行／選擇性連字號：%d 處" % found)

    def _refresh_quote(self):
        entries = [(os.path.splitext(os.path.basename(p))[0], r["chars_ws"])
                   for p, r, err in self._rows if err is None]
        rate = parse_rate(self.rate_var.get())
        if rate is None:
            return  # mid-edit (e.g. field momentarily empty) - leave existing text alone
        text = build_quote_text(entries, rate, self.rate_var.get().strip())
        self.quote_text.configure(state="normal")
        self.quote_text.delete("1.0", "end")
        self.quote_text.insert("1.0", text)
        self.quote_text.configure(state="disabled")

    # ----------------------------------------------------------- rate field

    def _validate_rate_keystroke(self, proposed):
        return proposed == "" or re.fullmatch(r"\d*\.?\d*", proposed) is not None

    def _normalize_rate(self, event=None):
        if parse_rate(self.rate_var.get()) is None:
            self.rate_var.set(DEFAULT_RATE)

    def _on_rate_changed(self, *args):
        if self._rows:
            self._refresh_quote()

    # ---------------------------------------------------------------- copy

    def _copy_total(self):
        if not self._rows:
            return
        self._to_clipboard(str(self._total_raw))
        self.status_var.set("已複製總字數：%s" % self._total_raw)

    def _copy_table(self):
        if not self._rows:
            return
        lines = ["\t".join(c[1] for c in COLUMNS)]
        for path, r, err in self._rows:
            name = os.path.basename(path)
            if err:
                lines.append("%s\t錯誤：%s" % (name, err))
                continue
            lines.append("\t".join(str(x) for x in (
                name, r["words"], r["chars"], r["chars_ws"], r["fareast"], r["paragraphs"],
            )))
        self._to_clipboard("\n".join(lines))
        self.status_var.set("已複製整份表格（可直接貼到 Excel）")

    def _copy_quote(self):
        text = self.quote_text.get("1.0", "end-1c")
        if not text:
            return
        self._to_clipboard(text)
        self.status_var.set("已複製報價文字")

    def _copy_selection(self, event=None):
        tree = event.widget if event is not None else self.tree
        sel = tree.selection()
        if not sel:
            return
        lines = ["\t".join(str(v) for v in tree.item(i, "values")) for i in sel]
        self._to_clipboard("\n".join(lines))
        self.status_var.set("已複製選取的 %d 列" % len(sel))

    def _to_clipboard(self, text):
        self.clipboard_clear()
        self.clipboard_append(text)
        self.update()  # keep clipboard content after the window loses focus


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
