# -*- coding: utf-8 -*-
from pathlib import Path
import re
import tkinter as tk
from tkinter import ttk, messagebox

BASE_DIR = Path(__file__).resolve().parent
JOURNAL_DIR = BASE_DIR / "업무일지"


class JournalViewer(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("완속충전기 업무일지")
        self.geometry("1280x820")
        self.minsize(920, 640)
        self.configure(bg="#f3f5f7")
        self.documents, self.filtered = [], []
        self.search_var = tk.StringVar()
        self._styles()
        self._layout()
        self._load()

    def _styles(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Side.TFrame", background="#102f2b")
        style.configure("Main.TFrame", background="#f3f5f7")
        style.configure("Card.TFrame", background="#ffffff")
        style.configure("Title.TLabel", background="#f3f5f7", foreground="#172026", font=("Malgun Gothic", 24, "bold"))
        style.configure("Sub.TLabel", background="#f3f5f7", foreground="#66747c", font=("Malgun Gothic", 10))
        style.configure("SideTitle.TLabel", background="#102f2b", foreground="#ffffff", font=("Malgun Gothic", 20, "bold"))
        style.configure("SideText.TLabel", background="#102f2b", foreground="#bdd0cc", font=("Malgun Gothic", 9))

    def _layout(self):
        side = ttk.Frame(self, style="Side.TFrame", width=310)
        side.pack(side="left", fill="y")
        side.pack_propagate(False)
        ttk.Label(side, text="WORK JOURNAL", style="SideText.TLabel").pack(anchor="w", padx=24, pady=(28, 2))
        ttk.Label(side, text="완속충전기\n업무일지", style="SideTitle.TLabel").pack(anchor="w", padx=24)
        ttk.Label(side, text="날짜별 업무와 대화별 목적·결과를\n편하게 찾아보세요.", style="SideText.TLabel", justify="left").pack(anchor="w", padx=24, pady=(8, 18))

        search = tk.Entry(side, textvariable=self.search_var, bg="#24443f", fg="white",
                          insertbackground="white", relief="flat", font=("Malgun Gothic", 10))
        search.pack(fill="x", padx=20, ipady=9)
        search.bind("<KeyRelease>", lambda _e: self._filter())
        tk.Label(side, text="날짜별 기록", bg="#102f2b", fg="#8fb0aa",
                 font=("Malgun Gothic", 9, "bold")).pack(anchor="w", padx=24, pady=(22, 8))

        list_wrap = tk.Frame(side, bg="#102f2b")
        list_wrap.pack(fill="both", expand=True, padx=15, pady=(0, 20))
        self.doc_list = tk.Listbox(list_wrap, bg="#102f2b", fg="#dce9e6",
                                   selectbackground="#2b5e55", selectforeground="white",
                                   relief="flat", highlightthickness=0, activestyle="none",
                                   font=("Malgun Gothic", 10), exportselection=False)
        list_scroll = ttk.Scrollbar(list_wrap, orient="vertical", command=self.doc_list.yview)
        self.doc_list.configure(yscrollcommand=list_scroll.set)
        self.doc_list.pack(side="left", fill="both", expand=True)
        list_scroll.pack(side="right", fill="y")
        self.doc_list.bind("<<ListboxSelect>>", self._select)

        main = ttk.Frame(self, style="Main.TFrame")
        main.pack(side="left", fill="both", expand=True)
        header = ttk.Frame(main, style="Main.TFrame")
        header.pack(fill="x", padx=42, pady=(34, 18))
        self.heading = ttk.Label(header, text="업무 기록 모아보기", style="Title.TLabel")
        self.heading.pack(anchor="w")
        self.subheading = ttk.Label(header, text="문서를 불러오는 중입니다.", style="Sub.TLabel")
        self.subheading.pack(anchor="w", pady=(4, 0))

        card = ttk.Frame(main, style="Card.TFrame")
        card.pack(fill="both", expand=True, padx=42, pady=(0, 38))
        self.text = tk.Text(card, wrap="word", bg="#ffffff", fg="#172026", relief="flat",
                            padx=34, pady=28, font=("Malgun Gothic", 10),
                            spacing1=2, spacing3=5, selectbackground="#b8ddd5")
        text_scroll = ttk.Scrollbar(card, orient="vertical", command=self.text.yview)
        self.text.configure(yscrollcommand=text_scroll.set)
        self.text.pack(side="left", fill="both", expand=True)
        text_scroll.pack(side="right", fill="y")
        self.text.tag_configure("h1", font=("Malgun Gothic", 20, "bold"), spacing1=8, spacing3=16)
        self.text.tag_configure("h2", font=("Malgun Gothic", 15, "bold"), foreground="#146b5d", spacing1=20, spacing3=10)
        self.text.tag_configure("h3", font=("Malgun Gothic", 12, "bold"), foreground="#8b5714", background="#fff8ed", spacing1=14, spacing3=8)
        self.text.tag_configure("bullet", lmargin1=18, lmargin2=34, spacing1=2, spacing3=3)
        self.text.tag_configure("table", font=("Consolas", 9), background="#f4f8f7")
        self.text.tag_configure("muted", foreground="#66747c")
        self.text.tag_configure("highlight", background="#ffe29d")
        self.text.configure(state="disabled")

    def _load(self):
        if not JOURNAL_DIR.exists():
            messagebox.showerror("업무일지 없음", f"폴더를 찾을 수 없습니다:\n{JOURNAL_DIR}")
            return
        for path in JOURNAL_DIR.glob("*.md"):
            try:
                content = path.read_text(encoding="utf-8")
            except UnicodeDecodeError:
                content = path.read_text(encoding="cp949", errors="replace")
            self.documents.append({"name": path.name, "text": content})
        self.documents.sort(key=lambda d: (0 if "종합" in d["name"] else 1, d["name"]), reverse=False)
        dated = [d for d in self.documents if "종합" not in d["name"]]
        summary = [d for d in self.documents if "종합" in d["name"]]
        self.documents = summary + list(reversed(dated))
        self.subheading.configure(text=f"{len(self.documents)}개 문서 · 검색하거나 날짜를 선택하세요.")
        self._filter()
        if self.filtered:
            self.doc_list.selection_set(0)
            self._show(self.filtered[0])

    def _filter(self):
        query = self.search_var.get().strip().lower()
        self.filtered = [d for d in self.documents if not query or query in (d["name"] + "\n" + d["text"]).lower()]
        self.doc_list.delete(0, "end")
        for doc in self.filtered:
            count = len(re.findall(r"^###\s+", doc["text"], flags=re.MULTILINE))
            label = "종합 업무일지" if "종합" in doc["name"] else doc["name"].replace(".md", "")
            self.doc_list.insert("end", f"  {label}   · {count}건")
        if not self.filtered:
            self._replace_text("검색 결과가 없습니다.", "muted")

    def _select(self, _event=None):
        selected = self.doc_list.curselection()
        if selected:
            self._show(self.filtered[selected[0]])

    def _show(self, doc):
        name = "종합 업무일지" if "종합" in doc["name"] else doc["name"].replace(".md", "")
        self.heading.configure(text=name)
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        table = []
        for raw in doc["text"].replace("\r", "").split("\n"):
            line = raw.rstrip()
            if line.startswith("|") and line.endswith("|"):
                table.append(line)
                continue
            if table:
                self._table(table)
                table = []
            if not line:
                self.text.insert("end", "\n")
            elif line.startswith("### "):
                self.text.insert("end", line[4:] + "\n", "h3")
            elif line.startswith("## "):
                self.text.insert("end", line[3:] + "\n", "h2")
            elif line.startswith("# "):
                self.text.insert("end", line[2:] + "\n", "h1")
            elif line.startswith("- "):
                self.text.insert("end", "• " + self._plain(line[2:]) + "\n", "bullet")
            else:
                self.text.insert("end", self._plain(line) + "\n")
        if table:
            self._table(table)
        query = self.search_var.get().strip()
        if query:
            point = "1.0"
            while True:
                found = self.text.search(query, point, stopindex="end", nocase=True)
                if not found:
                    break
                end = f"{found}+{len(query)}c"
                self.text.tag_add("highlight", found, end)
                point = end
        self.text.configure(state="disabled")
        self.text.yview_moveto(0)

    @staticmethod
    def _plain(value):
        value = re.sub(r"`([^`]+)`", r"\1", value)
        return re.sub(r"\*\*([^*]+)\*\*", r"\1", value)

    def _table(self, lines):
        rows = []
        for index, line in enumerate(lines):
            if index == 1 and re.fullmatch(r"\|[\s\-:|]+\|", line):
                continue
            rows.append([self._plain(c.strip()) for c in line.split("|")[1:-1]])
        if not rows:
            return
        widths = [0] * max(len(row) for row in rows)
        for row in rows:
            for i, cell in enumerate(row):
                widths[i] = min(max(widths[i], len(cell)), 34)
        for row in rows:
            values = [cell[:widths[i]].ljust(widths[i]) for i, cell in enumerate(row)]
            self.text.insert("end", "  ".join(values) + "\n", "table")
        self.text.insert("end", "\n")

    def _replace_text(self, value, tag):
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self.text.insert("end", value, tag)
        self.text.configure(state="disabled")


if __name__ == "__main__":
    JournalViewer().mainloop()
