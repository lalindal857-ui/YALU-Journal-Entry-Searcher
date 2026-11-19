import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook
from pathlib import Path

class App:
    def __init__(self, root):
        self.root = root
        root.title("YALU Journal Entry Searcher")
        root.geometry("520x190")
        root.minsize(420, 160)
        root.resizable(True, True)
        root.configure(bg="#eef2f7")
        self.rows, self.headers = [], []
        self.entry_col = self.desc_col = self.cost_col = None
        self.make_ui()
        root.bind("<plus>", self.plus); root.bind("<KP_Add>", self.plus)
        root.bind("<minus>", self.minus); root.bind("<KP_Subtract>", self.minus)

    def make_ui(self):
        bar = tk.Frame(self.root, bg="#eef2f7"); bar.pack(fill="x", padx=14, pady=(18, 5))
        tk.Button(bar, text="📂", command=self.open_file, font=("Segoe UI Emoji", 15),
                  width=4, bg="#3978bd", fg="white", relief="flat").pack(side="left")
        self.entry = tk.Entry(bar, font=("Segoe UI", 14), justify="center",
                              bg="white", fg="#202a36", relief="flat")
        self.entry.pack(side="left", fill="x", expand=True, padx=10, ipady=7)
        self.entry.bind("<Return>", lambda e: self.search_copy())
        tk.Button(bar, text="▤", command=self.search_copy, font=("Segoe UI Symbol", 16),
                  width=4, bg="#3978bd", fg="white", relief="flat").pack(side="right")
        self.cost = tk.StringVar(value="Cost Center: —")
        tk.Label(self.root, textvariable=self.cost, bg="#eef2f7", fg="#475569",
                 font=("Segoe UI", 8)).pack(pady=(3, 0))
        self.status = tk.StringVar(value="Select file")
        tk.Label(self.root, textvariable=self.status, bg="#eef2f7", fg="#667085",
                 font=("Segoe UI", 9)).pack(pady=(4, 0))
        tk.Label(self.root, text="Prepared By Lakshan", bg="#eef2f7", fg="#8a94a3",
                 font=("Segoe UI", 7)).pack(side="bottom", pady=6)

    def open_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xlsm")])
        if not path: return
        try:
            wb = load_workbook(path, data_only=True, read_only=True)
            values = list(wb.active.iter_rows(values_only=True))
            if not values: raise ValueError("Worksheet is empty.")
            self.headers = [str(v).strip().lower() if v is not None else "" for v in values[0]]
            self.rows = values[1:]
            self.entry_col = self.col(["entry no", "entry number", "journal no", "journal number", "document no", "doc no", "voucher no"])
            self.desc_col = self.col(["description", "memo", "details", "narration"])
            self.cost_col = self.col(["cost center", "costcentre", "cost centre"])
            if self.entry_col is None or self.desc_col is None:
                raise ValueError("Could not find Entry No. and Description headings in row 1.")
            self.status.set("Loaded: " + Path(path).name)
        except Exception as e: messagebox.showerror("Excel error", str(e))

    def col(self, names):
        for i, h in enumerate(self.headers):
            if any(n in h for n in names): return i
        return None

    def cell(self, row, i):
        return "" if i is None or i >= len(row) or row[i] is None else str(row[i]).strip()

    def search_copy(self):
        if not self.rows: self.status.set("Please select an Excel file"); return
        key = self.entry.get().strip()
        if not key: self.status.set("Enter Entry No."); return
        for row in self.rows:
            entry_no = self.cell(row, self.entry_col)
            if entry_no.casefold() == key.casefold():
                desc = self.cell(row, self.desc_col)
                if not desc: self.status.set("Description is empty"); return
                copied = f"{desc} ({entry_no})"
                self.root.clipboard_clear(); self.root.clipboard_append(copied); self.root.update()
                cost = self.cell(row, self.cost_col)
                self.cost.set("Cost Center: " + (cost or "—"))
                self.status.set("Prepared By Lakshan")
                self.notify()
                return
        self.cost.set("Cost Center: —"); self.status.set(f"Entry No. {key} not found")

    def notify(self):
        pop = tk.Toplevel(self.root); pop.title("Copied"); pop.configure(bg="#263b56")
        pop.resizable(False, False); pop.transient(self.root); pop.attributes("-topmost", True)
        tk.Label(pop, text="Prepared By Lakshan", bg="#263b56", fg="white",
                 font=("Segoe UI", 10, "bold"), padx=18, pady=12).pack()
        pop.update_idletasks()
        x = self.root.winfo_rootx() + max(0, (self.root.winfo_width()-pop.winfo_width())//2)
        y = self.root.winfo_rooty() + 20
        pop.geometry(f"+{x}+{y}"); pop.after(1500, pop.destroy)

    def plus(self, event=None):
        self.root.geometry(f"{min(self.root.winfo_width()+20,900)}x{min(self.root.winfo_height()+10,500)}")
    def minus(self, event=None):
        self.root.geometry(f"{max(self.root.winfo_width()-20,420)}x{max(self.root.winfo_height()-10,160)}")

if __name__ == "__main__":
    root = tk.Tk(); App(root); root.mainloop()
