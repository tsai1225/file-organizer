import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os, shutil, threading

APP_TITLE = "檔案整理工具"
BG = "#F8F8F6"
CARD = "#FFFFFF"
ACCENT = "#5348B7"
ACCENT_LIGHT = "#EEEDFE"
GREEN = "#0F6E56"
RED = "#993C1D"
GRAY = "#5F5E5A"
BORDER = "#D3D1C7"
FONT = ("微軟正黑體", 10)
FONT_B = ("微軟正黑體", 10, "bold")
FONT_H = ("微軟正黑體", 13, "bold")
FONT_SMALL = ("微軟正黑體", 9)

DEFAULT_EXTS = ".pdf, .docx, .xlsx, .doc, .xls"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("640x720")
        self.resizable(False, False)
        self.configure(bg=BG)
        self._plan = []
        self._build_ui()

    def _card(self, parent, **kw):
        return tk.Frame(parent, bg=CARD, relief="flat",
                        highlightbackground=BORDER, highlightthickness=1, **kw)

    def _build_ui(self):
        # Title bar
        hdr = tk.Frame(self, bg=ACCENT, height=52)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="  📁 " + APP_TITLE, bg=ACCENT, fg="white",
                 font=("微軟正黑體", 14, "bold")).pack(side="left", padx=12, pady=10)

        pad = tk.Frame(self, bg=BG)
        pad.pack(fill="both", expand=True, padx=16, pady=12)

        # Basic settings card
        cfg = self._card(pad)
        cfg.pack(fill="x", pady=(0, 10))
        tk.Label(cfg, text="基本設定", bg=CARD, fg=ACCENT, font=FONT_H
                 ).pack(anchor="w", padx=14, pady=(10, 6))
        tk.Frame(cfg, bg=BORDER, height=1).pack(fill="x", padx=14)

        inner = tk.Frame(cfg, bg=CARD)
        inner.pack(fill="x", padx=14, pady=8)

        self._lbl(inner, 0, "來源資料夾")
        self.src_var = tk.StringVar()
        src_row = tk.Frame(inner, bg=CARD)
        src_row.grid(row=0, column=1, sticky="ew", padx=(8, 0), pady=4)
        tk.Entry(src_row, textvariable=self.src_var, font=FONT,
                 relief="flat", bg="#F1EFE8", fg="#2C2C2A",
                 highlightbackground=BORDER, highlightthickness=1, width=34
                 ).pack(side="left", fill="x", expand=True, ipady=4)
        tk.Button(src_row, text="瀏覽…", font=FONT_SMALL, bg=ACCENT_LIGHT,
                  fg=ACCENT, relief="flat", cursor="hand2",
                  command=self._browse_src).pack(side="left", padx=(6, 0), ipady=4, ipadx=6)

        self._lbl(inner, 1, "目標資料夾名稱")
        self.dest_var = tk.StringVar(value="整理後")
        tk.Entry(inner, textvariable=self.dest_var, font=FONT,
                 relief="flat", bg="#F1EFE8", fg="#2C2C2A",
                 highlightbackground=BORDER, highlightthickness=1, width=24
                 ).grid(row=1, column=1, sticky="w", padx=(8, 0), pady=4, ipady=4)

        self._lbl(inner, 2, "副檔名篩選")
        ext_row = tk.Frame(inner, bg=CARD)
        ext_row.grid(row=2, column=1, sticky="ew", padx=(8, 0), pady=4)
        self.ext_var = tk.StringVar(value=DEFAULT_EXTS)
        tk.Entry(ext_row, textvariable=self.ext_var, font=FONT,
                 relief="flat", bg="#F1EFE8", fg="#2C2C2A",
                 highlightbackground=BORDER, highlightthickness=1, width=28
                 ).pack(side="left", ipady=4)
        tk.Label(ext_row, text="  留空 = 不篩選", bg=CARD, fg=GRAY,
                 font=FONT_SMALL).pack(side="left")

        inner.columnconfigure(1, weight=1)

        # Mode selection card
        mode_card = self._card(pad)
        mode_card.pack(fill="x", pady=(0, 10))
        tk.Label(mode_card, text="整理模式", bg=CARD, fg=ACCENT, font=FONT_H
                 ).pack(anchor="w", padx=14, pady=(10, 6))
        tk.Frame(mode_card, bg=BORDER, height=1).pack(fill="x", padx=14)

        self.mode_var = tk.StringVar(value="ext")
        modes = [
            ("ext",  "依副檔名分類",    "每種副檔名建一個子資料夾（PDF/、DOCX/…）"),
            ("name", "依檔名建立資料夾", "每個檔案各自放入以自身檔名命名的資料夾"),
            ("flat", "全部集中放入",     "直接放入目標資料夾，不再建子資料夾"),
        ]
        mode_inner = tk.Frame(mode_card, bg=CARD)
        mode_inner.pack(fill="x", padx=14, pady=(8, 12))
        for val, label, desc in modes:
            row = tk.Frame(mode_inner, bg=CARD)
            row.pack(anchor="w", pady=3)
            tk.Radiobutton(row, variable=self.mode_var, value=val,
                           text=label, font=FONT_B, bg=CARD,
                           activebackground=CARD, fg="#2C2C2A",
                           selectcolor=ACCENT_LIGHT).pack(side="left")
            tk.Label(row, text=f"  {desc}", font=FONT_SMALL,
                     bg=CARD, fg=GRAY).pack(side="left")

        # Dry run option
        opt_row = tk.Frame(pad, bg=BG)
        opt_row.pack(fill="x", pady=(0, 8))
        self.dryrun_var = tk.BooleanVar(value=True)
        f = tk.Frame(opt_row, bg=BG)
        f.pack(side="left")
        tk.Checkbutton(f, variable=self.dryrun_var, bg=BG,
                       activebackground=BG, font=FONT,
                       text="乾跑模式", fg="#2C2C2A",
                       selectcolor=ACCENT_LIGHT).pack(side="left")
        tk.Label(f, text="  （勾選時僅預覽，不移動檔案）", bg=BG,
                 fg=GRAY, font=FONT_SMALL).pack(side="left")

        # Buttons
        btn_row = tk.Frame(pad, bg=BG)
        btn_row.pack(fill="x", pady=(0, 8))
        self.preview_btn = tk.Button(btn_row, text="🔍 預覽", font=FONT_B,
                                     bg=ACCENT_LIGHT, fg=ACCENT, relief="flat",
                                     cursor="hand2", width=12,
                                     command=self._run_preview)
        self.preview_btn.pack(side="left", ipady=7, ipadx=4)
        self.exec_btn = tk.Button(btn_row, text="▶ 執行整理", font=FONT_B,
                                  bg=ACCENT, fg="white", relief="flat",
                                  cursor="hand2", width=14, state="disabled",
                                  command=self._run_execute)
        self.exec_btn.pack(side="left", padx=10, ipady=7, ipadx=4)
        tk.Button(btn_row, text="✕ 清除", font=FONT, bg=BG, fg=GRAY,
                  relief="flat", cursor="hand2", command=self._clear
                  ).pack(side="right", ipady=7)

        # Progress bar
        self.progress = ttk.Progressbar(pad, mode="determinate", maximum=100)
        self.progress.pack(fill="x", pady=(0, 8))

        # Log area
        log_card = self._card(pad)
        log_card.pack(fill="both", expand=True)
        log_hdr = tk.Frame(log_card, bg=CARD)
        log_hdr.pack(fill="x", padx=14, pady=(10, 4))
        tk.Label(log_hdr, text="執行記錄", bg=CARD, fg=ACCENT, font=FONT_B).pack(side="left")
        self.status_lbl = tk.Label(log_hdr, text="", bg=CARD, fg=GRAY, font=FONT_SMALL)
        self.status_lbl.pack(side="right")
        tk.Frame(log_card, bg=BORDER, height=1).pack(fill="x", padx=14)
        log_frame = tk.Frame(log_card, bg=CARD)
        log_frame.pack(fill="both", expand=True, padx=14, pady=8)
        self.log = tk.Text(log_frame, font=("Consolas", 9), bg="#F8F8F6",
                           fg="#2C2C2A", relief="flat", state="disabled",
                           wrap="word", highlightthickness=0, spacing3=2)
        sb = ttk.Scrollbar(log_frame, command=self.log.yview)
        self.log.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.log.pack(fill="both", expand=True)
        for tag, color in [("ok", GREEN), ("err", RED), ("info", ACCENT),
                            ("warn", "#BA7517"), ("dim", GRAY)]:
            self.log.tag_configure(tag, foreground=color)

    def _lbl(self, parent, row, text):
        tk.Label(parent, text=text, bg=CARD, fg="#2C2C2A",
                 font=FONT_B, width=16, anchor="w"
                 ).grid(row=row, column=0, sticky="w", pady=4)

    def _browse_src(self):
        d = filedialog.askdirectory(title="選擇來源資料夾")
        if d:
            self.src_var.set(d)

    def _log(self, msg, tag=""):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n", tag)
        self.log.see("end")
        self.log.configure(state="disabled")

    def _clear_log(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def _clear(self):
        self._plan = []
        self.exec_btn.configure(state="disabled")
        self.progress["value"] = 0
        self.status_lbl.configure(text="")
        self._clear_log()

    def _parse_extensions(self):
        raw = self.ext_var.get().strip()
        if not raw:
            return []
        exts = [e.strip().lower() for e in raw.replace("，", ",").split(",") if e.strip()]
        return [e if e.startswith(".") else "." + e for e in exts]

    def _validate(self):
        src = self.src_var.get().strip()
        if not src:
            messagebox.showwarning("缺少設定", "請選擇來源資料夾。")
            return False
        if not os.path.isdir(src):
            messagebox.showwarning("路徑錯誤", "來源資料夾不存在。")
            return False
        if not self.dest_var.get().strip():
            messagebox.showwarning("缺少設定", "請輸入目標資料夾名稱。")
            return False
        return True

    def _build_plan(self, src, dest_base, exts, mode):
        plan = []
        for fname in os.listdir(src):
            fpath = os.path.join(src, fname)
            if not os.path.isfile(fpath):
                continue
            stem, ext = os.path.splitext(fname)
            if exts and ext.lower() not in exts:
                continue

            if mode == "ext":
                sub = ext.lstrip(".").upper() if ext else "其他"
                target_dir = os.path.join(dest_base, sub)
            elif mode == "name":
                target_dir = os.path.join(dest_base, stem)
            else:
                target_dir = dest_base

            final_name = fname
            counter = 1
            while os.path.exists(os.path.join(target_dir, final_name)):
                final_name = f"{stem}_({counter}){ext}"
                counter += 1

            plan.append({
                "src": fpath, "fname": fname,
                "target_dir": target_dir, "final": final_name,
                "renamed": final_name != fname
            })
        return plan

    def _run_preview(self):
        if not self._validate():
            return
        self._clear_log()
        self._plan = []
        self.exec_btn.configure(state="disabled")
        self.progress["value"] = 0

        src = self.src_var.get().strip()
        dest_base = os.path.join(src, self.dest_var.get().strip())
        exts = self._parse_extensions()
        mode = self.mode_var.get()

        mode_labels = {"ext": "依副檔名分類", "name": "依檔名建立資料夾", "flat": "全部集中放入"}
        self._log(f"來源：{src}", "info")
        self._log(f"目標：{dest_base}", "info")
        self._log(f"模式：{mode_labels[mode]}", "info")
        self._log(f"篩選：{', '.join(exts) if exts else '（全部副檔名）'}", "info")
        self._log("─" * 60, "dim")

        plan = self._build_plan(src, dest_base, exts, mode)
        if not plan:
            self._log("⚠  未找到符合條件的檔案。", "warn")
            self.status_lbl.configure(text="0 個檔案")
            return

        for item in plan:
            rename_note = f"  [重新命名 → {item['final']}]" if item["renamed"] else ""
            self._log(f"  {item['fname']}{rename_note}", "")
            self._log(f"    → {item['target_dir']}", "dim")

        self._log("─" * 60, "dim")
        self._log(f"共 {len(plan)} 個檔案可整理。", "info")
        self.status_lbl.configure(text=f"預覽：{len(plan)} 個檔案")
        self._plan = plan
        self.exec_btn.configure(state="normal")

    def _run_execute(self):
        if not self._plan:
            return
        if self.dryrun_var.get():
            messagebox.showinfo("乾跑模式",
                                "目前為乾跑模式，不會實際移動檔案。\n請取消勾選「乾跑模式」後再執行。")
            return
        if not messagebox.askyesno("確認執行",
                                   f"確定要移動 {len(self._plan)} 個檔案嗎？\n此操作無法復原。"):
            return

        self.preview_btn.configure(state="disabled")
        self.exec_btn.configure(state="disabled")
        self._clear_log()
        self._log("開始執行整理…", "info")
        self._log("─" * 60, "dim")

        def worker():
            ok = 0; fail = 0
            total = len(self._plan)
            for i, item in enumerate(self._plan):
                try:
                    os.makedirs(item["target_dir"], exist_ok=True)
                    dst = os.path.join(item["target_dir"], item["final"])
                    shutil.move(item["src"], dst)
                    self._log(f"  ✓  {item['fname']}", "ok")
                    ok += 1
                except Exception as e:
                    self._log(f"  ✗  {item['fname']}：{e}", "err")
                    fail += 1
                self.progress["value"] = int((i + 1) / total * 100)
                self.update_idletasks()

            self._log("─" * 60, "dim")
            self._log(f"完成！成功 {ok} 個，失敗 {fail} 個。",
                      "ok" if fail == 0 else "warn")
            self.status_lbl.configure(text=f"完成：{ok} 成功 / {fail} 失敗")
            self.preview_btn.configure(state="normal")
            self._plan = []

        threading.Thread(target=worker, daemon=True).start()


if __name__ == "__main__":
    app = App()
    app.mainloop()
