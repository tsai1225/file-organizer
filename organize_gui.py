import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import shutil
import threading

# send2trash: 移到資源回收桶（若未安裝則停用該功能）
try:
    from send2trash import send2trash
    HAS_SEND2TRASH = True
except ImportError:
    HAS_SEND2TRASH = False

APP_TITLE = "檔案整理工具"
BG = "#F8F8F6"
CARD = "#FFFFFF"
ACCENT = "#5348B7"
ACCENT_LIGHT = "#EEEDFE"
GREEN = "#0F6E56"
RED = "#993C1D"
ORANGE = "#B75348"
GRAY = "#5F5E5A"
BORDER = "#D3D1C7"
FONT = ("微軟正黑體", 10)
FONT_B = ("微軟正黑體", 10, "bold")
FONT_SMALL = ("微軟正黑體", 9)
FONT_MONO = ("Consolas", 9)

DEFAULT_EXTS = ".pdf, .docx, .xlsx, .doc, .xls"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)

        # 修正重點 1：允許調整視窗大小，並降低預設高度，避免高 DPI 螢幕按鈕被擠出畫面。
        self.geometry("840x720")
        self.minsize(700, 520)
        self.resizable(True, True)
        self.configure(bg=BG)

        self._plan = []
        self._running = False
        self._build_ui()

    # ── 通用卡片容器 ────────────────────────────────────────────
    def _card(self, parent, **kw):
        f = tk.Frame(parent, bg=CARD, bd=1, relief="flat",
                     highlightbackground=BORDER, highlightthickness=1, **kw)
        return f

    def _label(self, parent, text, bold=False, color=None, size=10, **kw):
        font = ("微軟正黑體", size, "bold") if bold else ("微軟正黑體", size)
        return tk.Label(parent, text=text, bg=parent["bg"],
                        fg=color or "#1A1A18", font=font, **kw)

    def _btn(self, parent, text, cmd, color=ACCENT, fg="white", width=12, **kw):
        b = tk.Button(parent, text=text, command=cmd,
                      bg=color, fg=fg, activebackground=color,
                      font=FONT_B, relief="flat", bd=0,
                      padx=10, pady=6, cursor="hand2", width=width, **kw)
        return b

    # ── UI 建構 ─────────────────────────────────────────────────
    def _build_ui(self):
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # 標題
        hdr = tk.Frame(self, bg=ACCENT, height=52)
        hdr.grid(row=0, column=0, sticky="ew")
        hdr.grid_propagate(False)
        hdr.grid_columnconfigure(0, weight=1)
        tk.Label(hdr, text="📂  " + APP_TITLE, bg=ACCENT, fg="white",
                 font=("微軟正黑體", 14, "bold"), anchor="w").grid(
                     row=0, column=0, sticky="w", padx=18, pady=10)
        tk.Label(hdr, text="視窗可自由調整大小", bg=ACCENT, fg="white",
                 font=FONT_SMALL, anchor="e").grid(row=0, column=1, sticky="e", padx=18)

        # 中間內容改成可捲動，避免畫面縮小時下方按鈕消失。
        body = tk.Frame(self, bg=BG)
        body.grid(row=1, column=0, sticky="nsew")
        body.grid_rowconfigure(0, weight=1)
        body.grid_columnconfigure(0, weight=1)

        self.canvas = tk.Canvas(body, bg=BG, highlightthickness=0)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        ysb = ttk.Scrollbar(body, orient="vertical", command=self.canvas.yview)
        ysb.grid(row=0, column=1, sticky="ns")
        self.canvas.configure(yscrollcommand=ysb.set)

        outer = tk.Frame(self.canvas, bg=BG)
        self._canvas_window = self.canvas.create_window((0, 0), window=outer, anchor="nw")

        def _update_scroll_region(_event=None):
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

        def _fit_content_width(event):
            self.canvas.itemconfigure(self._canvas_window, width=event.width)

        outer.bind("<Configure>", _update_scroll_region)
        self.canvas.bind("<Configure>", _fit_content_width)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        content = tk.Frame(outer, bg=BG)
        content.pack(fill="both", expand=True, padx=16, pady=10)

        # ── 來源資料夾 ──
        c1 = self._card(content)
        c1.pack(fill="x", pady=(0, 8))
        self._label(c1, "來源資料夾", bold=True).pack(anchor="w", padx=12, pady=(10, 4))
        row1 = tk.Frame(c1, bg=CARD)
        row1.pack(fill="x", padx=12, pady=(0, 10))
        self.src_var = tk.StringVar()
        tk.Entry(row1, textvariable=self.src_var, font=FONT,
                 relief="solid", bd=1, bg="#FAFAF8").pack(side="left", fill="x", expand=True)
        self._btn(row1, "瀏覽…", self._browse_src, width=8).pack(side="left", padx=(6, 0))

        # ── 作業模式 ──
        c2 = self._card(content)
        c2.pack(fill="x", pady=(0, 8))
        self._label(c2, "作業模式", bold=True).pack(anchor="w", padx=12, pady=(10, 6))

        # --- 整理模式 ---
        self._label(c2, "▸ 整理模式", bold=True, color=ACCENT).pack(anchor="w", padx=16, pady=(0, 2))
        self.mode_var = tk.StringVar(value="ext")
        modes = [
            ("依副檔名建立資料夾  （例：.pdf → pdf資料夾）", "ext"),
            ("依檔名建立資料夾    （每個檔案各自一個資料夾）", "name"),
            ("平鋪收集到單一資料夾（全部移到同一個目的地）", "flat"),
        ]
        for txt, val in modes:
            rb = tk.Radiobutton(c2, text=txt, variable=self.mode_var, value=val,
                                bg=CARD, font=FONT, activebackground=CARD,
                                selectcolor=ACCENT_LIGHT, fg="#1A1A18",
                                command=self._on_mode_change)
            rb.pack(anchor="w", padx=28)

        # 平鋪模式的目的地資料夾
        self.flat_frame = tk.Frame(c2, bg=CARD)
        self.flat_frame.pack(fill="x", padx=28, pady=(2, 0))
        self._label(self.flat_frame, "目的地資料夾：", color=GRAY).pack(side="left")
        self.dst_var = tk.StringVar()
        tk.Entry(self.flat_frame, textvariable=self.dst_var, font=FONT,
                 relief="solid", bd=1, bg="#FAFAF8", width=32).pack(side="left", padx=4, fill="x", expand=True)
        self._btn(self.flat_frame, "瀏覽…", self._browse_dst, width=6).pack(side="left")
        self.flat_frame.pack_forget()

        # --- 刪除模式 ---
        sep = tk.Frame(c2, bg=BORDER, height=1)
        sep.pack(fill="x", padx=12, pady=8)
        self._label(c2, "▸ 刪除模式（獨立操作，不與整理模式並用）",
                    bold=True, color=ORANGE).pack(anchor="w", padx=16, pady=(0, 2))

        self.del_mode_var = tk.StringVar(value="none")
        del_modes = [
            ("不刪除（正常整理）", "none"),
            ("移到資源回收桶（可還原）", "recycle"),
            ("直接刪除（無法復原，請謹慎）", "perm"),
        ]
        for txt, val in del_modes:
            rb = tk.Radiobutton(c2, text=txt, variable=self.del_mode_var, value=val,
                                bg=CARD, font=FONT, activebackground=CARD,
                                selectcolor=ACCENT_LIGHT, fg="#1A1A18",
                                command=self._on_delete_mode_change)
            rb.pack(anchor="w", padx=28)

        self.mode_hint_lbl = self._label(
            c2,
            "提示：刪除模式若不是「不刪除」，程式會以刪除模式為主，不會同時執行上方整理模式。",
            color=GRAY, size=9)
        self.mode_hint_lbl.pack(anchor="w", padx=28, pady=(2, 0))

        if not HAS_SEND2TRASH:
            warn = self._label(
                c2,
                "⚠ 未安裝 send2trash，資源回收桶功能停用。請執行：pip install send2trash",
                color="#993C1D", size=9)
            warn.pack(anchor="w", padx=28, pady=(2, 0))

        # --- 乾跑模式 ---
        sep2 = tk.Frame(c2, bg=BORDER, height=1)
        sep2.pack(fill="x", padx=12, pady=8)
        self.dryrun_var = tk.BooleanVar(value=True)
        cb = tk.Checkbutton(
            c2,
            text="乾跑模式（Dry Run）：勾選時按「正式執行」也只會模擬，不會真的搬移／刪除檔案",
            variable=self.dryrun_var, bg=CARD, font=FONT_B,
            activebackground=CARD, selectcolor=ACCENT_LIGHT,
            fg=ACCENT)
        cb.pack(anchor="w", padx=16, pady=(0, 4))
        self._label(c2, "下方固定按鈕列也有「乾跑測試」按鈕，可直接測試而不動到檔案。",
                    color=GRAY, size=9).pack(anchor="w", padx=18, pady=(0, 8))

        # ── 副檔名篩選 ──
        c3 = self._card(content)
        c3.pack(fill="x", pady=(0, 8))
        self._label(c3, "副檔名篩選（空白 = 全部）", bold=True).pack(anchor="w", padx=12, pady=(10, 4))
        self.ext_var = tk.StringVar(value=DEFAULT_EXTS)
        tk.Entry(c3, textvariable=self.ext_var, font=FONT,
                 relief="solid", bd=1, bg="#FAFAF8").pack(fill="x", padx=12, pady=(0, 10))
        self._label(c3, "多個副檔名以逗號分隔，例：.pdf, .docx, .xlsx",
                    color=GRAY, size=9).pack(anchor="w", padx=12, pady=(0, 8))

        # ── 預覽 / 執行紀錄 ──
        c4 = self._card(content)
        c4.pack(fill="x", pady=(0, 8))
        hrow = tk.Frame(c4, bg=CARD)
        hrow.pack(fill="x", padx=12, pady=(10, 4))
        self._label(hrow, "預覽清單 / 執行紀錄", bold=True).pack(side="left")
        self.count_lbl = self._label(hrow, "", color=GRAY, size=9)
        self.count_lbl.pack(side="left", padx=8)

        lf = tk.Frame(c4, bg=CARD)
        lf.pack(fill="both", expand=True, padx=12, pady=(0, 6))
        sb_y = ttk.Scrollbar(lf, orient="vertical")
        sb_y.pack(side="right", fill="y")
        sb_x = ttk.Scrollbar(lf, orient="horizontal")
        sb_x.pack(side="bottom", fill="x")
        self.log = tk.Text(lf, font=FONT_MONO,
                           yscrollcommand=sb_y.set, xscrollcommand=sb_x.set,
                           bg="#FAFAF8", relief="solid", bd=1, wrap="none",
                           state="disabled", height=15)
        self.log.pack(fill="both", expand=True)
        sb_y.config(command=self.log.yview)
        sb_x.config(command=self.log.xview)
        self.log.tag_configure("ok", foreground=GREEN)
        self.log.tag_configure("err", foreground=RED)
        self.log.tag_configure("dim", foreground=GRAY)

        self.progress = ttk.Progressbar(c4, mode="determinate")
        self.progress.pack(fill="x", padx=12, pady=(0, 10))

        # 修正重點 2：按鈕列固定在視窗下方，不放進可捲動內容，因此不會再消失。
        footer = tk.Frame(self, bg=BG)
        footer.grid(row=2, column=0, sticky="ew", padx=16, pady=(6, 4))
        footer.grid_columnconfigure(4, weight=1)
        self.preview_btn = self._btn(footer, "🔍 預覽", self._preview, width=12)
        self.preview_btn.grid(row=0, column=0, sticky="w")
        self.dry_btn = self._btn(footer, "🧪 乾跑測試", self._execute_dry_run, color=ACCENT, width=14)
        self.dry_btn.grid(row=0, column=1, sticky="w", padx=(8, 0))
        self.execute_btn = self._btn(footer, "✅ 正式執行", self._execute_real, color=GREEN, width=14)
        self.execute_btn.grid(row=0, column=2, sticky="w", padx=(8, 0))
        self.clear_btn = self._btn(footer, "🗑 清除清單", self._clear, color=GRAY, width=12)
        self.clear_btn.grid(row=0, column=5, sticky="e")

        self.status_var = tk.StringVar(value="請先選擇來源資料夾，再按「預覽」；也可以直接按「乾跑測試」產生模擬結果。")
        tk.Label(self, textvariable=self.status_var, bg=BG,
                 fg=GRAY, font=FONT_SMALL, anchor="w").grid(
                     row=3, column=0, sticky="ew", padx=16, pady=(0, 8))

    def _on_mousewheel(self, event):
        # Windows / 一般滑鼠滾輪
        try:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except Exception:
            pass

    # ── 紀錄區操作 ─────────────────────────────────────────────
    def _log_clear(self):
        self.log.config(state="normal")
        self.log.delete("1.0", "end")
        self.log.config(state="disabled")

    def _log_write(self, text, tag=None):
        self.log.config(state="normal")
        if tag:
            self.log.insert("end", text + "\n", tag)
        else:
            self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.config(state="disabled")

    def _set_buttons_enabled(self, enabled):
        state = "normal" if enabled else "disabled"
        for btn in (self.preview_btn, self.dry_btn, self.execute_btn, self.clear_btn):
            btn.config(state=state)

    # ── 事件 ────────────────────────────────────────────────────
    def _browse_src(self):
        d = filedialog.askdirectory(title="選擇來源資料夾")
        if d:
            self.src_var.set(d)

    def _browse_dst(self):
        d = filedialog.askdirectory(title="選擇目的地資料夾")
        if d:
            self.dst_var.set(d)

    def _invalidate_plan(self):
        # 模式、來源或篩選條件變更後，舊預覽清單可能已不準確。
        self._plan = []
        if hasattr(self, "count_lbl"):
            self.count_lbl.config(text="")
        if hasattr(self, "progress"):
            self.progress.config(value=0)
        if hasattr(self, "status_var"):
            self.status_var.set("設定已變更，請重新按「預覽」或「乾跑測試」。")

    def _on_mode_change(self):
        if self.mode_var.get() == "flat":
            self.flat_frame.pack(fill="x", padx=28, pady=(2, 8))
        else:
            self.flat_frame.pack_forget()
        self._invalidate_plan()

    def _on_delete_mode_change(self):
        self._invalidate_plan()
        dm = self.del_mode_var.get()
        if dm == "none":
            self.status_var.set("目前為正常整理模式。請按「預覽」確認清單。")
        elif dm == "recycle":
            self.status_var.set("目前為移到資源回收桶模式。若乾跑模式仍勾選，只會模擬，不會真的移動檔案。")
        elif dm == "perm":
            self.status_var.set("目前為直接刪除模式。若乾跑模式仍勾選，只會模擬，不會真的刪除檔案。")

    def _parse_exts(self):
        raw = self.ext_var.get().strip()
        if not raw:
            return None  # None = 不篩選
        exts = set()
        for e in raw.split(","):
            e = e.strip().lower()
            if e and not e.startswith("."):
                e = "." + e
            if e:
                exts.add(e)
        return exts

    def _scan_files(self, src, exts):
        files = []
        for f in os.listdir(src):
            fp = os.path.join(src, f)
            if not os.path.isfile(fp):
                continue
            if exts is None or os.path.splitext(f)[1].lower() in exts:
                files.append((f, fp))
        files.sort(key=lambda x: x[0].lower())
        return files

    def _build_plan(self, files, src):
        mode = self.mode_var.get()
        del_mode = self.del_mode_var.get()
        plan = []

        if del_mode != "none":
            action_name = "資源回收桶" if del_mode == "recycle" else "永久刪除"
            for _fname, fpath in files:
                plan.append({
                    "src": fpath,
                    "dst": None,
                    "action": del_mode,
                    "label": f"[{action_name}]  {fpath}",
                })
        elif mode == "ext":
            for fname, fpath in files:
                ext = os.path.splitext(fname)[1].lower().lstrip(".") or "無副檔名"
                dst_dir = os.path.join(src, ext)
                dst_path = self._unique_path(dst_dir, fname)
                plan.append({
                    "src": fpath,
                    "dst": dst_path,
                    "action": "move",
                    "label": f"{fpath}\n      →  {dst_path}",
                })
        elif mode == "name":
            for fname, fpath in files:
                stem = os.path.splitext(fname)[0]
                dst_dir = os.path.join(src, stem)
                dst_path = self._unique_path(dst_dir, fname)
                plan.append({
                    "src": fpath,
                    "dst": dst_path,
                    "action": "move",
                    "label": f"{fpath}\n      →  {dst_path}",
                })
        elif mode == "flat":
            dst_root = self.dst_var.get().strip()
            if not dst_root:
                messagebox.showerror("錯誤", "平鋪模式需要指定目的地資料夾。")
                return None
            if not os.path.isdir(dst_root):
                messagebox.showerror("錯誤", "請選擇有效的目的地資料夾。")
                return None
            for fname, fpath in files:
                dst_path = self._unique_path(dst_root, fname)
                plan.append({
                    "src": fpath,
                    "dst": dst_path,
                    "action": "move",
                    "label": f"{fpath}\n      →  {dst_path}",
                })
        return plan

    def _unique_path(self, folder, fname):
        base, ext = os.path.splitext(fname)
        candidate = os.path.join(folder, fname)
        n = 1
        while os.path.exists(candidate):
            candidate = os.path.join(folder, f"{base}({n}){ext}")
            n += 1
        return candidate

    def _ensure_plan(self, show_perm_warning=True):
        src = self.src_var.get().strip()
        if not src or not os.path.isdir(src):
            messagebox.showerror("錯誤", "請先選擇有效的來源資料夾。")
            return False

        dm = self.del_mode_var.get()
        if dm == "perm" and show_perm_warning:
            if not messagebox.askyesno(
                "警告",
                "已選擇「直接刪除」模式！\n\n若正式執行，將永久刪除檔案，無法復原。\n確定繼續建立預覽清單嗎？",
                icon="warning"):
                return False
        elif dm == "recycle" and not HAS_SEND2TRASH:
            messagebox.showerror(
                "錯誤",
                "尚未安裝 send2trash，無法使用資源回收桶功能。\n請執行：pip install send2trash")
            return False

        exts = self._parse_exts()
        try:
            files = self._scan_files(src, exts)
        except Exception as e:
            messagebox.showerror("錯誤", f"掃描來源資料夾失敗：\n{e}")
            return False

        if not files:
            messagebox.showinfo("提示", "來源資料夾中沒有符合條件的檔案。")
            return False

        plan = self._build_plan(files, src)
        if plan is None:
            return False

        self._plan = plan
        return True

    def _preview(self):
        if not self._ensure_plan(show_perm_warning=True):
            return

        self._log_clear()
        self._log_write(f"=== 預覽結果（共 {len(self._plan)} 個檔案）===\n", "dim")
        for item in self._plan:
            self._log_write(item["label"])
        self.count_lbl.config(text=f"共 {len(self._plan)} 個檔案")
        self.progress.config(value=0, maximum=max(len(self._plan), 1))
        self.status_var.set(
            "預覽完成。可按「乾跑測試」確認模擬結果；確認無誤後再按「正式執行」。")

    def _execute_dry_run(self):
        if self._running:
            return
        # 每次都依目前畫面設定重新建立清單，避免使用者預覽後改了模式卻沿用舊清單。
        if not self._ensure_plan(show_perm_warning=False):
            return
        self._simulate_plan()

    def _simulate_plan(self):
        self._log_clear()
        self._log_write(
            f"=== 乾跑模式模擬結果（共 {len(self._plan)} 個檔案，未執行任何操作）===\n",
            "dim")
        for item in self._plan:
            self._log_write("（模擬，未實際執行）" + item["label"], "dim")
        self.count_lbl.config(text=f"共 {len(self._plan)} 個檔案")
        self.progress.config(value=0, maximum=max(len(self._plan), 1))
        self.status_var.set("乾跑測試完成：只是模擬結果，沒有搬移或刪除任何檔案。")

    def _execute_real(self):
        if self._running:
            return
        # 每次正式執行都依目前畫面設定重新建立清單，避免模式切換後沿用舊預覽。
        if not self._ensure_plan(show_perm_warning=True):
            return

        dm = self.del_mode_var.get()

        # 保留原本乾跑勾選邏輯：若使用者仍勾著乾跑，按正式執行也只模擬。
        if self.dryrun_var.get():
            self._simulate_plan()
            messagebox.showinfo(
                "乾跑模式",
                "目前仍勾選「乾跑模式」，所以沒有實際搬移或刪除檔案。\n\n若要正式執行，請取消勾選「乾跑模式」後再按「正式執行」。")
            return

        if dm == "perm":
            if not messagebox.askyesno(
                "最終確認",
                f"即將永久刪除 {len(self._plan)} 個檔案，此動作無法復原！\n\n確定執行嗎？",
                icon="warning"):
                return
        elif dm == "recycle":
            if not messagebox.askyesno(
                "確認",
                f"即將把 {len(self._plan)} 個檔案移到資源回收桶，確定嗎？"):
                return
        else:
            if not messagebox.askyesno(
                "確認",
                f"即將整理 {len(self._plan)} 個檔案，確定執行嗎？"):
                return

        self._running = True
        self._set_buttons_enabled(False)
        threading.Thread(target=self._run_plan, daemon=True).start()

    def _run_plan(self):
        ok = 0
        err = 0
        total = len(self._plan)
        plan = list(self._plan)

        self.after(0, self._log_clear)
        self.after(0, self._log_write, f"=== 開始執行（共 {total} 個檔案）===\n", "dim")
        self.after(0, lambda: self.progress.config(value=0, maximum=total))

        for i, item in enumerate(plan, 1):
            fname = os.path.basename(item["src"])
            self.after(0, self.status_var.set, f"處理中 {i}/{total}：{fname}")
            try:
                action = item["action"]
                if action == "move":
                    dst_dir = os.path.dirname(item["dst"])
                    os.makedirs(dst_dir, exist_ok=True)
                    shutil.move(item["src"], item["dst"])
                    if os.path.exists(item["src"]) or not os.path.exists(item["dst"]):
                        raise RuntimeError("移動後檢查失敗：來源仍存在或目的地不存在。")
                    self.after(0, self._log_write,
                               f"✓ 已整理：{item['src']}\n    →  {item['dst']}", "ok")
                elif action == "recycle":
                    send2trash(item["src"])
                    # send2trash 正常成功後，原路徑應不存在；若還存在，就不要回報成功。
                    if os.path.exists(item["src"]):
                        raise RuntimeError("已呼叫資源回收桶功能，但原檔案仍存在，可能是系統權限、同步資料夾或磁碟回收桶設定造成。")
                    self.after(0, self._log_write,
                               f"✓ 已移至資源回收桶：{item['src']}", "ok")
                elif action == "perm":
                    os.remove(item["src"])
                    if os.path.exists(item["src"]):
                        raise RuntimeError("刪除後檢查失敗：原檔案仍存在。")
                    self.after(0, self._log_write,
                               f"✓ 已永久刪除：{item['src']}", "ok")
                ok += 1
            except Exception as e:
                err += 1
                self.after(0, self._log_write,
                           f"✗ 失敗：{item['src']}\n    原因：{e}", "err")
            self.after(0, lambda v=i: self.progress.config(value=v))

        self.after(0, self._finish_run, ok, err)

    def _finish_run(self, ok, err):
        self._plan = []
        self._running = False
        self._set_buttons_enabled(True)

        msg = f"完成！成功 {ok} 個"
        if err:
            msg += f"，失敗 {err} 個"
        self._log_write(f"\n=== {msg} ===", "dim")
        self.status_var.set(msg)
        self.count_lbl.config(text="")

        if err:
            messagebox.showwarning("部分失敗", f"{msg}\n詳細內容請見上方紀錄。")
        else:
            messagebox.showinfo("完成", msg)

    def _clear(self):
        if self._running:
            return
        self._plan = []
        self._log_clear()
        self.count_lbl.config(text="")
        self.progress.config(value=0)
        self.status_var.set("清單已清除。")


if __name__ == "__main__":
    App().mainloop()
