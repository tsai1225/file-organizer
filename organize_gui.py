import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os, shutil, threading

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
FONT_H = ("微軟正黑體", 13, "bold")
FONT_SMALL = ("微軟正黑體", 9)
FONT_MONO = ("Consolas", 9)

DEFAULT_EXTS = ".pdf, .docx, .xlsx, .doc, .xls"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("760x740")
        self.minsize(700, 680)
        self.resizable(True, True)
        self.configure(bg=BG)
        self._plan = []
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
                      padx=10, pady=5, cursor="hand2", width=width, **kw)
        return b

    # ── UI 建構 ─────────────────────────────────────────────────
    def _build_ui(self):
        # 標題
        hdr = tk.Frame(self, bg=ACCENT, height=52)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="📂  " + APP_TITLE, bg=ACCENT, fg="white",
                 font=("微軟正黑體", 14, "bold")).pack(side="left", padx=18, pady=10)

        outer = tk.Frame(self, bg=BG)
        # 注意：outer 先建立但暫不 pack，等狀態列與按鈕列 pack 完
        # （保留好底部空間）之後才 pack，避免 expand 容器把空間佔滿。

        # ── 來源資料夾 ──
        c1 = self._card(outer)
        c1.pack(fill="x", pady=(0, 8))
        self._label(c1, "來源資料夾", bold=True).pack(anchor="w", padx=12, pady=(10, 4))
        row1 = tk.Frame(c1, bg=CARD)
        row1.pack(fill="x", padx=12, pady=(0, 10))
        self.src_var = tk.StringVar()
        tk.Entry(row1, textvariable=self.src_var, font=FONT,
                 relief="solid", bd=1, bg="#FAFAF8").pack(side="left", fill="x", expand=True)
        self._btn(row1, "瀏覽…", self._browse_src, width=8).pack(side="left", padx=(6, 0))

        # ── 作業模式 ──
        c2 = self._card(outer)
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
                 relief="solid", bd=1, bg="#FAFAF8", width=28).pack(side="left", padx=4)
        self._btn(self.flat_frame, "瀏覽…", self._browse_dst, width=6).pack(side="left")
        self.flat_frame.pack_forget()

        # --- 刪除模式 ---
        self._label(c2, "▸ 刪除模式（獨立操作，不與整理模式並用）",
                    bold=True, color=ORANGE).pack(anchor="w", padx=16, pady=(8, 2))

        self.del_mode_var = tk.StringVar(value="none")
        del_modes = [
            ("不刪除（正常整理）", "none"),
            ("移到資源回收桶（可還原）", "recycle"),
            ("直接刪除（無法復原，請謹慎）", "perm"),
        ]
        for txt, val in del_modes:
            rb = tk.Radiobutton(c2, text=txt, variable=self.del_mode_var, value=val,
                                bg=CARD, font=FONT, activebackground=CARD,
                                selectcolor=ACCENT_LIGHT, fg="#1A1A18")
            rb.pack(anchor="w", padx=28)

        # send2trash 警告（若未安裝）
        if not HAS_SEND2TRASH:
            warn = self._label(c2,
                "⚠ 未安裝 send2trash，資源回收桶功能停用。請執行：pip install send2trash",
                color="#993C1D", size=9)
            warn.pack(anchor="w", padx=28, pady=(2, 0))

        # --- 乾跑模式 ---
        self.dryrun_var = tk.BooleanVar(value=True)
        cb = tk.Checkbutton(c2, text="乾跑模式（Dry Run）：勾選時按「執行」只會模擬並顯示結果，不會真的搬移／刪除檔案",
                            variable=self.dryrun_var, bg=CARD, font=FONT_B,
                            activebackground=CARD, selectcolor=ACCENT_LIGHT,
                            fg=ACCENT)
        cb.pack(anchor="w", padx=16, pady=(8, 8))

        # ── 副檔名篩選 ──
        c3 = self._card(outer)
        c3.pack(fill="x", pady=(0, 8))
        self._label(c3, "副檔名篩選（空白 = 全部）", bold=True).pack(anchor="w", padx=12, pady=(10, 4))
        self.ext_var = tk.StringVar(value=DEFAULT_EXTS)
        tk.Entry(c3, textvariable=self.ext_var, font=FONT,
                 relief="solid", bd=1, bg="#FAFAF8").pack(fill="x", padx=12, pady=(0, 10))
        self._label(c3, "多個副檔名以逗號分隔，例：.pdf, .docx, .xlsx",
                    color=GRAY, size=9).pack(anchor="w", padx=12, pady=(0, 8))

        # ── 預覽 / 執行紀錄 ──
        c4 = self._card(outer)
        c4.pack(fill="both", expand=True, pady=(0, 8))
        hrow = tk.Frame(c4, bg=CARD)
        hrow.pack(fill="x", padx=12, pady=(10, 4))
        self._label(hrow, "預覽清單 / 執行紀錄", bold=True).pack(side="left")
        self.count_lbl = self._label(hrow, "", color=GRAY, size=9)
        self.count_lbl.pack(side="left", padx=8)

        lf = tk.Frame(c4, bg=CARD)
        lf.pack(fill="both", expand=True, padx=12, pady=(0, 6))
        sb = ttk.Scrollbar(lf)
        sb.pack(side="right", fill="y")
        self.log = tk.Text(lf, font=FONT_MONO, yscrollcommand=sb.set,
                           bg="#FAFAF8", relief="solid", bd=1, wrap="none",
                           height=8, state="disabled")
        self.log.pack(fill="both", expand=True)
        sb.config(command=self.log.yview)
        self.log.tag_configure("ok", foreground=GREEN)
        self.log.tag_configure("err", foreground=RED)
        self.log.tag_configure("dim", foreground=GRAY)

        # 進度條
        self.progress = ttk.Progressbar(c4, mode="determinate")
        self.progress.pack(fill="x", padx=12, pady=(0, 10))

        # ── 狀態列（先 pack，固定在最底部）──
        self.status_var = tk.StringVar(value="請先選擇來源資料夾，再按「預覽」")
        tk.Label(self, textvariable=self.status_var, bg=BG,
                 fg=GRAY, font=FONT_SMALL, anchor="w").pack(
            side="bottom", fill="x", padx=16, pady=(0, 8))

        # ── 按鈕列（pack 在狀態列之上）──
        brow = tk.Frame(self, bg=BG)
        brow.pack(side="bottom", fill="x", padx=16, pady=(0, 4))
        self._btn(brow, "🔍  預覽", self._preview, width=14).pack(side="left")
        self._btn(brow, "✅  執行", self._execute, color=GREEN, width=14).pack(side="left", padx=8)
        self._btn(brow, "🗑  清除清單", self._clear, color=GRAY, width=12).pack(side="right")

        # outer 最後才 pack，此時狀態列與按鈕列已經保留好底部空間，
        # outer（含可伸縮的預覽區）只會佔用剩餘空間。
        outer.pack(fill="both", expand=True, padx=16, pady=10)

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

    # ── 事件 ────────────────────────────────────────────────────
    def _browse_src(self):
        d = filedialog.askdirectory(title="選擇來源資料夾")
        if d:
            self.src_var.set(d)

    def _browse_dst(self):
        d = filedialog.askdirectory(title="選擇目的地資料夾")
        if d:
            self.dst_var.set(d)

    def _on_mode_change(self):
        if self.mode_var.get() == "flat":
            self.flat_frame.pack(fill="x", padx=28, pady=(2, 8))
        else:
            self.flat_frame.pack_forget()

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
        return files

    def _build_plan(self, files, src):
        mode = self.mode_var.get()
        del_mode = self.del_mode_var.get()
        plan = []

        if del_mode != "none":
            action_name = "資源回收桶" if del_mode == "recycle" else "永久刪除"
            for fname, fpath in files:
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

    def _preview(self):
        src = self.src_var.get().strip()
        if not src or not os.path.isdir(src):
            messagebox.showerror("錯誤", "請先選擇有效的來源資料夾。")
            return

        # 刪除模式警告
        dm = self.del_mode_var.get()
        if dm == "perm":
            if not messagebox.askyesno("警告",
                    "已選擇「直接刪除」模式！\n\n之後若取消勾選「乾跑模式」並執行，將永久刪除檔案，無法復原。\n確定繼續預覽嗎？"):
                return
        elif dm == "recycle" and not HAS_SEND2TRASH:
            messagebox.showerror("錯誤",
                "尚未安裝 send2trash，無法使用資源回收桶功能。\n請執行：pip install send2trash")
            return

        exts = self._parse_exts()
        files = self._scan_files(src, exts)
        if not files:
            messagebox.showinfo("提示", "來源資料夾中沒有符合條件的檔案。")
            return

        plan = self._build_plan(files, src)
        if plan is None:
            return

        self._plan = plan
        self._log_clear()
        self._log_write(f"=== 預覽結果（共 {len(plan)} 個檔案）===\n", "dim")
        for item in plan:
            self._log_write(item["label"])
        self.count_lbl.config(text=f"共 {len(plan)} 個檔案")
        self.progress.config(value=0, maximum=max(len(plan), 1))
        self.status_var.set(f"預覽完成，共 {len(plan)} 個檔案。確認無誤後按「執行」"
                            f"（{'乾跑模式：再按執行也只會模擬' if self.dryrun_var.get() else '注意：將實際執行'}）。")

    def _execute(self):
        if not self._plan:
            messagebox.showinfo("提示", "請先按「預覽」產生清單。")
            return

        dm = self.del_mode_var.get()
        dry = self.dryrun_var.get()

        # 乾跑模式：只模擬，不動檔案
        if dry:
            self._log_clear()
            self._log_write(f"=== 乾跑模式模擬結果（共 {len(self._plan)} 個檔案，未執行任何操作）===\n", "dim")
            for item in self._plan:
                self._log_write("（模擬）" + item["label"], "dim")
            self.status_var.set("乾跑模式：以上為模擬結果，未搬移或刪除任何檔案。取消勾選「乾跑模式」後可實際執行。")
            return

        # 非乾跑：執行前確認
        if dm == "perm":
            if not messagebox.askyesno("最終確認",
                    f"即將永久刪除 {len(self._plan)} 個檔案，此動作無法復原！\n\n確定執行嗎？",
                    icon="warning"):
                return
        elif dm == "recycle":
            if not messagebox.askyesno("確認",
                    f"即將把 {len(self._plan)} 個檔案移到資源回收桶，確定嗎？"):
                return
        else:
            if not messagebox.askyesno("確認",
                    f"即將整理 {len(self._plan)} 個檔案，確定執行嗎？"):
                return

        threading.Thread(target=self._run_plan, daemon=True).start()

    def _run_plan(self):
        ok = err = 0
        total = len(self._plan)
        plan = self._plan

        self._log_clear()
        self._log_write(f"=== 開始執行（共 {total} 個檔案）===\n", "dim")
        self.progress.config(value=0, maximum=total)

        for i, item in enumerate(plan, 1):
            fname = os.path.basename(item["src"])
            self.status_var.set(f"處理中 {i}/{total}：{fname}")
            try:
                action = item["action"]
                if action == "move":
                    dst_dir = os.path.dirname(item["dst"])
                    os.makedirs(dst_dir, exist_ok=True)
                    shutil.move(item["src"], item["dst"])
                    self._log_write(f"✓ {item['src']}\n    →  {item['dst']}", "ok")
                elif action == "recycle":
                    send2trash(item["src"])
                    self._log_write(f"✓ 已移至資源回收桶：{item['src']}", "ok")
                elif action == "perm":
                    os.remove(item["src"])
                    self._log_write(f"✓ 已永久刪除：{item['src']}", "ok")
                ok += 1
            except Exception as e:
                err += 1
                self._log_write(f"✗ 失敗：{item['src']}\n    原因：{e}", "err")
            self.progress.config(value=i)

        self._plan = []
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
        self._plan = []
        self._log_clear()
        self.count_lbl.config(text="")
        self.progress.config(value=0)
        self.status_var.set("清單已清除。")


if __name__ == "__main__":
    App().mainloop()
