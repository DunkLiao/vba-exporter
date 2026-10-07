# -*- coding: utf-8 -*-
"""
Excel VBA 模組匯出工具（免安裝版）

- 支援單檔 / 批次多檔匯出 VBA 模組
- 使用獨立 Excel 實例（不會關掉使用者正在使用的 Excel）
- 唯讀開啟活簿，不修改、不鎖定來源檔
- 匯出在背景執行緒進行，介面不凍結
- 記住上次選擇的路徑（設定檔存在執行檔旁邊）
"""

import base64
import ctypes
import json
import os
import queue
import sys
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

import pythoncom
import win32com.client

APP_TITLE = "Excel VBA 模組匯出工具"
CONFIG_FILENAME = "VBAExporter.json"

# 視窗圖示（48x48 PNG，base64 編碼）
ICON_PNG_B64 = (
    "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"
)


# ============================================================
# 工具函式
# ============================================================

def app_dir():
    """應用程式所在資料夾（免安裝版：設定檔跟著 exe 走，不放使用者的 AppData）"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


CONFIG_PATH = os.path.join(app_dir(), CONFIG_FILENAME)


def load_config():
    """讀取上次的使用狀態；回傳 (來源檔清單, 輸出資料夾)"""
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as fp:
            data = json.load(fp)
        files = [p for p in data.get("files", [])
                 if isinstance(p, str) and os.path.isfile(p)]
        out_dir = data.get("output_dir", "")
        return files, (out_dir if isinstance(out_dir, str) else "")
    except Exception:
        return [], ""


def save_config(files, out_dir):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as fp:
            json.dump({"files": list(files), "output_dir": out_dir},
                      fp, ensure_ascii=False, indent=2)
    except Exception:
        pass  # 唯讀資料夾等情況： silently 忽略，不影響主流程


class ExcelNotInstalledError(Exception):
    """本機沒有安裝 Microsoft Excel"""


# ============================================================
# 核心匯出邏輯（與 UI 無關，方便單獨測試）
# 必須在已呼叫 pythoncom.CoInitialize() 的執行緒中使用
# ============================================================

def export_files(file_paths, out_dir, log):
    """
    匯出一或多個活簿內的全部 VBA 模組。

    回傳 list[dict]，每個檔案一筆：
        {"path", "status", "count", "target_dir"}
        status: ok / noaccess / error / missing
    """
    results = []
    excel = None
    try:
        try:
            # DispatchEx：建立「獨立」的 Excel 實例，
            # 不會附著到使用者已經開啟的 Excel，結束時也不會把它關掉
            excel = win32com.client.DispatchEx("Excel.Application")
        except Exception as e:
            msg = str(e)
            if "800401f3" in msg.lower() or "invalid class string" in msg.lower():
                raise ExcelNotInstalledError(
                    "偵測不到本機安裝的 Microsoft Excel") from e
            raise

        excel.Visible = False
        excel.DisplayAlerts = False

        multi = len(file_paths) > 1
        for path in file_paths:
            if not os.path.isfile(path):
                log(f"[警告] 檔案不存在，略過：{path}")
                results.append({"path": path, "status": "missing",
                                "count": 0, "target_dir": ""})
                continue

            # 多檔時：每個活簿匯出到「以檔名命名」的子資料夾，
            # 避免不同活簿的同名模組互相覆蓋
            stem = os.path.splitext(os.path.basename(path))[0]
            target_dir = out_dir if not multi else os.path.join(out_dir, stem)
            os.makedirs(target_dir, exist_ok=True)
            log(f"開始處理：{path}")

            wb = None
            count = 0
            status = "ok"
            try:
                # 唯讀開啟 + 不更新外部連結：不鎖檔、不修改來源檔
                wb = excel.Workbooks.Open(os.path.abspath(path),
                                          UpdateLinks=0, ReadOnly=True)
                try:
                    comps = wb.VBProject.VBComponents
                except Exception:
                    status = "noaccess"
                    log("  [錯誤] 無法存取 VBA 專案"
                        "（信任中心未開放存取，或專案有密碼保護）")
                else:
                    for comp in comps:
                        # 1=標準模組 2=類別模組 3=表單 100=文件模組
                        if comp.Type == 1:
                            ext = ".txt"   # 標準模組匯出成 .txt
                        elif comp.Type == 2:
                            ext = ".cls"
                        elif comp.Type == 3:
                            ext = ".frm"
                        else:
                            if comp.CodeModule.CountOfLines == 0:
                                log(f"  略過空白模組：{comp.Name}")
                                continue
                            ext = ".txt"
                        target = os.path.join(target_dir, comp.Name + ext)
                        comp.Export(target)
                        count += 1
                        log(f"  已匯出：{comp.Name}{ext}")
            except Exception as e:
                status = "error"
                log(f"  [錯誤] 處理失敗：{e}")
            finally:
                try:
                    if wb is not None:
                        wb.Close(SaveChanges=False)
                except Exception:
                    pass

            if status == "ok":
                log(f"  此檔案共匯出 {count} 個模組 → {target_dir}")
            results.append({"path": path, "status": status,
                            "count": count, "target_dir": target_dir})
    finally:
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass
        excel = None   # 儘早釋放 COM 參照，避免殘留背景 EXCEL.EXE
    return results


# ============================================================
# GUI
# ============================================================

class VBAExporterApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.resizable(False, False)

        self._file_display = tk.StringVar()   # 來源檔欄位顯示用
        self.output_dir = tk.StringVar()
        self.files = []                       # 實際選到的來源檔清單
        self._running = False
        self._queue = queue.Queue()           # worker → 主執行緒 訊息通道

        self._build_ui()
        self._center_window(700, 560)
        self._load_last_state()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self._poll_queue()   # 主執行緒佇列輪詢（負責所有 UI 更新）

    # ---------- UI 建構 ----------
    def _build_ui(self):
        pad = {"padx": 8, "pady": 6}

        # --- 1. 來源檔案（可多選） ---
        frm1 = tk.LabelFrame(self.root, text="1. 選擇來源 Excel 檔（可多選 .xlsm）")
        frm1.pack(fill="x", **pad)
        row1 = tk.Frame(frm1)
        row1.pack(fill="x", padx=6, pady=(8, 0))
        tk.Entry(row1, textvariable=self._file_display, state="readonly",
                 width=64).pack(side="left", fill="x", expand=True)
        self.btn_file = tk.Button(row1, text="瀏覽…", command=self.choose_files)
        self.btn_file.pack(side="left", padx=6)
        self.files_hint = tk.Label(frm1, text="", fg="#666666", anchor="w",
                                   font=("Microsoft JhengHei", 9))
        self.files_hint.pack(fill="x", padx=8, pady=(2, 6))

        # --- 2. 輸出資料夾 ---
        frm2 = tk.LabelFrame(self.root, text="2. 選擇匯出資料夾")
        frm2.pack(fill="x", **pad)
        row2 = tk.Frame(frm2)
        row2.pack(fill="x", padx=6, pady=8)
        tk.Entry(row2, textvariable=self.output_dir, width=64).pack(
            side="left", fill="x", expand=True)
        self.btn_dir = tk.Button(row2, text="瀏覽…", command=self.choose_folder)
        self.btn_dir.pack(side="left", padx=6)

        # --- 執行按鈕列 ---
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=10)
        self.btn_export = tk.Button(
            btn_frame, text="開始匯出", height=2, width=18,
            bg="#2d7d46", fg="white", activebackground="#256639",
            disabledforeground="#c9d6cc",
            font=("Microsoft JhengHei", 11, "bold"), command=self.export)
        self.btn_export.pack(side="left", padx=8)
        self.btn_reset = tk.Button(
            btn_frame, text="🔄 重設", height=2, width=12,
            bg="#8a8a8a", fg="white", activebackground="#777777",
            disabledforeground="#d5d5d5",
            font=("Microsoft JhengHei", 11, "bold"), command=self.reset)
        self.btn_reset.pack(side="left", padx=8)

        # --- 執行紀錄 ---
        frm3 = tk.LabelFrame(self.root, text="執行紀錄")
        frm3.pack(fill="both", expand=True, **pad)
        self.log_box = scrolledtext.ScrolledText(
            frm3, height=10, state="disabled", font=("Consolas", 10))
        self.log_box.pack(fill="both", expand=True, padx=6, pady=6)

    # ---------- UI 輔助 ----------
    def _center_window(self, w, h):
        """視窗置中，並依高 DPI 螢幕等比放大，避免在高解析度螢幕上過小"""
        try:
            dpi = self.root.winfo_fpixels("1i")
            scale = max(1.0, min(dpi / 96.0, 2.0))
        except Exception:
            scale = 1.0
        w, h = int(w * scale), int(h * scale)
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = max(0, (sw - w) // 2)
        y = max(0, (sh - h) // 2)
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    def _load_last_state(self):
        files, out_dir = load_config()
        self.files = files
        self.output_dir.set(out_dir)
        self._refresh_file_entry()

    def _refresh_file_entry(self):
        if not self.files:
            self._file_display.set("")
            self.files_hint.config(text="")
        elif len(self.files) == 1:
            self._file_display.set(self.files[0])
            self.files_hint.config(text="")
        else:
            self._file_display.set(f"已選擇 {len(self.files)} 個檔案")
            names = "、".join(os.path.basename(p) for p in self.files)
            if len(names) > 88:
                names = names[:88] + "…"
            self.files_hint.config(
                text=f"{names}（將各自匯出到以檔名命名的子資料夾）")

    # ---------- 事件處理 ----------
    def choose_files(self):
        paths = filedialog.askopenfilenames(
            title="選擇 Excel 檔案（可按住 Ctrl / Shift 多選）",
            filetypes=[("啟用巨集的 Excel", "*.xlsm"),
                       ("所有 Excel 檔", "*.xls;*.xlsx;*.xlsm"),
                       ("所有檔案", "*.*")])
        if paths:
            self.files = [os.path.normpath(p) for p in paths]
            self._refresh_file_entry()
            if not self.output_dir.get():
                default_out = os.path.join(
                    os.path.dirname(self.files[0]), "vba_export")
                self.output_dir.set(default_out)

    def choose_folder(self):
        path = filedialog.askdirectory(title="選擇匯出資料夾")
        if path:
            self.output_dir.set(os.path.normpath(path))

    def log(self, msg):
        ts = time.strftime("%H:%M:%S")
        self.log_box.config(state="normal")
        self.log_box.insert("end", f"[{ts}] {msg}\n")
        self.log_box.see("end")
        self.log_box.config(state="disabled")

    def reset(self):
        self.files = []
        self._file_display.set("")
        self.output_dir.set("")
        self.files_hint.config(text="")
        self.log_box.config(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.config(state="disabled")

    def _on_close(self):
        if not self._running:
            save_config(self.files, self.output_dir.get().strip())
        self.root.destroy()

    # ---------- 匯出流程 ----------
    def export(self):
        if self._running:
            return

        out_dir = self.output_dir.get().strip()
        files = [p for p in (f.strip() for f in self.files) if p]

        if not files:
            messagebox.showwarning("提醒", "請先選擇來源 Excel 檔案。")
            return
        missing = [p for p in files if not os.path.isfile(p)]
        if len(missing) == len(files):
            messagebox.showwarning("提醒", "選擇的檔案都不存在，請重新選擇。")
            return
        if missing and not messagebox.askyesno(
                "提醒", f"有 {len(missing)} 個檔案已不存在，將予以略過。是否繼續？"):
            return
        if not out_dir:
            messagebox.showwarning("提醒", "請先選擇匯出資料夾。")
            return

        self._set_busy(True)
        self.log("─" * 46)
        self.log(f"開始匯出 {len(files)} 個檔案 → {out_dir}")
        threading.Thread(target=self._export_worker,
                         args=(files, out_dir), daemon=True).start()

    def _set_busy(self, busy):
        self._running = busy
        state = "disabled" if busy else "normal"
        for btn in (self.btn_export, self.btn_reset,
                    self.btn_file, self.btn_dir):
            btn.config(state=state)
        self.root.config(cursor="watch" if busy else "")

    def _export_worker(self, files, out_dir):
        """背景執行緒：實際進行匯出。只透過 queue 與主執行緒溝通。"""
        q = self._queue
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass
        try:
            try:
                results = export_files(
                    files, out_dir, log=lambda m: q.put(("log", m)))
            except ExcelNotInstalledError:
                q.put(("log", "發生錯誤：找不到本機安裝的 Microsoft Excel。"))
                q.put(("msgbox", "error", "找不到 Excel",
                       "偵測不到本機安裝的 Microsoft Excel。\n\n"
                       "此工具需要 Excel 才能運作，請確認這部電腦已安裝 Excel。"))
            except Exception as e:
                q.put(("log", f"發生未預期錯誤：{e}"))
                q.put(("msgbox", "error", "錯誤", str(e)))
            else:
                ok = [r for r in results if r["status"] == "ok"]
                failed = [r for r in results if r["status"] != "ok"]
                total = sum(r["count"] for r in ok)
                q.put(("log", f"完成！共處理 {len(ok)} 個檔案、"
                              f"匯出 {total} 個模組到：{out_dir}"))
                # 匯出完成後自動開啟資料夾
                try:
                    os.startfile(out_dir)
                except Exception as e:
                    q.put(("log", f"（無法自動開啟資料夾：{e}）"))

                if len(files) == 1 and failed and failed[0]["status"] == "noaccess":
                    q.put(("msgbox", "error", "無法存取 VBA 專案",
                           "請先到 Excel：\n"
                           "檔案 → 選項 → 信任中心 → 信任中心設定 → 巨集設定\n"
                           "勾選「信任存取 VBA 專案物件模型」後再試一次。\n\n"
                           "若此檔案的 VBA 專案有密碼保護，請先解除密碼再匯出。"))
                else:
                    text = (f"共處理 {len(ok)} 個檔案、匯出 {total} 個模組！\n"
                            f"已為你開啟資料夾。")
                    if failed:
                        text += (f"\n\n（{len(failed)} 個檔案失敗或無法存取，"
                                 f"詳見執行紀錄）")
                        q.put(("msgbox", "warning", "完成（部分失敗）", text))
                    else:
                        q.put(("msgbox", "info", "完成", text))
                q.put(("savecfg",))
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            q.put(("done",))

    # ---------- 主執行緒：處理 worker 送來的事件 ----------
    def _poll_queue(self):
        try:
            while True:
                item = self._queue.get_nowait()
                kind = item[0]
                if kind == "log":
                    self.log(item[1])
                elif kind == "msgbox":
                    _, box_type, title, text = item
                    if box_type == "info":
                        messagebox.showinfo(title, text)
                    elif box_type == "warning":
                        messagebox.showwarning(title, text)
                    else:
                        messagebox.showerror(title, text)
                elif kind == "savecfg":
                    save_config(self.files, self.output_dir.get().strip())
                elif kind == "done":
                    self._set_busy(False)
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)


# ============================================================
# 進入點
# ============================================================

def enable_dpi_awareness():
    """讓 Windows 知道本程式支援高 DPI，避免視窗模糊"""
    if sys.platform != "win32":
        return
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)   # Win 8.1+
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()    # 舊系統備援
        except Exception:
            pass


def main():
    enable_dpi_awareness()
    root = tk.Tk()
    try:
        icon = tk.PhotoImage(data=base64.b64decode(ICON_PNG_B64))
        root.iconphoto(True, icon)
    except Exception:
        icon = None
    app = VBAExporterApp(root)   # noqa: F841（保持參照，避免被 GC）
    app._icon_ref = icon
    root.mainloop()


if __name__ == "__main__":
    main()
