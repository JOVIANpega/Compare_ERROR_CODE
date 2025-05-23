"""
Excel Error Code 查詢工具 UI
- 支援選擇 Excel 檔案，讀取 Test Item All sheet
- 查詢字串模糊搜尋，支援中英文、多關鍵字（最多三個）
- 結果以表格顯示，error code 欄位高亮
- UI 可調整大小，按鈕 hand over style
- 支援鍵盤上下左右、字體大小調整、右鍵複製 row
- 字體大小、上次檔案路徑自動寫入/讀取 setup.txt
- 左右分割視窗，左：控制區，右：顯示區
- 程式碼有完整註解
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
from config_manager import ConfigManager

class ExcelErrorCodeSearchUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Error Code 查詢工具")
        self.root.geometry("1100x650")
        self.root.minsize(900, 400)
        self.root.resizable(True, True)
        # 初始化設定管理
        self.config_manager = ConfigManager()
        self.df = None  # 儲存 Test Item All sheet 的 DataFrame
        # 讀取字體大小與上次檔案路徑
        self.font_size = int(self.config_manager.get('FontSize', 12))
        self.last_excel_path = self.config_manager.get('LastExcelPath', os.getcwd())
        self._setup_ui()

    def _setup_ui(self):
        # 主分割區
        main_pane = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_pane.pack(fill=tk.BOTH, expand=True)

        # 左側控制區（無捲軸，內容由上往下）
        control_frame = ttk.Frame(main_pane, width=320)
        main_pane.add(control_frame, weight=0)

        # 右側顯示區
        display_frame = ttk.Frame(main_pane, padding=5)
        main_pane.add(display_frame, weight=1)

        # ===== 左側：控制區 =====
        # 選擇檔案按鈕
        self.file_btn = ttk.Button(control_frame, text="選擇 Excel 檔案", command=self.select_file, style="Custom.TButton")
        self.file_btn.pack(fill=tk.X, pady=5)
        self._add_hand_over(self.file_btn)

        # 顯示檔案路徑
        self.file_label = ttk.Label(control_frame, text="尚未選擇檔案", anchor="w")
        self.file_label.pack(fill=tk.X, pady=5)
        if self.last_excel_path and os.path.exists(self.last_excel_path):
            self.file_label.config(text=os.path.basename(self.last_excel_path))

        # 查詢欄位（最多三個）
        self.query_entries = []
        for i in range(3):
            entry = ttk.Entry(control_frame, width=18, font=("Microsoft JhengHei", self.font_size))
            entry.pack(fill=tk.X, pady=2)
            entry.bind("<Return>", lambda e: self.search())
            self.query_entries.append(entry)
        ttk.Label(control_frame, text="(可輸入1~3個關鍵字)").pack(pady=2)

        # 搜尋按鈕
        self.search_btn = ttk.Button(control_frame, text="搜尋", command=self.search, style="Custom.TButton")
        self.search_btn.pack(fill=tk.X, pady=8)
        self._add_hand_over(self.search_btn)

        # 字體大小調整
        fontsize_frame = ttk.Frame(control_frame)
        fontsize_frame.pack(pady=8)
        self.plus_btn = ttk.Button(fontsize_frame, text="＋", width=2, command=self.increase_fontsize)
        self.plus_btn.pack(side=tk.LEFT, padx=2)
        self.minus_btn = ttk.Button(fontsize_frame, text="－", width=2, command=self.decrease_fontsize)
        self.minus_btn.pack(side=tk.LEFT, padx=2)

        # ===== 右側：顯示區 =====
        # 結果表格
        self.tree = ttk.Treeview(display_frame, columns=[], show="headings", height=20, style="Custom.Treeview")
        self.tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.tree.tag_configure("highlight", background="#FFFACD")  # error code 高亮
        self.tree.bind("<Button-3>", self.copy_row_popup)  # 右鍵複製
        self.tree.bind("<Key>", self.on_tree_key)
        self.tree.bind("<Double-1>", self.copy_row_popup)  # 雙擊也可複製

        # 捲軸
        yscroll = ttk.Scrollbar(display_frame, orient="vertical", command=self.tree.yview)
        yscroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=yscroll.set)
        xscroll = ttk.Scrollbar(display_frame, orient="horizontal", command=self.tree.xview)
        xscroll.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.configure(xscrollcommand=xscroll.set)

        # 美化表格格線
        style = ttk.Style()
        style.configure("Custom.Treeview", rowheight=28, borderwidth=1, relief="solid")
        style.layout("Custom.Treeview", [
            ("Treeview.treearea", {'sticky': 'nswe'})
        ])
        style.map("Custom.Treeview", background=[('selected', '#3399FF')])
        style.configure("Custom.Treeview.Heading", borderwidth=1, relief="solid")

        # 若有上次檔案路徑自動載入
        if self.last_excel_path and os.path.exists(self.last_excel_path):
            try:
                df = pd.read_excel(self.last_excel_path, sheet_name="Test Item All")
                if df.shape[1] >= 7:
                    df = df.iloc[:, 2:7]
                self.df = df
                self._show_table(self.df)
            except Exception:
                pass

    def _add_hand_over(self, btn):
        # hand over style: 預設灰色，滑鼠經過變綠
        style = ttk.Style()
        style.configure("Custom.TButton", background="#cccccc")
        style.map("Custom.TButton",
                  background=[("active", "#28a745"), ("!active", "#cccccc")],
                  foreground=[("active", "#fff"), ("!active", "#000")])

    def select_file(self):
        # 選擇 Excel 檔案，讀取 Test Item All sheet，只取 C-G 欄
        file_path = filedialog.askopenfilename(
            title="選擇 Excel 檔案",
            filetypes=[("Excel files", "*.xlsx *.xls")],
            initialdir=os.path.dirname(self.last_excel_path) if self.last_excel_path else os.getcwd()
        )
        if not file_path:
            return
        try:
            df = pd.read_excel(file_path, sheet_name="Test Item All")
            # 只取 C-G 欄（index 2~6）
            if df.shape[1] >= 7:
                df = df.iloc[:, 2:7]
            self.df = df
            self.file_label.config(text=os.path.basename(file_path))
            self._show_table(self.df)
            # 更新設定檔
            self.last_excel_path = file_path
            self.config_manager.update_last_paths(excel_path=file_path)
        except Exception as e:
            messagebox.showerror("讀取失敗", f"無法讀取 Test Item All：\n{e}")
            self.df = None
            self.file_label.config(text="尚未選擇檔案")
            self._show_table(None)

    def search(self):
        # 執行多關鍵字查詢（AND 條件）
        if self.df is None:
            messagebox.showwarning("請先選擇檔案", "請先選擇 Excel 檔案並成功載入 Test Item All sheet！")
            return
        queries = [e.get().strip() for e in self.query_entries if e.get().strip()]
        if not queries:
            self._show_table(self.df)
            return
        mask = self.df.apply(lambda row: all(row.astype(str).str.contains(q, case=False, na=False).any() for q in queries), axis=1)
        result_df = self.df[mask]
        self._show_table(result_df)
        if result_df.empty:
            messagebox.showinfo("查無資料", f"找不到同時包含「{'、'.join(queries)}」的資料。")

    def _show_table(self, df):
        # 清空表格
        for col in self.tree["columns"]:
            self.tree.heading(col, text="")
            self.tree.column(col, width=100)
        self.tree.delete(*self.tree.get_children())
        if df is None or df.empty:
            self.tree["columns"] = []
            return
        # 只顯示 C-G 欄
        if df.shape[1] >= 7:
            df = df.iloc[:, 0:5]
        # 將 nan 轉成空字串，並過濾全空 row
        df = df.fillna("")
        df = df.loc[~df.apply(lambda row: all(str(cell).strip() == "" for cell in row), axis=1)]
        # 設定欄位，將 error code 欄位（如有）放最前面
        columns = list(df.columns)
        error_code_candidates = [c for c in columns if "error" in str(c).lower() or "code" in str(c).lower()]
        if error_code_candidates:
            first_col = error_code_candidates[0]
            columns.remove(first_col)
            columns = [first_col] + columns
        self.tree["columns"] = columns
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=180, anchor="w")
        # 插入資料，error code 欄位高亮，內容自動換行
        for _, row in df.iterrows():
            values = [str(row[col]).replace("\\n", "\n") for col in columns]
            tag = "highlight" if error_code_candidates and str(row[error_code_candidates[0]]).strip() else ""
            self.tree.insert("", "end", values=values, tags=(tag,))
        self._set_tree_fontsize()

    def _set_tree_fontsize(self):
        # 用 ttk.Style 設定整個 Treeview 字體
        font = ("Microsoft JhengHei", self.font_size)
        style = ttk.Style()
        style.configure("Treeview", font=font)
        style.configure("Treeview.Heading", font=font)
        self.tree.tag_configure("highlight", font=font)
        # 更新設定檔
        self.config_manager.set('FontSize', str(self.font_size))

    def increase_fontsize(self):
        self.font_size += 1
        self._set_tree_fontsize()

    def decrease_fontsize(self):
        if self.font_size > 8:
            self.font_size -= 1
            self._set_tree_fontsize()

    def copy_row_popup(self, event):
        # 右鍵選單複製 error code 欄位
        iid = self.tree.identify_row(event.y)
        if iid:
            row_values = self.tree.item(iid, "values")
            error_code = row_values[0] if row_values else ""
            self.root.clipboard_clear()
            self.root.clipboard_append(str(error_code))
            # 不彈窗，直接複製

    def on_tree_key(self, event):
        # 支援鍵盤上下左右移動與字體調整
        if event.keysym in ("Up", "Down", "Left", "Right"):
            return  # Treeview 原生支援
        if event.keysym == "plus" or event.char == '+':
            self.increase_fontsize()
        elif event.keysym == "minus" or event.char == '-':
            self.decrease_fontsize()

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelErrorCodeSearchUI(root)
    root.mainloop() 