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
    def __init__(self, parent=None, offset_x=0, offset_y=0):
        # 若有 parent 則用 Toplevel，否則用 Tk
        if parent is not None:
            self.root = tk.Toplevel(parent)
            self.root.transient(parent)
            self.root.lift()
            self.root.focus_force()
        else:
            self.root = tk.Tk()
        self.root.title("錯誤碼查詢工具 v1.5.0")
        self.root.geometry("1400x800")
        self.root.minsize(1200, 600)
        self.root.resizable(True, True)
        # 確保視窗有完整的控制按鈕
        self.root.wm_attributes('-topmost', False)
        # 預設最大化視窗
        self.root.state('zoomed')  # Windows 最大化
        # 初始化設定管理
        self.config_manager = ConfigManager()
        # 檢查 setup.txt 是否有 ExcelErrorCodeSearch_TIP，若無則自動寫入
        tip_key = 'ExcelErrorCodeSearch_TIP'
        default_tip = (
            "錯誤碼查詢工具使用說明：\n\n"
            "1. 檔案載入：\n"
            "   • 點選「選擇 Excel 檔案」載入包含 Test Item All 工作表的 Excel 檔案\n"
            "   • 系統會自動讀取 BCDE 欄位的資料（介面、內部錯誤代碼、描述、中文描述）\n\n"
            "2. 關鍵字搜尋：\n"
            "   • 可同時輸入 1-3 個關鍵字進行搜尋\n"
            "   • 支援中英文模糊搜尋，不區分大小寫\n"
            "   • 多個關鍵字會以 AND 條件搜尋（必須同時包含所有關鍵字）\n"
            "   • 按 Enter 鍵或點選「搜尋」按鈕執行搜尋\n\n"
            "3. 結果操作：\n"
            "   • 右鍵點選、雙擊或按 Ctrl+C 可複製 Interface\n"
            "   • 只複製介面欄位（第一欄位）\n"
            "   • 選中的行會以藍色高亮顯示\n"
            "   • 搜尋結果以藍色字體顯示，方便辨識\n"
            "   • 支援鍵盤上下左右移動瀏覽結果\n\n"
            "4. 字體調整：\n"
            "   • 使用 + - 按鈕調整字體大小（8-20）\n"
            "   • 字體大小會自動儲存，下次開啟時保持設定\n\n"
            "5. 其他功能：\n"
            "   • 「總計」顯示目前資料筆數\n"
            "   • 視窗可調整大小，支援最大化\n"
            "   • 表格支援水平和垂直捲軸"
        )
        if not self.config_manager.get(tip_key):
            self.config_manager.set(tip_key, default_tip)
        self.df = None  # 儲存 Test Item All sheet 的 DataFrame
        # 讀取字體大小與上次檔案路徑
        self.font_size = int(self.config_manager.get('SearchUIFontSize', self.config_manager.get('FontSize', 12)))
        self.last_excel_path = self.config_manager.get('LastExcelPath', os.getcwd())
        self._setup_ui()
        self.tip_window = None  # 用於 toggle 說明視窗
        self.center_window(offset_x, offset_y)

    def center_window(self, offset_x=0, offset_y=0):
        self.root.update_idletasks()
        w = self.root.winfo_width()
        h = self.root.winfo_height()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = (sw - w) // 2 + offset_x
        y = (sh - h) // 2 + offset_y
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    def _setup_ui(self):
        # 主分割區
        main_pane = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_pane.pack(fill=tk.BOTH, expand=True)

        # 左側控制區（無捲軸，內容由上往下）
        control_frame = ttk.Frame(main_pane, width=400)
        main_pane.add(control_frame, weight=0)

        # 右側顯示區
        display_frame = ttk.Frame(main_pane, padding=5)
        main_pane.add(display_frame, weight=1)

        # ===== 左側：控制區 =====
        # 選擇檔案按鈕
        self.file_btn = ttk.Button(control_frame, text="選擇 Excel 檔案", command=self.select_file, style="Custom.TButton")
        self.file_btn.pack(fill=tk.X, pady=5)
        self._add_hand_over(self.file_btn)

        # 顯示檔案路徑（初始顯示提示文字）
        self.file_label = ttk.Label(control_frame, text="請選擇 Excel 檔案", anchor="w", foreground="gray")
        self.file_label.pack(fill=tk.X, pady=5)

        # 查詢欄位標題
        ttk.Label(control_frame, text="關鍵字搜尋", font=("Microsoft JhengHei", self.font_size, "bold")).pack(pady=(10, 5))
        
        # 查詢欄位（最多三個）
        self.query_entries = []
        for i in range(3):
            entry_frame = ttk.Frame(control_frame)
            entry_frame.pack(fill=tk.X, pady=2)
            
            ttk.Label(entry_frame, text=f"關鍵字 {i+1}:", font=("Microsoft JhengHei", self.font_size)).pack(side=tk.LEFT, padx=(0, 5))
            entry = ttk.Entry(entry_frame, font=("Microsoft JhengHei", self.font_size))
            entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
            entry.bind("<Return>", lambda e: self.search())
            self.query_entries.append(entry)
        
        ttk.Label(control_frame, text="(支援中英文模糊搜尋，可同時使用多個關鍵字)", 
                 font=("Microsoft JhengHei", self.font_size-1), foreground="gray").pack(pady=2)

        # 搜尋按鈕
        self.search_btn = ttk.Button(control_frame, text="搜尋", command=self.search, style="Custom.TButton")
        self.search_btn.pack(fill=tk.X, pady=8)
        self._add_hand_over(self.search_btn)

        # 字體大小調整
        fontsize_frame = ttk.Frame(control_frame)
        fontsize_frame.pack(pady=8)
        
        ttk.Label(fontsize_frame, text="字體大小:", font=("Microsoft JhengHei", self.font_size)).pack(side=tk.LEFT, padx=(0, 5))
        self.minus_btn = ttk.Button(fontsize_frame, text="－", width=3, command=self.decrease_fontsize)
        self.minus_btn.pack(side=tk.LEFT, padx=2)
        
        self.font_label = ttk.Label(fontsize_frame, text=str(self.font_size), font=("Microsoft JhengHei", self.font_size, "bold"))
        self.font_label.pack(side=tk.LEFT, padx=5)
        
        self.plus_btn = ttk.Button(fontsize_frame, text="＋", width=3, command=self.increase_fontsize)
        self.plus_btn.pack(side=tk.LEFT, padx=2)

        # 資料筆數計數器（移到 + - 按鈕下方，置中顯示）
        total_label = self.config_manager.get('TotalCountLabel', '總計')
        count_unit = self.config_manager.get('CountUnit', '筆資料')
        self.count_label = ttk.Label(control_frame, text=f"{total_label}：0 {count_unit}", anchor="center", font=("Microsoft JhengHei", self.font_size+2, "bold"))
        self.count_label.pack(fill=tk.X, pady=16)

        # 使用說明按鈕（移到總計下方）
        self.tip_btn = ttk.Button(control_frame, text="使用說明", command=self.show_tip, style="Custom.TButton")
        self.tip_btn.pack(fill=tk.X, pady=5)
        self._add_hand_over(self.tip_btn)

        # ===== 右側：顯示區 =====
        # 用 grid 方式排版，讓卷軸緊貼表格右側
        display_frame.grid_rowconfigure(0, weight=1)
        display_frame.grid_columnconfigure(0, weight=1)
        self.tree = ttk.Treeview(display_frame, columns=[], show="headings", height=20, style="Custom.Treeview")
        self.tree.grid(row=0, column=0, sticky="nsew", padx=(5,0), pady=5)
        self.tree.tag_configure("highlight", background="#FFFACD", foreground="#000000")  # error code 高亮，黑色文字
        self.tree.tag_configure("search_blue", background="", foreground="#0070C0")  # 搜尋關鍵字 row 文字藍色
        self.tree.bind("<Button-3>", self.copy_row_popup)  # 右鍵複製
        self.tree.bind("<Key>", self.on_tree_key)
        self.tree.bind("<Double-1>", self.copy_row_popup)  # 雙擊也可複製
        # 添加 Ctrl+C 快捷鍵複製
        self.tree.bind("<Control-c>", self.copy_row_popup)

        # 垂直捲軸（加大寬度，緊貼表格右側）
        yscroll = ttk.Scrollbar(display_frame, orient="vertical", command=self.tree.yview, style="Vertical.TScrollbar")
        yscroll.grid(row=0, column=1, sticky="ns", pady=5)
        self.tree.configure(yscrollcommand=yscroll.set)
        # 水平捲軸
        xscroll = ttk.Scrollbar(display_frame, orient="horizontal", command=self.tree.xview, style="Horizontal.TScrollbar")
        xscroll.grid(row=1, column=0, sticky="ew", padx=(5,0))
        self.tree.configure(xscrollcommand=xscroll.set)

        # 美化表格格線
        style = ttk.Style()
        # 增加行高，讓文字不會太擠
        style.configure("Custom.Treeview", rowheight=40, borderwidth=1, relief="solid")
        style.layout("Custom.Treeview", [
            ("Treeview.treearea", {'sticky': 'nswe'})
        ])
        style.map("Custom.Treeview", 
                 background=[('selected', '#0078D4')],  # 更明顯的藍色
                 foreground=[('selected', 'white')])    # 選中時文字變白色
        style.configure("Custom.Treeview.Heading", borderwidth=1, relief="solid", font=("Microsoft JhengHei", self.font_size, "bold"))
        # 設定捲軸樣式
        style.configure("Vertical.TScrollbar", width=24)  # 加大垂直捲軸寬度
        style.configure("Horizontal.TScrollbar", height=20)  # 加大水平捲軸高度

        # 若有上次檔案路徑自動載入
        if self.last_excel_path and os.path.exists(self.last_excel_path):
            try:
                # 讀取 Excel 檔案，跳過前3行空行，第4行是標題
                df = pd.read_excel(self.last_excel_path, sheet_name="Test Item All", skiprows=3)
                # 重新命名欄位
                if len(df.columns) >= 8:
                    df.columns = [
                        'Main Function', 'Interface', 'Interenal Error Code', 
                        'Description', 'Chinese', 'Version', 'Error Code', 'Note'
                    ]
                # 只取 BCDE 欄（Interface, Interenal Error Code, Description, Chinese）
                if df.shape[1] >= 5:
                    df = df.iloc[:, 1:5]
                    # 重新命名為正確的欄位名稱
                    df.columns = ['Interface', 'Interenal Error Code', 'Description', 'Chinese']
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
        # 選擇 Excel 檔案，讀取 Test Item All sheet，只取 C-G 欄，header=0
        file_path = filedialog.askopenfilename(
            title="選擇 Excel 檔案",
            filetypes=[("Excel files", "*.xlsx *.xls")],
            initialdir=os.path.dirname(self.last_excel_path) if self.last_excel_path else os.getcwd()
        )
        if not file_path:
            return
        try:
            # 讀取 Excel 檔案，跳過前3行空行，第4行是標題
            df = pd.read_excel(file_path, sheet_name="Test Item All", skiprows=3)
            # 重新命名欄位
            if len(df.columns) >= 8:
                df.columns = [
                    'Main Function', 'Interface', 'Interenal Error Code', 
                    'Description', 'Chinese', 'Version', 'Error Code', 'Note'
                ]
            # 只取 BCDE 欄（Interface, Interenal Error Code, Description, Chinese）
            if df.shape[1] >= 5:
                df = df.iloc[:, 1:5]
                # 重新命名為正確的欄位名稱
                df.columns = ['Interface', 'Interenal Error Code', 'Description', 'Chinese']
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
            total_label = self.config_manager.get('TotalCountLabel', '總計')
            count_unit = self.config_manager.get('CountUnit', '筆資料')
            self.count_label.config(text=f"{total_label}：0 {count_unit}")
            return
        # 只顯示 BCDE 欄
        if df.shape[1] >= 4:
            df = df.iloc[:, 0:4]
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
        # 計算每一欄最大寬度
        col_widths = {col: max([len(str(col))] + [len(str(row[col])) for _, row in df.iterrows()]) for col in columns}
        # 取得搜尋關鍵字
        queries = [e.get().strip() for e in self.query_entries if e.get().strip()]
        for col in columns:
            self.tree.heading(col, text=col)
            # 增加欄位寬度，讓文字不會太擠
            width = min(max(col_widths[col]*20, 150), 600)
            self.tree.column(col, width=width, anchor="w")
        # 插入資料，error code 欄位高亮，搜尋關鍵字 row 文字顯示藍色
        for _, row in df.iterrows():
            # 處理資料，保持原始格式但增加適當的間距
            values = [str(row[col]).replace("\\n", "\n") for col in columns]
            
            tag = "highlight" if error_code_candidates and str(row[error_code_candidates[0]]).strip() else ""
            # 若有搜尋關鍵字，且該 row 有任一 cell 包含關鍵字，則加上 search_blue tag
            if queries and any(any(q in str(cell) for q in queries) for cell in row):
                tag = "search_blue"
            self.tree.insert("", "end", values=values, tags=(tag,))
        self._set_all_fontsize()
        # 更新資料筆數
        total_label = self.config_manager.get('TotalCountLabel', '總計')
        count_unit = self.config_manager.get('CountUnit', '筆資料')
        self.count_label.config(text=f"{total_label}：{len(df)} {count_unit}")

    def _set_all_fontsize(self):
        # 設定所有元件（左側控制區、表格等）的字體
        font = ("Microsoft JhengHei", self.font_size)
        style = ttk.Style()
        style.configure("Treeview", font=font)
        style.configure("Treeview.Heading", font=("Microsoft JhengHei", self.font_size, "bold"))
        style.configure("Custom.Treeview.Heading", font=("Microsoft JhengHei", self.font_size, "bold"))
        self.tree.tag_configure("highlight", font=font)
        self.tree.tag_configure("search_blue", font=font)
        
        # 更新字體大小顯示
        self.font_label.config(text=str(self.font_size))
        
        # 更新總計標籤字體
        self.count_label.config(font=("Microsoft JhengHei", self.font_size+2, "bold"))
        
        # 左側所有元件
        for widget in self.root.winfo_children():
            for child in widget.winfo_children():
                try:
                    child.configure(font=font)
                except Exception:
                    pass
        # 更新設定檔（SearchUIFontSize）
        self.config_manager.set('SearchUIFontSize', str(self.font_size))

    def increase_fontsize(self):
        if self.font_size < 20:  # 限制最大字體大小
            self.font_size += 1
            self._set_all_fontsize()

    def decrease_fontsize(self):
        if self.font_size > 8:  # 限制最小字體大小
            self.font_size -= 1
            self._set_all_fontsize()

    def copy_row_popup(self, event):
        # 右鍵選單複製 Interface 欄位
        iid = self.tree.identify_row(event.y)
        if iid:
            row_values = self.tree.item(iid, "values")
            # 複製第一個欄位（Interface）
            interface = row_values[0] if len(row_values) > 0 else ""
            if interface and interface.strip():
                self.root.clipboard_clear()
                self.root.clipboard_append(str(interface))
                # 顯示複製成功的提示
                self._show_copy_tooltip(event.x_root, event.y_root, f"已複製: {interface}")
            else:
                self._show_copy_tooltip(event.x_root, event.y_root, "無 Interface 可複製")
    
    def _show_copy_tooltip(self, x, y, message):
        """顯示複製提示"""
        tooltip = tk.Toplevel(self.root)
        tooltip.wm_overrideredirect(True)
        tooltip.wm_geometry(f"+{x+10}+{y+10}")
        
        label = tk.Label(tooltip, text=message, 
                        font=("Microsoft JhengHei", 9),
                        bg="#333333", fg="white", 
                        padx=8, pady=4,
                        relief="solid", borderwidth=1)
        label.pack()
        
        # 2秒後自動關閉
        tooltip.after(2000, tooltip.destroy)

    def on_tree_key(self, event):
        # 支援鍵盤上下左右移動與字體調整
        if event.keysym in ("Up", "Down", "Left", "Right"):
            return  # Treeview 原生支援
        if event.keysym == "plus" or event.char == '+':
            self.increase_fontsize()
        elif event.keysym == "minus" or event.char == '-':
            self.decrease_fontsize()

    def show_tip(self):
        # toggle 說明視窗
        if self.tip_window and self.tip_window.winfo_exists():
            self.tip_window.destroy()
            self.tip_window = None
            return
        tip = self.config_manager.get('ExcelErrorCodeSearch_TIP', '請洽管理員補充說明')
        tip = tip.replace('\\n', '\n').replace('\r\n', '\n').replace('\n', '\n')
        win = tk.Toplevel(self.root)
        win.title("錯誤碼查詢工具 - 使用說明")
        win.geometry("600x500")
        win.resizable(True, True)  # 允許調整大小
        win.minsize(500, 400)  # 設定最小尺寸
        # 置中於查詢UI
        self.root.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - 600) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - 500) // 2
        win.geometry(f"600x500+{x}+{y}")
        
        # 創建文字區域和捲軸
        text_frame = tk.Frame(win)
        text_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        text_widget = tk.Text(text_frame, font=("Microsoft JhengHei", 11), 
                             wrap=tk.WORD, state=tk.DISABLED, bg="#f8f9fa")
        text_widget.pack(side=tk.LEFT, fill="both", expand=True)
        
        scrollbar = tk.Scrollbar(text_frame, orient="vertical", command=text_widget.yview)
        scrollbar.pack(side=tk.RIGHT, fill="y")
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        # 插入文字
        text_widget.config(state=tk.NORMAL)
        text_widget.insert(tk.END, tip)
        text_widget.config(state=tk.DISABLED)
        
        # 按鈕框架
        btn_frame = tk.Frame(win)
        btn_frame.pack(pady=10)
        btn = ttk.Button(btn_frame, text="確定", command=win.destroy, style="Custom.TButton")
        btn.pack()
        self.tip_window = win
        win.protocol("WM_DELETE_WINDOW", lambda: (win.destroy(), setattr(self, 'tip_window', None)))

# if __name__ == "__main__":
#     root = tk.Tk()
#     app = ExcelErrorCodeSearchUI(root)
#     root.mainloop() 