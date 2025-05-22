"""
UI管理模組
負責處理所有UI相關的操作，包括視窗、按鈕、標籤等元件的建立和管理
"""
import tkinter as tk
import ttkbootstrap as tb
from ttkbootstrap.constants import *
import tkinter.font as tkfont
from tkinter import messagebox, filedialog
from typing import Callable, Optional
import logging
import os

logger = logging.getLogger(__name__)

class ToolTip:
    """工具提示類別"""
    def __init__(self, widget: tk.Widget, text: str):
        self.widget = widget
        self.text = text
        self.tipwindow = None
        widget.bind("<Enter>", self.show_tip)
        widget.bind("<Leave>", self.hide_tip)

    def show_tip(self, event=None):
        if self.tipwindow or not self.text:
            return
        x, y, cx, cy = self.widget.bbox("insert") if hasattr(self.widget, 'bbox') else (0,0,0,0)
        x = x + self.widget.winfo_rootx() + 30
        y = y + self.widget.winfo_rooty() + 20
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(tw, text=self.text, justify='left',
                        background="#ffffe0", relief='solid',
                        borderwidth=1, font=("Calibri", 10))
        label.pack(ipadx=4)

    def hide_tip(self, event=None):
        tw = self.tipwindow
        self.tipwindow = None
        if tw:
            tw.destroy()

class UIManager:
    def __init__(self, root: tk.Tk, config_manager):
        self.root = root
        self.config_manager = config_manager
        self.excel1_path: Optional[str] = None
        self.excel2_path: Optional[str] = None
        self.selected_sheet: Optional[str] = None
        self.sheet_load_callback = None  # 新增 callback 屬性
        
        # 初始化UI
        self._init_ui()
        self._setup_window()
        logger.info("UI初始化完成")

    def _init_ui(self):
        """初始化UI元件"""
        # 建立樣式
        self.style = tb.Style()
        self._setup_styles()
        
        # 建立主框架
        self.main_frame = tb.Frame(self.root, padding=10, style="Main.TFrame")
        self.main_frame.pack(fill=BOTH, expand=YES)
        
        # 建立UI元件
        self._create_widgets()
        
        # 設定字體大小
        self.font_size = int(self.config_manager.get('FontSize', 11))
        self.set_all_font_size(self.font_size)

    def _setup_styles(self):
        """設定UI樣式"""
        self.style.configure("Big.TButton",
                           font=("Microsoft JhengHei", 16),
                           padding=10,
                           foreground="#fff",
                           background="#3399ff")
        
        self.style.map("Big.TButton",
                      background=[("active", "#28a745")],
                      foreground=[("active", "#fff")])
        
        self.style.configure("Main.TFrame", background="#e6f2ff")

    def _setup_window(self):
        """設定視窗屬性"""
        self.root.title("Error Code Comparer")
        width = int(self.config_manager.get('WindowWidth', 540))
        height = int(self.config_manager.get('WindowHeight', 340))
        self.root.geometry(f"{width}x{height}")
        self.root.resizable(True, True)
        self.center_window()
        self.root.configure(bg="#e6f2ff")

    def _create_widgets(self):
        """建立UI元件"""
        # Error Code XML
        self._create_file_selection_row(0, "ErrorCodeXMLLabel", self.browse_excel1)
        
        # Source Excel
        self._create_file_selection_row(1, "SourceExcelLabel", self.browse_excel2)
        
        # Sheet select
        self._create_sheet_selection_row(2)
        
        # Compare button
        self._create_compare_button(3)

    def _create_file_selection_row(self, row: int, label_key: str, browse_command: Callable):
        """建立檔案選擇行"""
        lbl = tb.Label(self.main_frame, text=self.config_manager.get(label_key))
        lbl.grid(row=row, column=0, sticky=W, pady=10)
        
        entry = tb.Entry(self.main_frame, width=45)
        entry.grid(row=row, column=1, padx=5, pady=10)
        
        btn = tb.Button(self.main_frame,
                       text=self.config_manager.get('BrowseButton'),
                       bootstyle="outline-primary",
                       command=browse_command,
                       style="Big.TButton")
        btn.grid(row=row, column=2, padx=5, pady=10)
        
        if label_key == "ErrorCodeXMLLabel":
            self.excel1_entry = entry
            self.browse1_btn = btn
            ToolTip(btn, self.config_manager.get('BrowseXMLTooltip'))
        else:
            self.excel2_entry = entry
            self.browse2_btn = btn
            ToolTip(btn, self.config_manager.get('BrowseExcelTooltip'))

    def _create_sheet_selection_row(self, row: int):
        """建立工作表選擇行"""
        lbl = tb.Label(self.main_frame, text=self.config_manager.get('SelectSheetLabel'))
        lbl.grid(row=row, column=0, sticky=W, pady=10)
        
        self.sheet_combobox = tb.Combobox(self.main_frame, width=42)
        self.sheet_combobox.grid(row=row, column=1, padx=5, pady=10)
        
        self.browse3_btn = tb.Button(self.main_frame,
                                   text=self.config_manager.get('BrowseButton'),
                                   bootstyle="outline-primary",
                                   state='disabled',
                                   style="Big.TButton")
        self.browse3_btn.grid(row=row, column=2, padx=5, pady=10)
        ToolTip(self.browse3_btn, self.config_manager.get('BrowseSheetTooltip'))

    def _create_compare_button(self, row: int):
        """建立比對按鈕"""
        self.compare_btn = tb.Button(self.main_frame,
                                   text=self.config_manager.get('CompareButton'),
                                   bootstyle="outline-success",
                                   style="Big.TButton")
        self.compare_btn.grid(row=row, column=0, columnspan=3, pady=25, sticky='ew')

    def set_all_font_size(self, size: int):
        """設定所有元件的字體大小"""
        font = tkfont.Font(size=size, family='Microsoft JhengHei')
        for widget in self.root.winfo_children():
            for child in widget.winfo_children():
                if isinstance(child, (tk.Label, tk.Entry, tk.Button, tb.Label, tb.Entry, tb.Button, tb.Combobox)):
                    try:
                        child.configure(font=font)
                    except Exception as e:
                        logger.warning(f"設定字體大小時發生錯誤: {str(e)}")

    def center_window(self):
        """將視窗置中"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def browse_excel1(self):
        """瀏覽錯誤碼XML檔案"""
        filename = filedialog.askopenfilename(
            title=self.config_manager.get('ErrorCodeXMLLabel'),
            filetypes=[("Excel files", "*.xlsx *.xls")],
            initialdir=self.config_manager.get('LastXMLPath')
        )
        if filename:
            self.excel1_path = filename
            self.excel1_entry.delete(0, 'end')
            self.excel1_entry.insert(0, filename)
            self.config_manager.update_last_paths(xml_path=os.path.dirname(filename))
            logger.info(f"選擇錯誤碼XML檔案: {filename}")

    def browse_excel2(self):
        """瀏覽來源Excel檔案"""
        filename = filedialog.askopenfilename(
            title=self.config_manager.get('SourceExcelLabel'),
            filetypes=[("Excel files", "*.xlsx *.xls")],
            initialdir=self.config_manager.get('LastExcelPath')
        )
        if filename:
            self.excel2_path = filename
            self.excel2_entry.delete(0, 'end')
            self.excel2_entry.insert(0, filename)
            self.config_manager.update_last_paths(excel_path=os.path.dirname(filename))
            self.browse3_btn.config(state='normal')
            logger.info(f"選擇來源Excel檔案: {filename}")
            if self.sheet_load_callback:
                self.sheet_load_callback(filename)

    def update_sheet_list(self, sheets: list):
        """更新工作表列表"""
        self.sheet_combobox['values'] = sheets
        if sheets:
            self.sheet_combobox.set(sheets[0])
            logger.info(f"更新工作表列表: {sheets}")

    def get_selected_sheet(self) -> str:
        """獲取選擇的工作表"""
        return self.sheet_combobox.get()

    def set_compare_command(self, command: Callable):
        """設定比對按鈕的命令"""
        self.compare_btn.config(command=command)

    def show_error(self, title: str, message: str):
        """顯示錯誤訊息"""
        messagebox.showerror(title, message)
        logger.error(f"{title}: {message}")

    def show_info(self, title: str, message: str):
        """顯示資訊訊息"""
        messagebox.showinfo(title, message)
        logger.info(f"{title}: {message}")

    def ask_yes_no(self, title: str, message: str) -> bool:
        """顯示是/否對話框"""
        result = messagebox.askyesno(title, message)
        logger.info(f"{title}: {message} - 使用者選擇: {result}")
        return result

    def set_sheet_load_callback(self, callback: Callable):
        """設定 sheet 載入 callback"""
        self.sheet_load_callback = callback 