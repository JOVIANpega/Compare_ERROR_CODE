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
import sys

logger = logging.getLogger(__name__)

class ToolTip:
    """工具提示類別，滑鼠移到元件上時顯示提示文字"""
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
    """UI 管理類別，負責建立和管理主視窗與所有互動元件"""
    def __init__(self, root: tk.Tk, config_manager):
        self.root = root
        self.config_manager = config_manager
        self.excel1_path: Optional[str] = None
        self.excel2_path: Optional[str] = None
        self.selected_sheet: Optional[str] = None
        self.sheet_load_callback = None  # 新增 callback 屬性
        # 取得 EXE 目錄
        self.exe_dir = self.get_exe_dir()
        # 初始化UI
        self._init_ui()
        self._setup_window()
        logger.info("UI初始化完成")

    def get_exe_dir(self):
        """取得 EXE 或 py 檔案所在目錄"""
        if getattr(sys, 'frozen', False):
            return os.path.dirname(sys.executable)
        else:
            return os.path.dirname(os.path.abspath(__file__))

    def _init_ui(self):
        """初始化UI元件與樣式"""
        self.style = tb.Style()
        self._setup_styles()
        self.main_frame = tb.Frame(self.root, padding=10, style="Main.TFrame")
        self.main_frame.pack(fill=BOTH, expand=YES)
        self._create_widgets()
        self.font_size = int(self.config_manager.get('FontSize', 11))
        self.set_all_font_size(self.font_size)

    def _setup_styles(self):
        """設定UI樣式（按鈕顏色、字型等）"""
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
        """設定主視窗屬性（標題、大小、可調整、置中）"""
        self.root.title("Error Code Comparer")
        width = int(self.config_manager.get('WindowWidth', 540))
        height = int(self.config_manager.get('WindowHeight', 340))
        self.root.geometry(f"{width}x{height}")
        self.root.resizable(True, True)
        self.center_window()
        self.root.configure(bg="#e6f2ff")
        # 綁定視窗大小調整事件，關閉時記錄大小
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.bind("<Configure>", self._on_resize)
        self._last_size = (width, height)

    def _on_resize(self, event):
        # 只在大小有變化時記錄
        if event.widget == self.root:
            size = (self.root.winfo_width(), self.root.winfo_height())
            if size != getattr(self, '_last_size', None):
                self.config_manager.update_window_size(*size)
                self._last_size = size

    def _on_close(self):
        # 關閉視窗時記錄目前大小
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        self.config_manager.update_window_size(width, height)
        self.root.destroy()

    def _create_widgets(self):
        """建立所有UI元件（檔案選擇、sheet選擇、比對按鈕）"""
        self._create_file_selection_row(0, "ErrorCodeXMLLabel", self.browse_excel1)
        self._create_file_selection_row(1, "SourceExcelLabel", self.browse_excel2)
        self._create_sheet_selection_row(2)
        self._create_compare_button(3)

    def _create_file_selection_row(self, row: int, label_key: str, browse_command: Callable):
        """建立檔案選擇行（標籤、輸入框、瀏覽按鈕）"""
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
        """瀏覽錯誤碼XML檔案，選完後記錄目錄供下次用"""
        # 先讀取設定檔的LastXMLPath，沒有就用EXE目錄
        initial_dir = self.config_manager.get('LastXMLPath') or self.get_exe_dir()
        filename = filedialog.askopenfilename(
            title=self.config_manager.get('ErrorCodeXMLLabel'),
            filetypes=[("Excel files", "*.xlsx *.xls")],
            initialdir=initial_dir
        )
        if filename:
            self.excel1_path = filename
            self.excel1_entry.delete(0, 'end')
            self.excel1_entry.insert(0, filename)
            self.config_manager.update_last_paths(xml_path=os.path.dirname(filename))
            self.last_dir = os.path.dirname(filename)  # 仍保留給excel2用
            logger.info(f"選擇錯誤碼XML檔案: {filename}")

    def browse_excel2(self):
        """瀏覽來源Excel檔案，預設用上次選的目錄"""
        # 先讀取設定檔的LastExcelPath，沒有就用EXE目錄
        initial_dir = self.config_manager.get('LastExcelPath') or self.get_exe_dir()
        filename = filedialog.askopenfilename(
            title=self.config_manager.get('SourceExcelLabel'),
            filetypes=[("Excel files", "*.xlsx *.xls")],
            initialdir=initial_dir
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
        """更新下拉選單的工作表列表"""
        self.sheet_combobox['values'] = sheets
        if sheets:
            self.sheet_combobox.set(sheets[0])
            logger.info(f"更新工作表列表: {sheets}")

    def get_selected_sheet(self) -> str:
        """取得目前選擇的工作表名稱"""
        return self.sheet_combobox.get()

    def set_compare_command(self, command: Callable):
        """設定比對按鈕的 callback"""
        self.compare_btn.config(command=command)

    def show_error(self, title: str, message: str):
        """顯示錯誤訊息視窗"""
        messagebox.showerror(title, message)
        logger.error(f"{title}: {message}")

    def show_info(self, title: str, message: str):
        """顯示資訊訊息視窗"""
        messagebox.showinfo(title, message)
        logger.info(f"{title}: {message}")

    def ask_yes_no(self, title: str, message: str) -> bool:
        """顯示是/否對話框，回傳布林值"""
        result = messagebox.askyesno(title, message)
        logger.info(f"{title}: {message} - 使用者選擇: {result}")
        return result

    def set_sheet_load_callback(self, callback: Callable):
        """設定 sheet 載入 callback"""
        self.sheet_load_callback = callback 