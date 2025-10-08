"""
主程式檔案
整合所有模組，提供程式的主要入口點
包含錯誤碼比對和錯誤碼查詢兩個功能
"""
import tkinter as tk
from tkinter import ttk
import ttkbootstrap as tb
import logging
from pathlib import Path
from config_manager import ConfigManager
from ui_manager import UIManager
from excel_handler import ExcelHandler
from guide_popup.guide import show_guide
from excel_errorcode_search_ui import ExcelErrorCodeSearchUI
from ai_recommendation_engine import AIRecommendationEngine
from ai_prompt_templates import AIPromptTemplates
from file_finder import FileFinder
import pandas as pd
import threading
import subprocess
import platform
import os

# 設定日誌
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('app.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class ErrorCodeTool:
    """主流程控制類別，整合錯誤碼比對和查詢功能"""
    def __init__(self):
        # 初始化設定管理器
        self.config_manager = ConfigManager()
        
        # 初始化主視窗
        self.root = tb.Window(themename="cosmo")
        self.root.title("錯誤碼工具集")
        # 優先讀取 setup.txt 的視窗大小
        width = int(self.config_manager.get('WindowWidth', 927))
        height = int(self.config_manager.get('WindowHeight', 425))
        self.root.geometry(f"{width}x{height}")
        
        # 初始化UI管理器（錯誤碼比對功能），傳遞字體大小
        # font_size = int(self.config_manager.get('FontSize', 10))
        self.ui_manager = UIManager(self.root, self.config_manager)
        self.ui_manager.set_search_callback(self.toggle_search_ui)
        
        # 初始化Excel處理器
        self.excel_handler = ExcelHandler()
        
        # 初始化AI推薦引擎
        self.ai_engine = AIRecommendationEngine()
        self.prompt_templates = AIPromptTemplates()
        
        # 初始化錯誤碼查詢UI
        self.search_ui = ExcelErrorCodeSearchUI(parent=self.root, offset_x=100, offset_y=80)
        self.search_ui.root.withdraw()
        self.search_ui_visible = False
        # 設定查詢視窗的關閉事件
        self.search_ui.root.protocol("WM_DELETE_WINDOW", self.hide_search_ui)
        
        # 設定比對按鈕的命令
        self.ui_manager.set_compare_command(self.compare_files)
        
        # 設定AI推薦按鈕的命令
        self.ui_manager.set_ai_recommend_callback(self.ai_recommend_analysis)
        
        # 設定開啟結果檔案按鈕的命令
        self.ui_manager.set_open_result_callback(self.open_result_files)
        
        # 綁定 sheet 載入 callback
        self.ui_manager.set_sheet_load_callback(self.load_sheets)
        
        # 先顯示主UI，再顯示導覽，避免閃爍
        self.root.deiconify()
        show_guide(self.root, 'setup.txt', "錯誤碼工具集導覽")
        
        self.root.protocol("WM_DELETE_WINDOW", self.root.destroy)
        
        logger.info("程式初始化完成")

    def toggle_search_ui(self):
        """切換查詢 UI 浮動視窗顯示/隱藏，若視窗已被關閉則重建"""
        try:
            if not self.search_ui.root.winfo_exists():
                self.search_ui = ExcelErrorCodeSearchUI(parent=self.root, offset_x=100, offset_y=80)
                self.search_ui.root.withdraw()
                self.search_ui.root.protocol("WM_DELETE_WINDOW", self.hide_search_ui)
                self.search_ui_visible = False
        except Exception:
            self.search_ui = ExcelErrorCodeSearchUI(parent=self.root, offset_x=100, offset_y=80)
            self.search_ui.root.withdraw()
            self.search_ui.root.protocol("WM_DELETE_WINDOW", self.hide_search_ui)
            self.search_ui_visible = False

        if self.search_ui_visible:
            self.search_ui.root.withdraw()
            self.search_ui_visible = False
        else:
            self.search_ui.center_window(offset_x=100, offset_y=80)
            self.search_ui.root.deiconify()
            self.search_ui.root.lift()
            self.search_ui.root.focus_force()
            self.search_ui_visible = True

    def hide_search_ui(self):
        self.search_ui.root.withdraw()
        self.search_ui_visible = False

    def compare_files(self):
        """比對檔案（在背景執行緒執行，避免UI卡住）"""
        def do_compare():
            try:
                # 更新狀態列
                self.ui_manager.update_status("正在進行檔案比對...", "orange")
                
                # 檢查必要檔案是否已選擇
                if not all([self.ui_manager.excel1_path, 
                           self.ui_manager.excel2_path, 
                           self.ui_manager.get_selected_sheet()]):
                    self.ui_manager.show_error(
                        self.config_manager.get('ErrorTitle'),
                        "請選擇所有必要的檔案和工作表"
                    )
                    return

                # 載入錯誤碼檔案
                if not self.excel_handler.load_error_codes(self.ui_manager.excel1_path):
                    self.ui_manager.show_error(
                        self.config_manager.get('ErrorTitle'),
                        "載入錯誤碼檔案失敗"
                    )
                    return

                # 載入來源工作表
                df_source = self.excel_handler.load_source_sheet(
                    self.ui_manager.excel2_path,
                    self.ui_manager.get_selected_sheet()
                )
                if df_source is None:
                    self.ui_manager.show_error(
                        self.config_manager.get('ErrorTitle'),
                        "載入來源工作表失敗"
                    )
                    return

                # 優化比對：用 merge 取代 for 迴圈
                df_error_codes = pd.read_excel(self.ui_manager.excel1_path, sheet_name="Test Item All")
                # 取 C欄(TestID)、D欄(Description)、E欄(ChineseDesc)
                df_error_codes = df_error_codes.iloc[:, [2, 3, 4]]
                df_error_codes.columns = ['TestID', 'Description', 'ChineseDesc']
                # 找來源的 Description, TestID 欄位
                desc_col = self.excel_handler.find_column(df_source, 'Description')
                testid_col = self.excel_handler.find_column(df_source, 'TestID')
                if not desc_col or not testid_col:
                    self.ui_manager.show_error(
                        self.config_manager.get('ErrorTitle'),
                        f"找不到 Description 或 TestID 欄位，實際欄位: {df_source.columns.tolist()}"
                    )
                    return
                df_result = df_source[[desc_col, testid_col]].copy()
                df_result.columns = ['你的 description', '你寫的 Error Code']
                # merge
                df_merge = pd.merge(
                    df_result,
                    df_error_codes,
                    how='left',
                    left_on='你寫的 Error Code',
                    right_on='TestID'
                )
                df_merge['Test Item 文件的 description'] = df_merge['Description'].fillna(self.config_manager.get('NotFound'))
                df_merge['Test Item 的 Error Code'] = df_merge['ChineseDesc'].fillna(self.config_manager.get('NotFoundCN'))
                df_merge = df_merge[['你的 description', '你寫的 Error Code', 'Test Item 文件的 description', 'Test Item 的 Error Code']]
                # 準備輸出路徑
                output_path = str(Path(self.ui_manager.excel2_path).with_name(
                    f"{Path(self.ui_manager.excel2_path).stem}_compare_ERRORCODE.xlsx"
                ))
                # 檢查檔案是否存在
                if Path(output_path).exists():
                    if not self.ui_manager.ask_yes_no(
                        self.config_manager.get('FileExistsTitle'),
                        self.config_manager.get('FileExistsMsg').format(output_path=output_path)
                    ):
                        self.ui_manager.show_info(
                            self.config_manager.get('CancelTitle'),
                            self.config_manager.get('CancelMsg')
                        )
                        return
                # 儲存結果（含反白）
                if self.excel_handler.save_result(
                    df_merge,
                    pd.read_excel(self.ui_manager.excel1_path, sheet_name="Test Item All"),
                    output_path,
                    self.ui_manager.get_selected_sheet()
                ):
                    self.ui_manager.update_status(f"比對完成！結果已儲存於：{os.path.basename(output_path)}", "green")
                    self.ui_manager.show_info(
                        self.config_manager.get('SuccessTitle'),
                        f"{self.config_manager.get('SuccessMsg')}\n{output_path}"
                    )
                    # 更新最後使用的輸出目錄
                    self.config_manager.update_last_paths(
                        output_dir=str(Path(output_path).parent)
                    )
                else:
                    self.ui_manager.update_status("儲存結果失敗", "red")
                    self.ui_manager.show_error(
                        self.config_manager.get('ErrorTitle'),
                        "儲存結果失敗"
                    )
            except Exception as e:
                logger.error(f"比對檔案時發生錯誤: {str(e)}")
                self.ui_manager.update_status(f"比對失敗: {str(e)[:100]}", "red")
                self.ui_manager.show_error(
                    self.config_manager.get('ErrorTitle'),
                    f"{self.config_manager.get('CompareError').format(error=str(e))}"
                )
            finally:
                self.root.config(cursor="")
        # 執行時顯示處理中游標
        self.root.config(cursor="wait")
        threading.Thread(target=do_compare, daemon=True).start()

    def load_sheets(self, excel_path):
        """載入 Excel 檔案的所有 sheet 名稱並更新 UI"""
        try:
            sheets = self.excel_handler.get_sheet_names(excel_path)
            if sheets:
                self.ui_manager.update_sheet_list(sheets)
                logger.info(f"成功載入 {len(sheets)} 個工作表")
            else:
                logger.warning(f"無法從 {excel_path} 載入工作表")
                self.ui_manager.show_error(
                    "載入工作表失敗",
                    f"無法從檔案載入工作表：\n{os.path.basename(excel_path)}"
                )
        except Exception as e:
            logger.error(f"載入工作表時發生錯誤: {str(e)}")
            self.ui_manager.show_error(
                "載入工作表失敗",
                f"載入工作表時發生錯誤：{str(e)}"
            )

    def ai_recommend_analysis(self):
        """AI 推薦分析功能 - 自動檢查並執行比對"""
        def do_ai_analysis():
            try:
                # 更新狀態列
                self.ui_manager.update_status("正在進行 AI 推薦分析，請稍候...", "orange")
                
                # 檢查檔案設定
                if not self.ui_manager.excel1_path or not os.path.exists(self.ui_manager.excel1_path):
                    self.ui_manager.show_error("檔案錯誤", "請先選擇 Error Code 參考檔案")
                    return
                
                if not self.ui_manager.excel2_path or not os.path.exists(self.ui_manager.excel2_path):
                    self.ui_manager.show_error("檔案錯誤", "請先選擇要分析的 Excel 檔案")
                    return
                
                # 檢查是否已有比對結果檔案
                output_file = self._get_expected_output_file()
                
                if not os.path.exists(output_file):
                    # 如果沒有比對結果檔案，先執行比對
                    self.ui_manager.update_status("未找到比對結果檔案，正在自動執行比對...", "orange")
                    
                    # 執行比對功能
                    success = self._perform_comparison()
                    if not success:
                        self.ui_manager.update_status("自動比對失敗，無法進行 AI 推薦分析", "red")
                        self.ui_manager.show_error("比對失敗", "自動比對失敗，無法進行 AI 推薦分析")
                        return
                
                # 現在進行 AI 推薦分析
                self._perform_ai_recommendation(output_file)
                
            except Exception as e:
                logger.error(f"AI 推薦分析時發生錯誤: {str(e)}")
                self.ui_manager.update_status(f"AI 推薦分析失敗: {str(e)[:100]}", "red")
                self.ui_manager.show_error("分析失敗", f"AI 推薦分析時發生錯誤：{str(e)}")
            finally:
                self.root.config(cursor="")
        
        self.root.config(cursor="wait")
        threading.Thread(target=do_ai_analysis, daemon=True).start()

    def _get_expected_output_file(self):
        """取得預期的輸出檔案路徑"""
        base_name = os.path.splitext(os.path.basename(self.ui_manager.excel2_path))[0]
        return os.path.join("EXCEL", f"{base_name}_compare_ERRORCODE.xlsx")

    def _perform_comparison(self):
        """執行比對功能"""
        try:
            # 使用現有的比對邏輯
            return self.compare_files()
        except Exception as e:
            logger.error(f"執行比對時發生錯誤: {str(e)}")
            return False

    def _perform_ai_recommendation(self, output_file):
        """執行 AI 推薦分析"""
        try:
            # 載入參考資料
            if not self.ai_engine.load_reference_data(self.ui_manager.excel1_path):
                self.ui_manager.show_error("載入失敗", "無法載入 Error Code 參考資料")
                return
            
            # 讀取比對結果檔案
            # 先檢查檔案有哪些工作表
            try:
                excel_file = pd.ExcelFile(output_file)
                available_sheets = excel_file.sheet_names
                logger.info(f"比對結果檔案的工作表: {available_sheets}")
                
                # 嘗試使用原始工作表名稱，如果沒有則使用第一個工作表
                if self.ui_manager.selected_sheet in available_sheets:
                    sheet_name = self.ui_manager.selected_sheet
                else:
                    sheet_name = available_sheets[0] if available_sheets else None
                
                if not sheet_name:
                    self.ui_manager.show_error("檔案錯誤", "比對結果檔案沒有可用的工作表")
                    return
                
                df_result = pd.read_excel(output_file, sheet_name=sheet_name)
                logger.info(f"使用工作表: {sheet_name}")
                
            except Exception as e:
                logger.error(f"讀取比對結果檔案時發生錯誤: {str(e)}")
                self.ui_manager.show_error("讀取失敗", f"讀取比對結果檔案時發生錯誤：{str(e)}")
                return
            
            if 'Description' not in df_result.columns:
                # 檢查可能的欄位名稱變體
                possible_desc_columns = [
                    'Description', 'description', 'DESCRIPTION', 'O', 'O欄位',
                    '你的 description', 'Description', '描述', 'desc'
                ]
                desc_column = None
                
                for col in possible_desc_columns:
                    if col in df_result.columns:
                        desc_column = col
                        break
                
                if not desc_column:
                    # 顯示實際的欄位名稱
                    actual_columns = list(df_result.columns)
                    self.ui_manager.show_error("欄位錯誤", 
                        f"比對結果檔案沒有 Description 欄位\n"
                        f"實際欄位: {actual_columns}\n"
                        f"請檢查比對結果檔案是否正確")
                    return
                
                # 重新命名欄位為 Description
                df_result = df_result.rename(columns={desc_column: 'Description'})
                logger.info(f"找到描述欄位: {desc_column}，已重新命名為 Description")
            
            # 提取描述並生成推薦
            descriptions = df_result['Description'].fillna('').astype(str).tolist()
            recommendations = self.ai_engine.generate_recommendations_with_search(descriptions)
            
            # 更新檔案
            self._update_file_with_recommendations(output_file, df_result, recommendations)
            
            self.ui_manager.update_status(f"AI 推薦分析完成！結果已更新到：{os.path.basename(output_file)}", "green")
            self.ui_manager.show_info("分析完成", f"AI 推薦分析完成！\n結果已更新到：{os.path.basename(output_file)}")
            
        except Exception as e:
            logger.error(f"執行 AI 推薦時發生錯誤: {str(e)}")
            self.ui_manager.show_error("推薦失敗", f"執行 AI 推薦時發生錯誤：{str(e)}")

    def _update_file_with_recommendations(self, file_path, df_result, recommendations):
        """更新檔案並添加 AI 推薦"""
        try:
            # 使用 excel_handler 的現有功能
            self.excel_handler.add_ai_recommendations_to_existing_file(file_path, recommendations)
        except Exception as e:
            logger.error(f"更新檔案時發生錯誤: {str(e)}")
            raise

    def _show_ai_prompt(self, descriptions):
        """顯示 AI PROMPT 供使用者手動處理"""
        try:
            # 生成 PROMPT
            prompt = self.ai_engine.get_prompt_for_descriptions(descriptions, "excel")
            
            # 創建新視窗顯示 PROMPT
            prompt_window = tk.Toplevel(self.root)
            prompt_window.title("AI 推薦 PROMPT")
            prompt_window.geometry("800x600")
            prompt_window.transient(self.root)
            prompt_window.grab_set()
            
            # 創建文字區域
            text_frame = tk.Frame(prompt_window)
            text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            # 標題
            title_label = tk.Label(text_frame, text="請將以下 PROMPT 複製到 AI 工具中：", 
                                 font=("Microsoft JhengHei", 12, "bold"))
            title_label.pack(pady=(0, 10))
            
            # 文字區域
            text_widget = tk.Text(text_frame, wrap=tk.WORD, font=("Consolas", 10))
            text_widget.pack(fill=tk.BOTH, expand=True)
            text_widget.insert(tk.END, prompt)
            text_widget.config(state=tk.DISABLED)
            
            # 滾動條
            scrollbar = tk.Scrollbar(text_widget)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            text_widget.config(yscrollcommand=scrollbar.set)
            scrollbar.config(command=text_widget.yview)
            
            # 按鈕框架
            button_frame = tk.Frame(prompt_window)
            button_frame.pack(pady=10)
            
            # 複製按鈕
            copy_btn = tk.Button(button_frame, text="複製 PROMPT", 
                               command=lambda: self._copy_to_clipboard(prompt))
            copy_btn.pack(side=tk.LEFT, padx=5)
            
            # 關閉按鈕
            close_btn = tk.Button(button_frame, text="關閉", 
                                command=prompt_window.destroy)
            close_btn.pack(side=tk.LEFT, padx=5)
            
        except Exception as e:
            logger.error(f"顯示 AI PROMPT 時發生錯誤: {str(e)}")
            self.ui_manager.show_error(
                self.config_manager.get('ErrorTitle'),
                f"顯示 AI PROMPT 失敗: {str(e)}"
            )

    def _copy_to_clipboard(self, text):
        """複製文字到剪貼簿"""
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            self.ui_manager.show_info("成功", "PROMPT 已複製到剪貼簿")
        except Exception as e:
            logger.error(f"複製到剪貼簿時發生錯誤: {str(e)}")
            self.ui_manager.show_error("錯誤", "複製到剪貼簿失敗")

    def open_result_files(self):
        """開啟 compare_ERRORCODE.xlsx 檔案"""
        try:
            # 搜尋 compare 檔案
            compare_files = FileFinder.find_compare_files()
            
            if not compare_files:
                self.ui_manager.show_error(
                    "未找到結果檔案",
                    "沒有找到任何 compare_ERRORCODE.xlsx 檔案\n請先執行比對功能"
                )
                return
            
            # 顯示找到的檔案資訊
            logger.info(f"找到 {len(compare_files)} 個 compare 檔案:")
            for i, file_path in enumerate(compare_files, 1):
                logger.info(f"  {i}. {file_path}")
            
            if len(compare_files) == 1:
                # 只有一個檔案，直接開啟
                file_path = compare_files[0]
                self._open_file(file_path)
                self.ui_manager.show_info(
                    "檔案已開啟",
                    f"已開啟結果檔案：\n{os.path.basename(file_path)}\n\n檔案路徑：\n{file_path}"
                )
            else:
                # 多個檔案，讓使用者選擇
                self._show_file_selection_dialog(compare_files)
                
        except Exception as e:
            logger.error(f"開啟結果檔案時發生錯誤: {str(e)}")
            self.ui_manager.show_error(
                "開啟檔案失敗",
                f"開啟結果檔案時發生錯誤：{str(e)}"
            )

    def _show_file_selection_dialog(self, files):
        """顯示檔案選擇對話框"""
        try:
            # 創建選擇視窗
            selection_window = tk.Toplevel(self.root)
            selection_window.title("選擇要開啟的結果檔案")
            selection_window.geometry("600x400")
            selection_window.transient(self.root)
            selection_window.grab_set()
            
            # 標題
            title_label = tk.Label(selection_window, text="找到多個結果檔案，請選擇要開啟的檔案：", 
                                 font=("Microsoft JhengHei", 12, "bold"))
            title_label.pack(pady=10)
            
            # 創建列表框
            list_frame = tk.Frame(selection_window)
            list_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
            
            # 列表框和滾動條
            listbox = tk.Listbox(list_frame, font=("Consolas", 10))
            scrollbar = tk.Scrollbar(list_frame, orient=tk.VERTICAL, command=listbox.yview)
            listbox.configure(yscrollcommand=scrollbar.set)
            
            # 添加檔案到列表框
            formatted_files = FileFinder.format_file_list(files)
            for i, file_info in enumerate(formatted_files):
                listbox.insert(tk.END, f"{i+1}. {file_info}")
            
            # 預設選擇第一個
            listbox.selection_set(0)
            
            # 布局
            listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            # 按鈕框架
            button_frame = tk.Frame(selection_window)
            button_frame.pack(pady=10)
            
            # 開啟按鈕
            def open_selected():
                selection = listbox.curselection()
                if selection:
                    index = selection[0]
                    file_path = files[index]
                    self._open_file(file_path)
                    selection_window.destroy()
                    self.ui_manager.show_info(
                        "檔案已開啟",
                        f"已開啟結果檔案：\n{os.path.basename(file_path)}"
                    )
                else:
                    self.ui_manager.show_error("請選擇檔案", "請先選擇要開啟的檔案")
            
            open_btn = tk.Button(button_frame, text="開啟", command=open_selected)
            open_btn.pack(side=tk.LEFT, padx=5)
            
            # 取消按鈕
            cancel_btn = tk.Button(button_frame, text="取消", command=selection_window.destroy)
            cancel_btn.pack(side=tk.LEFT, padx=5)
            
            # 雙擊開啟
            def on_double_click(event):
                open_selected()
            
            listbox.bind('<Double-1>', on_double_click)
            
        except Exception as e:
            logger.error(f"顯示檔案選擇對話框時發生錯誤: {str(e)}")
            self.ui_manager.show_error(
                "顯示選擇對話框失敗",
                f"顯示檔案選擇對話框時發生錯誤：{str(e)}"
            )

    def _open_file(self, file_path):
        """開啟檔案"""
        try:
            system = platform.system()
            
            if system == "Windows":
                # Windows 系統
                os.startfile(file_path)
            elif system == "Darwin":
                # macOS 系統
                subprocess.run(["open", file_path])
            else:
                # Linux 系統
                subprocess.run(["xdg-open", file_path])
                
            logger.info(f"成功開啟檔案: {file_path}")
            
        except Exception as e:
            logger.error(f"開啟檔案失敗: {str(e)}")
            # 如果系統開啟失敗，嘗試用預設程式開啟
            try:
                subprocess.run(["start", file_path], shell=True, check=True)
            except:
                self.ui_manager.show_error(
                    "開啟檔案失敗",
                    f"無法開啟檔案：{file_path}\n請手動開啟檔案"
                )

    def run(self):
        """啟動主視窗事件迴圈"""
        try:
            # self.show_compare_ui()  # 已無此方法，移除
            self.root.mainloop()
        except Exception as e:
            logger.error(f"程式執行時發生錯誤: {str(e)}")
            raise

if __name__ == "__main__":
    # 程式進入點
    app = ErrorCodeTool()
    app.run() 