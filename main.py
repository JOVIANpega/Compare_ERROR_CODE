"""
主程式檔案
整合所有模組，提供程式的主要入口點
"""
import tkinter as tk
import ttkbootstrap as tb
import logging
from pathlib import Path
from config_manager import ConfigManager
from ui_manager import UIManager
from excel_handler import ExcelHandler
from guide_popup.guide import show_guide
import pandas as pd
import threading

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

class ErrorCodeComparer:
    """主流程控制類別，負責整合UI、Excel處理、設定管理等"""
    def __init__(self):
        # 初始化設定管理器
        self.config_manager = ConfigManager()
        
        # 初始化主視窗
        self.root = tb.Window(themename="cosmo")
        
        # 初始化UI管理器
        self.ui_manager = UIManager(self.root, self.config_manager)
        
        # 初始化Excel處理器
        self.excel_handler = ExcelHandler()
        
        # 設定比對按鈕的命令
        self.ui_manager.set_compare_command(self.compare_files)
        
        # 綁定 sheet 載入 callback
        self.ui_manager.set_sheet_load_callback(self.load_sheets)
        
        # 顯示導覽
        show_guide(self.root, 'setup.txt', "錯誤碼比對工具導覽")
        
        logger.info("程式初始化完成")

    def compare_files(self):
        """比對檔案（在背景執行緒執行，避免UI卡住）"""
        def do_compare():
            try:
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
                df_result.columns = ['AB', 'AC']
                # merge
                df_merge = pd.merge(
                    df_result,
                    df_error_codes,
                    how='left',
                    left_on='AC',
                    right_on='TestID'
                )
                df_merge['CD'] = df_merge['Description'].fillna(self.config_manager.get('NotFound'))
                df_merge['CE'] = df_merge['ChineseDesc'].fillna(self.config_manager.get('NotFoundCN'))
                df_merge = df_merge[['AB', 'AC', 'CD', 'CE']]
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
                    self.ui_manager.show_info(
                        self.config_manager.get('SuccessTitle'),
                        f"{self.config_manager.get('SuccessMsg')}\n{output_path}"
                    )
                    # 更新最後使用的輸出目錄
                    self.config_manager.update_last_paths(
                        output_dir=str(Path(output_path).parent)
                    )
                else:
                    self.ui_manager.show_error(
                        self.config_manager.get('ErrorTitle'),
                        "儲存結果失敗"
                    )
            except Exception as e:
                logger.error(f"比對檔案時發生錯誤: {str(e)}")
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
        sheets = self.excel_handler.get_sheet_names(excel_path)
        self.ui_manager.update_sheet_list(sheets)

    def run(self):
        """啟動主視窗事件迴圈"""
        try:
            self.root.mainloop()
        except Exception as e:
            logger.error(f"程式執行時發生錯誤: {str(e)}")
            raise

if __name__ == "__main__":
    # 程式進入點
    app = ErrorCodeComparer()
    app.run() 