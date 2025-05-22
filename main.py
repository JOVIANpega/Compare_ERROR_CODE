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
        """比對檔案"""
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

            # 比對資料
            df_result = self.excel_handler.compare_data(
                df_source,
                self.config_manager.get('NotFound'),
                self.config_manager.get('NotFoundCN')
            )
            if df_result is None:
                self.ui_manager.show_error(
                    self.config_manager.get('ErrorTitle'),
                    "比對資料失敗"
                )
                return

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

            # 儲存結果
            if self.excel_handler.save_result(
                df_result,
                pd.read_excel(self.ui_manager.excel1_path, sheet_name="Test Item All"),
                output_path,
                self.ui_manager.get_selected_sheet()
            ):
                self.ui_manager.show_info(
                    self.config_manager.get('SuccessTitle'),
                    f"{self.config_manager.get('SuccessMsg')} {output_path}"
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

    def load_sheets(self, excel_path):
        """載入 Excel 檔案的所有 sheet 名稱並更新 UI"""
        sheets = self.excel_handler.get_sheet_names(excel_path)
        self.ui_manager.update_sheet_list(sheets)

    def run(self):
        """執行程式"""
        try:
            self.root.mainloop()
        except Exception as e:
            logger.error(f"程式執行時發生錯誤: {str(e)}")
            raise

if __name__ == "__main__":
    app = ErrorCodeComparer()
    app.run() 