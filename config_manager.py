"""
設定管理模組
負責處理所有設定相關的操作，包括讀取、寫入和更新設定
"""
import os
import logging
from typing import Dict, Any

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

class ConfigManager:
    """設定管理類別，負責讀取、寫入、更新 setup.txt"""
    def __init__(self, setup_file: str = 'setup.txt'):
        self.setup_file = setup_file
        # 預設設定內容
        self.default_text = {
            'ErrorCodeXMLLabel': '錯誤碼 XML 檔案：',
            'SourceExcelLabel': '來源 Excel 檔案：',
            'SelectSheetLabel': '選擇工作表：',
            'BrowseButton': '瀏覽',
            'CompareButton': '比對',
            'SuccessTitle': '成功',
            'SuccessMsg': '比對完成，結果已儲存於：',
            'CancelTitle': '取消',
            'CancelMsg': '已取消操作。',
            'FileExistsTitle': '檔案已存在',
            'FileExistsMsg': '檔案 {output_path} 已存在，是否要覆蓋？',
            'ErrorTitle': '錯誤',
            'SheetLoadError': '載入工作表時發生錯誤：{error}',
            'CompareError': '比對時發生錯誤：{error}',
            'NotFound': '查無說明',
            'NotFoundCN': '查無中文說明',
            'WindowWidth': '540',
            'WindowHeight': '340',
            'FontSize': '11',
            'BrowseXMLTooltip': '選擇錯誤碼XML檔案',
            'BrowseExcelTooltip': '選擇來源Excel檔案',
            'BrowseSheetTooltip': '選擇工作表',
            'LastExcelPath': '',
            'LastXMLPath': '',
            'LastOutputDir': '',
        }
        self.config = self.load_config()

    def load_config(self) -> Dict[str, Any]:
        """載入設定檔，若不存在則建立預設設定"""
        try:
            if not os.path.exists(self.setup_file):
                logger.info(f"設定檔不存在，創建新的設定檔: {self.setup_file}")
                self._create_default_config()
                return self.default_text.copy()
            config = {}
            with open(self.setup_file, 'r', encoding='utf-8') as f:
                for line in f:
                    if '=' in line:
                        k, v = line.strip().split('=', 1)
                        config[k] = v
            # 確保所有預設值都存在
            for k, v in self.default_text.items():
                if k not in config:
                    config[k] = v
            logger.info("成功載入設定檔")
            return config
        except Exception as e:
            logger.error(f"載入設定檔時發生錯誤: {str(e)}")
            return self.default_text.copy()

    def _create_default_config(self):
        """創建預設設定檔"""
        try:
            with open(self.setup_file, 'w', encoding='utf-8') as f:
                for k, v in self.default_text.items():
                    f.write(f'{k}={v}\n')
            logger.info("成功創建預設設定檔")
        except Exception as e:
            logger.error(f"創建預設設定檔時發生錯誤: {str(e)}")

    def save_config(self, config: Dict[str, Any]):
        """儲存設定到檔案"""
        try:
            with open(self.setup_file, 'w', encoding='utf-8') as f:
                for k, v in config.items():
                    f.write(f'{k}={v}\n')
            self.config = config
            logger.info("成功儲存設定")
        except Exception as e:
            logger.error(f"儲存設定時發生錯誤: {str(e)}")

    def get(self, key: str, default: Any = None) -> Any:
        """取得指定設定值"""
        return self.config.get(key, default)

    def set(self, key: str, value: Any):
        """設定指定設定值並儲存"""
        self.config[key] = value
        self.save_config(self.config)

    def update_window_size(self, width: int, height: int):
        """更新視窗大小設定"""
        self.set('WindowWidth', str(width))
        self.set('WindowHeight', str(height))

    def update_last_paths(self, excel_path: str = None, xml_path: str = None, output_dir: str = None):
        """更新最後使用的路徑設定"""
        if excel_path:
            self.set('LastExcelPath', excel_path)
        if xml_path:
            self.set('LastXMLPath', xml_path)
        if output_dir:
            self.set('LastOutputDir', output_dir) 