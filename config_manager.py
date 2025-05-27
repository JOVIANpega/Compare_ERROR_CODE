"""
設定管理模組
負責處理所有設定相關的操作，包括讀取、寫入和更新設定
"""
import sys
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
    """設定管理類別，負責讀取、寫入、更新 setup.txt，保留註解與分隔線"""
    def __init__(self, setup_file: str = 'setup.txt'):
        # 自動偵測 EXE 或 py 路徑，確保 setup.txt 路徑正確
        if hasattr(sys, '_MEIPASS'):
            base_dir = sys._MEIPASS
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))
        self.setup_file = os.path.join(base_dir, setup_file)
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
        self.config = {}
        self.lines = []  # 保留原始所有行
        self._load_config_and_lines()

    def _load_config_and_lines(self):
        """同時載入設定檔內容與所有原始行"""
        self.config = {}
        self.lines = []
        if not os.path.exists(self.setup_file):
            self._create_default_config()
        try:
            with open(self.setup_file, 'r', encoding='utf-8') as f:
                for line in f:
                    self.lines.append(line.rstrip('\n'))
                    if '=' in line and not line.strip().startswith('#'):
                        k, v = line.strip().split('=', 1)
                        self.config[k] = v
            # 確保所有預設值都存在
            for k, v in self.default_text.items():
                if k not in self.config:
                    self.config[k] = v
        except Exception as e:
            logger.error(f"載入設定檔時發生錯誤: {str(e)}")
            self.config = self.default_text.copy()

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
        """儲存設定到檔案，保留註解、分隔線、空行"""
        # 先建立 key->value 的最新對照
        new_config = config.copy()
        written_keys = set()
        new_lines = []
        for line in self.lines:
            if '=' in line and not line.strip().startswith('#'):
                k, _ = line.strip().split('=', 1)
                if k in new_config:
                    new_lines.append(f'{k}={new_config[k]}')
                    written_keys.add(k)
                else:
                    new_lines.append(line)
            else:
                new_lines.append(line)
        # 新增還沒出現過的 key
        for k, v in new_config.items():
            if k not in written_keys:
                new_lines.append(f'{k}={v}')
        try:
            with open(self.setup_file, 'w', encoding='utf-8') as f:
                for line in new_lines:
                    f.write(line + '\n')
            self.lines = new_lines
            self.config = new_config
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