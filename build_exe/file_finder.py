"""
檔案搜尋工具
用於搜尋特定格式的檔案
"""
import os
import glob
from pathlib import Path
from typing import List, Optional
import logging

logger = logging.getLogger(__name__)

class FileFinder:
    """檔案搜尋類別"""
    
    @staticmethod
    def find_compare_files(search_dirs: List[str] = None) -> List[str]:
        """
        搜尋所有 compare_ERRORCODE.xlsx 檔案
        
        Args:
            search_dirs: 搜尋目錄列表，如果為 None 則搜尋當前目錄和子目錄
            
        Returns:
            List[str]: 找到的檔案路徑列表
        """
        compare_files = []
        
        if search_dirs is None:
            # 預設搜尋當前目錄和常見的輸出目錄
            search_dirs = [
                ".",  # 當前目錄
                "EXCEL",  # EXCEL 目錄
                "dist",  # dist 目錄
                "output",  # output 目錄
            ]
        
        for search_dir in search_dirs:
            if not os.path.exists(search_dir):
                continue
                
            # 搜尋模式：包含 compare_ERRORCODE.xlsx 的檔案
            pattern = os.path.join(search_dir, "**", "*compare_ERRORCODE.xlsx")
            files = glob.glob(pattern, recursive=True)
            
            # 將相對路徑轉換為絕對路徑，避免重複
            abs_files = [os.path.abspath(f) for f in files]
            compare_files.extend(abs_files)
            
            logger.info(f"在 {search_dir} 中找到 {len(files)} 個 compare 檔案")
        
        # 去重並排序（使用絕對路徑去重）
        compare_files = sorted(list(set(compare_files)))
        logger.info(f"總共找到 {len(compare_files)} 個 compare 檔案")
        
        return compare_files
    
    @staticmethod
    def get_file_info(file_path: str) -> dict:
        """
        獲取檔案資訊
        
        Args:
            file_path: 檔案路徑
            
        Returns:
            dict: 檔案資訊
        """
        try:
            path = Path(file_path)
            stat = path.stat()
            
            return {
                "name": path.name,
                "path": str(path.absolute()),
                "size": stat.st_size,
                "modified": stat.st_mtime,
                "exists": path.exists()
            }
        except Exception as e:
            logger.error(f"獲取檔案資訊失敗: {e}")
            return {
                "name": os.path.basename(file_path),
                "path": file_path,
                "size": 0,
                "modified": 0,
                "exists": False
            }
    
    @staticmethod
    def find_latest_compare_file(search_dirs: List[str] = None) -> Optional[str]:
        """
        找到最新的 compare_ERRORCODE.xlsx 檔案
        
        Args:
            search_dirs: 搜尋目錄列表
            
        Returns:
            Optional[str]: 最新檔案的路徑，如果沒找到則返回 None
        """
        files = FileFinder.find_compare_files(search_dirs)
        
        if not files:
            return None
        
        # 按修改時間排序，返回最新的
        files_with_time = []
        for file_path in files:
            info = FileFinder.get_file_info(file_path)
            if info["exists"]:
                files_with_time.append((file_path, info["modified"]))
        
        if not files_with_time:
            return None
        
        # 按修改時間降序排序
        files_with_time.sort(key=lambda x: x[1], reverse=True)
        return files_with_time[0][0]
    
    @staticmethod
    def format_file_list(files: List[str]) -> List[str]:
        """
        格式化檔案列表用於顯示
        
        Args:
            files: 檔案路徑列表
            
        Returns:
            List[str]: 格式化後的檔案資訊列表
        """
        formatted_files = []
        
        for file_path in files:
            info = FileFinder.get_file_info(file_path)
            if info["exists"]:
                # 格式化檔案大小
                size_mb = info["size"] / (1024 * 1024)
                size_str = f"{size_mb:.1f}MB" if size_mb >= 1 else f"{info['size']}B"
                
                # 格式化修改時間
                import time
                mod_time = time.strftime("%Y-%m-%d %H:%M", time.localtime(info["modified"]))
                
                formatted_info = f"{info['name']} ({size_str}, {mod_time})"
                formatted_files.append(formatted_info)
            else:
                formatted_files.append(f"{info['name']} (檔案不存在)")
        
        return formatted_files
