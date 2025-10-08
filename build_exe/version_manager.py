#!/usr/bin/env python3
"""
版本管理模組
用於管理 Error Code Comparer 的版本號和相關資訊
"""

import os
import json
from pathlib import Path
from datetime import datetime
from typing import Dict, Any

class VersionManager:
    """版本管理類別"""
    
    def __init__(self, version_file: str = "version.json"):
        """
        初始化版本管理器
        
        Args:
            version_file: 版本檔案路徑
        """
        self.version_file = version_file
        self.version_data = self._load_version_data()
    
    def _load_version_data(self) -> Dict[str, Any]:
        """載入版本資料"""
        if os.path.exists(self.version_file):
            try:
                with open(self.version_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                print(f"載入版本檔案失敗: {e}")
                return self._get_default_version()
        else:
            return self._get_default_version()
    
    def _get_default_version(self) -> Dict[str, Any]:
        """取得預設版本資料"""
        return {
            "version": "1.5.0",
            "build_number": 1,
            "release_date": datetime.now().strftime("%Y-%m-%d"),
            "last_modified": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "changelog": [
                {
                    "version": "1.5.0",
                    "date": datetime.now().strftime("%Y-%m-%d"),
                    "changes": [
                        "初始版本 1.5.0",
                        "新增 AI 推薦分析功能",
                        "新增自動檔案選擇功能",
                        "新增智能工作表過濾功能",
                        "新增快速開啟結果檔案功能"
                    ]
                }
            ],
            "features": {
                "ai_recommendation": True,
                "auto_file_selection": True,
                "smart_sheet_filtering": True,
                "quick_open_result": True,
                "excel_comparison": True,
                "error_code_search": True
            },
            "author": "JOVIANpega",
            "repository": "https://github.com/JOVIANpega/Compare_ERROR_CODE"
        }
    
    def get_version(self) -> str:
        """取得當前版本號"""
        return self.version_data.get("version", "1.5.0")
    
    def get_build_number(self) -> int:
        """取得建置編號"""
        return self.version_data.get("build_number", 1)
    
    def get_version_info(self) -> Dict[str, Any]:
        """取得完整版本資訊"""
        return self.version_data.copy()
    
    def update_version(self, new_version: str, changes: list = None) -> bool:
        """
        更新版本號
        
        Args:
            new_version: 新版本號 (格式: x.y.z)
            changes: 變更列表
            
        Returns:
            bool: 是否成功更新
        """
        try:
            # 驗證版本號格式
            if not self._validate_version_format(new_version):
                print(f"無效的版本號格式: {new_version}")
                return False
            
            # 更新版本資料
            old_version = self.version_data["version"]
            self.version_data["version"] = new_version
            self.version_data["build_number"] += 1
            self.version_data["last_modified"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            # 新增變更記錄
            if changes:
                changelog_entry = {
                    "version": new_version,
                    "date": datetime.now().strftime("%Y-%m-%d"),
                    "changes": changes
                }
                self.version_data["changelog"].insert(0, changelog_entry)
            
            # 儲存版本檔案
            self._save_version_data()
            
            print(f"版本已從 {old_version} 更新到 {new_version}")
            print(f"建置編號: {self.version_data['build_number']}")
            return True
            
        except Exception as e:
            print(f"更新版本失敗: {e}")
            return False
    
    def add_changelog_entry(self, version: str, changes: list) -> bool:
        """
        新增變更記錄
        
        Args:
            version: 版本號
            changes: 變更列表
            
        Returns:
            bool: 是否成功新增
        """
        try:
            changelog_entry = {
                "version": version,
                "date": datetime.now().strftime("%Y-%m-%d"),
                "changes": changes
            }
            self.version_data["changelog"].insert(0, changelog_entry)
            self.version_data["last_modified"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            self._save_version_data()
            print(f"已為版本 {version} 新增變更記錄")
            return True
            
        except Exception as e:
            print(f"新增變更記錄失敗: {e}")
            return False
    
    def update_feature_status(self, feature: str, enabled: bool) -> bool:
        """
        更新功能狀態
        
        Args:
            feature: 功能名稱
            enabled: 是否啟用
            
        Returns:
            bool: 是否成功更新
        """
        try:
            if "features" not in self.version_data:
                self.version_data["features"] = {}
            
            self.version_data["features"][feature] = enabled
            self.version_data["last_modified"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            self._save_version_data()
            print(f"功能 '{feature}' 已{'啟用' if enabled else '停用'}")
            return True
            
        except Exception as e:
            print(f"更新功能狀態失敗: {e}")
            return False
    
    def _validate_version_format(self, version: str) -> bool:
        """驗證版本號格式"""
        try:
            parts = version.split('.')
            if len(parts) != 3:
                return False
            
            for part in parts:
                int(part)  # 檢查是否為數字
            
            return True
        except ValueError:
            return False
    
    def _save_version_data(self) -> bool:
        """儲存版本資料"""
        try:
            with open(self.version_file, 'w', encoding='utf-8') as f:
                json.dump(self.version_data, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            print(f"儲存版本檔案失敗: {e}")
            return False
    
    def print_version_info(self):
        """列印版本資訊"""
        print("=" * 50)
        print("Error Code Comparer 版本資訊")
        print("=" * 50)
        print(f"版本號: {self.get_version()}")
        print(f"建置編號: {self.get_build_number()}")
        print(f"發布日期: {self.version_data.get('release_date', 'N/A')}")
        print(f"最後修改: {self.version_data.get('last_modified', 'N/A')}")
        print(f"作者: {self.version_data.get('author', 'N/A')}")
        print(f"倉庫: {self.version_data.get('repository', 'N/A')}")
        
        print("\n功能狀態:")
        features = self.version_data.get('features', {})
        for feature, enabled in features.items():
            status = "✓ 啟用" if enabled else "✗ 停用"
            print(f"  {feature}: {status}")
        
        print("\n變更記錄:")
        changelog = self.version_data.get('changelog', [])
        for entry in changelog[:5]:  # 只顯示最近5個版本
            print(f"  {entry['version']} ({entry['date']}):")
            for change in entry['changes']:
                print(f"    - {change}")
        print("=" * 50)
    
    def generate_version_header(self) -> str:
        """生成版本標頭字串"""
        version = self.get_version()
        build = self.get_build_number()
        date = self.version_data.get('last_modified', 'N/A')
        
        return f"""# Error Code Comparer v{version} (Build {build})
# 最後更新: {date}
# 作者: {self.version_data.get('author', 'JOVIANpega')}
# 倉庫: {self.version_data.get('repository', 'https://github.com/JOVIANpega/Compare_ERROR_CODE')}
"""

def main():
    """主程式 - 版本管理工具"""
    vm = VersionManager()
    
    print("Error Code Comparer 版本管理工具")
    print("=" * 40)
    
    while True:
        print("\n請選擇操作:")
        print("1. 查看當前版本資訊")
        print("2. 更新版本號")
        print("3. 新增變更記錄")
        print("4. 更新功能狀態")
        print("5. 生成版本標頭")
        print("0. 退出")
        
        choice = input("\n請輸入選項 (0-5): ").strip()
        
        if choice == "0":
            print("再見！")
            break
        elif choice == "1":
            vm.print_version_info()
        elif choice == "2":
            new_version = input("請輸入新版本號 (格式: x.y.z): ").strip()
            changes_input = input("請輸入變更說明 (多行，空行結束): ").strip()
            changes = []
            while changes_input:
                changes.append(changes_input)
                changes_input = input().strip()
            
            vm.update_version(new_version, changes if changes else None)
        elif choice == "3":
            version = input("請輸入版本號: ").strip()
            changes_input = input("請輸入變更說明 (多行，空行結束): ").strip()
            changes = []
            while changes_input:
                changes.append(changes_input)
                changes_input = input().strip()
            
            vm.add_changelog_entry(version, changes)
        elif choice == "4":
            print("可用功能:")
            features = vm.version_data.get('features', {})
            for i, feature in enumerate(features.keys(), 1):
                status = "啟用" if features[feature] else "停用"
                print(f"  {i}. {feature} ({status})")
            
            feature_name = input("請輸入功能名稱: ").strip()
            enabled_input = input("是否啟用? (y/n): ").strip().lower()
            enabled = enabled_input in ['y', 'yes', '是', '1', 'true']
            
            vm.update_feature_status(feature_name, enabled)
        elif choice == "5":
            header = vm.generate_version_header()
            print("\n版本標頭:")
            print(header)
            
            save = input("是否儲存到檔案? (y/n): ").strip().lower()
            if save in ['y', 'yes', '是', '1', 'true']:
                filename = input("請輸入檔案名稱 (預設: version_header.txt): ").strip()
                if not filename:
                    filename = "version_header.txt"
                
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(header)
                print(f"版本標頭已儲存到 {filename}")
        else:
            print("無效的選項，請重新選擇")

if __name__ == "__main__":
    main()
