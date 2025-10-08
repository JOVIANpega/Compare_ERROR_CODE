#!/usr/bin/env python3
"""
簡單版本管理工具
直接修改 VERSION.py 檔案來更新版本
"""

import os
import re
from datetime import datetime

def get_current_version():
    """取得當前版本"""
    try:
        from VERSION import VERSION, BUILD
        return VERSION, BUILD
    except:
        return "1.5.0", 1

def set_version(new_version, new_build=None):
    """設定新版本"""
    current_version, current_build = get_current_version()
    
    # 自動遞增建置編號
    if new_build is None:
        new_build = current_build + 1
    
    print(f"版本更新: {current_version} (Build {current_build}) -> {new_version} (Build {new_build})")
    
    # 讀取 VERSION.py
    with open('VERSION.py', 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 更新版本號
    content = re.sub(r'VERSION = "[^"]*"', f'VERSION = "{new_version}"', content)
    content = re.sub(r'BUILD = \d+', f'BUILD = {new_build}', content)
    content = re.sub(r'RELEASE_DATE = "[^"]*"', f'RELEASE_DATE = "{datetime.now().strftime("%Y-%m-%d")}"', content)
    
    # 寫回檔案
    with open('VERSION.py', 'w', encoding='utf-8') as f:
        f.write(content)
    
    print(f"[OK] 版本已更新為 {new_version} (Build {new_build})")

def add_change(change_text):
    """新增變更記錄"""
    with open('VERSION.py', 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 在 CHANGES 列表開頭新增變更記錄
    new_entry = f'    "{change_text}",'
    
    # 找到 CHANGES 列表的開始位置
    changes_start = content.find('CHANGES = [')
    if changes_start != -1:
        # 找到第一個變更記錄的位置
        first_change = content.find('"', changes_start + len('CHANGES = ['))
        if first_change != -1:
            # 在第一個變更記錄前插入新記錄
            content = content[:first_change] + new_entry + '\n    ' + content[first_change:]
            
            with open('VERSION.py', 'w', encoding='utf-8') as f:
                f.write(content)
            
            print(f"[OK] 已新增變更記錄: {change_text}")

def show_version():
    """顯示當前版本資訊"""
    try:
        from VERSION import VERSION, BUILD, RELEASE_DATE, AUTHOR, APP_NAME, DESCRIPTION, CHANGES
        print(f"{APP_NAME} v{VERSION} (Build {BUILD})")
        print(f"發布日期: {RELEASE_DATE}")
        print(f"作者: {AUTHOR}")
        print(f"描述: {DESCRIPTION}")
        print("\n變更記錄:")
        for change in CHANGES[:5]:  # 只顯示前5個
            print(f"  - {change}")
    except Exception as e:
        print(f"無法讀取版本資訊: {e}")

def main():
    """主程式"""
    import sys
    
    if len(sys.argv) < 2:
        print("使用方法:")
        print("  python version_tool.py show                    # 顯示版本資訊")
        print("  python version_tool.py set <version> [build]   # 設定版本號")
        print("  python version_tool.py add <change>            # 新增變更記錄")
        print("  python version_tool.py bump                    # 自動遞增修訂版本")
        return
    
    command = sys.argv[1].lower()
    
    if command == "show":
        show_version()
    elif command == "set":
        if len(sys.argv) < 3:
            print("請提供版本號")
            return
        new_version = sys.argv[2]
        new_build = int(sys.argv[3]) if len(sys.argv) > 3 else None
        set_version(new_version, new_build)
    elif command == "add":
        if len(sys.argv) < 3:
            print("請提供變更說明")
            return
        change_text = sys.argv[2]
        add_change(change_text)
    elif command == "bump":
        current_version, current_build = get_current_version()
        parts = current_version.split('.')
        parts[2] = str(int(parts[2]) + 1)
        new_version = '.'.join(parts)
        set_version(new_version)
    else:
        print(f"未知命令: {command}")

if __name__ == "__main__":
    main()
