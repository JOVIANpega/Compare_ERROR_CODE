#!/usr/bin/env python3
"""
版本更新腳本
快速更新 Error Code Comparer 的版本號
"""

import os
import re
from datetime import datetime

def update_version_files(new_version, new_build=None):
    """更新所有版本相關檔案"""
    
    # 自動遞增建置編號
    if new_build is None:
        try:
            from VERSION import BUILD
            new_build = BUILD + 1
        except:
            new_build = 1
    
    print(f"正在更新版本到 {new_version} (Build {new_build})...")
    
    # 更新 VERSION.py
    update_version_py(new_version, new_build)
    
    # 更新 main.py 中的版本資訊
    update_main_py(new_version, new_build)
    
    # 更新 README.md 中的版本資訊
    update_readme_md(new_version, new_build)
    
    print("版本更新完成！")

def update_version_py(new_version, new_build):
    """更新 VERSION.py"""
    try:
        with open('VERSION.py', 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 更新版本號
        content = re.sub(r'VERSION = "[^"]*"', f'VERSION = "{new_version}"', content)
        content = re.sub(r'BUILD = \d+', f'BUILD = {new_build}', content)
        content = re.sub(r'RELEASE_DATE = "[^"]*"', f'RELEASE_DATE = "{datetime.now().strftime("%Y-%m-%d")}"', content)
        
        with open('VERSION.py', 'w', encoding='utf-8') as f:
            f.write(content)
        
        print("✓ 已更新 VERSION.py")
    except Exception as e:
        print(f"✗ 更新 VERSION.py 失敗: {e}")

def update_main_py(new_version, new_build):
    """更新 main.py 中的版本資訊"""
    try:
        with open('main.py', 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 尋找並更新版本資訊
        version_pattern = r'version["\']?\s*[:=]\s*["\'][^"\']*["\']'
        build_pattern = r'build["\']?\s*[:=]\s*\d+'
        
        if re.search(version_pattern, content):
            content = re.sub(version_pattern, f'version = "{new_version}"', content)
        if re.search(build_pattern, content):
            content = re.sub(build_pattern, f'build = {new_build}', content)
        
        with open('main.py', 'w', encoding='utf-8') as f:
            f.write(content)
        
        print("✓ 已更新 main.py")
    except Exception as e:
        print(f"✗ 更新 main.py 失敗: {e}")

def update_readme_md(new_version, new_build):
    """更新 README.md 中的版本資訊"""
    try:
        with open('README.md', 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 更新標題中的版本號
        content = re.sub(r'# Error Code Comparer v[\d.]+', f'# Error Code Comparer v{new_version}', content)
        
        # 更新功能特點中的版本資訊
        content = re.sub(r'版本 [\d.]+', f'版本 {new_version}', content)
        
        with open('README.md', 'w', encoding='utf-8') as f:
            f.write(content)
        
        print("✓ 已更新 README.md")
    except Exception as e:
        print(f"✗ 更新 README.md 失敗: {e}")

def add_changelog_entry(new_version, changes):
    """新增變更記錄"""
    try:
        with open('VERSION.py', 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 在 CHANGES 列表開頭新增變更記錄
        new_entry = f'    "版本 {new_version} - {datetime.now().strftime("%Y-%m-%d")}",'
        for change in changes:
            new_entry += f'\n    "{change}",'
        
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
                
                print("✓ 已新增變更記錄")
    except Exception as e:
        print(f"✗ 新增變更記錄失敗: {e}")

def main():
    """主程式"""
    print("Error Code Comparer 版本更新工具")
    print("=" * 40)
    
    # 顯示當前版本
    try:
        from VERSION import VERSION, BUILD
        print(f"當前版本: {VERSION} (Build {BUILD})")
    except:
        print("無法讀取當前版本")
    
    print("\n請選擇操作:")
    print("1. 更新版本號")
    print("2. 新增變更記錄")
    print("3. 自動更新到下一版本")
    print("0. 退出")
    
    choice = input("\n請輸入選項 (0-3): ").strip()
    
    if choice == "0":
        print("再見！")
        return
    elif choice == "1":
        new_version = input("請輸入新版本號 (格式: x.y.z): ").strip()
        if not re.match(r'^\d+\.\d+\.\d+$', new_version):
            print("無效的版本號格式！")
            return
        
        new_build_input = input("請輸入建置編號 (留空自動遞增): ").strip()
        new_build = int(new_build_input) if new_build_input.isdigit() else None
        
        update_version_files(new_version, new_build)
    elif choice == "2":
        version = input("請輸入版本號: ").strip()
        print("請輸入變更說明 (多行，空行結束):")
        changes = []
        while True:
            change = input().strip()
            if not change:
                break
            changes.append(change)
        
        if changes:
            add_changelog_entry(version, changes)
        else:
            print("沒有輸入變更說明")
    elif choice == "3":
        try:
            from VERSION import VERSION
            # 自動更新到下一修訂版本
            parts = VERSION.split('.')
            parts[2] = str(int(parts[2]) + 1)
            new_version = '.'.join(parts)
            
            print(f"將更新到版本: {new_version}")
            confirm = input("確認更新? (y/n): ").strip().lower()
            if confirm in ['y', 'yes', '是', '1', 'true']:
                update_version_files(new_version)
            else:
                print("已取消更新")
        except Exception as e:
            print(f"自動更新失敗: {e}")
    else:
        print("無效的選項")

if __name__ == "__main__":
    main()
