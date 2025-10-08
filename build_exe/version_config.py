#!/usr/bin/env python3
"""
版本配置檔案
快速修改版本號和相關設定
"""

# ========================================
# 版本設定 - 修改這裡來更新版本
# ========================================

# 主要版本號 (格式: 主版本.次版本.修訂版本)
VERSION = "1.5.0"

# 建置編號 (每次發布時自動遞增)
BUILD_NUMBER = 1

# 發布日期
RELEASE_DATE = "2025-01-22"

# 作者資訊
AUTHOR = "JOVIANpega"

# 倉庫網址
REPOSITORY = "https://github.com/JOVIANpega/Compare_ERROR_CODE"

# 應用程式名稱
APP_NAME = "Error Code Comparer"

# 應用程式描述
APP_DESCRIPTION = "智能錯誤碼比對工具，支援 AI 推薦分析"

# ========================================
# 功能開關 - 啟用/停用特定功能
# ========================================

FEATURES = {
    "ai_recommendation": True,      # AI 推薦分析功能
    "auto_file_selection": True,    # 自動檔案選擇功能
    "smart_sheet_filtering": True,  # 智能工作表過濾功能
    "quick_open_result": True,      # 快速開啟結果檔案功能
    "excel_comparison": True,       # Excel 比對功能
    "error_code_search": True,      # 錯誤碼查詢功能
    "version_check": True,          # 版本檢查功能
    "auto_update": False,           # 自動更新功能
}

# ========================================
# 變更記錄 - 記錄各版本的變更
# ========================================

CHANGELOG = [
    {
        "version": "1.5.0",
        "date": "2025-01-22",
        "changes": [
            "初始版本 1.5.0",
            "新增 AI 推薦分析功能",
            "新增自動檔案選擇功能",
            "新增智能工作表過濾功能",
            "新增快速開啟結果檔案功能",
            "修復檔案搜尋重複問題",
            "優化 GUI 用戶體驗"
        ]
    }
]

# ========================================
# 版本資訊生成函數
# ========================================

def get_version_string():
    """取得版本字串"""
    return f"{VERSION} (Build {BUILD_NUMBER})"

def get_full_version_info():
    """取得完整版本資訊"""
    return {
        "version": VERSION,
        "build_number": BUILD_NUMBER,
        "release_date": RELEASE_DATE,
        "author": AUTHOR,
        "repository": REPOSITORY,
        "app_name": APP_NAME,
        "app_description": APP_DESCRIPTION,
        "features": FEATURES,
        "changelog": CHANGELOG
    }

def is_feature_enabled(feature_name):
    """檢查功能是否啟用"""
    return FEATURES.get(feature_name, False)

def get_version_header():
    """生成版本標頭"""
    return f"""# {APP_NAME} v{VERSION} (Build {BUILD_NUMBER})
# 發布日期: {RELEASE_DATE}
# 作者: {AUTHOR}
# 倉庫: {REPOSITORY}
# 描述: {APP_DESCRIPTION}
"""

def print_version_info():
    """列印版本資訊"""
    print("=" * 60)
    print(f"{APP_NAME} 版本資訊")
    print("=" * 60)
    print(f"版本號: {get_version_string()}")
    print(f"發布日期: {RELEASE_DATE}")
    print(f"作者: {AUTHOR}")
    print(f"倉庫: {REPOSITORY}")
    print(f"描述: {APP_DESCRIPTION}")
    
    print("\n功能狀態:")
    for feature, enabled in FEATURES.items():
        status = "✓ 啟用" if enabled else "✗ 停用"
        print(f"  {feature}: {status}")
    
    print("\n變更記錄:")
    for entry in CHANGELOG:
        print(f"  {entry['version']} ({entry['date']}):")
        for change in entry['changes']:
            print(f"    - {change}")
    print("=" * 60)

# ========================================
# 快速修改函數
# ========================================

def update_version(new_version, new_build=None):
    """快速更新版本號"""
    global VERSION, BUILD_NUMBER
    VERSION = new_version
    if new_build is not None:
        BUILD_NUMBER = new_build
    else:
        BUILD_NUMBER += 1
    print(f"版本已更新為: {get_version_string()}")

def add_changelog_entry(version, changes):
    """新增變更記錄"""
    global CHANGELOG
    from datetime import datetime
    entry = {
        "version": version,
        "date": datetime.now().strftime("%Y-%m-%d"),
        "changes": changes if isinstance(changes, list) else [changes]
    }
    CHANGELOG.insert(0, entry)
    print(f"已為版本 {version} 新增變更記錄")

def toggle_feature(feature_name):
    """切換功能開關"""
    global FEATURES
    if feature_name in FEATURES:
        FEATURES[feature_name] = not FEATURES[feature_name]
        status = "啟用" if FEATURES[feature_name] else "停用"
        print(f"功能 '{feature_name}' 已{status}")
    else:
        print(f"功能 '{feature_name}' 不存在")

# ========================================
# 主程式
# ========================================

if __name__ == "__main__":
    print("Error Code Comparer 版本配置工具")
    print("=" * 40)
    
    while True:
        print("\n請選擇操作:")
        print("1. 查看當前版本資訊")
        print("2. 更新版本號")
        print("3. 新增變更記錄")
        print("4. 切換功能開關")
        print("5. 生成版本標頭")
        print("0. 退出")
        
        choice = input("\n請輸入選項 (0-5): ").strip()
        
        if choice == "0":
            print("再見！")
            break
        elif choice == "1":
            print_version_info()
        elif choice == "2":
            new_version = input("請輸入新版本號 (格式: x.y.z): ").strip()
            new_build = input("請輸入建置編號 (留空自動遞增): ").strip()
            new_build = int(new_build) if new_build.isdigit() else None
            update_version(new_version, new_build)
        elif choice == "3":
            version = input("請輸入版本號: ").strip()
            changes_input = input("請輸入變更說明 (多行，空行結束): ").strip()
            changes = []
            while changes_input:
                changes.append(changes_input)
                changes_input = input().strip()
            add_changelog_entry(version, changes)
        elif choice == "4":
            print("可用功能:")
            for i, feature in enumerate(FEATURES.keys(), 1):
                status = "啟用" if FEATURES[feature] else "停用"
                print(f"  {i}. {feature} ({status})")
            
            feature_name = input("請輸入功能名稱: ").strip()
            toggle_feature(feature_name)
        elif choice == "5":
            header = get_version_header()
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
