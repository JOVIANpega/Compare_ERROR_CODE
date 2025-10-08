# ========================================
# Error Code Comparer 版本設定
# 修改這個檔案來更新版本號
# ========================================

# 版本號 (格式: 主版本.次版本.修訂版本)
VERSION = "1.5.0"

# 建置編號
BUILD = 1

# 發布日期
RELEASE_DATE = "2025-01-22"

# 作者
AUTHOR = "JOVIANpega"

# 倉庫
REPOSITORY = "https://github.com/JOVIANpega/Compare_ERROR_CODE"

# 應用程式名稱
APP_NAME = "Error Code Comparer"

# 應用程式描述
DESCRIPTION = "智能錯誤碼比對工具，支援 AI 推薦分析"

# ========================================
# 功能開關
# ========================================

# AI 推薦分析功能
AI_RECOMMENDATION = True

# 自動檔案選擇功能
AUTO_FILE_SELECTION = True

# 智能工作表過濾功能
SMART_SHEET_FILTERING = True

# 快速開啟結果檔案功能
QUICK_OPEN_RESULT = True

# Excel 比對功能
EXCEL_COMPARISON = True

# 錯誤碼查詢功能
ERROR_CODE_SEARCH = True

# 版本檢查功能
VERSION_CHECK = True

# 自動更新功能
AUTO_UPDATE = False

# ========================================
# 變更記錄
# ========================================

CHANGES = [
    "新增版本管理系統",
    "初始版本 1.5.0",
    "新增 AI 推薦分析功能",
    "新增自動檔案選擇功能", 
    "新增智能工作表過濾功能",
    "新增快速開啟結果檔案功能",
    "修復檔案搜尋重複問題",
    "優化 GUI 用戶體驗"
]

# ========================================
# 版本資訊函數
# ========================================

def get_version():
    """取得版本字串"""
    return f"{VERSION} (Build {BUILD})"

def get_info():
    """取得版本資訊"""
    return {
        "version": VERSION,
        "build": BUILD,
        "release_date": RELEASE_DATE,
        "author": AUTHOR,
        "repository": REPOSITORY,
        "app_name": APP_NAME,
        "description": DESCRIPTION,
        "changes": CHANGES
    }

def print_info():
    """列印版本資訊"""
    print(f"{APP_NAME} v{VERSION} (Build {BUILD})")
    print(f"發布日期: {RELEASE_DATE}")
    print(f"作者: {AUTHOR}")
    print(f"倉庫: {REPOSITORY}")
    print(f"描述: {DESCRIPTION}")
    print("\n變更記錄:")
    for change in CHANGES:
        print(f"  - {change}")

if __name__ == "__main__":
    print_info()