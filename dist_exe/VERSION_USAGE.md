# 版本管理使用說明

## 概述

Error Code Comparer 提供了多種版本管理工具，方便你快速修改和維護版本號。

## 檔案說明

### 1. VERSION.py - 主要版本配置檔案
這是主要的版本配置檔案，包含所有版本相關的設定：

```python
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
```

### 2. version_tool.py - 簡單版本管理工具
提供命令列介面來快速修改版本：

```bash
# 顯示當前版本資訊
python version_tool.py show

# 設定版本號
python version_tool.py set 1.6.0

# 設定版本號和建置編號
python version_tool.py set 1.6.0 5

# 新增變更記錄
python version_tool.py add "修復 AI 推薦分析問題"

# 自動遞增修訂版本
python version_tool.py bump
```

### 3. version_config.py - 完整版本管理工具
提供更完整的版本管理功能，包括功能開關和變更記錄管理。

### 4. version_manager.py - 進階版本管理工具
提供 JSON 格式的版本資料儲存和更複雜的版本管理功能。

## 快速使用指南

### 方法一：直接修改 VERSION.py
1. 開啟 `VERSION.py` 檔案
2. 修改 `VERSION` 變數為新版本號
3. 修改 `BUILD` 變數為新建置編號
4. 在 `CHANGES` 列表開頭新增變更記錄

### 方法二：使用 version_tool.py
```bash
# 更新到版本 1.6.0
python version_tool.py set 1.6.0

# 新增變更記錄
python version_tool.py add "新增功能 X"
python version_tool.py add "修復問題 Y"
```

### 方法三：使用 version_config.py
```bash
python version_config.py
```
然後按照選單提示操作。

## 版本號格式

- **主版本號**：重大功能更新或架構變更
- **次版本號**：新功能添加或重要改進
- **修訂版本號**：錯誤修復或小改進

範例：
- `1.5.0` - 初始版本
- `1.5.1` - 修復小問題
- `1.6.0` - 新增功能
- `2.0.0` - 重大更新

## 建置編號

建置編號用於追蹤每次發布，建議每次更新版本時自動遞增。

## 變更記錄

在 `VERSION.py` 的 `CHANGES` 列表中記錄每次版本的變更：

```python
CHANGES = [
    "版本 1.6.0 - 2025-01-23",
    "修復 AI 推薦分析問題",
    "優化檔案搜尋邏輯",
    "改進錯誤處理機制",
    "初始版本 1.5.0",
    "新增 AI 推薦分析功能",
    # ... 更多變更記錄
]
```

## 功能開關

在 `VERSION.py` 中可以控制各個功能的啟用狀態：

```python
# AI 推薦分析功能
AI_RECOMMENDATION = True

# 自動檔案選擇功能
AUTO_FILE_SELECTION = True

# 智能工作表過濾功能
SMART_SHEET_FILTERING = True
```

## 最佳實踐

1. **版本號管理**：
   - 每次發布前更新版本號
   - 使用語義化版本號
   - 記錄所有變更

2. **建置編號**：
   - 每次發布時遞增
   - 用於追蹤發布歷史

3. **變更記錄**：
   - 記錄所有重要變更
   - 使用清晰的描述
   - 按時間倒序排列

4. **功能開關**：
   - 用於控制實驗性功能
   - 方便快速啟用/停用功能

## 範例工作流程

1. **準備新版本**：
   ```bash
   python version_tool.py show
   ```

2. **更新版本號**：
   ```bash
   python version_tool.py set 1.6.0
   ```

3. **新增變更記錄**：
   ```bash
   python version_tool.py add "修復 AI 推薦分析問題"
   python version_tool.py add "優化檔案搜尋邏輯"
   ```

4. **確認版本資訊**：
   ```bash
   python version_tool.py show
   ```

5. **提交到 Git**：
   ```bash
   git add VERSION.py
   git commit -m "更新版本到 1.6.0"
   git push
   ```

這樣你就可以輕鬆管理 Error Code Comparer 的版本了！
