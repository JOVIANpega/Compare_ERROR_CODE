# Error Code Comparer v1.5.0 (Build 1)

這是一個用於比對 Error Code 的工具，可以將 Excel 文件中的 Error Code 與參考文件進行比對，並生成新的比對結果文件。

## 版本資訊

- **當前版本**：v1.5.0 (Build 1)
- **發布日期**：2025-01-22
- **作者**：JOVIANpega
- **倉庫**：[https://github.com/JOVIANpega/Compare_ERROR_CODE](https://github.com/JOVIANpega/Compare_ERROR_CODE)
- **描述**：智能錯誤碼比對工具，支援 AI 推薦分析

## 功能特點

- 支持選擇 Error Code XML 文件和源 Excel 文件
- 可以選擇要處理的工作表
- 自動比對 Error Code 並生成結果
- 生成新的 Excel 文件，包含原始數據和比對結果
- 支持多種 Error Code 格式的比對
- **AI 推薦分析**：基於 Description 自動推薦最適合的 Test ID
- **智能 PROMPT 生成**：為 AI 工具生成專業的比對 PROMPT
- **快速開啟結果檔案**：一鍵開啟比對結果檔案，支援多檔案選擇
- **自動檔案選擇**：啟動時自動搜尋並選擇 Test Item Code 文件
- **智能工作表過濾**：自動排除不需要的工作表，預設選擇第一個有效工作表
- **版本管理系統**：完整的版本控制和變更記錄管理
- **檔案搜尋優化**：修復檔案搜尋重複問題，提升搜尋效率

## 使用方法

### 基本比對功能
1. 運行 ErrorCodeComparer.exe
2. **自動選擇 Error Code 文件**：程式會自動搜尋並選擇 Test Item Code 文件
3. 點擊 "Browse" 選擇源 Excel 文件
4. **自動載入工作表**：選擇 Excel 文件後會自動載入所有工作表
5. **智能工作表過濾**：自動排除 properties、DUTs、Switch、Instrument 等不需要的工作表
6. **預設選擇第一個工作表**：自動選擇第一個有效的工作表（如 FWDL）
7. 點擊 "Compare" 開始比對
8. 比對完成後，會在源文件相同目錄下生成新的 Excel 文件

### AI 推薦分析功能
1. 完成基本比對後，點擊 "AI推薦分析" 按鈕
2. 系統會自動分析 Description 並推薦最適合的 Test ID
3. 推薦結果會自動填入 E、F 欄位（AI推薦 test ID 1、AI推薦 test ID 2）
4. 如果沒有現有比對結果，系統會顯示 PROMPT 供手動處理

### 錯誤碼查詢功能
1. 點擊 "錯誤碼查詢" 按鈕開啟查詢視窗
2. 選擇 Excel 檔案並載入 Test Item All 工作表
3. 輸入關鍵字進行搜尋
4. 查看匹配的錯誤碼和描述

### 開啟結果檔案功能
1. 點擊 "開啟結果檔案" 按鈕
2. 系統會自動搜尋所有 `compare_ERRORCODE.xlsx` 檔案
3. 如果只有一個檔案，會直接開啟
4. 如果有多個檔案，會顯示選擇對話框讓使用者選擇
5. 支援顯示檔案大小和修改時間等詳細資訊

## 參考資料格式說明

### Test Item Code CSV 檔案結構

參考資料 `Test Item Code V2.00_20241106_CSV.csv` 包含以下欄位：

| 欄位 | 說明 | 範例 |
|------|------|------|
| Main Function | 主要功能分類 | AFM, Audio, BOARD Measure |
| Interface | 介面類型 | Audio RCA L, Audio Jack R |
| Interenal Error Code | 內部錯誤代碼 | AFFY001, ADRL001 |
| Description | 英文描述 | AFM Frequency Fail |
| 中文描述 | 中文描述 | AFM Frequency 頻率測試失敗 |
| Version | 版本資訊 | Rev. 1.3, Rev. 2.00 |
| Error Code | 實際錯誤代碼 | BSFTUJ, E00IO000008 |
| Note | 備註 | 開發者資訊和版本說明 |

### 支援的 Error Code 格式

1. **BSF 系列**：如 `BSFTUJ`, `BSFR8L`, `BSFA04`
2. **E00 系列**：如 `E00IO000008`, `E00AT000164`, `E00FW000238`

### 功能分類

- **AFM**：AFM 功能測試
- **Audio**：音頻相關測試（RCA L/R, Jack L/R, SCART, SPDIF）
- **BOARD Measure**：板卡測量（電壓、系統、電源）
- **CA FUNCTION**：CA 功能測試
- **其他功能模組**：包含各種硬體和軟體測試項目

## AI 推薦功能說明

### 推薦邏輯
- **關鍵字匹配**：基於 Description 中的關鍵字進行匹配
- **功能分類分析**：考慮 AFM、Audio、BOARD Measure 等功能分類
- **語義理解**：支援中英文描述的語義分析
- **模糊匹配**：即使不完全匹配也能找到相關的 Test ID

### 輸出格式
AI 推薦結果會新增到比對結果檔案的 E、F 欄位：
- **E 欄位**：AI推薦 test ID 1（最匹配的推薦）
- **F 欄位**：AI推薦 test ID 2（次匹配的推薦）

### PROMPT 模板
系統提供多種 PROMPT 模板：
- **基本分析**：適用於一般 Description 分析
- **單一分析**：適用於單個 Description 的詳細分析
- **批量分析**：適用於多個 Description 的批量處理
- **Excel 整合**：適用於直接填入 Excel 欄位

## 專案結構

```
Error Code Comparer/
├── main.py                          # 主程式入口
├── ui_manager.py                    # GUI 介面管理
├── excel_handler.py                 # Excel 檔案處理
├── ai_recommendation_engine.py      # AI 推薦引擎
├── ai_prompt_templates.py           # AI PROMPT 模板
├── file_finder.py                   # 檔案搜尋工具
├── config_manager.py                # 配置管理
├── error_code_compare.py            # 錯誤碼比對核心
├── excel_errorcode_search_ui.py     # 錯誤碼查詢 UI
├── VERSION.py                       # 版本配置檔案
├── version_tool.py                  # 版本管理工具
├── version_config.py                # 完整版本管理工具
├── version_manager.py               # 進階版本管理工具
├── update_version.py                # 版本更新腳本
├── VERSION_USAGE.md                 # 版本管理使用說明
├── AI_RECOMMENDATION_USAGE.md       # AI 推薦功能說明
├── README.md                        # 專案說明文件
├── requirements.txt                 # Python 依賴
├── ErrorCodeComparer.spec           # PyInstaller 配置
├── pal.ico                          # 應用程式圖示
├── guide_popup/                     # GUI 引導圖片
│   ├── guide.py
│   ├── guide1.png
│   ├── guide2.png
│   ├── guide3.png
│   ├── guide4.png
│   └── guide5.png
├── EXCEL/                           # 範例 Excel 檔案
│   ├── Test Item Code V2.00_20241106.xlsx
│   ├── Test Item Code V2.00_20241106_CSV.csv
│   ├── MU310_TestFlow_FWdownload_20250902.xlsx
│   └── MU310_TestFlow_FWdownload_20250902_compare_ERRORCODE.xlsx
├── dist/                            # 編譯後的執行檔
│   └── ErrorCodeComparer.exe
└── build/                           # 編譯暫存檔案
```

## 版本管理

### 版本控制工具

本專案提供多種版本管理工具：

1. **VERSION.py** - 主要版本配置檔案
   - 包含版本號、建置編號、發布日期等資訊
   - 功能開關控制
   - 變更記錄管理

2. **version_tool.py** - 簡單版本管理工具
   ```bash
   # 顯示當前版本資訊
   python version_tool.py show
   
   # 設定版本號
   python version_tool.py set 1.6.0
   
   # 新增變更記錄
   python version_tool.py add "修復 AI 推薦分析問題"
   
   # 自動遞增修訂版本
   python version_tool.py bump
   ```

3. **version_config.py** - 完整版本管理工具
   - 提供互動式版本管理介面
   - 支援功能開關管理
   - 變更記錄編輯

4. **VERSION_USAGE.md** - 版本管理使用說明
   - 詳細的使用指南
   - 最佳實踐建議
   - 範例工作流程

### 變更記錄

#### v1.5.0 (2025-01-22)
- 初始版本 1.5.0
- 新增 AI 推薦分析功能
- 新增自動檔案選擇功能
- 新增智能工作表過濾功能
- 新增快速開啟結果檔案功能
- 修復檔案搜尋重複問題
- 優化 GUI 用戶體驗
- 新增版本管理系統

## 注意事項

- 確保 Error Code XML 文件包含 "Test Item All" 工作表
- 源 Excel 文件必須包含 O 行（Description）和 P 行（TestID）
- 生成的結果文件會自動命名為原文件名加上 "_compare_ERRORCODE.xlsx"
- 參考資料檔案包含約 5,000+ 個測試項目和對應的錯誤代碼
- 支援中英文雙語描述，方便不同語言環境使用
- AI 推薦功能需要先完成基本比對，或手動處理 PROMPT
- 版本管理工具需要 Python 環境才能使用 