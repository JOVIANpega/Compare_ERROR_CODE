# Error Code Comparer

這是一個用於比對 Error Code 的工具，可以將 Excel 文件中的 Error Code 與參考文件進行比對，並生成新的比對結果文件。

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

## 注意事項

- 確保 Error Code XML 文件包含 "Test Item All" 工作表
- 源 Excel 文件必須包含 O 行（Description）和 P 行（TestID）
- 生成的結果文件會自動命名為原文件名加上 "_compare ERRORCODE.xml"
- 參考資料檔案包含約 5,000+ 個測試項目和對應的錯誤代碼
- 支援中英文雙語描述，方便不同語言環境使用
- AI 推薦功能需要先完成基本比對，或手動處理 PROMPT 