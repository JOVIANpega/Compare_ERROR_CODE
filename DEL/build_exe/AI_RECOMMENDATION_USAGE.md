# AI 推薦功能使用指南

## 功能概述

AI 推薦功能可以根據 Excel 檔案中的 Description 自動推薦最適合的 Test ID，並將推薦結果填入 E、F 欄位。

## 使用步驟

### 步驟 1：準備檔案
1. 確保有參考資料檔案（Test Item Code V2.00_20241106_CSV.csv）
2. 準備包含 Description 欄位的 Excel 檔案
3. 確保 Description 欄位包含有意義的測試描述

### 步驟 2：執行基本比對
1. 開啟 ErrorCodeComparer.exe
2. 選擇參考資料檔案（Test Item Code XML 文件）
3. 選擇來源 Excel 檔案
4. 選擇要處理的工作表
5. 點擊 "Compare" 執行比對

### 步驟 3：執行 AI 推薦分析
1. 比對完成後，點擊 "AI推薦分析" 按鈕
2. 系統會自動分析所有 Description
3. 推薦結果會自動填入 E、F 欄位

## 輸出結果

### Excel 檔案結構
比對結果檔案會包含以下欄位：
- **A 欄位**：你的 description（原始 Description）
- **B 欄位**：你寫的 Error Code（原始 Test ID）
- **C 欄位**：Test Item 文件的 description（匹配的英文描述）
- **D 欄位**：Test Item 的 Error Code（匹配的中文描述）
- **E 欄位**：AI推薦 test ID 1（最匹配的推薦）
- **F 欄位**：AI推薦 test ID 2（次匹配的推薦）

### 範例輸出
```
| 你的 description | 你寫的 Error Code | Test Item 文件的 description | Test Item 的 Error Code | AI推薦 test ID 1 | AI推薦 test ID 2 |
|------------------|-------------------|------------------------------|-------------------------|-------------------|-------------------|
| PC#-#Show SSN to UI | B7PL025 | Test program can not running | 測試程式無反應 | OTFX064 | B7PL025 |
| PC#-#CheckRoute | OTFX064 | SFIS Link Fail | SFIS連接錯誤 | B7PL025 | OTFX064 |
```

## PROMPT 手動處理

如果沒有現有比對結果檔案，系統會顯示 PROMPT 供手動處理：

### 使用 PROMPT
1. 複製系統生成的 PROMPT
2. 貼到 AI 工具中（如 ChatGPT、Claude 等）
3. 獲取 AI 推薦結果
4. 手動填入 Excel 檔案的 E、F 欄位

### PROMPT 範例
```
你是一個專業的 Error Code 分析助手，基於 Error Code Comparer 工具的工作原理。

【任務描述】
- 讀取 Description 列表並分析
- 參考資料：EXCEL/Test Item Code V2.00_20241106_CSV.csv
- 輸出目標：為每個 Description 推薦 2 個最適合的 Test ID

【Description 列表】
1. PC#-#Show SSN to UI
2. PC#-#CheckRoute
3. DUT#-#Check_MO

【輸出格式】
為每個 Description 提供 2 個推薦 Test ID，格式如下：
1. [Test ID 1] | [Test ID 2]
2. [Test ID 1] | [Test ID 2]
3. [Test ID 1] | [Test ID 2]
```

## 推薦邏輯說明

### 匹配策略
1. **完全匹配**：Description 與參考資料完全一致
2. **關鍵字匹配**：基於 Description 中的關鍵字進行匹配
3. **功能分類匹配**：考慮功能分類（AFM、Audio、BOARD Measure 等）
4. **語義匹配**：支援中英文描述的語義理解

### 推薦優先順序
1. 完全匹配的 Test ID
2. 關鍵字高度相似的 Test ID
3. 功能分類相關的 Test ID
4. 語義相關的 Test ID

## 注意事項

### 使用前準備
- 確保參考資料檔案格式正確
- 確保 Description 欄位包含有意義的內容
- 建議先完成基本比對再使用 AI 推薦

### 推薦品質
- AI 推薦基於內建邏輯，可能不是 100% 準確
- 建議人工檢查推薦結果
- 可以根據實際需求調整推薦邏輯

### 檔案格式
- 支援 CSV 和 Excel 格式的參考資料
- 自動處理不同編碼的 CSV 檔案
- 支援中英文混合內容

## 故障排除

### 常見問題
1. **無法載入參考資料**：檢查檔案路徑和格式
2. **推薦結果為空**：檢查 Description 內容是否有效
3. **編碼錯誤**：系統會自動嘗試多種編碼方式

### 解決方案
1. 確認檔案存在且可讀取
2. 檢查 Description 欄位是否包含有效內容
3. 嘗試重新載入參考資料
4. 使用 PROMPT 手動處理

## 進階使用

### 自定義推薦邏輯
可以修改 `ai_recommendation_engine.py` 中的推薦邏輯：
- 調整關鍵字匹配權重
- 修改功能分類優先順序
- 增加自定義匹配規則

### 批量處理
- 支援一次處理多個 Description
- 自動生成批量處理 PROMPT
- 提供統計資訊和推薦品質評估

### 整合其他 AI 工具
- 支援多種 PROMPT 格式
- 可整合 ChatGPT、Claude 等 AI 工具
- 提供標準化的輸出格式
