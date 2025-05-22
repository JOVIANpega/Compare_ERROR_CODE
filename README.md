# Error Code Comparer

這是一個用於比對 Error Code 的工具，可以將 Excel 文件中的 Error Code 與參考文件進行比對，並生成新的比對結果文件。

## 功能特點

- 支持選擇 Error Code XML 文件和源 Excel 文件
- 可以選擇要處理的工作表
- 自動比對 Error Code 並生成結果
- 生成新的 Excel 文件，包含原始數據和比對結果

## 使用方法

1. 運行 ErrorCodeComparer.exe
2. 點擊 "Browse" 選擇 Error Code XML 文件
3. 點擊 "Browse" 選擇源 Excel 文件
4. 從下拉選單中選擇要處理的工作表
5. 點擊 "Compare" 開始比對
6. 比對完成後，會在源文件相同目錄下生成新的 Excel 文件

## 注意事項

- 確保 Error Code XML 文件包含 "Test Item All" 工作表
- 源 Excel 文件必須包含 O 行（Description）和 P 行（TestID）
- 生成的結果文件會自動命名為原文件名加上 "_compare ERRORCODE.xml" 