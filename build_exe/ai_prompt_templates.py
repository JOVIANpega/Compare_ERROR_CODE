"""
AI PROMPT 模板模組
提供各種 AI 推薦 Test ID 的 PROMPT 模板
"""
import logging

logger = logging.getLogger(__name__)

class AIPromptTemplates:
    """AI PROMPT 模板類別"""
    
    @staticmethod
    def get_basic_analysis_prompt(descriptions, reference_file_path):
        """
        基本分析 PROMPT 模板
        
        Args:
            descriptions: 要分析的 Description 列表
            reference_file_path: 參考資料檔案路徑
            
        Returns:
            str: 格式化的 PROMPT
        """
        prompt = f"""
你是一個專業的 Error Code 分析助手，基於 Error Code Comparer 工具的工作原理。

【任務描述】
- 讀取 Description 列表並分析
- 參考資料：{reference_file_path}
- 輸出目標：為每個 Description 推薦 2 個最適合的 Test ID

【Description 列表】
"""
        for i, desc in enumerate(descriptions, 1):
            prompt += f"{i}. {desc}\n"
        
        prompt += """
【輸出格式】
為每個 Description 提供 2 個推薦 Test ID，格式如下：
1. [Test ID 1] | [Test ID 2]
2. [Test ID 1] | [Test ID 2]
3. [Test ID 1] | [Test ID 2]
...

【匹配策略】
- 完全匹配：10/10
- 高度相似：8-9/10  
- 部分匹配：6-7/10
- 功能相關：4-5/10
- 優先考慮 BSF 系列和 E00 系列格式

【參考資料結構】
- Main Function: 主要功能分類 (AFM, Audio, BOARD Measure)
- Interface: 介面類型 (Audio RCA L, Audio Jack R)
- Interenal Error Code: 內部錯誤代碼 (AFFY001, ADRL001)
- Description: 英文描述
- 中文描述: 中文描述
- Version: 版本資訊
- Error Code: 實際錯誤代碼 (BSF系列, E00系列)
- Note: 備註
"""
        return prompt

    @staticmethod
    def get_single_analysis_prompt(description, reference_file_path):
        """
        單一 Description 分析 PROMPT 模板
        
        Args:
            description: 要分析的 Description
            reference_file_path: 參考資料檔案路徑
            
        Returns:
            str: 格式化的 PROMPT
        """
        prompt = f"""
請分析以下 Description 並推薦 2 個最適合的 Test ID：

【Description】
{description}

【參考資料】
{reference_file_path} (包含約 5,000+ 個測試項目)

【分析要求】
1. 理解 Description 的功能和語義
2. 在參考資料中搜尋最匹配的項目
3. 考慮功能分類、介面類型、描述相似性
4. 提供 2 個推薦 Test ID（按匹配度排序）

【輸出格式】
原始 Description: {description}
推薦 Test ID 1: [代碼] - [匹配度]/10 - [理由]
推薦 Test ID 2: [代碼] - [匹配度]/10 - [理由]

【匹配標準】
- 完全匹配：10/10
- 高度相似：8-9/10
- 部分匹配：6-7/10
- 功能相關：4-5/10
"""
        return prompt

    @staticmethod
    def get_batch_analysis_prompt(descriptions, reference_file_path):
        """
        批量分析 PROMPT 模板
        
        Args:
            descriptions: 要分析的 Description 列表
            reference_file_path: 參考資料檔案路徑
            
        Returns:
            str: 格式化的 PROMPT
        """
        prompt = f"""
基於你的 Error Code Comparer 工具，請執行以下批量分析：

【處理流程】
1. 讀取 Description 列表
2. 分析：每個 Description 的語義和功能
3. 匹配：在 Test Item Code 資料庫中搜尋
4. 推薦：每個 Description 提供 2 個最佳 Test ID
5. 輸出：生成可直接填入 E、F 欄位的結果

【Description 列表】
"""
        for i, desc in enumerate(descriptions, 1):
            prompt += f"{i}. {desc}\n"
        
        prompt += f"""
【參考資料】
{reference_file_path}

【輸出要求】
為每個 Description 提供 2 個推薦 Test ID，格式如下：
1. [Test ID 1] | [Test ID 2]
2. [Test ID 1] | [Test ID 2]
3. [Test ID 1] | [Test ID 2]
...

【特殊要求】
- 優先考慮功能分類匹配
- 考慮中英文描述相似性
- 提供匹配度評分和推薦理由
- 支援模糊匹配和語義理解
"""
        return prompt

    @staticmethod
    def get_excel_integration_prompt(descriptions, reference_file_path):
        """
        Excel 整合 PROMPT 模板（用於直接填入 E、F 欄位）
        
        Args:
            descriptions: 要分析的 Description 列表
            reference_file_path: 參考資料檔案路徑
            
        Returns:
            str: 格式化的 PROMPT
        """
        prompt = f"""
請基於你的 Error Code Comparer 工具，為現有的比對結果檔案新增 AI 推薦欄位。

【現有檔案結構】
- 檔案名稱：XXX_compare_ERRORCODE.xlsx
- 現有欄位：A=你的 description, B=你寫的 Error Code, C=Test Item 文件的 description, D=Test Item 的 Error Code
- 新增欄位：E=AI推薦 test ID 1, F=AI推薦 test ID 2

【Description 列表】
"""
        for i, desc in enumerate(descriptions, 1):
            prompt += f"{i}. {desc}\n"
        
        prompt += f"""
【處理邏輯】
1. 讀取 Description 列表
2. 分析每個 Description 的語義和功能
3. 在 Test Item Code 資料庫中搜尋最匹配的項目
4. 為每個 Description 推薦 2 個最佳 Test ID
5. 將推薦結果填入 E、F 欄位

【參考資料】
{reference_file_path}

【輸出格式】
直接提供可填入 E、F 欄位的 Test ID 列表：
1. [Test ID 1] | [Test ID 2]
2. [Test ID 1] | [Test ID 2]
3. [Test ID 1] | [Test ID 2]
...

【匹配標準】
- 完全匹配：優先推薦
- 功能相似：次優先推薦
- 語義相關：備選推薦
"""
        return prompt

    @staticmethod
    def parse_ai_response(response_text):
        """
        解析 AI 回應，提取 Test ID 推薦
        
        Args:
            response_text: AI 回應文字
            
        Returns:
            list: 包含 (test_id_1, test_id_2) 元組的列表
        """
        recommendations = []
        lines = response_text.strip().split('\n')
        
        for line in lines:
            line = line.strip()
            if not line or not line[0].isdigit():
                continue
                
            # 移除行號
            if '. ' in line:
                line = line.split('. ', 1)[1]
            
            # 檢查是否包含分隔符
            if ' | ' in line:
                parts = line.split(' | ')
                if len(parts) >= 2:
                    test_id_1 = parts[0].strip()
                    test_id_2 = parts[1].strip()
                    recommendations.append((test_id_1, test_id_2))
            elif '[' in line and ']' in line:
                # 處理 [Test ID 1] | [Test ID 2] 格式
                import re
                matches = re.findall(r'\[([^\]]+)\]', line)
                if len(matches) >= 2:
                    recommendations.append((matches[0], matches[1]))
        
        return recommendations

    @staticmethod
    def get_error_handling_prompt():
        """
        錯誤處理 PROMPT 模板
        
        Returns:
            str: 錯誤處理 PROMPT
        """
        return """
如果無法找到合適的 Test ID 推薦，請：
1. 提供最接近的 2 個選項
2. 在推薦理由中說明匹配度較低的原因
3. 建議使用者檢查 Description 是否完整或準確
4. 提供相關的功能分類建議

格式：
1. [最接近的 Test ID 1] | [最接近的 Test ID 2] (匹配度較低)
2. [相關的 Test ID 1] | [相關的 Test ID 2] (功能相關)
"""
