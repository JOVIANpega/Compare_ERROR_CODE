"""
AI 推薦引擎模組
負責處理 AI 推薦 Test ID 的核心邏輯
"""
import logging
import pandas as pd
from pathlib import Path
from typing import List, Tuple, Optional
from ai_prompt_templates import AIPromptTemplates

logger = logging.getLogger(__name__)

class AIRecommendationEngine:
    """AI 推薦引擎類別"""
    
    def __init__(self):
        self.prompt_templates = AIPromptTemplates()
        self.reference_data = None
        self.reference_file_path = None
    
    def load_reference_data(self, file_path: str) -> bool:
        """
        載入參考資料
        
        Args:
            file_path: 參考資料檔案路徑
            
        Returns:
            bool: 是否成功載入
        """
        try:
            if file_path.endswith('.csv'):
                # 嘗試不同的編碼方式讀取 CSV
                encodings = ['utf-8', 'big5', 'gbk', 'cp950', 'latin1']
                for encoding in encodings:
                    try:
                        self.reference_data = pd.read_csv(file_path, encoding=encoding)
                        logger.info(f"成功使用 {encoding} 編碼讀取 CSV 檔案")
                        break
                    except UnicodeDecodeError:
                        continue
                else:
                    raise Exception("無法使用任何編碼讀取 CSV 檔案")
            else:
                # 讀取 Excel 檔案的 "Test Item All" 工作表
                # 跳過前3行空行，第4行是標題
                self.reference_data = pd.read_excel(file_path, sheet_name="Test Item All", skiprows=3)
                # 重新命名欄位
                if len(self.reference_data.columns) >= 8:
                    self.reference_data.columns = [
                        'Main Function', 'Interface', 'Interenal Error Code', 
                        'Description', 'Chinese', 'Version', 'Error Code', 'Note'
                    ]
            
            self.reference_file_path = file_path
            logger.info(f"成功載入參考資料: {file_path}")
            return True
        except Exception as e:
            logger.error(f"載入參考資料時發生錯誤: {str(e)}")
            return False
    
    def generate_recommendations(self, descriptions: List[str], ai_response: str = None) -> List[Tuple[str, str]]:
        """
        生成 AI 推薦
        
        Args:
            descriptions: Description 列表
            ai_response: AI 回應文字（如果提供，將直接解析）
            
        Returns:
            List[Tuple[str, str]]: 推薦的 Test ID 對列表
        """
        try:
            if ai_response:
                # 直接解析 AI 回應
                recommendations = self.prompt_templates.parse_ai_response(ai_response)
            else:
                # 使用內建邏輯生成推薦
                recommendations = self._generate_recommendations_internal(descriptions)
            
            if progress_callback:
                progress_callback(total, total, "AI 推薦分析完成")
            
            logger.info(f"成功生成 {len(recommendations)} 個 AI 推薦")
            return recommendations
        except Exception as e:
            logger.error(f"生成 AI 推薦時發生錯誤: {str(e)}")
            return []

    def generate_recommendations_with_search(self, descriptions: List[str], progress_callback=None) -> List[Tuple[str, str, str, str]]:
        """
        使用錯誤碼查詢邏輯生成推薦
        
        Args:
            descriptions: 描述列表
            progress_callback: 進度回調函數，格式為 callback(current, total, message)
            
        Returns:
            List[Tuple[str, str, str, str]]: 推薦的 (Test ID 1, Test ID 2, 中文描述 1, 中文描述 2)
        """
        recommendations = []
        total = len(descriptions)
        
        for i, description in enumerate(descriptions):
            if progress_callback:
                progress_callback(i, total, f"分析描述 {i+1}/{total}: {description[:30]}...")
            
            if not description or not description.strip():
                recommendations.append(("", "", "", ""))
                continue
                
            # 提取關鍵字
            keywords = self._extract_keywords(description)
            
            if not keywords:
                recommendations.append(("", "", "", ""))
                continue
            
            # 使用錯誤碼查詢邏輯搜尋
            matches = self._search_with_keywords(keywords)
            
            # 從搜尋結果中提取 Test ID 和中文描述
            test_data = self._extract_test_data_from_matches(matches)
            
            if len(test_data) >= 2:
                recommendations.append((test_data[0][0], test_data[1][0], test_data[0][1], test_data[1][1]))
            elif len(test_data) == 1:
                recommendations.append((test_data[0][0], "", test_data[0][1], ""))
            else:
                recommendations.append(("", "", "", ""))
        
        logger.info(f"使用搜尋邏輯生成 {len(recommendations)} 個推薦")
        return recommendations

    def _extract_keywords(self, description: str) -> List[str]:
        """
        從描述中提取關鍵字
        
        Args:
            description: 描述文字
            
        Returns:
            List[str]: 關鍵字列表
        """
        import re
        
        # 清理描述
        description = description.strip()
        if not description:
            return []
        
        # 移除特殊字符
        cleaned = re.sub(r'[^\w\s#-]', ' ', description)
        words = cleaned.lower().split()
        
        # 停用詞
        stop_words = {
            'the', 'a', 'an', 'and', 'or', 'but', 'in', 'on', 'at', 'to', 'for', 
            'of', 'with', 'by', 'is', 'are', 'was', 'were', 'be', 'been', 'being',
            'have', 'has', 'had', 'do', 'does', 'did', 'will', 'would', 'could',
            'should', 'may', 'might', 'can', 'must', 'shall', 'this', 'that',
            'these', 'those', 'i', 'you', 'he', 'she', 'it', 'we', 'they'
        }
        
        # 過濾停用詞和短詞
        keywords = [
            word.strip('.,!?()[]{}') 
            for word in words 
            if word not in stop_words and len(word) > 2
        ]
        
        # 如果沒有找到關鍵字，嘗試更寬鬆的條件
        if not keywords:
            keywords = [
                word.strip('.,!?()[]{}') 
                for word in words 
                if len(word) > 1  # 降低長度要求
            ]
        
        # 如果還是沒有，嘗試分割特殊字符
        if not keywords and description:
            # 處理像 "PC#-#Show SSN to UI" 這樣的描述
            import re
            # 分割特殊字符但保留重要部分
            parts = re.split(r'[#\-_]+', description.lower())
            keywords = [part.strip() for part in parts if part.strip() and len(part.strip()) > 1]
        
        return keywords[:3]  # 最多3個關鍵字

    def _search_with_keywords(self, keywords: List[str]) -> pd.DataFrame:
        """
        使用關鍵字搜尋參考資料
        
        Args:
            keywords: 關鍵字列表
            
        Returns:
            pd.DataFrame: 搜尋結果
        """
        if self.reference_data is None or self.reference_data.empty:
            return pd.DataFrame()
        
        # 如果沒有關鍵字，返回空結果
        if not keywords:
            return pd.DataFrame()
        
        # 使用 OR 條件搜尋（包含任一關鍵字即可）
        mask = self.reference_data.apply(
            lambda row: any(
                row.astype(str).str.contains(keyword, case=False, na=False).any() 
                for keyword in keywords
            ), axis=1
        )
        
        result = self.reference_data[mask]
        
        # 如果 OR 條件沒有結果，嘗試更寬鬆的搜尋
        if result.empty:
            # 嘗試部分匹配
            for keyword in keywords:
                partial_mask = self.reference_data.apply(
                    lambda row: any(
                        str(cell).lower().find(keyword.lower()) != -1
                        for cell in row if pd.notna(cell)
                    ), axis=1
                )
                if partial_mask.any():
                    result = self.reference_data[partial_mask]
                    break
        
        return result

    def _extract_test_data_from_matches(self, matches: pd.DataFrame) -> List[Tuple[str, str]]:
        """
        從搜尋結果中提取 Test ID 和中文描述
        
        Args:
            matches: 搜尋結果
            
        Returns:
            List[Tuple[str, str]]: (Test ID, 中文描述) 列表
        """
        test_data = []
        
        for _, row in matches.iterrows():
            test_id = self._get_test_id_from_row(row)
            chinese_desc = self._get_chinese_desc_from_row(row)
            
            if test_id and test_id not in [data[0] for data in test_data]:
                test_data.append((test_id, chinese_desc))
                if len(test_data) >= 2:
                    break
        
        return test_data
    
    def _extract_test_ids_from_matches(self, matches: pd.DataFrame) -> List[str]:
        """
        從搜尋結果中提取 Test ID（保留向後相容性）
        
        Args:
            matches: 搜尋結果
            
        Returns:
            List[str]: Test ID 列表
        """
        test_ids = []
        
        for _, row in matches.iterrows():
            test_id = self._get_test_id_from_row(row)
            if test_id and test_id not in test_ids:
                test_ids.append(test_id)
                if len(test_ids) >= 2:
                    break
        
        return test_ids
    
    def _generate_recommendations_internal(self, descriptions: List[str]) -> List[Tuple[str, str]]:
        """
        使用內建邏輯生成推薦（備用方案）
        
        Args:
            descriptions: Description 列表
            
        Returns:
            List[Tuple[str, str]]: 推薦的 Test ID 對列表
        """
        recommendations = []
        
        if self.reference_data is None:
            logger.warning("參考資料未載入，無法生成推薦")
            return []
        
        for desc in descriptions:
            # 簡單的關鍵字匹配邏輯
            best_matches = self._find_best_matches(desc)
            if len(best_matches) >= 2:
                recommendations.append((best_matches[0], best_matches[1]))
            elif len(best_matches) == 1:
                recommendations.append((best_matches[0], ""))
            else:
                recommendations.append(("", ""))
        
        return recommendations
    
    def _find_best_matches(self, description: str) -> List[str]:
        """
        為單一 Description 找到最佳匹配
        
        Args:
            description: 要匹配的 Description
            
        Returns:
            List[str]: 匹配的 Test ID 列表
        """
        if self.reference_data is None:
            return []
        
        matches = []
        desc_lower = description.lower()
        
        # 根據實際 Excel 結構搜尋
        # Excel 結構：Main Function, Interface, Interenal Error Code, Description, Chinese, Version, Error Code, Note
        logger.info(f"搜尋 Description: {description}")
        logger.info(f"參考資料欄位: {list(self.reference_data.columns)}")
        
        # 搜尋 Description 欄位
        if 'Description' in self.reference_data.columns:
            logger.info("使用 Description 欄位進行搜尋")
            
            for idx, row in self.reference_data.iterrows():
                if pd.notna(row['Description']) and str(row['Description']).strip():
                    ref_desc = str(row['Description']).lower()
                    # 檢查關鍵字匹配 - 改進匹配邏輯
                    keywords = [k for k in desc_lower.split() if len(k) > 2]  # 只考慮長度大於2的關鍵字
                    
                    # 計算匹配分數
                    match_score = 0
                    for keyword in keywords:
                        if keyword in ref_desc:
                            match_score += 1
                    
                    # 如果匹配分數大於0，或者包含重要關鍵字
                    important_keywords = ['fail', 'error', 'test', 'check', 'get', 'set', 'pc', 'dut']
                    has_important_keyword = any(kw in ref_desc for kw in important_keywords if kw in desc_lower)
                    
                    if match_score > 0 or has_important_keyword:
                        test_id = self._get_test_id_from_row(row)
                        if test_id:
                            matches.append(test_id)
                            logger.info(f"找到匹配: {ref_desc} -> {test_id}")
        
        # 搜尋中文描述欄位
        if 'Chinese' in self.reference_data.columns:
            logger.info("使用 Chinese 欄位進行搜尋")
            
            for idx, row in self.reference_data.iterrows():
                if pd.notna(row['Chinese']) and str(row['Chinese']).strip():
                    ref_desc = str(row['Chinese']).lower()
                    keywords = [k for k in desc_lower.split() if len(k) > 2]
                    
                    # 計算匹配分數
                    match_score = 0
                    for keyword in keywords:
                        if keyword in ref_desc:
                            match_score += 1
                    
                    # 如果匹配分數大於0，或者包含重要關鍵字
                    important_keywords = ['fail', 'error', 'test', 'check', 'get', 'set', 'pc', 'dut']
                    has_important_keyword = any(kw in ref_desc for kw in important_keywords if kw in desc_lower)
                    
                    if match_score > 0 or has_important_keyword:
                        test_id = self._get_test_id_from_row(row)
                        if test_id and test_id not in matches:
                            matches.append(test_id)
                            logger.info(f"找到中文匹配: {ref_desc} -> {test_id}")
        
        logger.info(f"總共找到 {len(matches)} 個匹配")
        return matches[:2]  # 最多返回 2 個匹配
    
    def _get_test_id_from_row(self, row) -> Optional[str]:
        """
        從資料行中提取 Test ID
        
        Args:
            row: 資料行
            
        Returns:
            Optional[str]: Test ID 或 None
        """
        # 根據實際 Excel 結構提取 Test ID
        # Excel 結構：Main Function, Interface, Interenal Error Code, Description, Chinese, Version, Error Code, Note
        # 我們需要的是 "Interenal Error Code" 或 "Error Code"
        
        # 優先使用 "Interenal Error Code"（內部錯誤代碼）
        if 'Interenal Error Code' in self.reference_data.columns:
            if pd.notna(row['Interenal Error Code']) and str(row['Interenal Error Code']).strip():
                test_id = str(row['Interenal Error Code']).strip()
                if test_id and test_id != 'nan' and test_id != 'Interenal Error Code':
                    logger.info(f"找到內部錯誤代碼: {test_id}")
                    return test_id
        
        # 如果沒有內部錯誤代碼，使用 "Error Code"
        if 'Error Code' in self.reference_data.columns:
            if pd.notna(row['Error Code']) and str(row['Error Code']).strip():
                test_id = str(row['Error Code']).strip()
                if test_id and test_id != 'nan' and test_id != 'Error Code':
                    logger.info(f"找到錯誤代碼: {test_id}")
                    return test_id
        
        return None
    
    def _get_chinese_desc_from_row(self, row) -> str:
        """
        從資料行中提取中文描述
        
        Args:
            row: 資料行
            
        Returns:
            str: 中文描述，如果沒有則返回空字串
        """
        # 根據實際 Excel 結構提取中文描述
        # Excel 結構：Main Function, Interface, Interenal Error Code, Description, Chinese, Version, Error Code, Note
        # 我們需要的是 "Chinese" 欄位（E 欄位）
        
        if 'Chinese' in self.reference_data.columns:
            if pd.notna(row['Chinese']) and str(row['Chinese']).strip():
                chinese_desc = str(row['Chinese']).strip()
                if chinese_desc and chinese_desc != 'nan' and chinese_desc != 'Chinese':
                    logger.info(f"找到中文描述: {chinese_desc}")
                    return chinese_desc
        
        return ""
    
    def get_prompt_for_descriptions(self, descriptions: List[str], prompt_type: str = "basic") -> str:
        """
        為 Description 列表生成 PROMPT
        
        Args:
            descriptions: Description 列表
            prompt_type: PROMPT 類型 ("basic", "single", "batch", "excel")
            
        Returns:
            str: 格式化的 PROMPT
        """
        if not self.reference_file_path:
            logger.warning("參考資料檔案路徑未設定")
            return ""
        
        if prompt_type == "basic":
            return self.prompt_templates.get_basic_analysis_prompt(descriptions, self.reference_file_path)
        elif prompt_type == "single" and len(descriptions) == 1:
            return self.prompt_templates.get_single_analysis_prompt(descriptions[0], self.reference_file_path)
        elif prompt_type == "batch":
            return self.prompt_templates.get_batch_analysis_prompt(descriptions, self.reference_file_path)
        elif prompt_type == "excel":
            return self.prompt_templates.get_excel_integration_prompt(descriptions, self.reference_file_path)
        else:
            return self.prompt_templates.get_basic_analysis_prompt(descriptions, self.reference_file_path)
    
    def validate_recommendations(self, recommendations: List[Tuple[str, str]], descriptions: List[str]) -> bool:
        """
        驗證推薦結果
        
        Args:
            recommendations: 推薦結果
            descriptions: 原始 Description 列表
            
        Returns:
            bool: 是否有效
        """
        if len(recommendations) != len(descriptions):
            logger.warning(f"推薦數量 ({len(recommendations)}) 與 Description 數量 ({len(descriptions)}) 不一致")
            return False
        
        # 檢查是否有空的推薦
        empty_count = sum(1 for rec in recommendations if not rec[0] and not rec[1])
        if empty_count > len(recommendations) * 0.5:  # 如果超過 50% 是空的
            logger.warning(f"過多空推薦 ({empty_count}/{len(recommendations)})")
            return False
        
        return True
    
    def get_recommendation_statistics(self, recommendations: List[Tuple[str, str]]) -> dict:
        """
        獲取推薦統計資訊
        
        Args:
            recommendations: 推薦結果
            
        Returns:
            dict: 統計資訊
        """
        total = len(recommendations)
        valid_1 = sum(1 for rec in recommendations if rec[0])
        valid_2 = sum(1 for rec in recommendations if rec[1])
        both_valid = sum(1 for rec in recommendations if rec[0] and rec[1])
        
        return {
            "total_recommendations": total,
            "valid_first_recommendation": valid_1,
            "valid_second_recommendation": valid_2,
            "both_valid": both_valid,
            "first_recommendation_rate": valid_1 / total if total > 0 else 0,
            "second_recommendation_rate": valid_2 / total if total > 0 else 0,
            "both_valid_rate": both_valid / total if total > 0 else 0
        }
