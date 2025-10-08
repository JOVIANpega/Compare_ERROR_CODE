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
            
            logger.info(f"成功生成 {len(recommendations)} 個 AI 推薦")
            return recommendations
        except Exception as e:
            logger.error(f"生成 AI 推薦時發生錯誤: {str(e)}")
            return []
    
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
                if test_id and test_id != 'nan':
                    logger.info(f"找到內部錯誤代碼: {test_id}")
                    return test_id
        
        # 如果沒有內部錯誤代碼，使用 "Error Code"
        if 'Error Code' in self.reference_data.columns:
            if pd.notna(row['Error Code']) and str(row['Error Code']).strip():
                test_id = str(row['Error Code']).strip()
                if test_id and test_id != 'nan':
                    logger.info(f"找到錯誤代碼: {test_id}")
                    return test_id
        
        return None
    
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
