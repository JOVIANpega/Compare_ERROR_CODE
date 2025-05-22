"""
Excel處理模組
負責處理所有Excel相關的操作，包括讀取、比對和寫入Excel檔案
"""
import os
import logging
import pandas as pd
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill
from typing import Tuple, Dict, Optional

logger = logging.getLogger(__name__)

class ExcelHandler:
    def __init__(self):
        self.error_code_map: Dict[str, Tuple[str, str]] = {}
        self.current_sheet: Optional[str] = None

    def load_error_codes(self, file_path: str) -> bool:
        """載入錯誤碼Excel檔案"""
        try:
            df_error_codes = pd.read_excel(file_path, sheet_name="Test Item All")
            self.error_code_map = {
                str(k).strip(): (str(v1).strip(), str(v2).strip())
                for k, v1, v2 in zip(df_error_codes.iloc[:, 2], df_error_codes.iloc[:, 3], df_error_codes.iloc[:, 4])
            }
            logger.info(f"成功載入錯誤碼檔案: {file_path}")
            return True
        except Exception as e:
            logger.error(f"載入錯誤碼檔案時發生錯誤: {str(e)}")
            return False

    def load_source_sheet(self, file_path: str, sheet_name: str) -> Optional[pd.DataFrame]:
        """載入來源Excel檔案的工作表"""
        try:
            df_source = pd.read_excel(file_path, sheet_name=sheet_name)
            self.current_sheet = sheet_name
            logger.info(f"成功載入來源工作表: {sheet_name}")
            return df_source
        except Exception as e:
            logger.error(f"載入來源工作表時發生錯誤: {str(e)}")
            return None

    def find_column(self, df: pd.DataFrame, target: str) -> Optional[str]:
        """在DataFrame中尋找目標欄位"""
        for col in df.columns:
            if str(col).strip().lower() == target.lower():
                return col
        return None

    def compare_data(self, df_source: pd.DataFrame, not_found_text: str, not_found_cn_text: str) -> Optional[pd.DataFrame]:
        """比對資料"""
        try:
            desc_col = self.find_column(df_source, 'Description')
            testid_col = self.find_column(df_source, 'TestID')

            if not desc_col or not testid_col:
                logger.error(f"找不到必要欄位，實際欄位: {df_source.columns.tolist()}")
                return None

            df_result = df_source[[desc_col, testid_col]].copy()
            df_result.columns = ['AB', 'AC']

            # 過濾空值
            df_result = df_result[
                df_result['AB'].notna() & 
                df_result['AC'].notna() & 
                (df_result['AB'].astype(str).str.strip() != '') & 
                (df_result['AC'].astype(str).str.strip() != '')
            ].reset_index(drop=True)

            # 比對錯誤碼
            for idx, row in df_result.iterrows():
                test_id = str(row['AC']).strip()
                if test_id in self.error_code_map:
                    description, chinese_desc = self.error_code_map[test_id]
                    df_result.at[idx, 'CD'] = description
                    df_result.at[idx, 'CE'] = chinese_desc
                else:
                    df_result.at[idx, 'CD'] = not_found_text
                    df_result.at[idx, 'CE'] = not_found_cn_text

            logger.info("成功完成資料比對")
            return df_result

        except Exception as e:
            logger.error(f"比對資料時發生錯誤: {str(e)}")
            return None

    def save_result(self, df_result: pd.DataFrame, df_error_codes: pd.DataFrame, 
                   output_path: str, sheet_name: str) -> bool:
        """儲存比對結果"""
        try:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                df_result.to_excel(writer, index=False, sheet_name=sheet_name)
                df_error_codes.to_excel(writer, index=False, sheet_name='Test Item All')

            self._format_excel(output_path)
            logger.info(f"成功儲存比對結果: {output_path}")
            return True

        except Exception as e:
            logger.error(f"儲存比對結果時發生錯誤: {str(e)}")
            return False

    def _format_excel(self, file_path: str):
        """格式化Excel檔案"""
        try:
            wb = load_workbook(file_path)
            calibri_font = Font(name='Calibri', size=11)
            bold_font = Font(name='Calibri', size=11, bold=True)
            thin = Side(border_style="thin", color="000000")
            border = Border(left=thin, right=thin, top=thin, bottom=thin)

            for ws in wb.worksheets:
                # 設定標題列樣式
                for cell in ws[1]:
                    cell.font = bold_font
                    cell.border = border
                    cell.alignment = Alignment(horizontal='center', vertical='center')

                # 設定資料列樣式
                for row in ws.iter_rows(min_row=2):
                    for cell in row:
                        cell.font = calibri_font
                        cell.border = border
                        cell.alignment = Alignment(vertical='center')

                # 自動調整欄寬
                for column in ws.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = (max_length + 2)
                    ws.column_dimensions[column_letter].width = adjusted_width

            wb.save(file_path)
            logger.info("成功格式化Excel檔案")

        except Exception as e:
            logger.error(f"格式化Excel檔案時發生錯誤: {str(e)}")

    def get_sheet_names(self, file_path: str) -> list:
        """獲取Excel檔案中的所有工作表名稱"""
        try:
            excel_file = pd.ExcelFile(file_path)
            return excel_file.sheet_names
        except Exception as e:
            logger.error(f"獲取工作表名稱時發生錯誤: {str(e)}")
            return [] 