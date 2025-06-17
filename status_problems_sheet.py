"""
seofrog/exporters/sheets/status_problems_sheet.py
Aba de problemas de status HTTP
"""

import pandas as pd
from .base_sheet import BaseSheet

class StatusProblemsSheet(BaseSheet):
    """Aba com problemas de status HTTP (não 200)"""
    
    def get_sheet_name(self) -> str:
        return 'Erros HTTP'
    
    def create_sheet(self, df: pd.DataFrame, writer) -> None:
        try:
            if 'status_code' not in df.columns:
                self._create_error_sheet(writer, 'Coluna status_code não encontrada')
                return
            
            # Filtra problemas de status
            status_problems = self._safe_filter(df, 'status_code', 
                df['status_code'].fillna(0) != 200)
            
            if status_problems.empty:
                self._create_success_sheet(writer, '✅ Nenhum erro HTTP encontrado!')
                return
            
            # Ordena por status code
            status_problems = status_problems.sort_values('status_code')
            
            # Define colunas para exportar
            columns = ['url', 'status_code', 'final_url', 'response_time']
            self._export_dataframe(status_problems, writer, columns)
            
        except Exception as e:
            self.logger.error(f"Erro criando aba de erros HTTP: {e}")
            self._create_error_sheet(writer, f'Erro: {str(e)}')