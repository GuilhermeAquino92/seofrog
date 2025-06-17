"""
seofrog/exporters/sheets/base_sheet.py
Classe base para todas as abas do Excel
"""

from abc import ABC, abstractmethod
import pandas as pd
from typing import Dict, Any, List
from seofrog.utils.logger import get_logger

class BaseSheet(ABC):
    """Classe base para todas as abas do Excel"""
    
    def __init__(self):
        self.logger = get_logger(self.__class__.__name__)
    
    @abstractmethod
    def get_sheet_name(self) -> str:
        """Retorna o nome da aba"""
        pass
    
    @abstractmethod
    def create_sheet(self, df: pd.DataFrame, writer) -> None:
        """Cria a aba no Excel"""
        pass
    
    def _safe_filter(self, df: pd.DataFrame, column: str, condition) -> pd.DataFrame:
        """Filtra DataFrame de forma segura"""
        if column not in df.columns:
            return pd.DataFrame()
        try:
            return df[condition].copy()
        except Exception as e:
            self.logger.warning(f"Erro filtrando coluna {column}: {e}")
            return pd.DataFrame()
    
    def _safe_get_column(self, df: pd.DataFrame, column: str, default_value=0):
        """Obtém coluna de forma segura"""
        if column in df.columns:
            return df[column].fillna(default_value)
        else:
            return pd.Series([default_value] * len(df), name=column)
    
    def _create_error_sheet(self, writer, error_msg: str):
        """Cria aba de erro padronizada"""
        try:
            error_df = pd.DataFrame([[error_msg]], columns=['Erro'])
            error_df.to_excel(writer, sheet_name=self.get_sheet_name(), index=False)
        except Exception as e:
            self.logger.error(f"Erro criando sheet de erro: {e}")
    
    def _create_success_sheet(self, writer, success_msg: str):
        """Cria aba de sucesso padronizada"""
        try:
            success_df = pd.DataFrame([[success_msg]], columns=['Status'])
            success_df.to_excel(writer, sheet_name=self.get_sheet_name(), index=False)
        except Exception as e:
            self.logger.error(f"Erro criando sheet de sucesso: {e}")
    
    def _export_dataframe(self, df: pd.DataFrame, writer, columns: List[str] = None):
        """Exporta DataFrame para Excel de forma segura"""
        try:
            if df.empty:
                self._create_success_sheet(writer, f"✅ Nenhum problema encontrado na categoria {self.get_sheet_name()}!")
                return
            
            # Filtra colunas disponíveis
            if columns:
                available_cols = [col for col in columns if col in df.columns]
                if available_cols:
                    df = df[available_cols]
                else:
                    self._create_error_sheet(writer, f"Colunas necessárias não encontradas: {columns}")
                    return
            
            # Remove duplicatas
            if 'url' in df.columns:
                df = df.drop_duplicates(subset=['url'], keep='first')
            
            # Exporta
            df.to_excel(writer, sheet_name=self.get_sheet_name(), index=False)
            
            # Log
            self.logger.info(f"✅ {self.get_sheet_name()}: {len(df)} registros exportados")
            
        except Exception as e:
            self.logger.error(f"Erro exportando DataFrame: {e}")
            self._create_error_sheet(writer, f"Erro na exportação: {str(e)}")
    
    def _add_problem_column(self, df: pd.DataFrame, problem_name: str, criticality: str = 'MÉDIO') -> pd.DataFrame:
        """Adiciona colunas de problema e criticidade"""
        df['problema'] = problem_name
        df['criticidade'] = criticality
        return df
    
    def _sort_by_criticality(self, df: pd.DataFrame) -> pd.DataFrame:
        """Ordena DataFrame por criticidade"""
        if 'criticidade' not in df.columns:
            return df
        
        criticality_order = {
            'CRÍTICO MÁXIMO': 0,
            'CRÍTICO': 1,
            'ALTO': 2,
            'MÉDIO': 3,
            'BAIXO': 4
        }
        
        df['_sort_order'] = df['criticidade'].map(criticality_order).fillna(5)
        df = df.sort_values(['_sort_order', 'url']).drop('_sort_order', axis=1)
        return df