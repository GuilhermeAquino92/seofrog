"""
seofrog/exporters/sheets/title_problems_sheet.py
Aba específica para problemas de títulos
"""

import pandas as pd
from .base_sheet import BaseSheet

class TitleProblemsSheet(BaseSheet):
    """Aba com problemas de títulos (ausentes, longos, curtos)"""
    
    def get_sheet_name(self) -> str:
        return 'Problemas Títulos'
    
    def create_sheet(self, df: pd.DataFrame, writer) -> None:
        try:
            title_issues = []
            
            # URLs sem título
            no_title_issues = self._find_no_title_issues(df)
            title_issues.extend(no_title_issues)
            
            # Títulos muito longos
            long_title_issues = self._find_long_title_issues(df)
            title_issues.extend(long_title_issues)
            
            # Títulos muito curtos
            short_title_issues = self._find_short_title_issues(df)
            title_issues.extend(short_title_issues)
            
            # Títulos duplicados
            duplicate_title_issues = self._find_duplicate_title_issues(df)
            title_issues.extend(duplicate_title_issues)
            
            if title_issues:
                self._export_consolidated_issues(title_issues, writer)
            else:
                self._create_success_sheet(writer, '✅ Nenhum problema de título encontrado!')
                
        except Exception as e:
            self.logger.error(f"Erro criando aba de problemas de título: {e}")
            self._create_error_sheet(writer, f'Erro: {str(e)}')
    
    def _find_no_title_issues(self, df: pd.DataFrame) -> list:
        """Encontra URLs sem título"""
        if 'title' not in df.columns:
            return []
        
        no_title = self._safe_filter(df, 'title', df['title'].fillna('') == '')
        issues = []
        
        for _, row in no_title.iterrows():
            issues.append({
                'url': row.get('url', ''),
                'problema': 'Sem título',
                'criticidade': 'CRÍTICO',
                'title': '',
                'title_length': 0,
                'title_words': 0,
                'recomendacao': 'Adicionar título único e descritivo (30-60 chars)',
                'impacto_seo': 'Muito alto - Título é fundamental para SEO'
            })
        
        return issues
    
    def _find_long_title_issues(self, df: pd.DataFrame) -> list:
        """Encontra títulos muito longos (>60 chars)"""
        if 'title_length' not in df.columns:
            return []
        
        long_titles = self._safe_filter(df, 'title_length', 
            df['title_length'].fillna(0) > 60)
        issues = []
        
        for _, row in long_titles.iterrows():
            title_length = row.get('title_length', 0)
            excess_chars = title_length - 60
            
            issues.append({
                'url': row.get('url', ''),
                'problema': f'Título muito longo ({title_length} chars)',
                'criticidade': 'ALTO' if title_length > 70 else 'MÉDIO',
                'title': row.get('title', ''),
                'title_length': title_length,
                'title_words': row.get('title_words', 0),
                'recomendacao': f'Reduzir em {excess_chars} caracteres (ideal: 30-60)',
                'impacto_seo': 'Google pode truncar na SERP'
            })
        
        return issues
    
    def _find_short_title_issues(self, df: pd.DataFrame) -> list:
        """Encontra títulos muito curtos (<30 chars)"""
        if 'title_length' not in df.columns or 'title' not in df.columns:
            return []
        
        short_titles = self._safe_filter(df, 'title_length',
            (df['title_length'].fillna(100) < 30) & 
            (df['title'].fillna('') != ''))  # Não considera vazios
        issues = []
        
        for _, row in short_titles.iterrows():
            title_length = row.get('title_length', 0)
            needed_chars = 30 - title_length
            
            issues.append({
                'url': row.get('url', ''),
                'problema': f'Título muito curto ({title_length} chars)',
                'criticidade': 'MÉDIO',
                'title': row.get('title', ''),
                'title_length': title_length,
                'title_words': row.get('title_words', 0),
                'recomendacao': f'Adicionar {needed_chars} caracteres (ideal: 30-60)',
                'impacto_seo': 'Pouco aproveitamento do espaço na SERP'
            })
        
        return issues
    
    def _find_duplicate_title_issues(self, df: pd.DataFrame) -> list:
        """Encontra títulos duplicados"""
        if 'title' not in df.columns:
            return []
        
        # Filtra apenas títulos não vazios
        non_empty_titles = df[df['title'].fillna('') != ''].copy()
        if non_empty_titles.empty:
            return []
        
        # Encontra duplicatas
        duplicate_titles = non_empty_titles[
            non_empty_titles.duplicated(subset=['title'], keep=False)
        ].copy()
        
        issues = []
        
        # Agrupa por título duplicado
        for title, group in duplicate_titles.groupby('title'):
            urls_with_same_title = group['url'].tolist()
            duplicate_count = len(urls_with_same_title)
            
            for _, row in group.iterrows():
                issues.append({
                    'url': row.get('url', ''),
                    'problema': f'Título duplicado ({duplicate_count} páginas)',
                    'criticidade': 'ALTO' if duplicate_count > 3 else 'MÉDIO',
                    'title': title,
                    'title_length': row.get('title_length', 0),
                    'title_words': row.get('title_words', 0),
                    'recomendacao': 'Criar título único para cada página',
                    'impacto_seo': 'Google pode não indexar todas as páginas'
                })
        
        return issues
    
    def _export_consolidated_issues(self, issues: list, writer):
        """Exporta problemas consolidados com ordenação por criticidade"""
        if not issues:
            return
        
        issues_df = pd.DataFrame(issues)
        
        # Remove duplicatas (mesma URL pode ter múltiplos problemas)
        issues_df = issues_df.drop_duplicates(subset=['url', 'problema'], keep='first')
        
        # Ordena por criticidade e URL
        issues_df = self._sort_by_criticality(issues_df)
        
        # Define colunas para exportar
        columns = [
            'url', 'problema', 'criticidade', 'title', 'title_length', 
            'title_words', 'recomendacao', 'impacto_seo'
        ]
        
        self._export_dataframe(issues_df, writer, columns)
        
        # Log estatísticas
        total_issues = len(issues_df)
        critical_issues = len(issues_df[issues_df['criticidade'] == 'CRÍTICO'])
        high_issues = len(issues_df[issues_df['criticidade'] == 'ALTO'])
        
        self.logger.info(f"✅ {self.get_sheet_name()}: {total_issues} problemas "
                        f"({critical_issues} críticos, {high_issues} altos)")