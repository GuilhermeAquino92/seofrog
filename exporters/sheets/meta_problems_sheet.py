"""
seofrog/exporters/sheets/meta_problems_sheet.py
Aba específica para problemas de meta description
"""

import pandas as pd
from .base_sheet import BaseSheet

class MetaProblemsSheet(BaseSheet):
    """Aba com problemas de meta description (ausentes, longas, curtas, duplicadas)"""
    
    def get_sheet_name(self) -> str:
        return 'Problemas Meta'
    
    def create_sheet(self, df: pd.DataFrame, writer) -> None:
        try:
            meta_issues = []
            
            # URLs sem meta description
            no_meta_issues = self._find_no_meta_issues(df)
            meta_issues.extend(no_meta_issues)
            
            # Meta descriptions muito longas
            long_meta_issues = self._find_long_meta_issues(df)
            meta_issues.extend(long_meta_issues)
            
            # Meta descriptions muito curtas
            short_meta_issues = self._find_short_meta_issues(df)
            meta_issues.extend(short_meta_issues)
            
            # Meta descriptions duplicadas
            duplicate_meta_issues = self._find_duplicate_meta_issues(df)
            meta_issues.extend(duplicate_meta_issues)
            
            if meta_issues:
                self._export_consolidated_issues(meta_issues, writer)
            else:
                self._create_success_sheet(writer, '✅ Nenhum problema de meta description!')
                
        except Exception as e:
            self.logger.error(f"Erro criando aba de problemas de meta: {e}")
            self._create_error_sheet(writer, f'Erro: {str(e)}')
    
    def _find_no_meta_issues(self, df: pd.DataFrame) -> list:
        """Encontra URLs sem meta description"""
        if 'meta_description' not in df.columns:
            return []
        
        no_meta = self._safe_filter(df, 'meta_description', 
            df['meta_description'].fillna('') == '')
        issues = []
        
        for _, row in no_meta.iterrows():
            issues.append({
                'url': row.get('url', ''),
                'problema': 'Sem meta description',
                'criticidade': 'ALTO',
                'meta_description': '',
                'meta_description_length': 0,
                'recomendacao': 'Adicionar meta description única (120-160 chars)',
                'impacto_seo': 'Google pode gerar snippet automaticamente',
                'exemplo': 'Descrição atrativa que resume o conteúdo da página'
            })
        
        return issues
    
    def _find_long_meta_issues(self, df: pd.DataFrame) -> list:
        """Encontra meta descriptions muito longas (>160 chars)"""
        if 'meta_description_length' not in df.columns:
            return []
        
        long_meta = self._safe_filter(df, 'meta_description_length',
            df['meta_description_length'].fillna(0) > 160)
        issues = []
        
        for _, row in long_meta.iterrows():
            meta_length = row.get('meta_description_length', 0)
            excess_chars = meta_length - 160
            
            issues.append({
                'url': row.get('url', ''),
                'problema': f'Meta description muito longa ({meta_length} chars)',
                'criticidade': 'MÉDIO' if meta_length < 180 else 'ALTO',
                'meta_description': row.get('meta_description', ''),
                'meta_description_length': meta_length,
                'recomendacao': f'Reduzir em {excess_chars} caracteres (máximo: 160)',
                'impacto_seo': 'Google truncará na SERP com "..."',
                'exemplo': f'Resumir mantendo as palavras-chave principais'
            })
        
        return issues
    
    def _find_short_meta_issues(self, df: pd.DataFrame) -> list:
        """Encontra meta descriptions muito curtas (<120 chars)"""
        if 'meta_description_length' not in df.columns or 'meta_description' not in df.columns:
            return []
        
        short_meta = self._safe_filter(df, 'meta_description_length',
            (df['meta_description_length'].fillna(200) < 120) & 
            (df['meta_description'].fillna('') != ''))  # Não considera vazias
        issues = []
        
        for _, row in short_meta.iterrows():
            meta_length = row.get('meta_description_length', 0)
            needed_chars = 120 - meta_length
            
            issues.append({
                'url': row.get('url', ''),
                'problema': f'Meta description muito curta ({meta_length} chars)',
                'criticidade': 'BAIXO',
                'meta_description': row.get('meta_description', ''),
                'meta_description_length': meta_length,
                'recomendacao': f'Expandir em {needed_chars} caracteres (ideal: 120-160)',
                'impacto_seo': 'Pouco aproveitamento do espaço na SERP',
                'exemplo': 'Adicionar mais detalhes sobre o conteúdo'
            })
        
        return issues
    
    def _find_duplicate_meta_issues(self, df: pd.DataFrame) -> list:
        """Encontra meta descriptions duplicadas"""
        if 'meta_description' not in df.columns:
            return []
        
        # Filtra apenas meta descriptions não vazias
        non_empty_meta = df[df['meta_description'].fillna('') != ''].copy()
        if non_empty_meta.empty:
            return []
        
        # Encontra duplicatas
        duplicate_meta = non_empty_meta[
            non_empty_meta.duplicated(subset=['meta_description'], keep=False)
        ].copy()
        
        issues = []
        
        # Agrupa por meta description duplicada
        for meta_desc, group in duplicate_meta.groupby('meta_description'):
            urls_with_same_meta = group['url'].tolist()
            duplicate_count = len(urls_with_same_meta)
            
            for _, row in group.iterrows():
                issues.append({
                    'url': row.get('url', ''),
                    'problema': f'Meta description duplicada ({duplicate_count} páginas)',
                    'criticidade': 'ALTO' if duplicate_count > 2 else 'MÉDIO',
                    'meta_description': meta_desc,
                    'meta_description_length': row.get('meta_description_length', 0),
                    'recomendacao': 'Criar meta description única para cada página',
                    'impacto_seo': 'Google pode não distinguir o conteúdo das páginas',
                    'exemplo': 'Destacar o diferencial específico de cada página'
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
            'url', 'problema', 'criticidade', 'meta_description', 
            'meta_description_length', 'recomendacao', 'impacto_seo', 'exemplo'
        ]
        
        self._export_dataframe(issues_df, writer, columns)
        
        # Log estatísticas
        total_issues = len(issues_df)
        high_issues = len(issues_df[issues_df['criticidade'] == 'ALTO'])
        medium_issues = len(issues_df[issues_df['criticidade'] == 'MÉDIO'])
        
        self.logger.info(f"✅ {self.get_sheet_name()}: {total_issues} problemas "
                        f"({high_issues} altos, {medium_issues} médios)")