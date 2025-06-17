"""
seofrog/exporters/sheets/empty_headings_sheet.py
Aba específica para headings vazias e escondidas por CSS
"""

import pandas as pd
from typing import List, Dict
from .base_sheet import BaseSheet

class EmptyHeadingsSheet(BaseSheet):
    """Aba para headings vazias e escondidas - versão consolidada"""
    
    def get_sheet_name(self) -> str:
        return '🔍 Headings Vazias'
    
    def create_sheet(self, df: pd.DataFrame, writer) -> None:
        try:
            consolidated_problems = []
            
            for _, row in df.iterrows():
                url = row.get('url', '')
                empty_count = row.get('empty_headings_count', 0)
                hidden_count = row.get('hidden_headings_count', 0)
                
                # Só processa URLs que tenham problemas
                if empty_count > 0 or hidden_count > 0:
                    problem_data = self._process_heading_problems(row, url, empty_count, hidden_count)
                    if problem_data:
                        consolidated_problems.append(problem_data)
            
            if consolidated_problems:
                self._export_consolidated_problems(consolidated_problems, writer)
            else:
                self._create_success_sheet(writer, '✅ Nenhuma heading vazia ou escondida encontrada!')
                
        except Exception as e:
            self.logger.error(f"Erro criando aba de headings vazias: {e}")
            self._create_error_sheet(writer, f'Erro: {str(e)}')
    
    def _process_heading_problems(self, row, url: str, empty_count: int, hidden_count: int) -> Dict:
        """Processa problemas de headings para uma URL"""
        empty_by_level = {}
        hidden_by_level = {}
        hidden_methods = []
        
        # Processa detalhes das headings vazias
        empty_details = row.get('empty_headings_details', [])
        if empty_details and isinstance(empty_details, list):
            for empty_heading in empty_details:
                if isinstance(empty_heading, dict):
                    level = empty_heading.get('level', 'H?')
                    empty_by_level[level] = empty_by_level.get(level, 0) + 1
        
        # Processa detalhes das headings escondidas
        hidden_details = row.get('hidden_headings_details', [])
        if hidden_details and isinstance(hidden_details, list):
            methods_set = set()
            for hidden_heading in hidden_details:
                if isinstance(hidden_heading, dict):
                    level = hidden_heading.get('level', 'H?')
                    method = hidden_heading.get('css_issue', 'CSS')
                    hidden_by_level[level] = hidden_by_level.get(level, 0) + 1
                    methods_set.add(method)
            hidden_methods = list(methods_set)
        
        # FALLBACK: Usa resumos se não há detalhes
        if not empty_details and empty_count > 0:
            empty_by_level['H? (não especificado)'] = int(empty_count)
        
        if not hidden_details and hidden_count > 0:
            hidden_summary = row.get('hidden_headings_summary', '')
            if hidden_summary:
                self._parse_hidden_summary(hidden_summary, hidden_by_level, hidden_methods)
            else:
                hidden_by_level['H? (não especificado)'] = int(hidden_count)
                hidden_methods = ['CSS escondido']
        
        # Cria descrições consolidadas
        descriptions = []
        
        if empty_by_level:
            empty_parts = [f"{count} {level}" for level, count in empty_by_level.items()]
            descriptions.append(f"Vazias: {', '.join(empty_parts)}")
        
        if hidden_by_level:
            hidden_parts = [f"{count} {level}" for level, count in hidden_by_level.items()]
            methods_text = ', '.join(set(hidden_methods))
            descriptions.append(f"Escondidas: {', '.join(hidden_parts)} ({methods_text})")
        
        return {
            'url': url,
            'total_vazias': int(empty_count),
            'total_escondidas': int(hidden_count),
            'total_problemas': int(empty_count) + int(hidden_count),
            'detalhamento': ' | '.join(descriptions),
            'metodos_css': ', '.join(set(hidden_methods)) if hidden_methods else ''
        }
    
    def _parse_hidden_summary(self, hidden_summary: str, hidden_by_level: Dict, hidden_methods: List):
        """Parse do resumo de headings escondidas"""
        methods = hidden_summary.split(';')
        for method in methods:
            if ':' in method:
                level = method.split(':', 1)[0].strip()
                css_method = method.split(':', 1)[1].strip()
                hidden_by_level[level] = hidden_by_level.get(level, 0) + 1
                if css_method not in hidden_methods:
                    hidden_methods.append(css_method)
    
    def _export_consolidated_problems(self, problems: List[Dict], writer):
        """Exporta problemas consolidados"""
        problems_df = pd.DataFrame(problems)
        
        # Ordena por total de problemas (maior primeiro)
        problems_df = problems_df.sort_values('total_problemas', ascending=False)
        
        # Exporta
        problems_df.to_excel(writer, sheet_name=self.get_sheet_name(), index=False)
        
        # Log estatísticas
        total_urls = len(problems_df)
        total_issues = problems_df['total_problemas'].sum()
        self.logger.info(f"✅ {self.get_sheet_name()}: {total_urls} URLs com {total_issues} problemas consolidados")