"""
seofrog/exporters/sheets/h1_h2_missing_sheet.py
Aba específica para URLs sem H1 e/ou H2 - Análise focada
"""

import pandas as pd
from .base_sheet import BaseSheet

class H1H2MissingSheet(BaseSheet):
    """Aba específica para análise focada de H1/H2 ausentes"""
    
    def get_sheet_name(self) -> str:
        return 'H1 H2 Ausentes'
    
    def create_sheet(self, df: pd.DataFrame, writer) -> None:
        try:
            h1_h2_issues = []
            
            # URLs sem H1 (crítico máximo)
            no_h1_issues = self._find_no_h1_issues(df)
            h1_h2_issues.extend(no_h1_issues)
            
            # URLs sem H2 (alto)
            no_h2_issues = self._find_no_h2_issues(df)
            h1_h2_issues.extend(no_h2_issues)
            
            # URLs sem H1 NEM H2 (crítico máximo)
            no_h1_h2_issues = self._find_no_h1_h2_issues(df)
            h1_h2_issues.extend(no_h1_h2_issues)
            
            if h1_h2_issues:
                self._export_consolidated_issues(h1_h2_issues, writer)
            else:
                self._create_success_sheet(writer, '✅ Estrutura de H1/H2 adequada!')
                
        except Exception as e:
            self.logger.error(f"Erro criando aba H1/H2 ausentes: {e}")
            self._create_error_sheet(writer, f'Erro: {str(e)}')
    
    def _find_no_h1_issues(self, df: pd.DataFrame) -> list:
        """Encontra URLs sem H1 - CRÍTICO"""
        if 'h1_count' not in df.columns:
            return []
        
        no_h1 = self._safe_filter(df, 'h1_count', df['h1_count'].fillna(0) == 0)
        issues = []
        
        for _, row in no_h1.iterrows():
            issues.append({
                'url': row.get('url', ''),
                'problema_h1_h2': 'Sem H1',
                'criticidade': 'CRÍTICO',
                'prioridade': 1,
                'h1_count': 0,
                'h1_text': '',
                'h2_count': row.get('h2_count', 0),
                'title': row.get('title', ''),
                'page_type': self._detect_page_type(row),
                'action_required': 'URGENTE: Adicionar H1 único',
                'business_impact': 'SEO comprometido - Página não será bem rankeada',
                'technical_fix': '<h1>Título Principal da Página</h1>'
            })
        
        return issues
    
    def _find_no_h2_issues(self, df: pd.DataFrame) -> list:
        """Encontra URLs sem H2 - ALTO"""
        if 'h2_count' not in df.columns:
            return []
        
        no_h2 = self._safe_filter(df, 'h2_count', df['h2_count'].fillna(0) == 0)
        issues = []
        
        for _, row in no_h2.iterrows():
            # Só considera se tem H1 (senão é problema do H1)
            h1_count = row.get('h1_count', 0)
            if h1_count > 0:
                issues.append({
                    'url': row.get('url', ''),
                    'problema_h1_h2': 'Sem H2',
                    'criticidade': 'ALTO',
                    'prioridade': 2,
                    'h1_count': h1_count,
                    'h1_text': row.get('h1_text', ''),
                    'h2_count': 0,
                    'title': row.get('title', ''),
                    'page_type': self._detect_page_type(row),
                    'action_required': 'Adicionar H2s para estruturar conteúdo',
                    'business_impact': 'Estrutura de conteúdo confusa para usuários',
                    'technical_fix': '<h2>Seção Principal</h2> + subsections'
                })
        
        return issues
    
    def _find_no_h1_h2_issues(self, df: pd.DataFrame) -> list:
        """Encontra URLs sem H1 NEM H2 - CRÍTICO MÁXIMO"""
        if 'h1_count' not in df.columns or 'h2_count' not in df.columns:
            return []
        
        no_h1_h2 = self._safe_filter(df, 'h1_count',
            (df['h1_count'].fillna(0) == 0) & (df['h2_count'].fillna(0) == 0))
        issues = []
        
        for _, row in no_h1_h2.iterrows():
            issues.append({
                'url': row.get('url', ''),
                'problema_h1_h2': 'Sem H1 E sem H2',
                'criticidade': 'CRÍTICO MÁXIMO',
                'prioridade': 0,
                'h1_count': 0,
                'h1_text': '',
                'h2_count': 0,
                'title': row.get('title', ''),
                'page_type': self._detect_page_type(row),
                'action_required': 'URGENTÍSSIMO: Estrutura de headings completa',
                'business_impact': 'Página invisível para SEO - Zero estrutura',
                'technical_fix': 'Implementar hierarquia completa H1>H2>H3'
            })
        
        return issues
    
    def _detect_page_type(self, row) -> str:
        """Detecta tipo da página baseado na URL"""
        url = row.get('url', '').lower()
        
        if '/blog/' in url or '/artigo/' in url:
            return 'Blog/Artigo'
        elif '/produto/' in url or '/product/' in url:
            return 'Produto'
        elif '/categoria/' in url or '/category/' in url:
            return 'Categoria'
        elif '/sobre' in url or '/about' in url:
            return 'Institucional'
        elif '/contato' in url or '/contact' in url:
            return 'Contato'
        elif url.endswith('/') and url.count('/') <= 3:
            return 'Homepage'
        else:
            return 'Conteúdo'
    
    def _export_consolidated_issues(self, issues: list, writer):
        """Exporta problemas com foco em ação executiva"""
        if not issues:
            return
        
        issues_df = pd.DataFrame(issues)
        
        # Remove duplicatas por URL (mantém o mais crítico)
        issues_df = issues_df.drop_duplicates(subset=['url'], keep='first')
        
        # Ordena por prioridade (crítico máximo primeiro)
        issues_df = issues_df.sort_values(['prioridade', 'url'])
        
        # Define colunas para exportar
        columns = [
            'url', 'problema_h1_h2', 'criticidade', 'h1_count', 'h2_count', 
            'h1_text', 'title', 'page_type', 'action_required', 
            'business_impact', 'technical_fix'
        ]
        
        self._export_dataframe(issues_df, writer, columns)
        
        # Log estatísticas detalhadas
        total_issues = len(issues_df)
        critical_max = len(issues_df[issues_df['criticidade'] == 'CRÍTICO MÁXIMO'])
        critical = len(issues_df[issues_df['criticidade'] == 'CRÍTICO'])
        high = len(issues_df[issues_df['criticidade'] == 'ALTO'])
        
        self.logger.info(f"✅ {self.get_sheet_name()}: {total_issues} URLs problemáticas "
                        f"({critical_max} crítico máximo, {critical} críticos, {high} altos)")
        
        # Log por tipo de página
        page_types = issues_df['page_type'].value_counts()
        for page_type, count in page_types.items():
            self.logger.info(f"   📄 {page_type}: {count} páginas")