"""
seofrog/exporters/sheets/heading_problems_sheet.py
Aba específica para problemas gerais de headings
"""

import pandas as pd
from .base_sheet import BaseSheet

class HeadingProblemsSheet(BaseSheet):
    """Aba com problemas gerais de headings (estrutura hierárquica)"""
    
    def get_sheet_name(self) -> str:
        return 'Problemas Headings'
    
    def create_sheet(self, df: pd.DataFrame, writer) -> None:
        try:
            heading_issues = []
            
            # URLs sem H1
            no_h1_issues = self._find_no_h1_issues(df)
            heading_issues.extend(no_h1_issues)
            
            # URLs com múltiplos H1
            multiple_h1_issues = self._find_multiple_h1_issues(df)
            heading_issues.extend(multiple_h1_issues)
            
            # URLs sem estrutura de headings (sem H2)
            no_h2_issues = self._find_no_h2_issues(df)
            heading_issues.extend(no_h2_issues)
            
            # Problemas de hierarquia (H3 sem H2, etc.)
            hierarchy_issues = self._find_hierarchy_issues(df)
            heading_issues.extend(hierarchy_issues)
            
            # H1 muito longo
            long_h1_issues = self._find_long_h1_issues(df)
            heading_issues.extend(long_h1_issues)
            
            if heading_issues:
                self._export_consolidated_issues(heading_issues, writer)
            else:
                self._create_success_sheet(writer, '✅ Estrutura de headings adequada!')
                
        except Exception as e:
            self.logger.error(f"Erro criando aba de problemas de headings: {e}")
            self._create_error_sheet(writer, f'Erro: {str(e)}')
    
    def _find_no_h1_issues(self, df: pd.DataFrame) -> list:
        """Encontra URLs sem H1"""
        if 'h1_count' not in df.columns:
            return []
        
        no_h1 = self._safe_filter(df, 'h1_count', df['h1_count'].fillna(0) == 0)
        issues = []
        
        for _, row in no_h1.iterrows():
            issues.append({
                'url': row.get('url', ''),
                'problema': 'Sem H1',
                'criticidade': 'CRÍTICO',
                'h1_count': 0,
                'h1_text': '',
                'h2_count': row.get('h2_count', 0),
                'h3_count': row.get('h3_count', 0),
                'estrutura_headings': self._get_heading_structure(row),
                'recomendacao': 'Adicionar H1 único e descritivo para a página',
                'impacto_seo': 'Muito alto - H1 é fundamental para SEO'
            })
        
        return issues
    
    def _find_multiple_h1_issues(self, df: pd.DataFrame) -> list:
        """Encontra URLs com múltiplos H1"""
        if 'h1_count' not in df.columns:
            return []
        
        multiple_h1 = self._safe_filter(df, 'h1_count', df['h1_count'].fillna(0) > 1)
        issues = []
        
        for _, row in multiple_h1.iterrows():
            h1_count = row.get('h1_count', 0)
            
            issues.append({
                'url': row.get('url', ''),
                'problema': f'Múltiplos H1 ({h1_count} H1s)',
                'criticidade': 'ALTO',
                'h1_count': h1_count,
                'h1_text': row.get('h1_text', ''),
                'h2_count': row.get('h2_count', 0),
                'h3_count': row.get('h3_count', 0),
                'estrutura_headings': self._get_heading_structure(row),
                'recomendacao': f'Manter apenas 1 H1, converter outros {h1_count-1} para H2',
                'impacto_seo': 'Confunde hierarquia e importância do conteúdo'
            })
        
        return issues
    
    def _find_no_h2_issues(self, df: pd.DataFrame) -> list:
        """Encontra URLs sem H2 (estrutura de conteúdo)"""
        if 'h2_count' not in df.columns:
            return []
        
        no_h2 = self._safe_filter(df, 'h2_count', df['h2_count'].fillna(0) == 0)
        issues = []
        
        for _, row in no_h2.iterrows():
            # Só considera problema se tem H1 mas não tem H2
            h1_count = row.get('h1_count', 0)
            if h1_count > 0:
                issues.append({
                    'url': row.get('url', ''),
                    'problema': 'Sem H2 (estrutura incompleta)',
                    'criticidade': 'MÉDIO',
                    'h1_count': h1_count,
                    'h1_text': row.get('h1_text', ''),
                    'h2_count': 0,
                    'h3_count': row.get('h3_count', 0),
                    'estrutura_headings': self._get_heading_structure(row),
                    'recomendacao': 'Adicionar H2s para estruturar o conteúdo',
                    'impacto_seo': 'Estrutura de conteúdo pouco clara'
                })
        
        return issues
    
    def _find_hierarchy_issues(self, df: pd.DataFrame) -> list:
        """Encontra problemas de hierarquia (H3 sem H2, H4 sem H3, etc.)"""
        issues = []
        
        for _, row in df.iterrows():
            url = row.get('url', '')
            h1_count = row.get('h1_count', 0)
            h2_count = row.get('h2_count', 0)
            h3_count = row.get('h3_count', 0)
            h4_count = row.get('h4_count', 0)
            h5_count = row.get('h5_count', 0)
            h6_count = row.get('h6_count', 0)
            
            hierarchy_problems = []
            
            # H3 sem H2
            if h3_count > 0 and h2_count == 0:
                hierarchy_problems.append('H3 sem H2')
            
            # H4 sem H3
            if h4_count > 0 and h3_count == 0:
                hierarchy_problems.append('H4 sem H3')
            
            # H5 sem H4
            if h5_count > 0 and h4_count == 0:
                hierarchy_problems.append('H5 sem H4')
            
            # H6 sem H5
            if h6_count > 0 and h5_count == 0:
                hierarchy_problems.append('H6 sem H5')
            
            if hierarchy_problems:
                issues.append({
                    'url': url,
                    'problema': f'Hierarquia quebrada: {", ".join(hierarchy_problems)}',
                    'criticidade': 'MÉDIO',
                    'h1_count': h1_count,
                    'h1_text': row.get('h1_text', ''),
                    'h2_count': h2_count,
                    'h3_count': h3_count,
                    'estrutura_headings': self._get_heading_structure(row),
                    'recomendacao': 'Corrigir hierarquia: cada nível deve ter o anterior',
                    'impacto_seo': 'Estrutura confusa para crawlers e usuários'
                })
        
        return issues
    
    def _find_long_h1_issues(self, df: pd.DataFrame) -> list:
        """Encontra H1s muito longos"""
        if 'h1_length' not in df.columns:
            return []
        
        long_h1 = self._safe_filter(df, 'h1_length', df['h1_length'].fillna(0) > 70)
        issues = []
        
        for _, row in long_h1.iterrows():
            h1_length = row.get('h1_length', 0)
            
            issues.append({
                'url': row.get('url', ''),
                'problema': f'H1 muito longo ({h1_length} chars)',
                'criticidade': 'BAIXO',
                'h1_count': row.get('h1_count', 0),
                'h1_text': row.get('h1_text', ''),
                'h2_count': row.get('h2_count', 0),
                'h3_count': row.get('h3_count', 0),
                'estrutura_headings': self._get_heading_structure(row),
                'recomendacao': 'Encurtar H1 para 30-70 caracteres',
                'impacto_seo': 'H1s longos podem perder foco nas palavras-chave'
            })
        
        return issues
    
    def _get_heading_structure(self, row) -> str:
        """Cria resumo da estrutura de headings"""
        structure_parts = []
        
        for level in range(1, 7):
            count = row.get(f'h{level}_count', 0)
            if count > 0:
                structure_parts.append(f'H{level}:{count}')
        
        return ' | '.join(structure_parts) if structure_parts else 'Sem headings'
    
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
            'url', 'problema', 'criticidade', 'h1_count', 'h1_text', 
            'h2_count', 'h3_count', 'estrutura_headings', 'recomendacao', 'impacto_seo'
        ]
        
        self._export_dataframe(issues_df, writer, columns)
        
        # Log estatísticas
        total_issues = len(issues_df)
        critical_issues = len(issues_df[issues_df['criticidade'] == 'CRÍTICO'])
        high_issues = len(issues_df[issues_df['criticidade'] == 'ALTO'])
        
        self.logger.info(f"✅ {self.get_sheet_name()}: {total_issues} problemas "
                        f"({critical_issues} críticos, {high_issues} altos)")