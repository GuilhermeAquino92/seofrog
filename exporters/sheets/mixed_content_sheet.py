"""
seofrog/exporters/sheets/mixed_content_sheet.py
Aba específica para problemas de Mixed Content
"""

import pandas as pd
from typing import List, Dict
from .base_sheet import BaseSheet

class MixedContentSheet(BaseSheet):
    """Aba específica para problemas de Mixed Content - HTTPS/HTTP Security"""
    
    def get_sheet_name(self) -> str:
        return '🔒 Mixed Content'
    
    def create_sheet(self, df: pd.DataFrame, writer) -> None:
        try:
            mixed_content_issues = []
            
            # Processa páginas HTTPS com problemas
            https_issues = self._process_https_issues(df)
            mixed_content_issues.extend(https_issues)
            
            # Processa Links e Forms HTTP
            http_issues = self._process_http_issues(df)
            mixed_content_issues.extend(http_issues)
            
            if mixed_content_issues:
                self._export_issues(mixed_content_issues, writer)
            else:
                self._create_success_sheet(writer, 
                    '✅ Nenhum problema de Mixed Content encontrado!\n🔒 Todas as páginas HTTPS estão seguras')
                
        except Exception as e:
            self.logger.error(f"Erro criando aba Mixed Content: {e}")
            self._create_error_sheet(writer, f'Erro: {str(e)}')
    
    def _process_https_issues(self, df: pd.DataFrame) -> List[Dict]:
        """Processa problemas de Mixed Content em páginas HTTPS"""
        issues = []
        
        # Filtra páginas HTTPS com problemas
        https_pages = self._safe_filter(df, 'is_https_page', 
            (df['is_https_page'].fillna(False) == True) & 
            (df['total_mixed_content_count'].fillna(0) > 0))
        
        for _, row in https_pages.iterrows():
            url = row.get('url', '')
            active_count = row.get('active_mixed_content_count', 0)
            passive_count = row.get('passive_mixed_content_count', 0)
            
            # Processa Active Mixed Content (CRÍTICO)
            active_details = row.get('active_mixed_content_details', [])
            if active_details and isinstance(active_details, list):
                for item in active_details:
                    if isinstance(item, dict):
                        issues.append({
                            'url': url,
                            'tipo_mixed_content': 'ACTIVE (Crítico)',
                            'tipo_recurso': item.get('type', 'unknown').upper(),
                            'url_http': item.get('url', ''),
                            'risco': 'CRÍTICO - Bloqueado pelo browser',
                            'impacto': 'Quebra funcionalidade da página',
                            'solucao': f'Alterar {item.get("type")} para HTTPS'
                        })
            
            # Processa Passive Mixed Content (AVISO)
            passive_details = row.get('passive_mixed_content_details', [])
            if passive_details and isinstance(passive_details, list):
                for item in passive_details:
                    if isinstance(item, dict):
                        issues.append({
                            'url': url,
                            'tipo_mixed_content': 'PASSIVE (Aviso)',
                            'tipo_recurso': item.get('type', 'unknown').upper(),
                            'url_http': item.get('url', ''),
                            'risco': 'MÉDIO - Cadeado quebrado',
                            'impacto': 'Reduz confiança do usuário',
                            'solucao': f'Alterar {item.get("type")} para HTTPS'
                        })
            
            # FALLBACK: Se não há detalhes, usa resumos
            if (not active_details and active_count > 0) or (not passive_details and passive_count > 0):
                risk_level = row.get('mixed_content_risk', 'DESCONHECIDO')
                issues.append({
                    'url': url,
                    'tipo_mixed_content': 'MIXED CONTENT',
                    'tipo_recurso': 'MÚLTIPLOS',
                    'url_http': f'{active_count} active + {passive_count} passive',
                    'risco': risk_level,
                    'impacto': 'Problemas de segurança HTTPS',
                    'solucao': 'Verificar todos os recursos HTTP'
                })
        
        return issues
    
    def _process_http_issues(self, df: pd.DataFrame) -> List[Dict]:
        """Processa Links e Forms HTTP"""
        issues = []
        
        # Filtra páginas com Links ou Forms HTTP
        http_pages = self._safe_filter(df, 'http_links_count',
            (df['http_links_count'].fillna(0) > 0) | 
            (df['http_forms_count'].fillna(0) > 0))
        
        for _, row in http_pages.iterrows():
            url = row.get('url', '')
            links_count = row.get('http_links_count', 0)
            forms_count = row.get('http_forms_count', 0)
            
            # Links HTTP
            if links_count > 0:
                issues.append({
                    'url': url,
                    'tipo_mixed_content': 'HTTP LINKS',
                    'tipo_recurso': 'LINKS',
                    'url_http': f'{links_count} links HTTP',
                    'risco': 'BAIXO - Não é mixed content',
                    'impacto': 'Usuário pode sair do HTTPS',
                    'solucao': 'Alterar links para HTTPS quando possível'
                })
            
            # Forms HTTP
            if forms_count > 0:
                issues.append({
                    'url': url,
                    'tipo_mixed_content': 'HTTP FORMS',
                    'tipo_recurso': 'FORMS',
                    'url_http': f'{forms_count} forms HTTP',
                    'risco': 'MÉDIO - Dados não criptografados',
                    'impacto': 'Submissão de dados insegura',
                    'solucao': 'URGENTE: Alterar action para HTTPS'
                })
        
        return issues
    
    def _export_issues(self, issues: List[Dict], writer):
        """Exporta problemas para Excel com ordenação por criticidade"""
        issues_df = pd.DataFrame(issues)
        
        # Remove duplicatas
        issues_df = issues_df.drop_duplicates(subset=['url', 'tipo_mixed_content', 'url_http'], keep='first')
        
        # Ordena por criticidade
        risk_order = {
            'CRÍTICO - Bloqueado pelo browser': 1,
            'MÉDIO - Cadeado quebrado': 2,
            'MÉDIO - Dados não criptografados': 3,
            'BAIXO - Não é mixed content': 4
        }
        issues_df['_sort_order'] = issues_df['risco'].map(risk_order).fillna(5)
        issues_df = issues_df.sort_values(['_sort_order', 'url']).drop('_sort_order', axis=1)
        
        # Exporta
        issues_df.to_excel(writer, sheet_name=self.get_sheet_name(), index=False)
        
        # Log estatísticas
        total_issues = len(issues_df)
        critical_issues = len(issues_df[issues_df['risco'].str.contains('CRÍTICO')])
        medium_issues = len(issues_df[issues_df['risco'].str.contains('MÉDIO')])
        
        self.logger.info(f"✅ {self.get_sheet_name()}: {total_issues} problemas ({critical_issues} críticos, {medium_issues} médios)")