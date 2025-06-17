"""
seofrog/exporters/sheets/technical_problems_sheet.py
Aba específica para problemas técnicos de SEO
"""

import pandas as pd
from .base_sheet import BaseSheet

class TechnicalProblemsSheet(BaseSheet):
    """Aba com problemas técnicos (canonical, viewport, charset, schema, etc.)"""
    
    def get_sheet_name(self) -> str:
        return 'Problemas Técnicos'
    
    def create_sheet(self, df: pd.DataFrame, writer) -> None:
        try:
            technical_issues = []
            
            # Problemas de canonical
            canonical_issues = self._find_canonical_issues(df)
            technical_issues.extend(canonical_issues)
            
            # Problemas de viewport
            viewport_issues = self._find_viewport_issues(df)
            technical_issues.extend(viewport_issues)
            
            # Problemas de charset
            charset_issues = self._find_charset_issues(df)
            technical_issues.extend(charset_issues)
            
            # Problemas de structured data
            schema_issues = self._find_schema_issues(df)
            technical_issues.extend(schema_issues)
            
            # Problemas de meta robots
            robots_issues = self._find_robots_issues(df)
            technical_issues.extend(robots_issues)
            
            # Problemas de favicon
            favicon_issues = self._find_favicon_issues(df)
            technical_issues.extend(favicon_issues)
            
            # Problemas de Open Graph
            og_issues = self._find_og_issues(df)
            technical_issues.extend(og_issues)
            
            if technical_issues:
                self._export_consolidated_issues(technical_issues, writer)
            else:
                self._create_success_sheet(writer, '✅ Nenhum problema técnico encontrado!')
                
        except Exception as e:
            self.logger.error(f"Erro criando aba de problemas técnicos: {e}")
            self._create_error_sheet(writer, f'Erro: {str(e)}')
    
    def _find_canonical_issues(self, df: pd.DataFrame) -> list:
        """Encontra problemas de canonical URL"""
        issues = []
        
        # URLs sem canonical
        if 'canonical_url' in df.columns:
            no_canonical = self._safe_filter(df, 'canonical_url',
                df['canonical_url'].fillna('') == '')
            
            for _, row in no_canonical.iterrows():
                issues.append({
                    'url': row.get('url', ''),
                    'problema': 'Sem canonical URL',
                    'criticidade': 'MÉDIO',
                    'categoria': 'Canonical',
                    'valor_atual': 'Ausente',
                    'valor_esperado': 'URL canonical definida',
                    'impacto_seo': 'Possível conteúdo duplicado',
                    'action_required': 'Adicionar tag canonical',
                    'technical_fix': '<link rel="canonical" href="URL_CANONICAL">',
                    'priority_score': 6
                })
        
        # Canonical não é self
        if 'canonical_is_self' in df.columns:
            canonical_not_self = self._safe_filter(df, 'canonical_is_self',
                df['canonical_is_self'].fillna(True) == False)
            
            for _, row in canonical_not_self.iterrows():
                canonical_url = row.get('canonical_url', '')
                issues.append({
                    'url': row.get('url', ''),
                    'problema': 'Canonical aponta para outra URL',
                    'criticidade': 'BAIXO',
                    'categoria': 'Canonical',
                    'valor_atual': canonical_url,
                    'valor_esperado': 'Canonical = URL atual',
                    'impacto_seo': 'Pode indicar conteúdo duplicado ou redirecionamento',
                    'action_required': 'Verificar se canonical está correto',
                    'technical_fix': 'Confirmar se redirecionamento é intencional',
                    'priority_score': 3
                })
        
        return issues
    
    def _find_viewport_issues(self, df: pd.DataFrame) -> list:
        """Encontra problemas de viewport (mobile)"""
        if 'has_viewport' not in df.columns:
            return []
        
        no_viewport = self._safe_filter(df, 'has_viewport',
            df['has_viewport'].fillna(False) == False)
        issues = []
        
        for _, row in no_viewport.iterrows():
            issues.append({
                'url': row.get('url', ''),
                'problema': 'Sem meta viewport',
                'criticidade': 'ALTO',
                'categoria': 'Mobile',
                'valor_atual': 'Ausente',
                'valor_esperado': 'Meta viewport configurado',
                'impacto_seo': 'Página não otimizada para mobile - Core Web Vitals',
                'action_required': 'Adicionar meta viewport',
                'technical_fix': '<meta name="viewport" content="width=device-width, initial-scale=1">',
                'priority_score': 8
            })
        
        return issues
    
    def _find_charset_issues(self, df: pd.DataFrame) -> list:
        """Encontra problemas de charset"""
        if 'has_charset' not in df.columns:
            return []
        
        no_charset = self._safe_filter(df, 'has_charset',
            df['has_charset'].fillna(False) == False)
        issues = []
        
        for _, row in no_charset.iterrows():
            issues.append({
                'url': row.get('url', ''),
                'problema': 'Sem charset definido',
                'criticidade': 'MÉDIO',
                'categoria': 'Encoding',
                'valor_atual': 'Ausente',
                'valor_esperado': 'UTF-8',
                'impacto_seo': 'Possíveis problemas de codificação de caracteres',
                'action_required': 'Definir charset UTF-8',
                'technical_fix': '<meta charset="UTF-8">',
                'priority_score': 5
            })
        
        return issues
    
    def _find_schema_issues(self, df: pd.DataFrame) -> list:
        """Encontra problemas de structured data"""
        if 'schema_total_count' not in df.columns:
            return []
        
        no_schema = self._safe_filter(df, 'schema_total_count',
            df['schema_total_count'].fillna(0) == 0)
        issues = []
        
        for _, row in no_schema.iterrows():
            page_type = self._detect_page_type(row)
            
            # Só considera problema para tipos de página que se beneficiam de schema
            if page_type in ['Produto', 'Blog/Artigo', 'Categoria', 'Institucional']:
                criticality = 'MÉDIO' if page_type == 'Produto' else 'BAIXO'
                priority = 7 if page_type == 'Produto' else 4
                
                issues.append({
                    'url': row.get('url', ''),
                    'problema': 'Sem structured data (Schema)',
                    'criticidade': criticality,
                    'categoria': 'Structured Data',
                    'valor_atual': 'Ausente',
                    'valor_esperado': f'Schema.org para {page_type}',
                    'impacto_seo': 'Perda de rich snippets no Google',
                    'action_required': f'Implementar schema para {page_type}',
                    'technical_fix': 'JSON-LD com schema apropriado',
                    'priority_score': priority
                })
        
        return issues
    
    def _find_robots_issues(self, df: pd.DataFrame) -> list:
        """Encontra problemas de meta robots"""
        issues = []
        
        # Páginas com noindex
        if 'meta_robots_noindex' in df.columns:
            noindex_pages = self._safe_filter(df, 'meta_robots_noindex',
                df['meta_robots_noindex'].fillna(False) == True)
            
            for _, row in noindex_pages.iterrows():
                issues.append({
                    'url': row.get('url', ''),
                    'problema': 'Meta robots: noindex',
                    'criticidade': 'ALTO',
                    'categoria': 'Indexação',
                    'valor_atual': 'noindex',
                    'valor_esperado': 'index (ou remover tag)',
                    'impacto_seo': 'Página não será indexada pelo Google',
                    'action_required': 'Verificar se noindex é intencional',
                    'technical_fix': 'Remover noindex ou alterar para index',
                    'priority_score': 9
                })
        
        # Páginas com nofollow
        if 'meta_robots_nofollow' in df.columns:
            nofollow_pages = self._safe_filter(df, 'meta_robots_nofollow',
                df['meta_robots_nofollow'].fillna(False) == True)
            
            for _, row in nofollow_pages.iterrows():
                issues.append({
                    'url': row.get('url', ''),
                    'problema': 'Meta robots: nofollow',
                    'criticidade': 'MÉDIO',
                    'categoria': 'Link Equity',
                    'valor_atual': 'nofollow',
                    'valor_esperado': 'follow (ou remover tag)',
                    'impacto_seo': 'Links não passam autoridade',
                    'action_required': 'Verificar se nofollow é intencional',
                    'technical_fix': 'Remover nofollow ou alterar para follow',
                    'priority_score': 5
                })
        
        return issues
    
    def _find_favicon_issues(self, df: pd.DataFrame) -> list:
        """Encontra problemas de favicon"""
        if 'has_favicon' not in df.columns:
            return []
        
        no_favicon = self._safe_filter(df, 'has_favicon',
            df['has_favicon'].fillna(False) == False)
        issues = []
        
        for _, row in no_favicon.iterrows():
            issues.append({
                'url': row.get('url', ''),
                'problema': 'Sem favicon',
                'criticidade': 'BAIXO',
                'categoria': 'UX/Branding',
                'valor_atual': 'Ausente',
                'valor_esperado': 'Favicon configurado',
                'impacto_seo': 'Baixo impacto SEO, afeta confiança do usuário',
                'action_required': 'Adicionar favicon',
                'technical_fix': '<link rel="icon" href="/favicon.ico">',
                'priority_score': 2
            })
        
        return issues
    
    def _find_og_issues(self, df: pd.DataFrame) -> list:
        """Encontra problemas de Open Graph"""
        if 'og_tags_count' not in df.columns:
            return []
        
        no_og = self._safe_filter(df, 'og_tags_count',
            df['og_tags_count'].fillna(0) == 0)
        issues = []
        
        for _, row in no_og.iterrows():
            page_type = self._detect_page_type(row)
            
            # Prioriza páginas que se beneficiam mais de OG
            if page_type in ['Produto', 'Blog/Artigo', 'Homepage']:
                criticality = 'MÉDIO' if page_type in ['Produto', 'Homepage'] else 'BAIXO'
                priority = 6 if page_type in ['Produto', 'Homepage'] else 3
                
                issues.append({
                    'url': row.get('url', ''),
                    'problema': 'Sem Open Graph tags',
                    'criticidade': criticality,
                    'categoria': 'Social Media',
                    'valor_atual': 'Ausente',
                    'valor_esperado': 'OG tags configuradas',
                    'impacto_seo': 'Compartilhamentos sociais sem preview adequado',
                    'action_required': 'Implementar Open Graph básico',
                    'technical_fix': 'og:title, og:description, og:image, og:url',
                    'priority_score': priority
                })
        
        return issues
    
    def _detect_page_type(self, row) -> str:
        """Detecta tipo da página baseado na URL"""
        url = row.get('url', '').lower()
        
        if '/blog/' in url or '/artigo/' in url or '/post/' in url:
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
        """Exporta problemas técnicos com ordenação por prioridade"""
        if not issues:
            return
        
        issues_df = pd.DataFrame(issues)
        
        # Remove duplicatas (mesma URL pode ter múltiplos problemas)
        issues_df = issues_df.drop_duplicates(subset=['url', 'problema'], keep='first')
        
        # Ordena por priority_score (maior = mais prioritário) e criticidade
        issues_df = issues_df.sort_values(['priority_score', 'criticidade'], ascending=[False, True])
        
        # Define colunas para exportar
        columns = [
            'url', 'problema', 'criticidade', 'categoria', 'valor_atual',
            'valor_esperado', 'impacto_seo', 'action_required', 'technical_fix'
        ]
        
        self._export_dataframe(issues_df, writer, columns)
        
        # Log estatísticas por categoria
        category_stats = issues_df['categoria'].value_counts()
        criticality_stats = issues_df['criticidade'].value_counts()
        
        total_issues = len(issues_df)
        self.logger.info(f"✅ {self.get_sheet_name()}: {total_issues} problemas técnicos")
        
        # Log por categoria
        for category, count in category_stats.items():
            self.logger.info(f"   🔧 {category}: {count} problemas")
        
        # Log por criticidade
        for criticality, count in criticality_stats.items():
            self.logger.info(f"   ⚠️ {criticality}: {count} problemas")
        
        # Log top 3 problemas mais comuns
        top_problems = issues_df['problema'].value_counts().head(3)
        self.logger.info("   📊 Problemas mais comuns:")
        for problem, count in top_problems.items():
            self.logger.info(f"      • {problem}: {count} URLs")