"""
seofrog/exporters/sheets/performance_sheet.py
Aba específica para problemas de performance
"""

import pandas as pd
from .base_sheet import BaseSheet

class PerformanceSheet(BaseSheet):
    """Aba com problemas de performance (velocidade, tamanho, Core Web Vitals)"""
    
    def get_sheet_name(self) -> str:
        return 'Problemas Performance'
    
    def create_sheet(self, df: pd.DataFrame, writer) -> None:
        try:
            performance_issues = []
            
            # Páginas lentas
            slow_pages_issues = self._find_slow_pages_issues(df)
            performance_issues.extend(slow_pages_issues)
            
            # Páginas muito lentas
            very_slow_pages_issues = self._find_very_slow_pages_issues(df)
            performance_issues.extend(very_slow_pages_issues)
            
            # Páginas pesadas
            heavy_pages_issues = self._find_heavy_pages_issues(df)
            performance_issues.extend(heavy_pages_issues)
            
            # Páginas com muitos recursos
            resource_heavy_issues = self._find_resource_heavy_issues(df)
            performance_issues.extend(resource_heavy_issues)
            
            # Problemas de eficiência de conteúdo
            content_efficiency_issues = self._find_content_efficiency_issues(df)
            performance_issues.extend(content_efficiency_issues)
            
            if performance_issues:
                self._export_consolidated_issues(performance_issues, writer)
            else:
                self._create_success_sheet(writer, '✅ Nenhum problema de performance encontrado!')
                
        except Exception as e:
            self.logger.error(f"Erro criando aba de problemas de performance: {e}")
            self._create_error_sheet(writer, f'Erro: {str(e)}')
    
    def _find_slow_pages_issues(self, df: pd.DataFrame) -> list:
        """Encontra páginas lentas (3-5 segundos)"""
        if 'response_time' not in df.columns:
            return []
        
        slow_pages = self._safe_filter(df, 'response_time',
            (df['response_time'].fillna(0) > 3) & (df['response_time'].fillna(0) <= 5))
        issues = []
        
        for _, row in slow_pages.iterrows():
            response_time = row.get('response_time', 0)
            
            issues.append({
                'url': row.get('url', ''),
                'problema': f'Página lenta ({response_time:.2f}s)',
                'criticidade': 'ALTO',
                'categoria': 'Velocidade',
                'response_time': response_time,
                'content_length': row.get('content_length', 0),
                'content_length_mb': self._bytes_to_mb(row.get('content_length', 0)),
                'page_type': self._detect_page_type(row),
                'core_web_vitals_impact': 'LCP pode estar comprometido',
                'user_experience_impact': 'Usuários podem abandonar a página',
                'seo_impact': 'Google prioriza páginas mais rápidas',
                'action_required': 'Otimizar tempo de resposta do servidor',
                'technical_recommendations': self._get_speed_recommendations(response_time),
                'business_impact': 'Perda de conversões e engajamento',
                'priority_score': 8
            })
        
        return issues
    
    def _find_very_slow_pages_issues(self, df: pd.DataFrame) -> list:
        """Encontra páginas muito lentas (>5 segundos)"""
        if 'response_time' not in df.columns:
            return []
        
        very_slow_pages = self._safe_filter(df, 'response_time',
            df['response_time'].fillna(0) > 5)
        issues = []
        
        for _, row in very_slow_pages.iterrows():
            response_time = row.get('response_time', 0)
            
            issues.append({
                'url': row.get('url', ''),
                'problema': f'Página muito lenta ({response_time:.2f}s)',
                'criticidade': 'CRÍTICO',
                'categoria': 'Velocidade',
                'response_time': response_time,
                'content_length': row.get('content_length', 0),
                'content_length_mb': self._bytes_to_mb(row.get('content_length', 0)),
                'page_type': self._detect_page_type(row),
                'core_web_vitals_impact': 'LCP crítico - Experiência ruim',
                'user_experience_impact': 'Alta probabilidade de abandono',
                'seo_impact': 'Penalização severa no ranking',
                'action_required': 'URGENTE: Investigar e otimizar',
                'technical_recommendations': self._get_critical_speed_recommendations(),
                'business_impact': 'Perda significativa de receita',
                'priority_score': 10
            })
        
        return issues
    
    def _find_heavy_pages_issues(self, df: pd.DataFrame) -> list:
        """Encontra páginas pesadas (>1MB)"""
        if 'content_length' not in df.columns:
            return []
        
        heavy_pages = self._safe_filter(df, 'content_length',
            df['content_length'].fillna(0) > 1048576)  # 1MB
        issues = []
        
        for _, row in heavy_pages.iterrows():
            content_length = row.get('content_length', 0)
            content_mb = self._bytes_to_mb(content_length)
            
            # Determina criticidade baseada no tamanho
            if content_mb > 5:
                criticality = 'CRÍTICO'
                priority = 9
            elif content_mb > 3:
                criticality = 'ALTO'
                priority = 7
            else:
                criticality = 'MÉDIO'
                priority = 5
            
            issues.append({
                'url': row.get('url', ''),
                'problema': f'Página pesada ({content_mb:.1f} MB)',
                'criticidade': criticality,
                'categoria': 'Tamanho',
                'response_time': row.get('response_time', 0),
                'content_length': content_length,
                'content_length_mb': content_mb,
                'page_type': self._detect_page_type(row),
                'core_web_vitals_impact': 'FCP e LCP podem ser afetados',
                'user_experience_impact': 'Carregamento lento, especialmente mobile',
                'seo_impact': 'Mobile-first indexing penalizado',
                'action_required': 'Reduzir tamanho da página',
                'technical_recommendations': self._get_size_recommendations(content_mb),
                'business_impact': 'Custos de banda e UX prejudicada',
                'priority_score': priority
            })
        
        return issues
    
    def _find_resource_heavy_issues(self, df: pd.DataFrame) -> list:
        """Encontra páginas com muitos recursos (imagens, links)"""
        issues = []
        
        for _, row in df.iterrows():
            url = row.get('url', '')
            images_count = row.get('images_count', 0)
            total_links = row.get('total_links_count', 0)
            
            resource_problems = []
            total_resources = images_count + total_links
            
            # Muitas imagens
            if images_count > 50:
                resource_problems.append(f'{images_count} imagens')
            
            # Muitos links
            if total_links > 200:
                resource_problems.append(f'{total_links} links')
            
            if resource_problems and total_resources > 100:
                criticality = 'ALTO' if total_resources > 300 else 'MÉDIO'
                
                issues.append({
                    'url': url,
                    'problema': f'Muitos recursos: {", ".join(resource_problems)}',
                    'criticidade': criticality,
                    'categoria': 'Recursos',
                    'response_time': row.get('response_time', 0),
                    'content_length': row.get('content_length', 0),
                    'content_length_mb': self._bytes_to_mb(row.get('content_length', 0)),
                    'page_type': self._detect_page_type(row),
                    'core_web_vitals_impact': 'CLS e LCP podem ser impactados',
                    'user_experience_impact': 'Página complexa, possível lentidão',
                    'seo_impact': 'Crawl budget pode ser desperdiçado',
                    'action_required': 'Otimizar quantidade de recursos',
                    'technical_recommendations': 'Lazy loading, CDN, compressão',
                    'business_impact': 'UX degradada, especialmente mobile',
                    'priority_score': 6
                })
        
        return issues
    
    def _find_content_efficiency_issues(self, df: pd.DataFrame) -> list:
        """Encontra problemas de eficiência de conteúdo"""
        if 'text_ratio' not in df.columns:
            return []
        
        low_text_ratio = self._safe_filter(df, 'text_ratio',
            df['text_ratio'].fillna(1) < 0.1)  # Menos de 10% de texto
        issues = []
        
        for _, row in low_text_ratio.iterrows():
            text_ratio = row.get('text_ratio', 0)
            
            issues.append({
                'url': row.get('url', ''),
                'problema': f'Baixa eficiência de conteúdo ({text_ratio:.1%} texto)',
                'criticidade': 'BAIXO',
                'categoria': 'Eficiência',
                'response_time': row.get('response_time', 0),
                'content_length': row.get('content_length', 0),
                'content_length_mb': self._bytes_to_mb(row.get('content_length', 0)),
                'page_type': self._detect_page_type(row),
                'core_web_vitals_impact': 'Baixo impacto direto',
                'user_experience_impact': 'Conteúdo pode estar "diluído"',
                'seo_impact': 'Pouco conteúdo indexável',
                'action_required': 'Revisar estrutura da página',
                'technical_recommendations': 'Reduzir código, aumentar conteúdo útil',
                'business_impact': 'Baixa relevância para usuários',
                'priority_score': 3
            })
        
        return issues
    
    def _bytes_to_mb(self, bytes_value: int) -> float:
        """Converte bytes para MB"""
        return bytes_value / (1024 * 1024) if bytes_value else 0
    
    def _get_speed_recommendations(self, response_time: float) -> str:
        """Gera recomendações baseadas no tempo de resposta"""
        if response_time > 4:
            return "CDN, cache server, otimização DB, minificação"
        elif response_time > 3.5:
            return "Cache browser, compressão GZIP, otimização imagens"
        else:
            return "Fine-tuning server, async loading"
    
    def _get_critical_speed_recommendations(self) -> str:
        """Recomendações para páginas críticas"""
        return "URGENTE: CDN global, cache avançado, server upgrade, code splitting"
    
    def _get_size_recommendations(self, size_mb: float) -> str:
        """Gera recomendações baseadas no tamanho"""
        if size_mb > 5:
            return "Compressão agressiva, lazy loading, code splitting, WebP"
        elif size_mb > 3:
            return "Otimização imagens, minificação, tree shaking"
        else:
            return "Compressão GZIP, remoção código unused"
    
    def _detect_page_type(self, row) -> str:
        """Detecta tipo da página"""
        url = row.get('url', '').lower()
        
        if '/blog/' in url or '/artigo/' in url:
            return 'Blog/Artigo'
        elif '/produto/' in url or '/product/' in url:
            return 'Produto'
        elif '/categoria/' in url or '/category/' in url:
            return 'Categoria'
        elif url.endswith('/') and url.count('/') <= 3:
            return 'Homepage'
        else:
            return 'Conteúdo'
    
    def _export_consolidated_issues(self, issues: list, writer):
        """Exporta problemas de performance com ordenação por impacto"""
        if not issues:
            return
        
        issues_df = pd.DataFrame(issues)
        
        # Remove duplicatas
        issues_df = issues_df.drop_duplicates(subset=['url', 'problema'], keep='first')
        
        # Ordena por priority_score (maior = mais crítico)
        issues_df = issues_df.sort_values(['priority_score', 'response_time'], ascending=[False, False])
        
        # Define colunas para exportar
        columns = [
            'url', 'problema', 'criticidade', 'categoria', 'response_time',
            'content_length_mb', 'page_type', 'core_web_vitals_impact',
            'user_experience_impact', 'seo_impact', 'action_required',
            'technical_recommendations', 'business_impact'
        ]
        
        self._export_dataframe(issues_df, writer, columns)
        
        # Log estatísticas de performance
        total_issues = len(issues_df)
        critical_issues = len(issues_df[issues_df['criticidade'] == 'CRÍTICO'])
        
        # Estatísticas de velocidade
        if 'response_time' in issues_df.columns:
            avg_response_time = issues_df['response_time'].mean()
            max_response_time = issues_df['response_time'].max()
            self.logger.info(f"✅ {self.get_sheet_name()}: {total_issues} problemas de performance")
            self.logger.info(f"   ⏱️ Tempo médio: {avg_response_time:.2f}s | Máximo: {max_response_time:.2f}s")
        
        # Estatísticas por categoria
        category_stats = issues_df['categoria'].value_counts()
        for category, count in category_stats.items():
            self.logger.info(f"   📊 {category}: {count} problemas")
        
        if critical_issues > 0:
            self.logger.warning(f"   🚨 {critical_issues} problemas CRÍTICOS de performance!")