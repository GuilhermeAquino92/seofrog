"""
seofrog/exporters/sheets/image_problems_sheet.py
Aba específica para problemas de imagens
"""

import pandas as pd
from .base_sheet import BaseSheet

class ImageProblemsSheet(BaseSheet):
    """Aba com problemas de imagens (ALT, SRC, otimização)"""
    
    def get_sheet_name(self) -> str:
        return 'Problemas Imagens'
    
    def create_sheet(self, df: pd.DataFrame, writer) -> None:
        try:
            image_issues = []
            
            # URLs com imagens sem ALT
            no_alt_issues = self._find_no_alt_issues(df)
            image_issues.extend(no_alt_issues)
            
            # URLs com imagens sem SRC
            no_src_issues = self._find_no_src_issues(df)
            image_issues.extend(no_src_issues)
            
            # URLs com muitas imagens
            many_images_issues = self._find_many_images_issues(df)
            image_issues.extend(many_images_issues)
            
            # URLs sem imagens (quando deveria ter)
            no_images_issues = self._find_no_images_issues(df)
            image_issues.extend(no_images_issues)
            
            # Problemas de dimensões
            dimension_issues = self._find_dimension_issues(df)
            image_issues.extend(dimension_issues)
            
            if image_issues:
                self._export_consolidated_issues(image_issues, writer)
            else:
                self._create_success_sheet(writer, '✅ Nenhum problema de imagem encontrado!')
                
        except Exception as e:
            self.logger.error(f"Erro criando aba de problemas de imagens: {e}")
            self._create_error_sheet(writer, f'Erro: {str(e)}')
    
    def _find_no_alt_issues(self, df: pd.DataFrame) -> list:
        """Encontra URLs com imagens sem ALT"""
        if 'images_without_alt' not in df.columns:
            return []
        
        no_alt_pages = self._safe_filter(df, 'images_without_alt', 
            df['images_without_alt'].fillna(0) > 0)
        issues = []
        
        for _, row in no_alt_pages.iterrows():
            images_without_alt = row.get('images_without_alt', 0)
            total_images = row.get('images_count', 0)
            alt_percentage = ((total_images - images_without_alt) / total_images * 100) if total_images > 0 else 0
            
            criticality = 'CRÍTICO' if images_without_alt > 5 else 'ALTO'
            
            issues.append({
                'url': row.get('url', ''),
                'problema': f'{images_without_alt} imagens sem ALT',
                'criticidade': criticality,
                'images_count': total_images,
                'images_without_alt': images_without_alt,
                'images_with_alt': total_images - images_without_alt,
                'alt_coverage': f'{alt_percentage:.1f}%',
                'page_type': self._detect_page_type(row),
                'accessibility_impact': 'Alto - Leitores de tela não conseguem descrever',
                'seo_impact': 'Perda de indexação de imagens no Google Images',
                'action_required': f'Adicionar ALT em {images_without_alt} imagens',
                'technical_fix': '<img src="..." alt="Descrição da imagem">'
            })
        
        return issues
    
    def _find_no_src_issues(self, df: pd.DataFrame) -> list:
        """Encontra URLs com imagens sem SRC"""
        if 'images_without_src' not in df.columns:
            return []
        
        no_src_pages = self._safe_filter(df, 'images_without_src',
            df['images_without_src'].fillna(0) > 0)
        issues = []
        
        for _, row in no_src_pages.iterrows():
            images_without_src = row.get('images_without_src', 0)
            total_images = row.get('images_count', 0)
            
            issues.append({
                'url': row.get('url', ''),
                'problema': f'{images_without_src} imagens sem SRC',
                'criticidade': 'CRÍTICO',
                'images_count': total_images,
                'images_without_src': images_without_src,
                'images_with_alt': row.get('images_count', 0) - row.get('images_without_alt', 0),
                'alt_coverage': 'N/A',
                'page_type': self._detect_page_type(row),
                'accessibility_impact': 'Crítico - Imagens quebradas',
                'seo_impact': 'Imagens não carregam - Experiência ruim',
                'action_required': f'Corrigir SRC em {images_without_src} imagens',
                'technical_fix': 'Verificar URLs das imagens e CDN'
            })
        
        return issues
    
    def _find_many_images_issues(self, df: pd.DataFrame) -> list:
        """Encontra URLs com muitas imagens (performance)"""
        if 'images_count' not in df.columns:
            return []
        
        many_images = self._safe_filter(df, 'images_count',
            df['images_count'].fillna(0) > 50)
        issues = []
        
        for _, row in many_images.iterrows():
            images_count = row.get('images_count', 0)
            images_without_alt = row.get('images_without_alt', 0)
            
            criticality = 'ALTO' if images_count > 100 else 'MÉDIO'
            
            issues.append({
                'url': row.get('url', ''),
                'problema': f'Muitas imagens ({images_count} imagens)',
                'criticidade': criticality,
                'images_count': images_count,
                'images_without_alt': images_without_alt,
                'images_with_alt': images_count - images_without_alt,
                'alt_coverage': f'{((images_count - images_without_alt) / images_count * 100):.1f}%' if images_count > 0 else '0%',
                'page_type': self._detect_page_type(row),
                'accessibility_impact': 'Médio - Muitos elementos para navegar',
                'seo_impact': 'Possível impacto na velocidade de carregamento',
                'action_required': 'Otimizar quantidade e compressão das imagens',
                'technical_fix': 'Lazy loading + compressão + WebP'
            })
        
        return issues
    
    def _find_no_images_issues(self, df: pd.DataFrame) -> list:
        """Encontra URLs que deveriam ter imagens mas não têm"""
        if 'images_count' not in df.columns:
            return []
        
        no_images = self._safe_filter(df, 'images_count',
            df['images_count'].fillna(0) == 0)
        issues = []
        
        for _, row in no_images.iterrows():
            url = row.get('url', '').lower()
            page_type = self._detect_page_type(row)
            
            # Só considera problema para tipos de página que tipicamente têm imagens
            if page_type in ['Produto', 'Blog/Artigo', 'Categoria']:
                issues.append({
                    'url': row.get('url', ''),
                    'problema': 'Nenhuma imagem encontrada',
                    'criticidade': 'MÉDIO' if page_type == 'Produto' else 'BAIXO',
                    'images_count': 0,
                    'images_without_alt': 0,
                    'images_with_alt': 0,
                    'alt_coverage': 'N/A',
                    'page_type': page_type,
                    'accessibility_impact': 'Baixo - Conteúdo pode ser só textual',
                    'seo_impact': 'Perda de visibilidade no Google Images',
                    'action_required': 'Considerar adicionar imagens relevantes',
                    'technical_fix': 'Adicionar imagens que complementem o conteúdo'
                })
        
        return issues
    
    def _find_dimension_issues(self, df: pd.DataFrame) -> list:
        """Encontra problemas de dimensões de imagens"""
        if 'images_with_dimensions' not in df.columns or 'images_count' not in df.columns:
            return []
        
        issues = []
        
        for _, row in df.iterrows():
            images_count = row.get('images_count', 0)
            images_with_dimensions = row.get('images_with_dimensions', 0)
            
            # Só analisa se tem imagens
            if images_count > 0:
                images_without_dimensions = images_count - images_with_dimensions
                dimension_coverage = (images_with_dimensions / images_count * 100)
                
                # Considera problema se menos de 80% das imagens tem dimensões
                if dimension_coverage < 80 and images_without_dimensions > 0:
                    issues.append({
                        'url': row.get('url', ''),
                        'problema': f'{images_without_dimensions} imagens sem dimensões',
                        'criticidade': 'BAIXO',
                        'images_count': images_count,
                        'images_without_alt': row.get('images_without_alt', 0),
                        'images_with_alt': images_count - row.get('images_without_alt', 0),
                        'alt_coverage': f'{((images_count - row.get("images_without_alt", 0)) / images_count * 100):.1f}%',
                        'page_type': self._detect_page_type(row),
                        'accessibility_impact': 'Baixo - Pode causar layout shifts',
                        'seo_impact': 'Core Web Vitals - CLS pode ser afetado',
                        'action_required': f'Definir width/height em {images_without_dimensions} imagens',
                        'technical_fix': '<img src="..." width="400" height="300">'
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
        elif '/galeria/' in url or '/gallery/' in url:
            return 'Galeria'
        elif '/sobre' in url or '/about' in url:
            return 'Institucional'
        else:
            return 'Conteúdo'
    
    def _export_consolidated_issues(self, issues: list, writer):
        """Exporta problemas consolidados com ordenação por criticidade"""
        if not issues:
            return
        
        issues_df = pd.DataFrame(issues)
        
        # Remove duplicatas (mesma URL pode ter múltiplos problemas)
        issues_df = issues_df.drop_duplicates(subset=['url', 'problema'], keep='first')
        
        # Ordena por criticidade e número de problemas
        issues_df = self._sort_by_criticality(issues_df)
        
        # Define colunas para exportar
        columns = [
            'url', 'problema', 'criticidade', 'images_count', 'images_without_alt',
            'images_with_alt', 'alt_coverage', 'page_type', 'accessibility_impact',
            'seo_impact', 'action_required', 'technical_fix'
        ]
        
        self._export_dataframe(issues_df, writer, columns)
        
        # Log estatísticas por tipo de problema
        problem_types = issues_df['problema'].str.extract(r'(sem ALT|sem SRC|Muitas imagens|Nenhuma imagem|sem dimensões)')[0].value_counts()
        
        total_issues = len(issues_df)
        self.logger.info(f"✅ {self.get_sheet_name()}: {total_issues} problemas de imagem")
        
        for problem_type, count in problem_types.items():
            self.logger.info(f"   🖼️ {problem_type}: {count} URLs")