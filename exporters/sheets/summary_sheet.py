"""
seofrog/exporters/sheets/summary_sheet.py
Aba de Resumo Executivo
"""

import pandas as pd
from .base_sheet import BaseSheet

class SummarySheet(BaseSheet):
    """Aba de resumo executivo com estatísticas principais"""
    
    def get_sheet_name(self) -> str:
        return 'Resumo Executivo'
    
    def create_sheet(self, df: pd.DataFrame, writer) -> None:
        try:
            total_urls = len(df)
            
            if total_urls == 0:
                self._create_error_sheet(writer, 'Nenhum dado para resumir')
                return
            
            # Cria dados do resumo
            summary_data = [['Métrica', 'Valor', 'Percentual']]
            summary_data.append(['Total de URLs', total_urls, '100%'])
            
            # Adiciona estatísticas
            self._add_status_codes_stats(df, summary_data, total_urls)
            self._add_seo_problems_stats(df, summary_data, total_urls)
            self._add_technical_stats(df, summary_data, total_urls)
            self._add_mixed_content_stats(df, summary_data, total_urls)
            
            # Cria DataFrame e exporta
            if len(summary_data) > 1:
                summary_df = pd.DataFrame(summary_data[1:], columns=summary_data[0])
                summary_df.to_excel(writer, sheet_name=self.get_sheet_name(), index=False)
                self.logger.info(f"✅ {self.get_sheet_name()}: {len(summary_df)} métricas")
            else:
                self._create_error_sheet(writer, 'Erro na criação do resumo')
                
        except Exception as e:
            self.logger.error(f"Erro criando resumo: {e}")
            self._create_error_sheet(writer, f'Erro: {str(e)}')
    
    def _add_status_codes_stats(self, df: pd.DataFrame, summary_data: list, total_urls: int):
        """Adiciona estatísticas de status codes"""
        if 'status_code' not in df.columns:
            return
        
        try:
            status_counts = df['status_code'].value_counts()
            for status, count in status_counts.head(10).items():
                percentage = f"{count/total_urls*100:.1f}%"
                summary_data.append([f'Status {status}', count, percentage])
        except Exception as e:
            self.logger.warning(f"Erro processando status codes: {e}")
    
    def _add_seo_problems_stats(self, df: pd.DataFrame, summary_data: list, total_urls: int):
        """Adiciona estatísticas de problemas SEO"""
        # Títulos
        if 'title' in df.columns:
            no_title = len(df[df['title'].fillna('') == ''])
            if no_title > 0:
                summary_data.append(['URLs sem título', no_title, f"{no_title/total_urls*100:.1f}%"])
        
        # Meta descriptions
        if 'meta_description' in df.columns:
            no_meta = len(df[df['meta_description'].fillna('') == ''])
            if no_meta > 0:
                summary_data.append(['URLs sem meta description', no_meta, f"{no_meta/total_urls*100:.1f}%"])
        
        # H1s
        if 'h1_count' in df.columns:
            no_h1 = len(df[df['h1_count'].fillna(0) == 0])
            if no_h1 > 0:
                summary_data.append(['URLs sem H1', no_h1, f"{no_h1/total_urls*100:.1f}%"])
        
        # H2s
        if 'h2_count' in df.columns:
            no_h2 = len(df[df['h2_count'].fillna(0) == 0])
            if no_h2 > 0:
                summary_data.append(['URLs sem H2', no_h2, f"{no_h2/total_urls*100:.1f}%"])
        
        # Imagens sem ALT
        if 'images_without_alt' in df.columns:
            images_no_alt = len(df[df['images_without_alt'].fillna(0) > 0])
            if images_no_alt > 0:
                summary_data.append(['URLs com imagens sem ALT', images_no_alt, f"{images_no_alt/total_urls*100:.1f}%"])
    
    def _add_technical_stats(self, df: pd.DataFrame, summary_data: list, total_urls: int):
        """Adiciona estatísticas técnicas"""
        # Páginas sem canonical
        if 'canonical_url' in df.columns:
            no_canonical = len(df[df['canonical_url'].fillna('') == ''])
            if no_canonical > 0:
                summary_data.append(['URLs sem canonical', no_canonical, f"{no_canonical/total_urls*100:.1f}%"])
        
        # Páginas sem viewport
        if 'has_viewport' in df.columns:
            no_viewport = len(df[df['has_viewport'].fillna(False) == False])
            if no_viewport > 0:
                summary_data.append(['URLs sem viewport', no_viewport, f"{no_viewport/total_urls*100:.1f}%"])
        
        # Páginas lentas
        if 'response_time' in df.columns:
            slow_pages = len(df[df['response_time'].fillna(0) > 3])
            if slow_pages > 0:
                summary_data.append(['URLs lentas (>3s)', slow_pages, f"{slow_pages/total_urls*100:.1f}%"])
    
    def _add_mixed_content_stats(self, df: pd.DataFrame, summary_data: list, total_urls: int):
        """Adiciona estatísticas de Mixed Content"""
        # Total Mixed Content
        if 'total_mixed_content_count' in df.columns:
            pages_with_mixed = len(df[df['total_mixed_content_count'].fillna(0) > 0])
            if pages_with_mixed > 0:
                summary_data.append(['URLs com Mixed Content', pages_with_mixed, f"{pages_with_mixed/total_urls*100:.1f}%"])
        
        # Mixed Content crítico
        if 'active_mixed_content_count' in df.columns:
            critical_mixed = len(df[df['active_mixed_content_count'].fillna(0) > 0])
            if critical_mixed > 0:
                summary_data.append(['URLs com Mixed Content CRÍTICO', critical_mixed, f"{critical_mixed/total_urls*100:.1f}%"])
        
        # Páginas HTTPS
        if 'is_https_page' in df.columns:
            https_pages = len(df[df['is_https_page'].fillna(False) == True])
            if https_pages > 0:
                summary_data.append(['Total páginas HTTPS', https_pages, f"{https_pages/total_urls*100:.1f}%"])