"""
seofrog/exporters/sheets/__init__.py
Módulo de sheets especializadas para export Excel
"""

from .base_sheet import BaseSheet
from .summary_sheet import SummarySheet
from .status_problems_sheet import StatusProblemsSheet
from .title_problems_sheet import TitleProblemsSheet
from .meta_problems_sheet import MetaProblemsSheet
from .heading_problems_sheet import HeadingProblemsSheet
from .h1_h2_missing_sheet import H1H2MissingSheet
from .empty_headings_sheet import EmptyHeadingsSheet
from .image_problems_sheet import ImageProblemsSheet
from .technical_problems_sheet import TechnicalProblemsSheet
from .performance_sheet import PerformanceSheet
from .mixed_content_sheet import MixedContentSheet
from .technical_analysis_sheet import TechnicalAnalysisSheet

# Lista de todas as sheets disponíveis
ALL_SHEETS = [
    SummarySheet,
    StatusProblemsSheet,
    TitleProblemsSheet,
    MetaProblemsSheet,
    HeadingProblemsSheet,
    H1H2MissingSheet,
    EmptyHeadingsSheet,
    ImageProblemsSheet,
    TechnicalProblemsSheet,
    PerformanceSheet,
    MixedContentSheet,
    TechnicalAnalysisSheet
]

# Sheets por categoria
CRITICAL_SHEETS = [
    StatusProblemsSheet,
    H1H2MissingSheet,
    TitleProblemsSheet
]

SEO_SHEETS = [
    TitleProblemsSheet,
    MetaProblemsSheet,
    HeadingProblemsSheet,
    H1H2MissingSheet,
    EmptyHeadingsSheet
]

TECHNICAL_SHEETS = [
    TechnicalProblemsSheet,
    PerformanceSheet,
    MixedContentSheet,
    TechnicalAnalysisSheet
]

CONTENT_SHEETS = [
    ImageProblemsSheet,
    EmptyHeadingsSheet
]

__all__ = [
    'BaseSheet',
    'SummarySheet',
    'StatusProblemsSheet', 
    'TitleProblemsSheet',
    'MetaProblemsSheet',
    'HeadingProblemsSheet',
    'H1H2MissingSheet',
    'EmptyHeadingsSheet',
    'ImageProblemsSheet',
    'TechnicalProblemsSheet',
    'PerformanceSheet',
    'MixedContentSheet',
    'TechnicalAnalysisSheet',
    'ALL_SHEETS',
    'CRITICAL_SHEETS',
    'SEO_SHEETS',
    'TECHNICAL_SHEETS',
    'CONTENT_SHEETS'
]