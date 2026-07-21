"""Automated Financial Data Processing & Excel Summarization Pipeline.

A Layered ETL (Extract, Transform, Load) pipeline that replicates the manual
Excel workflow of:
    1. Text-to-Columns pipe parsing        -> IngestionEngine
    2. VLOOKUP reference enrichment         -> ReferenceEnricher
    3. PivotTable summarization             -> AggregationEngine
    4. Formatted report export              -> ExcelReportExporter
"""

__version__ = "1.0.0"
