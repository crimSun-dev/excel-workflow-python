"""Phase 1 unit tests: Ingestion, Enrichment, Aggregation (TDD Section 4)."""

from __future__ import annotations

from decimal import Decimal

import pandas as pd
import pytest

from src.aggregation import AggregationEngine
from src.enrichment import ReferenceEnricher, ReferenceEnrichmentError
from src.ingestion import DataIngestionError, IngestionEngine


# --------------------------------------------------------------------------- #
# Ingestion
# --------------------------------------------------------------------------- #
def test_ingest_reads_valid_pipe_delimited_file(raw_pipe_file):
    result = IngestionEngine().read_raw_data(raw_pipe_file)
    assert list(result.data.columns) == ["KODE_UKER", "SEGMEN", "VOLUME_IN_IDR"]
    assert result.total_rows == 4
    assert result.malformed_rows_count == 0
    assert result.data["VOLUME_IN_IDR"].dtype == "float64"


def test_ingest_handles_utf8_and_cp1252_encodings(raw_pipe_file_cp1252):
    result = IngestionEngine().read_raw_data(raw_pipe_file_cp1252)
    assert result.total_rows == 1
    assert "Wholesal" in result.data["SEGMEN"].iloc[0]


def test_ingest_strips_leading_trailing_whitespace(raw_pipe_file_whitespace):
    result = IngestionEngine().read_raw_data(raw_pipe_file_whitespace)
    assert result.data["KODE_UKER"].iloc[0] == "0001"
    assert result.data["SEGMEN"].iloc[0] == "Wholesale"


def test_ingest_skips_and_counts_malformed_rows(raw_pipe_file_malformed):
    result = IngestionEngine().read_raw_data(raw_pipe_file_malformed)
    assert result.total_rows == 2  # 3 data rows, 1 malformed skipped
    assert result.malformed_rows_count == 1


def test_ingest_empty_file_raises(tmp_path):
    empty = tmp_path / "empty.txt"
    empty.write_text("", encoding="utf-8")
    with pytest.raises(DataIngestionError):
        IngestionEngine().read_raw_data(empty)


def test_ingest_missing_file_raises(tmp_path):
    with pytest.raises(DataIngestionError):
        IngestionEngine().read_raw_data(tmp_path / "does_not_exist.txt")


# --------------------------------------------------------------------------- #
# Enrichment
# --------------------------------------------------------------------------- #
def test_enrichment_matches_valid_kode_uker_keys(raw_pipe_file, reference_file):
    raw = IngestionEngine().read_raw_data(raw_pipe_file).data
    result = ReferenceEnricher(reference_file).enrich(raw)
    row = result.data[result.data["KODE_UKER"] == "0001"].iloc[0]
    assert row["MAIN_CODE"] == "MC10"
    assert row["MAIN_BRANCH"] == "Jakarta Pusat"
    assert result.matched_count == 3  # two 0001 rows + one 0002


def test_enrichment_flags_unmapped_kode_uker_keys(raw_pipe_file, reference_file):
    raw = IngestionEngine().read_raw_data(raw_pipe_file).data
    result = ReferenceEnricher(reference_file).enrich(raw)
    assert result.unmapped_count == 1  # 0003 not in reference
    assert "0003" in result.unmapped_keys
    orphan = result.data[result.data["KODE_UKER"] == "0003"].iloc[0]
    assert orphan["MAIN_CODE"] == "UNMAPPED"
    assert orphan["MAIN_BRANCH"] == "UNMAPPED"


def test_enrichment_missing_lookup_key_raises(reference_file):
    df = pd.DataFrame({"WRONG_KEY": ["0001"], "VOLUME_IN_IDR": [1.0]})
    with pytest.raises(ReferenceEnrichmentError):
        ReferenceEnricher(reference_file).enrich(df)


def test_enrichment_missing_reference_file_raises(tmp_path, raw_pipe_file):
    raw = IngestionEngine().read_raw_data(raw_pipe_file).data
    with pytest.raises(ReferenceEnrichmentError):
        ReferenceEnricher(tmp_path / "nope.xlsx").enrich(raw)


# --------------------------------------------------------------------------- #
# Aggregation
# --------------------------------------------------------------------------- #
def test_aggregation_groups_by_code_and_sums_volume(raw_pipe_file, reference_file):
    raw = IngestionEngine().read_raw_data(raw_pipe_file).data
    enriched = ReferenceEnricher(reference_file).enrich(raw).data
    result = AggregationEngine().aggregate(enriched)

    mc10 = result.summary_data[result.summary_data["MAIN_CODE"] == "MC10"].iloc[0]
    # 1_250_000_000_000 + 250_000_000_000 = 1_500_000_000_000
    assert mc10["VOLUME_IN_IDR"] == 1_500_000_000_000.0
    assert result.branch_count == 3  # MC10, MC20, UNMAPPED


def test_aggregation_monetary_precision_matches_decimal(raw_pipe_file, reference_file):
    raw = IngestionEngine().read_raw_data(raw_pipe_file).data
    enriched = ReferenceEnricher(reference_file).enrich(raw).data
    result = AggregationEngine().aggregate(enriched)
    reference_total = Decimal("1250000000000") + Decimal("500000000.50") + Decimal(
        "250000000000"
    ) + Decimal("1000000")
    assert result.total_volume_idr == float(round(reference_total, 2))


def test_aggregation_applies_segment_filter(raw_pipe_file, reference_file):
    raw = IngestionEngine().read_raw_data(raw_pipe_file).data
    enriched = ReferenceEnricher(reference_file).enrich(raw).data
    result = AggregationEngine(segment_filter="Wholesale").aggregate(enriched)
    # Only the two Wholesale rows (both 0001 -> MC10) survive the filter.
    assert result.branch_count == 1
    assert result.summary_data.iloc[0]["MAIN_CODE"] == "MC10"
    assert result.total_volume_idr == 1_500_000_000_000.0
