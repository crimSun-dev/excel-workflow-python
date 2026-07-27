"""Reference Enricher (TDD Section 3.3).

Replicates Excel's VLOOKUP: left-joins the raw dataset against a reference
workbook (Excel/CSV) on KODE_UKER to attach MAIN_CODE and MAIN_BRANCH. Codes
absent from the reference are flagged as "UNMAPPED" and surfaced in an audit
payload so report generation is never blocked by minor mapping gaps.

Real-world BRIEFX/GL mapping workbooks use alternate column vocabulary (e.g.
``KANCA`` for branch codes, ``UNIQUE CODE`` / ``DESCRIPTION`` for the code/name
pair) and sometimes place the mapping table on a non-default sheet. To avoid
forcing users to rename columns by hand, the loader resolves the required
columns from an ordered alias registry and scans every sheet until all three
logical columns resolve, renaming matches to canonical names in-memory so the
downstream merge/aggregation/export logic stays unchanged.
"""

from __future__ import annotations

import logging
import re
from collections.abc import Sequence
from pathlib import Path

import pandas as pd

from .schemas import EnrichmentResult

logger = logging.getLogger(__name__)

MAIN_CODE = "MAIN_CODE"
MAIN_BRANCH = "MAIN_BRANCH"
UNMAPPED_SENTINEL = "UNMAPPED"

# Fraction of UNMAPPED rows above which a join is treated as "silently broken"
# and a diagnostic WARNING is emitted (never an abort - see warn_if_mostly_unmapped).
UNMAPPED_WARNING_THRESHOLD = 0.99

# Ordered alias registry (first match wins). Headers and aliases are both
# trimmed and upper-cased before comparison (see `resolve_alias_column`), so
# spaced/underscored variants such as `KODE UKER` vs `KODE_UKER` are distinct
# entries while casing and stray padding never matter.
#
# Order matters: an explicit *code* header always outranks a bare one. Real
# Briva exports carry BOTH `KODE UKER` (the branch code, e.g. 110) and `UKER`
# (the branch *name*, e.g. "KC Tulungagung"); resolving to the latter silently
# collapsed every row to UNMAPPED, so `UKER` stays last as a fallback for files
# that use it as the code.
LOOKUP_KEY_ALIASES = [
    "KODE_UKER",
    "KODE UNIT",
    "KANCA",
    "KODE SUB KANCA",
    "KODE KANCA",
    "KODE UKER",
    "UKER",
]
MAIN_CODE_ALIASES = [
    "MAIN_CODE",
    "MAIN CODE",
    "MAIN_CO",
    "KODE",
    "KODE KANCA",
    "KODE UNIT",
    "UNIQUE CODE",
    "KODE INDUK",
]
MAIN_BRANCH_ALIASES = [
    "MAIN_BRANCH",
    "MAIN BRANCH",
    "BRANCH",
    "DESC KANCA",
    "DESC UNIT",
    "SUB KANCA",
    "NAMA UKER",
    "DESCRIPTION",
]

# Ordered aliases for the master-data ID column (Qlola ID -> MAIN_CODE lookup).
# ``ID PRODUCT`` / ``ID_PRODUCT`` is the Qlola user-ID column in live
# ``MASTER DATA STATIS MALANG`` workbooks; ``CIF`` holds account codes and
# must never win when raw IDs are numeric Qlola keys.
ID_ALIASES = [
    "ID",
    "ID PRODUCT",
    "ID_PRODUCT",
    "ID USER",
    "USER ID",
    "USERID",
    "ID QLOLA",
    "CIF",
    "NO REKENING",
]

_EXCEL_SUFFIXES = (".xlsx", ".xls", ".xlsm")


class ReferenceEnrichmentError(Exception):
    """Raised when reference workbook loading fails."""


def normalize_header(name: object) -> str:
    """Canonical comparison form for a column header: trimmed and upper-cased."""
    return str(name).strip().upper()


# Pattern for pure-numeric strings with optional trailing `.0` (e.g. "1001.0", "1001.00").
_TRAILING_DOT_ZERO_RE = re.compile(r"^(\d+)\.0+$")


def canonicalize_join_id(value: object) -> str:
    """Canonicalize an ID value for join comparison.

    Applies three normalizations in order:
    1. ``str()`` + strip leading/trailing whitespace.
    2. Collapse trailing ``.0`` on pure-numeric strings (``"1001.0"`` → ``"1001"``).
       This handles the classic Excel int→float→str round-trip mismatch.
    3. For remaining pure-digit strings the value is already canonical after step 1.

    The result is always a non-empty stripped string suitable for dict-key comparison.
    """
    s = str(value).strip()
    if not s:
        return s
    m = _TRAILING_DOT_ZERO_RE.match(s)
    if m:
        return m.group(1)
    return s


def resolve_alias_column(df: pd.DataFrame, aliases: list[str]) -> str | None:
    """Returns the physical column matching the first alias that is present.

    The *full* alias list is walked in priority order, and both sides of the
    comparison are normalized (trim + case-fold via upper), so a header like
    ` kode uker ` still matches the `KODE UKER` alias.
    """
    by_normalized: dict[str, str] = {}
    for column in df.columns:
        by_normalized.setdefault(normalize_header(column), str(column))
    for alias in aliases:
        match = by_normalized.get(normalize_header(alias))
        if match is not None:
            return match
    return None


def build_unmapped_diagnostic(
    *,
    unmapped_count: int,
    total: int,
    raw_key_sample: list,
    reference_key_sample: list,
    aliases_attempted: list[str],
    context: str,
) -> str:
    """Formats the join-failure diagnostic shared by the log and the GUI dialog.

    The same text backs the WARNING record and the operator-facing messagebox,
    so what a support engineer reads in the log is what the operator saw.
    """
    fraction = (unmapped_count / total) if total else 0.0
    return (
        f"{context}: {fraction * 100:.1f}% of rows "
        f"({unmapped_count}/{total}) resolved to {UNMAPPED_SENTINEL} - the join "
        "key likely does not match.\n"
        f"  raw key sample       : {raw_key_sample}\n"
        f"  reference key sample : {reference_key_sample}\n"
        f"  aliases attempted    : {aliases_attempted}"
    )


def warn_if_mostly_unmapped(
    *,
    unmapped_count: int,
    total: int,
    raw_key_sample: list,
    reference_key_sample: list,
    aliases_attempted: list[str],
    context: str,
    threshold: float = UNMAPPED_WARNING_THRESHOLD,
) -> bool:
    """Logs a diagnostic WARNING when a join collapses almost entirely to UNMAPPED.

    Surfaces *why* the join failed (raw vs. reference key shape, aliases tried)
    without requiring a debugger. Purely advisory: it never raises, so the
    pipeline continues and still produces a report.

    Returns:
        True if the warning was emitted, False otherwise.
    """
    if total <= 0:
        return False
    fraction = unmapped_count / total
    if fraction < threshold:
        return False
    logger.warning(
        "%s",
        build_unmapped_diagnostic(
            unmapped_count=unmapped_count,
            total=total,
            raw_key_sample=raw_key_sample,
            reference_key_sample=reference_key_sample,
            aliases_attempted=aliases_attempted,
            context=context,
        ),
    )
    return True


class ReferenceEnricher:
    """Enriches the main dataset with reference mapping data (VLOOKUP equivalent)."""

    def __init__(
        self,
        reference_path: Path,
        lookup_key: str = "KODE_UKER",
        unmapped_warning_threshold: float = UNMAPPED_WARNING_THRESHOLD,
    ):
        self.reference_path = Path(reference_path)
        self.lookup_key = lookup_key.upper()
        self.unmapped_warning_threshold = unmapped_warning_threshold

    def enrich(self, raw_df: pd.DataFrame) -> EnrichmentResult:
        """Performs a left outer join to attach MAIN_CODE and MAIN_BRANCH.

        Args:
            raw_df: Ingested raw DataFrame.

        Returns:
            EnrichmentResult with the merged dataset and orphan mapping stats.

        Raises:
            ReferenceEnrichmentError: If the reference file is missing or the
                lookup key column is absent in either dataset.
        """
        ref_df = self._load_reference()

        # Resolve the join key on the raw frame too: real Briva exports keep the
        # spaced header `KODE UKER` while the configured key is `KODE_UKER`.
        raw_df = self._resolve_raw_lookup_key(raw_df)

        if self.lookup_key not in raw_df.columns:
            raise ReferenceEnrichmentError(
                f"Lookup key '{self.lookup_key}' absent from raw data. "
                f"Available columns: {list(raw_df.columns)}"
            )
        # _load_reference guarantees the canonical columns are present (or raises),
        # so these checks are a defensive safety net rather than the primary guard.
        if self.lookup_key not in ref_df.columns:
            raise ReferenceEnrichmentError(
                f"Lookup key '{self.lookup_key}' absent from reference file. "
                f"Available columns: {list(ref_df.columns)}"
            )
        for required in (MAIN_CODE, MAIN_BRANCH):
            if required not in ref_df.columns:
                raise ReferenceEnrichmentError(
                    f"Reference file missing required column '{required}'. "
                    f"Available columns: {list(ref_df.columns)}"
                )

        # Standardize the join key to string on both sides to avoid the classic
        # int-vs-str VLOOKUP mismatch (e.g. 001 vs 1).
        merged = raw_df.copy()
        merged[self.lookup_key] = merged[self.lookup_key].astype(str).str.strip()

        ref_lookup = ref_df[[self.lookup_key, MAIN_CODE, MAIN_BRANCH]].copy()
        ref_lookup[self.lookup_key] = (
            ref_lookup[self.lookup_key].astype(str).str.strip()
        )
        # Reference tables can contain duplicate keys; VLOOKUP takes the first.
        ref_lookup = ref_lookup.drop_duplicates(subset=self.lookup_key, keep="first")

        merged = merged.merge(ref_lookup, on=self.lookup_key, how="left")

        unmapped_mask = merged[MAIN_CODE].isna()
        unmapped_keys = sorted(
            merged.loc[unmapped_mask, self.lookup_key].unique().tolist()
        )
        merged[MAIN_CODE] = merged[MAIN_CODE].fillna(UNMAPPED_SENTINEL)
        merged[MAIN_BRANCH] = merged[MAIN_BRANCH].fillna(UNMAPPED_SENTINEL)

        unmapped_count = int(unmapped_mask.sum())
        matched_count = int(len(merged) - unmapped_count)

        # Advisory diagnostic only: a fully-UNMAPPED join still exports a report.
        warn_if_mostly_unmapped(
            unmapped_count=unmapped_count,
            total=len(merged),
            raw_key_sample=merged[self.lookup_key].head(5).tolist(),
            reference_key_sample=ref_lookup[self.lookup_key].head(5).tolist(),
            aliases_attempted=self._lookup_aliases(),
            context=f"Reference enrichment on '{self.lookup_key}'",
            threshold=self.unmapped_warning_threshold,
        )

        return EnrichmentResult(
            data=merged,
            matched_count=matched_count,
            unmapped_count=unmapped_count,
            unmapped_keys=unmapped_keys,
        )

    # ------------------------------------------------------------------ #
    # Loading & alias resolution
    # ------------------------------------------------------------------ #
    def _lookup_aliases(self) -> list[str]:
        """Alias list for the join key, with the configured key tried first."""
        aliases = [self.lookup_key]
        for alias in LOOKUP_KEY_ALIASES:
            if alias not in aliases:
                aliases.append(alias)
        return aliases

    def _resolve_raw_lookup_key(self, raw_df: pd.DataFrame) -> pd.DataFrame:
        """Renames a raw alias column to the configured lookup key when needed.

        Raw headers are already upper/stripped by ingestion, so `KODE UKER`
        stays spaced; if the canonical key is absent but a known alias is
        present, rename it in-memory so the downstream merge is unchanged.
        """
        if self.lookup_key in raw_df.columns:
            return raw_df
        match = resolve_alias_column(raw_df, self._lookup_aliases())
        if match is not None:
            return raw_df.rename(columns={match: self.lookup_key})
        return raw_df

    def _resolve_reference_columns(
        self, df: pd.DataFrame
    ) -> tuple[pd.DataFrame, dict[str, str]]:
        """Renames the first alias match for each logical column to its canonical name.

        Args:
            df: Reference sheet with already-normalized (upper/stripped) headers.

        Returns:
            A tuple of (possibly-renamed DataFrame, mapping of canonical column
            name -> physical column that satisfied it). A logical column absent
            from ``resolved`` could not be matched to any alias.
        """
        df = df.copy()
        resolved: dict[str, str] = {}
        for canonical, aliases in (
            (self.lookup_key, self._lookup_aliases()),
            (MAIN_CODE, MAIN_CODE_ALIASES),
            (MAIN_BRANCH, MAIN_BRANCH_ALIASES),
        ):
            match = resolve_alias_column(df, aliases)
            if match is None:
                continue
            resolved[canonical] = match
            if match != canonical:
                df = df.rename(columns={match: canonical})
        return df, resolved

    def _load_reference(self) -> pd.DataFrame:
        """Loads the reference workbook (.xlsx/.xls/.xlsm/.csv) into a DataFrame.

        Excel workbooks are scanned sheet-by-sheet; the first sheet whose columns
        resolve to the lookup key, MAIN_CODE, and MAIN_BRANCH is used.
        """
        if not self.reference_path.exists():
            raise ReferenceEnrichmentError(
                f"Reference file not found: {self.reference_path}"
            )
        suffix = self.reference_path.suffix.lower()
        try:
            if suffix in _EXCEL_SUFFIXES:
                return self._load_excel_reference()
            if suffix == ".csv":
                return self._load_csv_reference()
            raise ReferenceEnrichmentError(
                f"Unsupported reference file type: '{suffix}'"
            )
        except ReferenceEnrichmentError:
            raise
        except Exception as exc:  # noqa: BLE001 - re-wrapped for the caller
            raise ReferenceEnrichmentError(
                f"Failed to load reference file {self.reference_path}: {exc}"
            ) from exc

    def _load_csv_reference(self) -> pd.DataFrame:
        ref_df = pd.read_csv(self.reference_path, dtype=str)
        ref_df.columns = [str(c).strip().upper() for c in ref_df.columns]
        resolved_df, resolved = self._resolve_reference_columns(ref_df)
        if self._is_fully_resolved(resolved):
            return resolved_df
        raise self._unresolvable_error([("(csv)", list(ref_df.columns))])

    def _load_excel_reference(self) -> pd.DataFrame:
        excel_file = pd.ExcelFile(self.reference_path)
        attempts: list[tuple[str, list[str]]] = []
        candidates: list[tuple[int, str, pd.DataFrame, dict[str, str]]] = []
        for sheet_name in excel_file.sheet_names:
            sheet_df = pd.read_excel(excel_file, sheet_name=sheet_name, dtype=str)
            sheet_df.columns = [str(c).strip().upper() for c in sheet_df.columns]
            resolved_df, resolved = self._resolve_reference_columns(sheet_df)
            if self._is_fully_resolved(resolved):
                score = self._sheet_preference_score(str(sheet_name), resolved)
                candidates.append((score, str(sheet_name), resolved_df, resolved))
            else:
                attempts.append((str(sheet_name), list(sheet_df.columns)))
        if candidates:
            candidates.sort(key=lambda item: item[0], reverse=True)
            return candidates[0][2]
        raise self._unresolvable_error(attempts)

    def _sheet_preference_score(
        self, sheet_name: str, resolved: dict[str, str]
    ) -> int:
        """Rank resolving sheets so UKER/unit-level tabs beat generic Unit Kerja tabs."""
        score = 0
        upper_name = sheet_name.upper()
        if "UKER" in upper_name:
            score += 100
        if "UNIT KERJA" in upper_name:
            score -= 10
        lookup_source = resolved.get(self.lookup_key, "")
        lookup_rank = {
            "KODE UNIT": 40,
            "KODE SUB KANCA": 30,
            "KANCA": 20,
            "KODE KANCA": 10,
        }
        score += lookup_rank.get(lookup_source, 0)
        return score

    def _is_fully_resolved(self, resolved: dict[str, str]) -> bool:
        return all(col in resolved for col in (self.lookup_key, MAIN_CODE, MAIN_BRANCH))

    def _unresolvable_error(
        self, attempts: list[tuple[str, list[str]]]
    ) -> ReferenceEnrichmentError:
        """Builds an actionable error listing sheets scanned and aliases attempted."""
        sheet_lines = "\n".join(
            f"  - sheet '{name}': columns {cols}" for name, cols in attempts
        )
        return ReferenceEnrichmentError(
            "Unable to resolve required reference columns "
            f"(lookup key '{self.lookup_key}', '{MAIN_CODE}', '{MAIN_BRANCH}') "
            f"in {self.reference_path}.\n"
            f"Sheets scanned:\n{sheet_lines}\n"
            "Aliases attempted:\n"
            f"  - lookup key: {self._lookup_aliases()}\n"
            f"  - {MAIN_CODE}: {MAIN_CODE_ALIASES}\n"
            f"  - {MAIN_BRANCH}: {MAIN_BRANCH_ALIASES}"
        )


# When raw IDs are supplied, overlap with the master keys dominates sheet-name
# heuristics so a decoy ``ID`` column (e.g. CIF/account codes) cannot beat the
# real user-ID column on the same sheet.
_PROBE_OVERLAP_WEIGHT = 10_000


class MasterDataEnricher:
    """Resolves an ID -> MAIN_CODE mapping from a master-data workbook.

    Used by the Qlola workflow to key its final crosstab on the master
    MAIN_CODE rather than the UKER-enriched MAIN_CODE. Columns are resolved via
    the shared alias registries (ID and MAIN_CODE), scanning every sheet — the
    manual "VLOOKUP column index 21" note is treated as a discovery hint, never
    as the shipped contract.

    When multiple sheets or column pairs resolve ID + MAIN_CODE aliases, the
    enricher scores candidates by (1) overlap with optional ``probe_ids`` from
    the raw/aggregated side, then (2) sheet/column preference heuristics.
    """

    def __init__(self, master_path: Path):
        self.master_path = Path(master_path)

    def build_id_to_main_code(
        self, probe_ids: Sequence[str] | None = None
    ) -> dict[str, str]:
        """Returns a {canonicalized ID: MAIN_CODE} mapping (best sheet wins).

        Args:
            probe_ids: Optional raw-side user IDs. When provided, the column pair
                whose keys overlap the most with these IDs is preferred — this
                disambiguates workbooks that expose both a decoy ``ID`` column
                (CIF / account codes) and the real Qlola user-ID column.

        Raises:
            ReferenceEnrichmentError: if no sheet resolves both an ID column and
                a MAIN_CODE column via aliases, listing sheets and aliases tried.
        """
        if not self.master_path.exists():
            raise ReferenceEnrichmentError(
                f"Master-data file not found: {self.master_path}"
            )
        suffix = self.master_path.suffix.lower()
        try:
            if suffix in _EXCEL_SUFFIXES:
                frames = self._read_excel_sheets()
            elif suffix == ".csv":
                frames = [("(csv)", self._read_csv())]
            else:
                raise ReferenceEnrichmentError(
                    f"Unsupported master-data file type: '{suffix}'"
                )
        except ReferenceEnrichmentError:
            raise
        except Exception as exc:  # noqa: BLE001 - re-wrapped for the caller
            raise ReferenceEnrichmentError(
                f"Failed to load master-data file {self.master_path}: {exc}"
            ) from exc

        attempts: list[tuple[str, list[str]]] = []
        candidates: list[tuple[int, str, pd.DataFrame, str, str, dict[str, str]]] = []
        for sheet_name, df in frames:
            id_cols = self._all_alias_columns(df, ID_ALIASES)
            code_cols = self._all_alias_columns(df, MAIN_CODE_ALIASES)
            if id_cols and code_cols:
                for id_col in id_cols:
                    for code_col in code_cols:
                        mapping = self._mapping_from_columns(df, id_col, code_col)
                        overlap = self._probe_overlap_score(mapping, probe_ids)
                        preference = self._sheet_preference_score(
                            str(sheet_name), id_col, code_col
                        )
                        score = overlap * _PROBE_OVERLAP_WEIGHT + preference
                        candidates.append(
                            (score, str(sheet_name), df, id_col, code_col, mapping)
                        )
            else:
                attempts.append((sheet_name, list(df.columns)))

        if candidates:
            candidates.sort(key=lambda item: item[0], reverse=True)
            return candidates[0][5]

        raise self._unresolvable_error(attempts)

    @staticmethod
    def _mapping_from_columns(
        df: pd.DataFrame, id_col: str, code_col: str
    ) -> dict[str, str]:
        """Build a canonical {ID: MAIN_CODE} dict from one column pair."""
        frame = df[[id_col, code_col]].copy()
        frame[id_col] = frame[id_col].apply(canonicalize_join_id)
        frame[code_col] = frame[code_col].astype(str).str.strip()
        frame = frame[
            frame[id_col].notna()
            & (frame[id_col] != "")
            & (frame[id_col].str.lower() != "nan")
        ]
        frame = frame.drop_duplicates(subset=id_col, keep="first")
        return dict(zip(frame[id_col], frame[code_col]))

    @staticmethod
    def _probe_overlap_score(
        mapping: dict[str, str], probe_ids: Sequence[str] | None
    ) -> int:
        if not probe_ids:
            return 0
        probe_set = {
            canonicalize_join_id(pid)
            for pid in probe_ids
            if str(pid).strip() and str(pid).strip().lower() != "nan"
        }
        if not probe_set:
            return 0
        return sum(1 for pid in probe_set if pid in mapping)

    @staticmethod
    def _sheet_preference_score(
        sheet_name: str, id_col: str, code_col: str
    ) -> int:
        """Rank resolving sheets so the true ID→MAIN_CODE mapping tab wins.

        Higher score = better candidate. Heuristics mirror the UKER reference
        pattern: sheets whose names contain common master-data keywords are
        preferred, and more specific alias matches score higher.
        """
        score = 0
        upper_name = sheet_name.upper()
        # Sheets named after the master domain are preferred.
        if "MASTER" in upper_name or "STATIS" in upper_name:
            score += 100
        if "DATA" in upper_name:
            score += 50
        # Penalize generic cover/summary sheets.
        if "COVER" in upper_name or "SUMMARY" in upper_name:
            score -= 20
        # More specific ID alias matches outrank generic ones.
        id_rank = {
            "ID": 40,
            "ID PRODUCT": 38,
            "ID_PRODUCT": 38,
            "ID USER": 35,
            "USER ID": 30,
            "USERID": 25,
            "ID QLOLA": 20,
            "CIF": 5,
        }
        score += id_rank.get(id_col, 0)
        # More specific MAIN_CODE alias matches outrank generic ones.
        code_rank = {
            "MAIN_CODE": 40,
            "MAIN CODE": 35,
            "KODE INDUK": 30,
        }
        score += code_rank.get(code_col, 0)
        return score

    def _read_csv(self) -> pd.DataFrame:
        df = pd.read_csv(self.master_path, dtype=str)
        df.columns = [str(c).strip().upper() for c in df.columns]
        return df

    def _read_excel_sheets(self) -> list[tuple[str, pd.DataFrame]]:
        excel_file = pd.ExcelFile(self.master_path)
        frames: list[tuple[str, pd.DataFrame]] = []
        for sheet_name in excel_file.sheet_names:
            df = pd.read_excel(excel_file, sheet_name=sheet_name, dtype=str)
            df.columns = [str(c).strip().upper() for c in df.columns]
            frames.append((str(sheet_name), df))
        return frames

    @staticmethod
    def _first_alias(df: pd.DataFrame, aliases: list[str]) -> str | None:
        return resolve_alias_column(df, aliases)

    @staticmethod
    def _all_alias_columns(df: pd.DataFrame, aliases: list[str]) -> list[str]:
        """Physical columns matching ``aliases``, in alias priority order."""
        by_normalized: dict[str, str] = {}
        for column in df.columns:
            by_normalized.setdefault(normalize_header(column), str(column))
        found: list[str] = []
        seen: set[str] = set()
        for alias in aliases:
            match = by_normalized.get(normalize_header(alias))
            if match is not None and match not in seen:
                found.append(match)
                seen.add(match)
        return found

    def _unresolvable_error(
        self, attempts: list[tuple[str, list[str]]]
    ) -> ReferenceEnrichmentError:
        sheet_lines = "\n".join(
            f"  - sheet '{name}': columns {cols}" for name, cols in attempts
        )
        return ReferenceEnrichmentError(
            "Unable to resolve required master-data columns "
            f"('ID', '{MAIN_CODE}') in {self.master_path}.\n"
            f"Sheets scanned:\n{sheet_lines}\n"
            "Aliases attempted:\n"
            f"  - ID: {ID_ALIASES}\n"
            f"  - {MAIN_CODE}: {MAIN_CODE_ALIASES}"
        )

