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

from pathlib import Path

import pandas as pd

from .schemas import EnrichmentResult

MAIN_CODE = "MAIN_CODE"
MAIN_BRANCH = "MAIN_BRANCH"
UNMAPPED_SENTINEL = "UNMAPPED"

# Ordered alias registry (first match wins). All names are compared uppercase,
# matching the header normalization applied during load.
LOOKUP_KEY_ALIASES = [
    "KODE_UKER",
    "KODE UNIT",
    "KANCA",
    "KODE SUB KANCA",
    "KODE KANCA",
    "UKER",
    "KODE UKER",
]
MAIN_CODE_ALIASES = [
    "MAIN_CODE",
    "MAIN CODE",
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
ID_ALIASES = [
    "ID",
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


class ReferenceEnricher:
    """Enriches the main dataset with reference mapping data (VLOOKUP equivalent)."""

    def __init__(self, reference_path: Path, lookup_key: str = "KODE_UKER"):
        self.reference_path = Path(reference_path)
        self.lookup_key = lookup_key.upper()

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
        for alias in self._lookup_aliases():
            if alias in raw_df.columns:
                return raw_df.rename(columns={alias: self.lookup_key})
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
            match = next((alias for alias in aliases if alias in df.columns), None)
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


class MasterDataEnricher:
    """Resolves an ID -> MAIN_CODE mapping from a master-data workbook.

    Used by the Qlola workflow to key its final crosstab on the master
    MAIN_CODE rather than the UKER-enriched MAIN_CODE. Columns are resolved via
    the shared alias registries (ID and MAIN_CODE), scanning every sheet — the
    manual "VLOOKUP column index 21" note is treated as a discovery hint, never
    as the shipped contract.
    """

    def __init__(self, master_path: Path):
        self.master_path = Path(master_path)

    def build_id_to_main_code(self) -> dict[str, str]:
        """Returns a {stripped-str ID: MAIN_CODE} mapping (first key wins).

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
        for sheet_name, df in frames:
            id_col = self._first_alias(df, ID_ALIASES)
            code_col = self._first_alias(df, MAIN_CODE_ALIASES)
            if id_col is not None and code_col is not None:
                mapping = (
                    df[[id_col, code_col]]
                    .assign(
                        **{
                            id_col: df[id_col].astype(str).str.strip(),
                            code_col: df[code_col].astype(str).str.strip(),
                        }
                    )
                    .drop_duplicates(subset=id_col, keep="first")
                )
                return dict(zip(mapping[id_col], mapping[code_col]))
            attempts.append((sheet_name, list(df.columns)))

        raise self._unresolvable_error(attempts)

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
        return next((alias for alias in aliases if alias in df.columns), None)

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
