"""France report stage: ERA and Portefeuille workbooks.

This module bridges the main remarketing pipeline and the France inventory
calculations (``remarketing.france``). The main report already covers the
Remarketing (R1) tables, so only two additional workbooks are produced:

- ``france_era_<YYYY-MM>.xlsx`` — R2 tables (ERA INPUT/OUTPUT, taux)
- ``france_portefeuille_<YYYY-MM>.xlsx`` — R3 tables (comptes à échéance)

Both workbooks also carry the shared ``Resume`` and ``Exceptions`` sheets.
"""
from __future__ import annotations

import contextlib
import logging
from calendar import monthrange
from collections.abc import Iterator
from dataclasses import dataclass
from datetime import date
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.packaging.relationship import get_rel
from openpyxl.pivot.cache import CacheDefinition
from openpyxl.reader.workbook import WorkbookParser

from .errors import ReportGenerationError, RemarketingError
from .france import config as france_config
from .france import pipeline as france_pipeline

LOGGER = logging.getLogger(__name__)

SHARED_SHEETS = ("Resume", "Exceptions")


def _fast_pivot_caches(parser: object) -> dict:
    """Return pivot cache definitions without loading their record payloads."""
    try:
        caches: dict = {}
        for entry in parser.caches:  # type: ignore[attr-defined]
            cache = get_rel(
                parser.archive,  # type: ignore[attr-defined]
                parser.rels,  # type: ignore[attr-defined]
                id=entry.id,
                cls=CacheDefinition,
            )
            caches[entry.cacheId] = cache
        return caches
    except Exception:
        raise


@contextlib.contextmanager
def _without_pivot_cache_records() -> Iterator[None]:
    """Load workbooks while skipping pivot-cache records and external links.

    Some analyst workbooks embed very large pivot caches and external-link
    caches; openpyxl parses them on every ``load_workbook`` call, which can
    stall for many minutes. The France calculations only need cell values and
    fills and never touch pivot data or external links, so both are skipped
    while this context is active (openpyxl's supported ``keep_links`` switch
    plus a records-free pivot-cache loader). The stock behaviour is restored
    on exit. Not thread-safe by design (single-threaded runners).
    """
    try:
        original_caches = WorkbookParser.pivot_caches
        original_init = WorkbookParser.__init__
        WorkbookParser.pivot_caches = property(_fast_pivot_caches)

        def _init_without_links(self, archive, workbook_part_name, keep_links=True):
            try:
                original_init(self, archive, workbook_part_name, keep_links=False)
            except Exception:
                raise

        WorkbookParser.__init__ = _init_without_links  # type: ignore[method-assign]
        try:
            yield
        finally:
            WorkbookParser.pivot_caches = original_caches
            WorkbookParser.__init__ = original_init  # type: ignore[method-assign]
    except Exception:
        raise


@dataclass(frozen=True, slots=True)
class FranceReportsResult:
    """Paths of the two France workbooks produced for one processing run."""

    era_path: Path | None
    portefeuille_path: Path | None
    skipped: bool = False


def has_france_sheets(input_path: Path) -> bool:
    """Return True when the workbook holds the French Remarketing sheets."""
    try:
        with _without_pivot_cache_records():
            workbook = load_workbook(str(input_path), read_only=True, data_only=True)
        try:
            names = set(workbook.sheetnames)
        finally:
            workbook.close()
        return (
            france_config.FEUILLE_REMARKETING in names
            and france_config.FEUILLE_REPOSSESSION in names
        )
    except Exception:
        # An unreadable workbook cannot feed the France calculations; the main
        # pipeline stage has already validated readability before this point.
        return False


def _month_end(processing_date: date) -> date:
    try:
        last = monthrange(processing_date.year, processing_date.month)[1]
        return date(processing_date.year, processing_date.month, last)
    except Exception as exc:
        raise ReportGenerationError("Unable to resolve month-end date") from exc


def _split_tables(
    tables: dict[str, object],
) -> tuple[dict[str, object], dict[str, object]]:
    try:
        shared = {key: tables[key] for key in SHARED_SHEETS if key in tables}
        era = {key: value for key, value in tables.items() if key.startswith("R2")}
        era.update(shared)
        portefeuille = {
            key: value for key, value in tables.items() if key.startswith("R3")
        }
        return era, portefeuille
    except Exception as exc:
        raise ReportGenerationError("Unable to split France report tables") from exc


def generate_france_reports(
    *,
    input_path: Path,
    output_dir: Path,
    processing_date: date,
    portefeuille_path: Path | None = None,
    taux: float | None = None,
) -> FranceReportsResult:
    """Generate the ERA and Portefeuille workbooks from the analyst workbook.

    Raises ReportGenerationError when the France calculations fail; callers
    that want a graceful skip must check ``has_france_sheets`` first.
    """
    try:
        source = Path(input_path)
        out_dir = Path(output_dir)
        out_dir.mkdir(parents=True, exist_ok=True)
        month = processing_date.strftime("%Y-%m")
        with _without_pivot_cache_records():
            tables = france_pipeline.generer_rapports(
                str(source),
                _month_end(processing_date),
                str(portefeuille_path) if portefeuille_path is not None else None,
                taux,
            )
        era_tables, portefeuille_tables = _split_tables(tables)
        era_path = out_dir / f"france_era_{month}.xlsx"
        portefeuille_path_out = out_dir / f"france_portefeuille_{month}.xlsx"
        france_pipeline.ecrire_classeur(era_tables, str(era_path))
        france_pipeline.ecrire_classeur(portefeuille_tables, str(portefeuille_path_out))
        LOGGER.info(
            "France reports generated: %s, %s", era_path.name, portefeuille_path_out.name
        )
        return FranceReportsResult(
            era_path=era_path, portefeuille_path=portefeuille_path_out
        )
    except RemarketingError:
        raise
    except Exception as exc:
        raise ReportGenerationError(
            f"France report generation failed for {input_path}"
        ) from exc


def skipped_france_reports() -> FranceReportsResult:
    """Return the canonical skipped result for the France stage."""
    try:
        return FranceReportsResult(
            era_path=None, portefeuille_path=None, skipped=True
        )
    except Exception as exc:
        raise ReportGenerationError("Unable to build skipped France result") from exc
