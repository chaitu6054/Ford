"""Local runner that executes the same shared pipeline without Airflow or GCS credentials.

Stages:
  main    - original report + reconciliation (+ SUCCESS manifest)
  france  - France ERA + Portefeuille workbooks only (strict: the workbook must
            carry the French Remarketing/Repossession sheets)
  all     - both (default); the France stage is skipped gracefully when the
            workbook lacks the French sheets
"""
from __future__ import annotations

import argparse
import dataclasses
import shutil
import sys
from datetime import date
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
ASTRO_ROOT = REPO_ROOT / "astro"
INCLUDE_ROOT = ASTRO_ROOT / "include"
if str(INCLUDE_ROOT) not in sys.path:
    sys.path.insert(0, str(INCLUDE_ROOT))

from remarketing.config import load_config  # noqa: E402
from remarketing.errors import InputDiscoveryError, RemarketingError  # noqa: E402
from remarketing.france_reports import generate_france_reports  # noqa: E402
from remarketing.france_reports import has_france_sheets  # noqa: E402
from remarketing.pipeline import process_local_file  # noqa: E402
from remarketing.storage import LocalStorage  # noqa: E402


def _parse_date(value: str) -> date:
    try:
        return date.fromisoformat(value)
    except Exception as exc:
        raise argparse.ArgumentTypeError("processing date must be YYYY-MM-DD") from exc


def _parse_taux(value: str) -> float:
    try:
        taux = float(value)
        if taux <= 0:
            raise ValueError("exchange rate must be positive")
        return taux
    except argparse.ArgumentTypeError:
        raise
    except Exception as exc:
        raise argparse.ArgumentTypeError(
            f"exchange rate must be a positive number: {value}"
        ) from exc


def _parser() -> argparse.ArgumentParser:
    try:
        parser = argparse.ArgumentParser(description="Run remarketing reporting locally")
        parser.add_argument("--input", type=Path)
        parser.add_argument("--input-dir", type=Path, default=REPO_ROOT / "local" / "input")
        parser.add_argument("--output-dir", type=Path, default=REPO_ROOT / "local" / "output")
        parser.add_argument("--processing-date", type=_parse_date)
        parser.add_argument(
            "--config",
            type=Path,
            default=ASTRO_ROOT / "config" / "workbook.yml",
        )
        parser.add_argument("--full-fidelity", action="store_true")
        parser.add_argument(
            "--stages",
            choices=("main", "france", "all"),
            default="all",
            help="which reporting stages to run (default: all)",
        )
        parser.add_argument(
            "--portefeuille",
            type=Path,
            default=None,
            help="optional portefeuille workbook for the France ERA stage",
        )
        parser.add_argument(
            "--taux",
            type=_parse_taux,
            default=None,
            help="optional EUR->USD rate for the France ERA USD sheet",
        )
        return parser
    except Exception:
        raise


def _stage_explicit_input(source: Path, input_dir: Path) -> Path:
    try:
        if not source.exists() or source.suffix.lower() != ".xlsx":
            raise InputDiscoveryError(f"Input must be an existing .xlsx file: {source}")
        input_dir.mkdir(parents=True, exist_ok=True)
        destination = input_dir / source.name
        if source.resolve() != destination.resolve():
            if destination.exists():
                raise InputDiscoveryError(f"Staged input already exists: {destination}")
            shutil.copy2(source, destination)
        return destination
    except RemarketingError:
        raise
    except Exception as exc:
        raise InputDiscoveryError(f"Unable to stage explicit input {source}") from exc


def _validated_portefeuille(portefeuille: Path | None) -> Path | None:
    try:
        if portefeuille is None:
            return None
        if not portefeuille.exists() or portefeuille.suffix.lower() not in {".xlsx", ".xlsm"}:
            raise InputDiscoveryError(
                f"Portefeuille must be an existing .xlsx/.xlsm file: {portefeuille}"
            )
        return portefeuille
    except RemarketingError:
        raise
    except Exception as exc:
        raise InputDiscoveryError(f"Unable to validate portefeuille {portefeuille}") from exc


def _run_france_only(args: argparse.Namespace, source: Path) -> None:
    """Run only the France ERA/Portefeuille stage (strict sheet check)."""
    try:
        if not has_france_sheets(source):
            raise InputDiscoveryError(
                "France stage requires the French 'Remarketing' and "
                f"'Repossession' sheets in {source.name}"
            )
        processing_date = args.processing_date or date.today()
        result = generate_france_reports(
            input_path=source,
            output_dir=args.output_dir,
            processing_date=processing_date,
            portefeuille_path=_validated_portefeuille(args.portefeuille),
            taux=args.taux,
        )
        print(f"France ERA report: {result.era_path}")
        print(f"France Portefeuille report: {result.portefeuille_path}")
    except RemarketingError:
        raise
    except Exception as exc:
        raise InputDiscoveryError(f"France stage failed: {exc}") from exc


def _print_main_result(args: argparse.Namespace, result) -> None:
    try:
        print(f"Status: {result.status.value}")
        if result.business_reconciliation:
            print(f"Business reconciliation: {result.business_reconciliation}")
        if result.visual_fidelity:
            print(f"Visual fidelity: {result.visual_fidelity}")
        if result.report:
            print(f"Report: {args.output_dir / result.report}")
        if result.reconciliation:
            print(f"Reconciliation: {args.output_dir / result.reconciliation}")
        if result.france_reports_skipped:
            print("France reports: skipped")
        else:
            if result.france_era_report:
                print(f"France ERA report: {args.output_dir / result.france_era_report}")
            if result.france_portefeuille_report:
                print(
                    "France Portefeuille report: "
                    f"{args.output_dir / result.france_portefeuille_report}"
                )
        if result.manifest:
            print(f"Manifest: {args.output_dir / result.manifest}")
    except Exception as exc:
        raise InputDiscoveryError(f"Unable to print processing result: {exc}") from exc


def main(argv: list[str] | None = None) -> int:
    """Run one local workbook through the production-equivalent shared pipeline."""
    try:
        args = _parser().parse_args(argv)
        args.input_dir.mkdir(parents=True, exist_ok=True)
        args.output_dir.mkdir(parents=True, exist_ok=True)
        if args.input is not None:
            source = _stage_explicit_input(args.input, args.input_dir)
        else:
            storage = LocalStorage(input_dir=args.input_dir, output_root=args.output_dir)
            source = storage.resolve_single_input()
            if source is None:
                print(f"No XLSX file found in {args.input_dir}")
                return 0
        config = load_config(
            config_path=args.config,
            overrides={"full_fidelity": bool(args.full_fidelity)},
            require_bucket=False,
        )
        if args.stages == "france":
            _run_france_only(args, source)
            return 0
        run_config = (
            config
            if args.stages == "all"
            else dataclasses.replace(config, france_reports_enabled=False)
        )
        result = process_local_file(
            config=run_config,
            input_path=source,
            output_root=args.output_dir,
            processing_date=args.processing_date,
            portefeuille_path=_validated_portefeuille(args.portefeuille),
            taux_eur_usd=args.taux,
        )
        _print_main_result(args, result)
        return 0
    except InputDiscoveryError as exc:
        print(f"ERROR: {exc}")
        return 2
    except RemarketingError as exc:
        print(f"ERROR: {exc}")
        return 2
    except Exception as exc:
        print(f"ERROR: unexpected local processing failure: {exc}")
        return 2


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except SystemExit:
        raise
    except Exception as exc:
        print(f"ERROR: {exc}")
        raise
