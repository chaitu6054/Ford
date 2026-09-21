"""Application configuration loader shared by Airflow and local execution."""
from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import yaml

from .errors import ConfigurationError

DEFAULT_CONFIG_PATH = Path(__file__).resolve().parents[2] / "config" / "workbook.yml"
DEFAULT_GCS_INPUT_BUCKET = "prj-ml-cap-rv-uk-d-tg-rv-fr-inv-input"
DEFAULT_GCS_OUTPUT_BUCKET = "prj-ml-cap-rv-uk-d-tg-rv-fr-inv-output"


def _validated_config_path(config_path: Path | str) -> Path:
    try:
        path = Path(config_path).expanduser()
        if path.suffix.lower() not in {".yml", ".yaml"}:
            raise ConfigurationError("Configuration path must be a YAML file")
        if any(part == ".." for part in path.parts):
            raise ConfigurationError("Configuration path cannot include parent traversal")
        return path.resolve()
    except ConfigurationError:
        raise
    except Exception as exc:
        raise ConfigurationError("Unable to validate configuration path") from exc


@dataclass(frozen=True, slots=True)
class AppConfig:
    """Resolved runtime configuration for the reporting application."""

    bucket_name: str | None
    output_bucket_name: str | None
    input_prefix: str
    output_prefix: str
    watcher_schedule: str
    allowed_extensions: tuple[str, ...]
    required_sheets: tuple[str, ...]
    report_sheets: tuple[str, ...]
    technical_sheets: tuple[str, ...]
    fidelity_enabled: bool
    report_filename_pattern: str
    reconciliation_filename: str
    manifest_pattern: str
    reference_dir: Path
    template_path: Path
    full_fidelity: bool = False
    france_reports_enabled: bool = True


def _nested(data: Mapping[str, Any], section: str, key: str, default: Any) -> Any:
    try:
        value = data.get(section, {})
        if not isinstance(value, Mapping):
            return default
        return value.get(key, default)
    except Exception as exc:
        raise ConfigurationError(
            f"Unable to read configuration value {section}.{key}"
        ) from exc


def load_config(
    config_path: Path | str = DEFAULT_CONFIG_PATH,
    *,
    overrides: Mapping[str, Any] | None = None,
    require_bucket: bool = False,
) -> AppConfig:
    """Load YAML defaults and apply explicit overrides."""
    try:
        path = _validated_config_path(config_path)
        data: dict[str, Any] = {}
        if path.exists():
            loaded = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
            if not isinstance(loaded, dict):
                raise ConfigurationError("workbook.yml must contain a mapping")
            data = loaded

        custom = dict(overrides or {})

        bucket_name = (
            custom.get("bucket_name")
            or _nested(data, "storage", "bucket_name", None)
            or DEFAULT_GCS_INPUT_BUCKET
        )
        output_bucket_name = (
            custom.get("output_bucket_name")
            or _nested(data, "storage", "output_bucket_name", None)
            or DEFAULT_GCS_OUTPUT_BUCKET
        )
        input_prefix = str(
            custom.get("input_prefix")
            or _nested(data, "storage", "input_prefix", "input/")
        )
        output_prefix = str(
            custom.get("output_prefix")
            or _nested(data, "storage", "output_prefix", "output/")
        )
        watcher_schedule = str(
            custom.get("watcher_schedule")
            or _nested(data, "watcher", "schedule", "*/5 * * * *")
        )

        if require_bucket and not bucket_name:
            raise ConfigurationError("GCS input bucket is required")

        root = path.resolve().parent
        return AppConfig(
            bucket_name=None if not bucket_name else str(bucket_name),
            output_bucket_name=(
                None if not output_bucket_name else str(output_bucket_name)
            ),
            input_prefix=input_prefix,
            output_prefix=output_prefix,
            watcher_schedule=watcher_schedule,
            allowed_extensions=tuple(
                str(value)
                for value in _nested(
                    data,
                    "workbook",
                    "allowed_extensions",
                    [".xlsx"],
                )
            ),
            required_sheets=tuple(
                str(value)
                for value in _nested(
                    data,
                    "workbook",
                    "required_sheets",
                    ["Remarketing"],
                )
            ),
            report_sheets=tuple(
                str(value)
                for value in _nested(data, "workbook", "report_sheets", [])
            ),
            technical_sheets=tuple(
                str(value)
                for value in _nested(data, "workbook", "technical_sheets", ["Remarketing"])
            ),
            fidelity_enabled=bool(_nested(data, "fidelity", "enabled", True)),
            report_filename_pattern=str(
                _nested(
                    data,
                    "report",
                    "filename_pattern",
                    "REMARKETING_REPOSSESSION_REPORT_{date}.xlsx",
                )
            ),
            reconciliation_filename=str(
                _nested(
                    data,
                    "report",
                    "reconciliation_filename",
                    "reconciliation.xlsx",
                )
            ),
            manifest_pattern=str(
                _nested(
                    data,
                    "report",
                    "manifest_pattern",
                    "run_summary_{date}.json",
                )
            ),
            reference_dir=(root / "reference").resolve(),
            template_path=(root / "report_template.xlsx").resolve(),
            full_fidelity=bool(custom.get("full_fidelity", False)),
            france_reports_enabled=bool(
                custom.get(
                    "france_reports_enabled",
                    _nested(data, "france_reports", "enabled", True),
                )
            ),
        )
    except ConfigurationError:
        raise
    except Exception as exc:
        raise ConfigurationError(f"Unable to load configuration from {config_path}") from exc
