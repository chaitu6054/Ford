"""Local and generation-safe Google Cloud Storage adapters."""

from __future__ import annotations

import os
import re
from dataclasses import dataclass
from pathlib import Path
from shutil import copy2
from typing import Any

from .errors import ConfigurationError, InputDiscoveryError, RemarketingError, StorageError
from .models import GCSObjectIdentity, RunManifest

XLSX_SUFFIX = ".xlsx"

try:
    from google.cloud import storage as google_storage
except ImportError:  # pragma: no cover - optional outside production.
    google_storage = None


@dataclass(frozen=True, slots=True)
class GCSObjectInfo:
    """Stable metadata for an exact GCS object generation."""

    name: str
    generation: int
    size: int = 0


class GCSStorage:
    """GCS adapter that uses generation preconditions for every mutation."""

    def __init__(self, bucket_name: str, client: Any | None = None):
        try:
            if not str(bucket_name).strip():
                raise ConfigurationError("GCS bucket name is required")
            resolved_client = client
            if resolved_client is None:
                if google_storage is None:
                    raise ConfigurationError(
                        "google-cloud-storage is required for production GCS processing"
                    )
                resolved_client = google_storage.Client()
            self._client = resolved_client
            self._bucket = self._client.bucket(bucket_name)
            self.bucket_name = str(bucket_name)
        except Exception as exc:
            if isinstance(exc, RemarketingError):
                raise
            raise ConfigurationError(f"Unable to configure GCS bucket {bucket_name!r}") from exc

    @staticmethod
    def _object_info(blob: Any) -> GCSObjectInfo:
        try:
            return GCSObjectInfo(
                name=str(blob.name),
                generation=int(blob.generation),
                size=int(blob.size or 0),
            )
        except Exception as exc:
            name = getattr(blob, "name", "<unknown>")
            raise StorageError(f"Unable to read GCS metadata for {name}") from exc

    def _list_objects(self, prefix: str) -> list[GCSObjectInfo]:
        try:
            blobs = self._client.list_blobs(self._bucket, prefix=prefix)
            return sorted(
                (
                    self._object_info(blob)
                    for blob in blobs
                    if not str(blob.name).endswith("/")
                ),
                key=lambda item: item.name,
            )
        except Exception as exc:
            if isinstance(exc, RemarketingError):
                raise
            raise StorageError(f"Unable to list GCS objects under {prefix!r}") from exc

    def list_xlsx_candidates(self, prefix: str = "input/") -> list[GCSObjectInfo]:
        """Return direct-child XLSX objects beneath the configured input prefix."""
        try:
            normalized = prefix if prefix.endswith("/") else f"{prefix}/"
            candidates: list[GCSObjectInfo] = []
            for item in self._list_objects(normalized):
                relative = (
                    item.name[len(normalized) :]
                    if item.name.startswith(normalized)
                    else item.name
                )
                if "/" not in relative and relative.lower().endswith(XLSX_SUFFIX):
                    candidates.append(item)
            return candidates
        except Exception as exc:
            if isinstance(exc, RemarketingError):
                raise
            raise StorageError(
                f"Unable to discover XLSX candidates under {prefix!r}"
            ) from exc

    def list_input_candidates(self, prefix: str = "input/") -> list[GCSObjectIdentity]:
        """Return watcher-safe object identities for direct-child XLSX inputs."""
        try:
            return [
                GCSObjectIdentity(
                    bucket_name=self.bucket_name,
                    object_name=item.name,
                    generation=item.generation,
                    size=item.size,
                )
                for item in self.list_xlsx_candidates(prefix)
            ]
        except Exception as exc:
            if isinstance(exc, RemarketingError):
                raise
            raise StorageError("Unable to list remarketing input candidates") from exc

    def list_xlsx_objects(self, prefix: str) -> list[GCSObjectInfo]:
        """Return all XLSX objects recursively beneath a prefix."""
        try:
            normalized = prefix if prefix.endswith("/") else f"{prefix}/"
            return [
                item
                for item in self._list_objects(normalized)
                if item.name.lower().endswith(XLSX_SUFFIX)
            ]
        except Exception as exc:
            if isinstance(exc, RemarketingError):
                raise
            raise StorageError(f"Unable to list XLSX objects under {prefix!r}") from exc

    def download_exact(self, object_info: GCSObjectInfo, destination: Path | str) -> Path:
        """Download exactly the observed GCS generation."""
        try:
            destination_path = Path(destination)
            destination_path.parent.mkdir(parents=True, exist_ok=True)
            blob = self._bucket.blob(object_info.name)
            blob.download_to_filename(
                str(destination_path),
                if_generation_match=object_info.generation,
            )
            return destination_path
        except Exception as exc:
            raise StorageError(
                f"Unable to download GCS object {object_info.name!r} at generation "
                f"{object_info.generation}"
            ) from exc

    def copy_exact(
        self,
        source: GCSObjectInfo,
        destination_name: str,
        destination_generation: int = 0,
    ) -> GCSObjectInfo:
        """Copy an exact source generation with a destination precondition."""
        try:
            source_blob = self._bucket.blob(source.name)
            copied = self._bucket.copy_blob(
                source_blob,
                self._bucket,
                new_name=destination_name,
                if_source_generation_match=source.generation,
                if_generation_match=destination_generation,
            )
            copied.reload()
            return self._object_info(copied)
        except Exception as exc:
            if isinstance(exc, RemarketingError):
                raise
            raise StorageError(
                f"Unable to copy GCS object {source.name!r} to {destination_name!r}"
            ) from exc

    def delete_exact(self, object_info: GCSObjectInfo) -> None:
        """Delete exactly the observed GCS generation."""
        try:
            blob = self._bucket.blob(object_info.name)
            blob.delete(if_generation_match=object_info.generation)
        except Exception as exc:
            raise StorageError(
                f"Unable to delete GCS object {object_info.name!r} at generation "
                f"{object_info.generation}"
            ) from exc

    def archive_exact(
        self,
        source: GCSObjectIdentity,
        destination_name: str,
    ) -> GCSObjectIdentity:
        """Copy, verify, then generation-delete one exact input object."""
        try:
            source_info = GCSObjectInfo(
                name=source.object_name,
                generation=source.generation,
                size=source.size,
            )
            existing = self.exists(destination_name)
            destination_generation = 0 if existing is None else existing.generation
            copied = self.copy_exact(
                source_info,
                destination_name,
                destination_generation=destination_generation,
            )
            verified = self.exists(destination_name)
            if verified is None or verified.generation != copied.generation:
                raise StorageError(f"Unable to verify archived object {destination_name!r}")
            self.delete_exact(source_info)
            return GCSObjectIdentity(
                bucket_name=self.bucket_name,
                object_name=copied.name,
                generation=copied.generation,
                size=copied.size,
            )
        except Exception as exc:
            if isinstance(exc, RemarketingError):
                raise
            raise StorageError(
                f"Unable to archive exact GCS object {source.object_name!r}"
            ) from exc

    def upload_file(
        self,
        source_path: Path | str,
        destination_name: str,
        *,
        if_generation_match: int,
    ) -> GCSObjectInfo:
        """Upload a file using an explicit destination generation precondition."""
        try:
            blob = self._bucket.blob(destination_name)
            blob.upload_from_filename(
                str(source_path),
                if_generation_match=if_generation_match,
            )
            blob.reload()
            return self._object_info(blob)
        except Exception as exc:
            if isinstance(exc, RemarketingError):
                raise
            raise StorageError(
                f"Unable to upload file to GCS object {destination_name!r}"
            ) from exc

    def upload_text(
        self,
        text: str,
        destination_name: str,
        *,
        if_generation_match: int,
    ) -> GCSObjectInfo:
        """Upload UTF-8 JSON/text with a destination generation precondition."""
        try:
            blob = self._bucket.blob(destination_name)
            blob.upload_from_string(
                text,
                content_type="application/json; charset=utf-8",
                if_generation_match=if_generation_match,
            )
            blob.reload()
            return self._object_info(blob)
        except Exception as exc:
            if isinstance(exc, RemarketingError):
                raise
            raise StorageError(
                f"Unable to upload text to GCS object {destination_name!r}"
            ) from exc

    def read_text(self, object_info: GCSObjectInfo) -> str:
        """Read exactly the observed object generation as text."""
        try:
            blob = self._bucket.blob(object_info.name)
            return blob.download_as_text(if_generation_match=object_info.generation)
        except Exception as exc:
            raise StorageError(
                f"Unable to read GCS object {object_info.name!r} at generation "
                f"{object_info.generation}"
            ) from exc

    def exists(self, object_name: str) -> GCSObjectInfo | None:
        """Return object metadata when an exact object name exists."""
        try:
            for item in self._list_objects(object_name):
                if item.name == object_name:
                    return item
            return None
        except Exception as exc:
            if isinstance(exc, RemarketingError):
                raise
            raise StorageError(f"Unable to inspect GCS object {object_name!r}") from exc

    def list_manifest_objects(self, output_prefix: str = "output/") -> list[GCSObjectInfo]:
        """Return dated SUCCESS-manifest object candidates."""
        try:
            normalized = (
                output_prefix if output_prefix.endswith("/") else f"{output_prefix}/"
            )
            pattern = re.compile(
                rf"^{re.escape(normalized)}\d{{4}}-\d{{2}}-\d{{2}}/output/"
                r"run_summary_\d{8}\.json$"
            )
            return [
                item
                for item in self._list_objects(normalized)
                if pattern.match(item.name)
            ]
        except Exception as exc:
            if isinstance(exc, RemarketingError):
                raise
            raise StorageError(
                f"Unable to list production manifests under {output_prefix!r}"
            ) from exc

    def list_success_manifests(self, output_prefix: str = "output/") -> list[RunManifest]:
        """Parse all readable SUCCESS manifests beneath the output prefix."""
        try:
            manifests: list[RunManifest] = []
            for object_info in self.list_manifest_objects(output_prefix):
                manifest = RunManifest.from_json(self.read_text(object_info))
                if manifest.is_success:
                    manifests.append(manifest)
            return manifests
        except Exception as exc:
            if isinstance(exc, RemarketingError):
                raise
            raise StorageError("Unable to list successful processing manifests") from exc


class LocalStorage:
    """Filesystem adapter that mirrors production dated output semantics."""

    def __init__(self, *, input_dir: Path | str, output_root: Path | str):
        try:
            self.input_dir = Path(input_dir)
            self.output_root = Path(output_root)
        except Exception as exc:
            raise StorageError("Unable to configure local storage") from exc

    _SAFE_SEGMENT = re.compile(r"^[A-Za-z0-9._ -]+$")

    def _output_root(self) -> str:
        return os.path.abspath(str(self.output_root))

    def _safe_output_path(self, destination_relative: str) -> str:
        """Build an absolute path under output_root from an untrusted relative path.

        Each segment is safelisted (no separators, no '..'), then the joined
        absolute path is verified to sit under output_root before it is returned.
        """
        try:
            raw = str(destination_relative)
            if os.path.isabs(raw) or raw.startswith(("/", "\\")):
                raise StorageError("Destination path must be relative")
            segments = [s for s in raw.replace("\\", "/").split("/") if s]
            if not segments:
                raise StorageError("Destination path is empty")
            for segment in segments:
                if segment in (".", "..") or not self._SAFE_SEGMENT.match(segment):
                    raise StorageError(f"Destination path contains an unsafe segment: {segment!r}")
            base_directory = self._output_root()
            candidate = os.path.abspath(os.path.join(base_directory, *segments))
            if not candidate.startswith(base_directory + os.sep):
                raise StorageError("Destination path escapes output root")
            return candidate
        except Exception as exc:
            if isinstance(exc, RemarketingError):
                raise
            raise StorageError("Unable to validate destination path") from exc

    def _safe_input_path(self, source: Path | str) -> str:
        """Accept only a direct-child .xlsx of input_dir; return its absolute path."""
        try:
            input_root = os.path.abspath(str(self.input_dir))
            file_name = os.path.basename(str(source))
            if not file_name or not self._SAFE_SEGMENT.match(file_name):
                raise StorageError("Input file name contains unsafe characters")
            if not file_name.lower().endswith(XLSX_SUFFIX):
                raise StorageError("Input file must have an XLSX extension")
            candidate = os.path.abspath(os.path.join(input_root, file_name))
            if not candidate.startswith(input_root + os.sep) or not os.path.isfile(candidate):
                raise StorageError("Input must be a direct-child file of the input directory")
            return candidate
        except Exception as exc:
            if isinstance(exc, RemarketingError):
                raise
            raise StorageError("Unable to validate input file path") from exc

    def resolve_single_input(self) -> Path | None:
        """Return zero or one direct-child XLSX, rejecting multiple candidates."""
        try:
            self.input_dir.mkdir(parents=True, exist_ok=True)
            candidates = sorted(
                path
                for path in self.input_dir.iterdir()
                if path.is_file() and path.suffix.lower() == XLSX_SUFFIX
            )
            if not candidates:
                return None
            if len(candidates) > 1:
                raise InputDiscoveryError(
                    f"Expected at most one XLSX, found {len(candidates)}"
                )
            return candidates[0]
        except Exception as exc:
            if isinstance(exc, RemarketingError):
                raise
            raise StorageError("Unable to discover local input") from exc

    def archive_input(self, source: Path, paths: Any) -> Path:
        """Archive and remove the staged local input."""
        try:
            base_directory = self._output_root()
            safe_source = self._safe_input_path(source)
            safe_destination = self._safe_output_path(str(paths.archived_input))
            if safe_destination.startswith(base_directory + os.sep):
                os.makedirs(os.path.dirname(safe_destination), exist_ok=True)
                copy2(safe_source, safe_destination)
                if safe_source != safe_destination:
                    os.remove(safe_source)
                return Path(safe_destination)
            raise StorageError("Archive destination escapes output root")
        except Exception as exc:
            if isinstance(exc, RemarketingError):
                raise
            raise StorageError(f"Unable to archive local input {source}") from exc

    def publish_file(self, source: Path, destination_relative: str) -> Path:
        """Copy a generated file into the local production-equivalent layout."""
        try:
            base_directory = self._output_root()
            safe_destination = self._safe_output_path(destination_relative)
            if safe_destination.startswith(base_directory + os.sep):
                os.makedirs(os.path.dirname(safe_destination), exist_ok=True)
                copy2(source, safe_destination)
                return Path(safe_destination)
            raise StorageError("Publish destination escapes output root")
        except Exception as exc:
            if isinstance(exc, RemarketingError):
                raise
            raise StorageError(
                f"Unable to publish local file {destination_relative}"
            ) from exc

    def publish_text(self, text: str, destination_relative: str) -> Path:
        """Write text into the local production-equivalent layout."""
        try:
            base_directory = self._output_root()
            safe_destination = self._safe_output_path(destination_relative)
            if safe_destination.startswith(base_directory + os.sep):
                os.makedirs(os.path.dirname(safe_destination), exist_ok=True)
                with open(safe_destination, "w", encoding="utf-8") as handle:
                    handle.write(text)
                return Path(safe_destination)
            raise StorageError("Publish destination escapes output root")
        except Exception as exc:
            if isinstance(exc, RemarketingError):
                raise
            raise StorageError(
                f"Unable to publish local text {destination_relative}"
            ) from exc

    def list_success_manifests(self) -> list[RunManifest]:
        """Parse local SUCCESS manifests from prior dated runs."""
        try:
            manifests: list[RunManifest] = []
            if not self.output_root.exists():
                return manifests
            for path in self.output_root.glob("*/output/run_summary_*.json"):
                manifest = RunManifest.from_json(path.read_text(encoding="utf-8"))
                if manifest.is_success:
                    manifests.append(manifest)
            return manifests
        except Exception as exc:
            if isinstance(exc, RemarketingError):
                raise
            raise StorageError("Unable to read local success manifests") from exc
