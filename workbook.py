from __future__ import annotations


from collections.abc import Iterator
from dataclasses import dataclass
from pathlib import Path, PurePosixPath
import zipfile
from xml.etree import ElementTree as ET
from defusedxml.ElementTree import ParseError, fromstring as safe_fromstring

from openpyxl.styles.numbers import BUILTIN_FORMATS

from .errors import WorkbookOpenError
from .models import CellValue, ExcelError

_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_NS = {"m": _MAIN_NS, "r": _REL_NS, "pr": _PKG_REL_NS}

_WORKBOOK_PART = "xl/workbook.xml"
_WORKBOOK_RELS_PART = "xl/_rels/workbook.xml.rels"
_SHARED_STRINGS_PART = "xl/sharedStrings.xml"
_STYLES_PART = "xl/styles.xml"
_GENERAL_FORMAT = "General"


@dataclass(frozen=True, slots=True)
class OOXMLCell:
    """Cached OOXML cell value with formula/style metadata."""

    coordinate: str
    value: CellValue
    formula: str | None
    style_id: int
    data_type: str | None


def _parse_number(raw: str) -> int | float | str:
    try:
        if not any(ch in raw for ch in ".eE"):
            return int(raw)
        return float(raw)
    except ValueError:
        return raw


def _collapse_posix_parts(resolved: str) -> str:
    """Collapse '.' and '..' segments of a package path without host-path semantics."""
    parts: list[str] = []
    for part in PurePosixPath(resolved).parts:
        if part == "..":
            if parts:
                parts.pop()
        elif part not in (".", "/"):
            parts.append(part)
    return "/".join(parts)


def _resolve_sheet_target(target: str) -> str:
    """Turn a relationship Target into a package part path under xl/."""
    target = target.replace("\\", "/")
    if target.startswith("/"):
        resolved = target.lstrip("/")
    else:
        resolved = str(PurePosixPath("xl") / target)
    return _collapse_posix_parts(resolved)


def _text_of(element: ET.Element | None) -> str:
    """Concatenate all <t> nodes below an element (shared or inline string)."""
    if element is None:
        return ""
    return "".join(t.text or "" for t in element.findall(".//m:t", _NS))


class OOXMLWorkbook:
    """Read cached Excel values directly from an XLSX OOXML package."""

    def __init__(self, path: Path | str):
        self.path = Path(path)
        try:
            self._zip = zipfile.ZipFile(self.path, "r")
        except (OSError, zipfile.BadZipFile) as exc:
            raise WorkbookOpenError(f"could not open workbook '{self.path}': {exc}") from exc
        if _WORKBOOK_PART not in self._zip.namelist():
            self._zip.close()
            raise WorkbookOpenError(f"workbook is missing {_WORKBOOK_PART}")
        self._sheet_paths = self._load_sheet_paths()
        self._shared_strings: list[str] | None = None
        self._style_formats: list[str] | None = None
        self._sheet_cache: dict[str, ET.Element] = {}

    def close(self) -> None:
        """Close the underlying XLSX ZIP package."""
        self._zip.close()

    def __enter__(self) -> "OOXMLWorkbook":
        return self

    def __exit__(self, *_args) -> None:
        self.close()

    # ------------------------------------------------------------------ parts
    def _read_part(self, part: str, what: str) -> ET.Element:
        """Read and parse one XML part of the package."""
        try:
            return safe_fromstring(self._zip.read(part))
        except (KeyError, ParseError) as exc:
            raise WorkbookOpenError(f"could not read {what} ({part}): {exc}") from exc

    def _relationship_targets(self) -> dict[str, str]:
        rels = self._read_part(_WORKBOOK_RELS_PART, "workbook relationships")
        return {
            rel.attrib["Id"]: rel.attrib.get("Target", "")
            for rel in rels.findall("pr:Relationship", _NS)
        }

    def _load_sheet_paths(self) -> dict[str, str]:
        workbook = self._read_part(_WORKBOOK_PART, "workbook")
        targets = self._relationship_targets()
        result: dict[str, str] = {}
        sheets = workbook.find("m:sheets", _NS)
        if sheets is None:
            return result
        for sheet in sheets.findall("m:sheet", _NS):
            rid = sheet.attrib.get(f"{{{_REL_NS}}}id")
            if rid and rid in targets:
                result[sheet.attrib["name"]] = _resolve_sheet_target(targets[rid])
        return result

    # ----------------------------------------------------------------- sheets
    def sheet_names(self) -> tuple[str, ...]:
        """Return workbook worksheet names in source order."""
        return tuple(self._sheet_paths)

    def sheet_path(self, name: str) -> str:
        """Resolve a worksheet name to its OOXML package part path."""
        try:
            return self._sheet_paths[name]
        except KeyError as exc:
            raise KeyError(f"worksheet '{name}' not found") from exc

    def _sheet_root(self, name: str) -> ET.Element:
        if name not in self._sheet_cache:
            self._sheet_cache[name] = self._read_part(self.sheet_path(name), f"worksheet '{name}'")
        return self._sheet_cache[name]

    # ---------------------------------------------------------------- strings
    def _load_shared_strings(self) -> list[str]:
        if self._shared_strings is not None:
            return self._shared_strings
        if _SHARED_STRINGS_PART not in self._zip.namelist():
            self._shared_strings = []
            return self._shared_strings
        root = self._read_part(_SHARED_STRINGS_PART, "shared strings")
        self._shared_strings = [_text_of(si) for si in root.findall("m:si", _NS)]
        return self._shared_strings

    def shared_string(self, index: int) -> str:
        """Resolve a shared-string table index to its text value."""
        return self._load_shared_strings()[index]

    # ------------------------------------------------------------------ cells
    def _cell_value(self, element: ET.Element, data_type: str | None, raw_v: str | None) -> CellValue:
        if data_type == "s" and raw_v is not None:
            return self.shared_string(int(raw_v))
        if data_type == "inlineStr":
            return _text_of(element.find("m:is", _NS))
        if data_type == "b":
            return raw_v == "1"
        if data_type == "e":
            return ExcelError(raw_v or "#ERROR!")
        if data_type == "str" or raw_v is None:
            return raw_v
        return _parse_number(raw_v)

    def _parse_cell(self, element: ET.Element) -> OOXMLCell:
        coordinate = element.attrib["r"]
        data_type = element.attrib.get("t")
        style_id = int(element.attrib.get("s", "0"))
        formula_el = element.find("m:f", _NS)
        formula = formula_el.text if formula_el is not None else None
        value_el = element.find("m:v", _NS)
        raw_v = value_el.text if value_el is not None else None
        value = self._cell_value(element, data_type, raw_v)
        return OOXMLCell(coordinate, value, formula, style_id, data_type)

    def read_cell(self, sheet_name: str, coordinate: str) -> OOXMLCell:
        """Read one cached cell from the requested worksheet."""
        root = self._sheet_root(sheet_name)
        element = root.find(f".//m:c[@r='{coordinate}']", _NS)
        if element is None:
            return OOXMLCell(coordinate, None, None, 0, None)
        return self._parse_cell(element)

    def iter_sheet_rows(
        self,
        sheet_name: str,
        *,
        min_row: int = 1,
    ) -> Iterator[tuple[OOXMLCell, ...]]:
        """Yield parsed worksheet rows starting at the requested row number."""
        root = self._sheet_root(sheet_name)
        sheet_data = root.find("m:sheetData", _NS)
        if sheet_data is None:
            return
        for row in sheet_data.findall("m:row", _NS):
            if int(row.attrib.get("r", "0")) < min_row:
                continue
            yield tuple(self._parse_cell(c) for c in row.findall("m:c", _NS))

    # ----------------------------------------------------------------- styles
    def _load_style_formats(self) -> list[str]:
        if self._style_formats is not None:
            return self._style_formats
        if _STYLES_PART not in self._zip.namelist():
            self._style_formats = [_GENERAL_FORMAT]
            return self._style_formats
        root = self._read_part(_STYLES_PART, "styles")
        custom = self._custom_number_formats(root)
        formats: list[str] = []
        xfs = root.find("m:cellXfs", _NS)
        if xfs is not None:
            for xf in xfs.findall("m:xf", _NS):
                fmt_id = int(xf.attrib.get("numFmtId", "0"))
                formats.append(custom.get(fmt_id, BUILTIN_FORMATS.get(fmt_id, _GENERAL_FORMAT)))
        self._style_formats = formats or [_GENERAL_FORMAT]
        return self._style_formats

    @staticmethod
    def _custom_number_formats(root: ET.Element) -> dict[int, str]:
        numfmts = root.find("m:numFmts", _NS)
        if numfmts is None:
            return {}
        return {
            int(fmt.attrib["numFmtId"]): fmt.attrib["formatCode"]
            for fmt in numfmts.findall("m:numFmt", _NS)
        }

    def style_number_format(self, style_id: int) -> str:
        """Return the number-format code for a workbook style identifier."""
        formats = self._load_style_formats()
        if 0 <= style_id < len(formats):
            return formats[style_id]
        return _GENERAL_FORMAT


WorkbookReader = OOXMLWorkbook
