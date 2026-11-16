"""Small read-only XLSX adapter for the eval's four-column input contract.

It intentionally supports the common cell types needed by raw data exports and
keeps the project dependency-free. Workbook creation is handled separately by
the spreadsheet artifact builder.
"""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path

from models import RevenueRecord


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"x": MAIN_NS, "r": REL_NS}
REQUIRED_HEADERS = {
    "customername": "customer_name",
    "month": "month",
    "amount": "amount",
    "product": "product",
}
BUILTIN_DATE_FORMAT_IDS = set(range(14, 23)) | {27, 30, 36, 45, 46, 47, 50, 57}
BUILTIN_NUMBER_FORMATS = {
    0: "General",
    1: "0",
    2: "0.00",
    3: "#,##0",
    4: "#,##0.00",
    9: "0%",
    10: "0.00%",
    11: "0.00E+00",
    12: "# ?/?",
    13: "# ??/??",
    14: "mm-dd-yy",
    15: "d-mmm-yy",
    16: "d-mmm",
    17: "mmm-yy",
    18: "h:mm AM/PM",
    19: "h:mm:ss AM/PM",
    20: "h:mm",
    21: "h:mm:ss",
    22: "m/d/yy h:mm",
    37: "#,##0;[Red](#,##0)",
    38: "#,##0;[Red](#,##0)",
    39: "#,##0.00;[Red](#,##0.00)",
    40: "#,##0.00;[Red](#,##0.00)",
    45: "mm:ss",
    46: "[h]:mm:ss",
    47: "mmss.0",
    48: "##0.0E+0",
    49: "@",
}


@dataclass(frozen=True, slots=True)
class WorkbookCell:
    """A value/formula/number-format tuple read from one XLSX cell."""

    value: object
    formula: str | None = None
    number_format: str | None = None


WorkbookCells = dict[str, dict[str, WorkbookCell]]


def _column_index(cell_reference: str) -> int:
    letters = re.match(r"[A-Z]+", cell_reference.upper())
    if not letters:
        raise ValueError(f"Invalid cell reference: {cell_reference!r}")
    result = 0
    for character in letters.group(0):
        result = result * 26 + ord(character) - ord("A") + 1
    return result - 1


def _shared_strings(archive: zipfile.ZipFile) -> list[str]:
    try:
        root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    return [
        "".join(node.text or "" for node in item.findall(".//x:t", NS))
        for item in root.findall("x:si", NS)
    ]


def _style_metadata(
    archive: zipfile.ZipFile,
) -> tuple[set[int], dict[int, str]]:
    try:
        root = ET.fromstring(archive.read("xl/styles.xml"))
    except KeyError:
        return set(), {}
    custom_formats = {
        int(item.attrib["numFmtId"]): item.attrib.get("formatCode", "")
        for item in root.findall("x:numFmts/x:numFmt", NS)
    }
    cell_formats = root.find("x:cellXfs", NS)
    if cell_formats is None:
        return set(), {}
    date_styles: set[int] = set()
    number_formats: dict[int, str] = {}
    for index, item in enumerate(cell_formats.findall("x:xf", NS)):
        format_id = int(item.attrib.get("numFmtId", "0"))
        raw_format_code = custom_formats.get(
            format_id, BUILTIN_NUMBER_FORMATS.get(format_id, "")
        )
        if raw_format_code and raw_format_code != "General":
            number_formats[index] = raw_format_code
        format_code = raw_format_code.lower()
        simplified = re.sub(r'"[^\"]*"|\\.|\[[^]]*\]', "", format_code)
        if format_id in BUILTIN_DATE_FORMAT_IDS or re.search(r"[dmy]", simplified):
            date_styles.add(index)
    return date_styles, number_formats


def _sheet_paths(archive: zipfile.ZipFile) -> list[tuple[str, str]]:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    sheets = workbook.findall("x:sheets/x:sheet", NS)
    if not sheets:
        raise ValueError("Workbook has no worksheets")
    relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {
        relationship.attrib.get("Id"): relationship.attrib["Target"]
        for relationship in relationships.findall(
            f"{{{PKG_REL_NS}}}Relationship"
        )
    }
    result: list[tuple[str, str]] = []
    for sheet in sheets:
        relationship_id = sheet.attrib[f"{{{REL_NS}}}id"]
        if relationship_id not in targets:
            raise ValueError(f"Could not resolve worksheet {sheet.attrib['name']!r}")
        target = targets[relationship_id].lstrip("/")
        sheet_path = target if target.startswith("xl/") else f"xl/{target}"
        result.append((sheet.attrib["name"], sheet_path))
    return result


def _workbook_uses_1904_dates(archive: zipfile.ZipFile) -> bool:
    root = ET.fromstring(archive.read("xl/workbook.xml"))
    properties = root.find("x:workbookPr", NS)
    return properties is not None and properties.attrib.get("date1904") in {
        "1",
        "true",
        "True",
    }


def _excel_date(value: float, uses_1904_dates: bool) -> date:
    epoch = datetime(1904, 1, 1) if uses_1904_dates else datetime(1899, 12, 30)
    return (epoch + timedelta(days=value)).date()


def _cell_value(
    cell: ET.Element,
    shared_strings: list[str],
    date_styles: set[int],
    uses_1904_dates: bool,
) -> object:
    cell_type = cell.attrib.get("t")
    style_index = int(cell.attrib.get("s", "0"))
    value_node = cell.find("x:v", NS)
    raw = value_node.text if value_node is not None else None

    if cell_type == "inlineStr":
        return "".join(node.text or "" for node in cell.findall(".//x:t", NS))
    if raw is None:
        return ""
    if cell_type == "s":
        return shared_strings[int(raw)]
    if cell_type in ("str", "e"):
        return raw
    if cell_type == "b":
        return raw == "1"
    if cell_type == "d":
        parsed = datetime.fromisoformat(raw.replace("Z", "+00:00"))
        return parsed.date()
    number = float(raw)
    if style_index in date_styles:
        return _excel_date(number, uses_1904_dates)
    return int(number) if number.is_integer() else number


def _validate_xlsx_path(path: Path) -> None:
    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        if path.suffix.lower() == ".xls":
            raise ValueError(
                "Legacy .xls files are not ZIP/XML workbooks. Save the file as "
                ".xlsx before scoring."
            )
        raise ValueError(f"Expected an .xlsx or .xlsm workbook: {path}")


def read_workbook_cells(path: Path) -> WorkbookCells:
    """Read every worksheet into ``sheet -> A1 address -> value/formula``."""
    _validate_xlsx_path(path)
    with zipfile.ZipFile(path) as archive:
        shared = _shared_strings(archive)
        date_styles, number_formats = _style_metadata(archive)
        uses_1904_dates = _workbook_uses_1904_dates(archive)
        workbook: WorkbookCells = {}
        for sheet_name, sheet_path in _sheet_paths(archive):
            sheet = ET.fromstring(archive.read(sheet_path))
            cells: dict[str, WorkbookCell] = {}
            for cell in sheet.findall("x:sheetData/x:row/x:c", NS):
                reference = cell.attrib.get("r")
                if not reference:
                    continue
                formula_node = cell.find("x:f", NS)
                formula = None
                if formula_node is not None:
                    formula = "=" + (formula_node.text or "")
                cells[reference.upper()] = WorkbookCell(
                    value=_cell_value(
                        cell, shared, date_styles, uses_1904_dates
                    ),
                    formula=formula,
                    number_format=number_formats.get(
                        int(cell.attrib.get("s", "0"))
                    ),
                )
            workbook[sheet_name] = cells
    return workbook


def read_first_sheet(path: Path) -> list[list[object]]:
    """Read the first worksheet into a rectangular list of values."""
    workbook = read_workbook_cells(path)
    first_sheet = next(iter(workbook.values()))
    row_values: dict[int, dict[int, object]] = {}
    for reference, cell in first_sheet.items():
        row_match = re.search(r"\d+$", reference)
        if not row_match:
            continue
        row_index = int(row_match.group(0))
        row_values.setdefault(row_index, {})[_column_index(reference)] = cell.value
    rows: list[list[object]] = []
    for row_index in sorted(row_values):
        values = row_values[row_index]
        width = max(values) + 1
        rows.append([values.get(index, "") for index in range(width)])
    return rows


def _normalized_header(value: object) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value).strip().lower())


def _parse_month(value: object) -> date:
    if isinstance(value, datetime):
        return date(value.year, value.month, 1)
    if isinstance(value, date):
        return date(value.year, value.month, 1)
    text = str(value).strip()
    formats = ("%Y-%m-%d", "%Y-%m", "%m/%d/%Y", "%m/%Y", "%b %Y", "%B %Y")
    for format_string in formats:
        try:
            parsed = datetime.strptime(text, format_string)
            return date(parsed.year, parsed.month, 1)
        except ValueError:
            continue
    raise ValueError(f"Unsupported Month value: {value!r}")


def load_revenue_records(path: Path) -> tuple[RevenueRecord, ...]:
    """Validate and normalize the raw Customer name/Month/Amount/Product file."""
    rows = read_first_sheet(path)
    if not rows:
        raise ValueError("Input workbook is empty")
    headers = {_normalized_header(value): index for index, value in enumerate(rows[0])}
    missing = set(REQUIRED_HEADERS) - set(headers)
    extras = set(headers) - set(REQUIRED_HEADERS)
    if missing or extras or len(headers) != 4:
        raise ValueError(
            "First sheet must contain exactly Customer name, Month, Amount, Product; "
            f"missing={sorted(missing)}, extra={sorted(extras)}"
        )

    records: list[RevenueRecord] = []
    for row_number, row in enumerate(rows[1:], start=2):
        if not any(str(value).strip() for value in row):
            continue

        def value(header: str) -> object:
            index = headers[header]
            return row[index] if index < len(row) else ""

        customer = str(value("customername")).strip()
        product = str(value("product")).strip()
        amount_value = value("amount")
        if not customer or not product:
            raise ValueError(f"Blank customer or product on Excel row {row_number}")
        if isinstance(amount_value, bool):
            raise ValueError(f"Amount is not numeric on Excel row {row_number}")
        try:
            amount = float(amount_value)
        except (TypeError, ValueError) as error:
            raise ValueError(
                f"Amount is not numeric on Excel row {row_number}: {amount_value!r}"
            ) from error
        records.append(
            RevenueRecord(
                customer_name=customer,
                month=_parse_month(value("month")),
                amount=amount,
                product=product,
            )
        )
    if not records:
        raise ValueError("Input workbook has headers but no data rows")
    return tuple(records)
