from __future__ import annotations

import json
import os
import posixpath
import re
from dataclasses import dataclass
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import unquote
from xml.etree import ElementTree as ET
from zipfile import ZipFile

BASE_DIR = Path(__file__).parent
XLSX_PATH = BASE_DIR / "CASH FLOW SOURCE-APRIL26.xlsx"
STATIC_DIR = BASE_DIR / "static"

NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


@dataclass
class SheetData:
    name: str
    columns: list[str]
    rows: list[dict[str, Any]]


def _column_index(cell_ref: str) -> int:
    letters = "".join(ch for ch in cell_ref if ch.isalpha())
    idx = 0
    for char in letters:
        idx = idx * 26 + (ord(char.upper()) - ord("A") + 1)
    return idx - 1


def _cell_value(cell: ET.Element, shared_strings: list[str]) -> Any:
    cell_type = cell.attrib.get("t")
    value_node = cell.find("main:v", NS)
    inline_node = cell.find("main:is/main:t", NS)

    if inline_node is not None:
        return inline_node.text or ""

    if value_node is None or value_node.text is None:
        return ""

    raw = value_node.text

    if cell_type == "s":
        return shared_strings[int(raw)] if raw.isdigit() else raw
    if cell_type == "b":
        return raw == "1"

    if raw and re.fullmatch(r"-?\d+", raw):
        return int(raw)

    if raw and re.fullmatch(r"-?\d+\.\d+", raw):
        return float(raw)

    return raw


def _read_shared_strings(zip_file: ZipFile) -> list[str]:
    try:
        content = zip_file.read("xl/sharedStrings.xml")
    except KeyError:
        return []

    root = ET.fromstring(content)
    strings: list[str] = []
    for item in root.findall("main:si", NS):
        texts = [node.text or "" for node in item.findall(".//main:t", NS)]
        strings.append("".join(texts))
    return strings


def _sheet_map(zip_file: ZipFile) -> list[tuple[str, str]]:
    workbook = ET.fromstring(zip_file.read("xl/workbook.xml"))
    rels = ET.fromstring(zip_file.read("xl/_rels/workbook.xml.rels"))

    rel_lookup = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in rels.findall("rel:Relationship", NS)
    }

    sheets: list[tuple[str, str]] = []
    for sheet in workbook.findall("main:sheets/main:sheet", NS):
        name = sheet.attrib["name"]
        rel_id = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
        target = rel_lookup[rel_id].lstrip("/")
        if not target.startswith("xl/"):
            target = f"xl/{target}"
        sheets.append((name, target))
    return sheets


def _sheet_rows(zip_file: ZipFile, path: str, shared_strings: list[str]) -> list[list[Any]]:
    root = ET.fromstring(zip_file.read(path))
    all_rows: list[list[Any]] = []

    for row in root.findall("main:sheetData/main:row", NS):
        cells = row.findall("main:c", NS)
        if not cells:
            continue

        expanded_row: list[Any] = []
        current_col = 0

        for cell in cells:
            ref = cell.attrib.get("r", "")
            target_col = _column_index(ref) if ref else current_col
            while current_col < target_col:
                expanded_row.append("")
                current_col += 1

            expanded_row.append(_cell_value(cell, shared_strings))
            current_col += 1

        if any(value != "" for value in expanded_row):
            all_rows.append(expanded_row)

    return all_rows


def load_workbook_data() -> list[SheetData]:
    if not XLSX_PATH.exists():
        raise FileNotFoundError(f"Spreadsheet not found: {XLSX_PATH}")

    with ZipFile(XLSX_PATH) as zip_file:
        shared_strings = _read_shared_strings(zip_file)
        sheet_paths = _sheet_map(zip_file)

        sheets: list[SheetData] = []
        for sheet_name, path in sheet_paths:
            raw_rows = _sheet_rows(zip_file, path, shared_strings)
            if not raw_rows:
                sheets.append(SheetData(name=sheet_name, columns=[], rows=[]))
                continue

            header_row = next((row for row in raw_rows if any(str(c).strip() for c in row)), [])
            headers = [str(col).strip() if str(col).strip() else f"Column {i + 1}" for i, col in enumerate(header_row)]

            data_rows: list[dict[str, Any]] = []
            for row in raw_rows[raw_rows.index(header_row) + 1 :]:
                padded = row + [""] * (len(headers) - len(row))
                values = padded[: len(headers)]
                if all(value == "" for value in values):
                    continue
                data_rows.append({headers[i]: values[i] for i in range(len(headers))})

            sheets.append(SheetData(name=sheet_name, columns=headers, rows=data_rows))

        return sheets


class SpreadsheetHandler(BaseHTTPRequestHandler):
    def _send_json(self, payload: Any, status: int = 200) -> None:
        data = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _serve_file(self, filepath: Path) -> None:
        if not filepath.exists() or not filepath.is_file():
            self.send_error(404, "Not found")
            return

        content = filepath.read_bytes()
        content_type = "text/plain; charset=utf-8"
        if filepath.suffix == ".html":
            content_type = "text/html; charset=utf-8"
        elif filepath.suffix == ".js":
            content_type = "application/javascript; charset=utf-8"
        elif filepath.suffix == ".css":
            content_type = "text/css; charset=utf-8"

        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(content)))
        self.end_headers()
        self.wfile.write(content)

    def do_GET(self) -> None:  # noqa: N802
        path = unquote(self.path.split("?", 1)[0])

        if path == "/api/data":
            try:
                sheets = load_workbook_data()
            except FileNotFoundError as exc:
                self._send_json({"error": str(exc)}, status=404)
                return
            except Exception as exc:  # pragma: no cover
                self._send_json({"error": f"Failed to parse spreadsheet: {exc}"}, status=500)
                return

            payload = {
                "workbook": XLSX_PATH.name,
                "sheets": [
                    {
                        "name": sheet.name,
                        "columns": sheet.columns,
                        "rows": sheet.rows,
                    }
                    for sheet in sheets
                ],
            }
            self._send_json(payload)
            return

        if path == "/":
            self._serve_file(STATIC_DIR / "index.html")
            return

        safe_path = posixpath.normpath(path).lstrip("/")
        if safe_path.startswith(".."):
            self.send_error(400, "Invalid path")
            return

        self._serve_file(STATIC_DIR / safe_path)


def run_server(host: str = "127.0.0.1", port: int = 8000) -> None:
    server = HTTPServer((host, port), SpreadsheetHandler)
    print(f"Server started at http://{host}:{port}")
    print(f"Reading workbook: {XLSX_PATH.name}")
    server.serve_forever()


if __name__ == "__main__":
    run_server(host=os.environ.get("HOST", "127.0.0.1"), port=int(os.environ.get("PORT", "8000")))
