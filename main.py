from __future__ import annotations

import os
import posixpath
import platform
import re
import subprocess
import sys
import tempfile
import tkinter as tk
import xml.etree.ElementTree as ET
import zipfile
from collections.abc import Iterable, Sequence
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from xml.sax.saxutils import escape


REPORTS_FOLDER_NAME = "отчеты"
COMMENT_HEADER = "комментарий"
PLATE_HEADER = "госномер"
NOTE_HEADER = "ПРИМЕЧАНИЕ"
REPORT_SECTIONS = {"ДУТ", "УДКМ", "БК", NOTE_HEADER}
SUPPORTED_EXCEL_SUFFIXES = {".xlsx", ".xlsm"}
SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOCUMENT_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CELL_ELEMENT_PATTERN = re.compile(
    rb"<(?P<prefix>(?:[A-Za-z_][\w.-]*:)?)c\b(?P<attrs>[^>]*)>"
    rb"(?P<body>.*?)</(?P=prefix)c\s*>"
    rb"|<(?P<self_prefix>(?:[A-Za-z_][\w.-]*:)?)c\b"
    rb"(?P<self_attrs>[^>]*)/\s*>",
    re.DOTALL,
)
ROW_ELEMENT_PATTERN = re.compile(
    rb"<(?P<prefix>(?:[A-Za-z_][\w.-]*:)?)row\b(?P<attrs>[^>]*)>"
    rb"(?P<body>.*?)</(?P=prefix)row\s*>",
    re.DOTALL,
)


@dataclass(frozen=True)
class UpdateResult:
    comments_written: int
    rows_without_report: int
    rows_without_interpretation: int
    rows_skipped_merged_cells: int
    rows_skipped_existing_comments: int
    worksheets_updated: int
    vehicle_folders_found: int
    report_folders_found: int
    unique_report_plates: int
    report_plates_not_in_workbook: int
    extract_path: Path


@dataclass(frozen=True)
class ReportFolderRecord:
    plate: str
    plate_key: str
    folder_path: str | Path
    report_path: str | Path
    comment: str


def normalize_key(value: object) -> str:
    return "".join(character for character in str(value).casefold() if character.isalnum())


def read_report(path: Path) -> str:
    return _decode_report(path.read_bytes())


def _decode_report(contents: bytes) -> str:
    for encoding in ("utf-8-sig", "utf-8", "cp1251", "cp866"):
        try:
            return contents.decode(encoding)
        except UnicodeDecodeError:
            continue
    return contents.decode("utf-8", errors="replace")


def _is_separator(line: str) -> bool:
    return bool(line.strip()) and re.fullmatch(r"[-_=—\s]+", line.strip()) is not None


def _parse_report_sections(report_text: str) -> dict[str, list[str]]:
    sections: dict[str, list[str]] = {}
    current_section: str | None = None
    section_names = {name.casefold(): name for name in REPORT_SECTIONS}

    for line in report_text.splitlines():
        label = line.strip().rstrip(":").strip()
        section = section_names.get(label.casefold())
        if section is not None:
            current_section = section
            sections.setdefault(section, [])
        elif current_section is not None and not _is_separator(line):
            sections[current_section].append(line.strip())
    return sections


def interpret_report(report_text: str) -> str:
    """Составляет комментарий в формате сводной Excel-таблицы."""
    simple_statuses: dict[str, bool] = {}
    for line in report_text.splitlines():
        match = re.fullmatch(r"\s*(гео|дут)\s*([+-])\s*", line, re.IGNORECASE)
        if match:
            simple_statuses.setdefault(
                match.group(1).casefold(),
                match.group(2) == "+",
            )
    if simple_statuses:
        geo_status = "есть" if simple_statuses.get("гео", False) else "НЕТ"
        dut_status = "есть" if simple_statuses.get("дут", False) else "нет"
        if geo_status == "НЕТ" and dut_status == "нет":
            return "Гео НЕТ, ДУТ НЕТ."
        return f"Гео {geo_status}, ДУТ {dut_status}."

    sections = _parse_report_sections(report_text)

    geo_status = "НЕТ"
    for line in sections.get("БК", []):
        match = re.search(
            r"\bданные\s*:\s*(есть|нету|нет|отсутствуют?)\b",
            line,
            re.IGNORECASE,
        )
        if match:
            geo_status = "есть" if match.group(1).casefold() == "есть" else "НЕТ"
            break

    dut_data_statuses: list[bool] = []
    for line in sections.get("ДУТ", []):
        data_match = re.search(
            r"\bданные\s*:\s*(есть|нету|нет|отсутствуют?)\b",
            line,
            re.IGNORECASE,
        )
        if data_match:
            dut_data_statuses.append(data_match.group(1).casefold() == "есть")

    dut_status = "отсутствует"
    if dut_data_statuses:
        dut_status = "есть" if any(dut_data_statuses) else "отсутствует"
    else:
        for line in sections.get("ДУТ", []):
            match = re.search(r"\bкол[\s\-‑–—]*во\s*:\s*(\d+(?:[.,]\d+)?)", line, re.IGNORECASE)
            if match:
                dut_status = "есть" if float(match.group(1).replace(",", ".")) > 0 else "отсутствует"
                break

    parts: list[str] = []
    parts.append(f"Гео {geo_status}")
    parts.append("ДУТ есть" if dut_status == "есть" else "ДУТ нет")

    summary = ", ".join(parts)
    if summary and geo_status == "НЕТ" and dut_status == "отсутствует":
        summary = "Гео НЕТ, ДУТ НЕТ"

    return f"{summary}."


def _is_vehicle_report_file(path: Path, vehicle_folder_name: str) -> bool:
    return _is_vehicle_report_name(path.name, vehicle_folder_name)


def _is_vehicle_report_name(report_name: str, vehicle_folder_name: str) -> bool:
    report_path = Path(report_name)
    return (
        report_path.suffix.casefold() == ".txt"
        and (
            report_path.stem.casefold().startswith("отчет_")
            or normalize_key(report_path.stem) == normalize_key(vehicle_folder_name)
        )
    )


def _zip_path_parts(member_name: str) -> tuple[str, ...]:
    return tuple(
        part
        for part in member_name.replace("\\", "/").split("/")
        if part and part != "."
    )


def _find_zip_reports_prefix(archive: zipfile.ZipFile) -> tuple[str, ...]:
    reports_prefixes = {
        parts[:index + 1]
        for info in archive.infolist()
        for parts in [_zip_path_parts(info.filename)]
        for index, part in enumerate(parts)
        if part.casefold() == REPORTS_FOLDER_NAME
        and (info.is_dir() or index + 1 < len(parts))
    }
    if len(reports_prefixes) > 1:
        raise ValueError("В ZIP-архиве найдено несколько папок «Отчеты».")
    if reports_prefixes:
        return next(iter(reports_prefixes))

    for info in archive.infolist():
        parts = _zip_path_parts(info.filename)
        if (
            not info.is_dir()
            and len(parts) == 2
            and _is_vehicle_report_name(parts[-1], parts[0])
        ):
            return ()
    raise ValueError(
        "В ZIP-архиве нет папки «Отчеты» или подпапок с отчетами ТС."
    )


def find_reports_folder(selected_path: Path) -> Path:
    if selected_path.is_file() and selected_path.suffix.casefold() == ".zip":
        with zipfile.ZipFile(selected_path) as archive:
            _find_zip_reports_prefix(archive)
        return selected_path
    if not selected_path.is_dir():
        raise ValueError(f"Выбранная папка или ZIP-архив не существует: {selected_path}")

    if selected_path.name.casefold() == REPORTS_FOLDER_NAME:
        return selected_path

    matches = [
        child
        for child in selected_path.iterdir()
        if child.is_dir() and child.name.casefold() == REPORTS_FOLDER_NAME
    ]
    if len(matches) > 1:
        raise ValueError(f"Найдено несколько папок «Отчеты» в: {selected_path}")
    if matches:
        return matches[0]
    for child in selected_path.iterdir():
        if child.is_dir() and any(
            _is_vehicle_report_file(path, child.name)
            for path in child.iterdir()
        ):
            return selected_path
    raise ValueError(
        f"В выбранной папке нет папки «Отчеты» и подпапок с отчетами: {selected_path}"
    )


def _read_folder_reports(
    reports_paths: Path | Iterable[Path],
) -> tuple[dict[str, str], int, int, list[ReportFolderRecord]]:
    selected_paths = [reports_paths] if isinstance(reports_paths, Path) else list(reports_paths)
    if not selected_paths:
        raise ValueError("Добавьте хотя бы одну папку с отчетами.")

    latest_reports: dict[str, tuple[float, str]] = {}
    vehicle_folders_found = 0
    report_folders_found = 0
    report_records: list[ReportFolderRecord] = []

    def record_report(
        plate: str,
        folder_path: str | Path,
        report_path: str | Path,
        modified_at: float,
        report_text: str,
    ) -> None:
        key = normalize_key(plate)
        comment = interpret_report(report_text)
        report_records.append(
            ReportFolderRecord(
                plate=plate,
                plate_key=key,
                folder_path=folder_path,
                report_path=report_path,
                comment=comment,
            )
        )
        existing = latest_reports.get(key)
        if existing is None or modified_at >= existing[0]:
            latest_reports[key] = (modified_at, comment)

    for selected_path in selected_paths:
        selected_path = selected_path.expanduser().resolve()
        reports_source = find_reports_folder(selected_path)
        if selected_path.suffix.casefold() == ".zip" and selected_path.is_file():
            with zipfile.ZipFile(selected_path) as archive:
                reports_prefix = _find_zip_reports_prefix(archive)
                groups: dict[str, list[zipfile.ZipInfo]] = {}
                prefix_length = len(reports_prefix)
                for info in archive.infolist():
                    parts = _zip_path_parts(info.filename)
                    if len(parts) < prefix_length + 2:
                        continue
                    if parts[:prefix_length] != reports_prefix:
                        continue
                    groups.setdefault(parts[prefix_length], []).append(info)

                vehicle_folders_found += len(groups)
                for vehicle_name, members in sorted(
                    groups.items(),
                    key=lambda item: item[0].casefold(),
                ):
                    report_files = [
                        info
                        for info in members
                        if len(_zip_path_parts(info.filename)) == prefix_length + 2
                        and not info.is_dir()
                        and _is_vehicle_report_name(
                            _zip_path_parts(info.filename)[-1],
                            vehicle_name,
                        )
                    ]
                    if not report_files:
                        continue
                    report_folders_found += 1
                    report_info = max(
                        report_files,
                        key=lambda info: (info.date_time, info.filename.casefold()),
                    )
                    internal_path = "/".join(_zip_path_parts(report_info.filename))
                    folder_path = f"{selected_path}!/{'/'.join((*reports_prefix, vehicle_name))}"
                    report_path = f"{selected_path}!/{internal_path}"
                    modified_at = datetime(*report_info.date_time).timestamp()
                    record_report(
                        vehicle_name,
                        folder_path,
                        report_path,
                        modified_at,
                        _decode_report(archive.read(report_info)),
                    )
            continue

        reports_folder = reports_source
        for vehicle_folder in sorted(reports_folder.iterdir(), key=lambda item: item.name.casefold()):
            if not vehicle_folder.is_dir():
                continue
            vehicle_folders_found += 1

            report_files = [
                path
                for path in vehicle_folder.iterdir()
                if _is_vehicle_report_file(path, vehicle_folder.name)
            ]
            if not report_files:
                continue
            report_folders_found += 1

            report_path = max(
                report_files,
                key=lambda path: (path.stat().st_mtime, path.name.casefold()),
            )
            record_report(
                vehicle_folder.name,
                vehicle_folder,
                report_path,
                report_path.stat().st_mtime,
                read_report(report_path),
            )
    return (
        {key: comment for key, (_, comment) in latest_reports.items()},
        vehicle_folders_found,
        report_folders_found,
        report_records,
    )


def read_folder_reports(reports_paths: Path | Iterable[Path]) -> dict[str, str]:
    reports, _, _, _ = _read_folder_reports(reports_paths)
    return reports


def _write_reconciliation_extract(
    output_path: Path,
    report_records: list[ReportFolderRecord],
    workbook_plate_keys: set[str],
    duplicate_plate_keys: set[str],
    merged_cell_rows: list[tuple[str, str, int, str, str]],
    existing_comment_rows: list[tuple[str, str, int, str, str]],
    reserved_paths: Iterable[Path] = (),
) -> Path:
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font
    except ImportError as error:
        raise RuntimeError(
            "Не установлена библиотека openpyxl. Установите зависимости из requirements.txt."
        ) from error

    extract_path = output_path.with_name(f"{output_path.stem}_выписка.xlsx")
    reserved = {path.expanduser().resolve() for path in reserved_paths}
    reserved.add(output_path.expanduser().resolve())
    suffix = 2
    while extract_path.expanduser().resolve() in reserved:
        extract_path = output_path.with_name(
            f"{output_path.stem}_выписка_{suffix}.xlsx"
        )
        suffix += 1
    workbook = Workbook()
    default_sheet = workbook.active
    if default_sheet is not None:
        workbook.remove(default_sheet)

    sheets: list[tuple[str, list[str], Sequence[Sequence[object]]]] = [
        (
            "Распознанные отчеты",
            ["Госномер", "Папка ТС", "Файл отчета", "Комментарий"],
            [
                (
                    record.plate,
                    str(record.folder_path),
                    str(record.report_path),
                    record.comment,
                )
                for record in report_records
            ],
        ),
        (
            "Отсутствуют в Excel",
            ["Госномер", "Папка ТС", "Файл отчета", "Комментарий"],
            [
                (
                    record.plate,
                    str(record.folder_path),
                    str(record.report_path),
                    record.comment,
                )
                for record in report_records
                if record.plate_key not in workbook_plate_keys
            ],
        ),
        (
            "Дубли папок",
            ["Госномер", "Число папок", "Папка ТС", "Файл отчета"],
            [
                (
                    record.plate,
                    sum(item.plate_key == record.plate_key for item in report_records),
                    str(record.folder_path),
                    str(record.report_path),
                )
                for record in report_records
                if record.plate_key in duplicate_plate_keys
            ],
        ),
        (
            "Объединенные ячейки",
            [
                "Госномер Excel",
                "Лист",
                "Строка",
                "Ячейка комментария",
                "Комментарий отчета",
            ],
            merged_cell_rows,
        ),
        (
            "Существующие комментарии",
            [
                "Госномер Excel",
                "Лист",
                "Строка",
                "Ячейка комментария",
                "Текущий комментарий",
            ],
            existing_comment_rows,
        ),
    ]

    for title, headers, rows in sheets:
        sheet = workbook.create_sheet(title)
        sheet.append(headers)
        for cell in sheet[1]:
            cell.font = Font(bold=True)
        for row in rows:
            sheet.append(row)
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for column in sheet.columns:
            values = [str(cell.value or "") for cell in column]
            width = min(max(max((len(value) for value in values), default=0) + 2, 12), 72)
            sheet.column_dimensions[column[0].column_letter].width = width

    temporary_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            prefix=f".{extract_path.stem}_",
            suffix=extract_path.suffix,
            dir=extract_path.parent,
            delete=False,
        ) as temporary_file:
            temporary_path = Path(temporary_file.name)
        workbook.save(temporary_path)
        os.replace(temporary_path, extract_path)
        temporary_path = None
    finally:
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)
    return extract_path


def find_header_columns(sheet) -> tuple[int, int, int] | None:
    for row_index, row in enumerate(sheet.iter_rows(values_only=True), start=1):
        columns = {
            normalize_key(value): column_index
            for column_index, value in enumerate(row, start=1)
            if value is not None
        }
        plate_column = columns.get(PLATE_HEADER)
        comment_column = columns.get(COMMENT_HEADER)
        if plate_column is not None and comment_column is not None:
            return row_index, plate_column, comment_column
    return None


def _worksheet_paths(archive: zipfile.ZipFile) -> dict[str, str]:
    workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
    relationships_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {
        relationship.attrib["Id"]: relationship.attrib["Target"]
        for relationship in relationships_root.findall(f"{{{PACKAGE_REL_NS}}}Relationship")
    }

    paths: dict[str, str] = {}
    sheets = workbook_root.find(f"{{{SPREADSHEET_NS}}}sheets")
    if sheets is None:
        raise ValueError("В книге Excel не найден список листов.")
    for sheet in sheets.findall(f"{{{SPREADSHEET_NS}}}sheet"):
        relationship_id = sheet.attrib.get(f"{{{DOCUMENT_REL_NS}}}id")
        if relationship_id not in targets:
            raise ValueError(f"Не найден XML-файл листа «{sheet.attrib.get('name', '')}».")
        target = targets[relationship_id]
        worksheet_path = (
            target.lstrip("/")
            if target.startswith("/")
            else posixpath.normpath(posixpath.join("xl", target))
        )
        paths[sheet.attrib["name"]] = worksheet_path
    return paths


def _inline_string_cell(
    prefix: bytes,
    attributes: bytes,
    value: str,
) -> bytes:
    attributes = re.sub(rb"\s+t\s*=\s*(?:\"[^\"]*\"|'[^']*')", b"", attributes)
    tag = b"<" + prefix + b"c" + attributes + b' t="inlineStr">'
    text = escape(value).encode("utf-8")
    preserve_space = value != value.strip()
    space_attribute = b' xml:space="preserve"' if preserve_space else b""
    return tag + b"<is><t" + space_attribute + b">" + text + b"</t></is></" + prefix + b"c>"


def _shared_string_cell(prefix: bytes, attributes: bytes, string_index: int) -> bytes:
    attributes = re.sub(rb"\s+t\s*=\s*(?:\"[^\"]*\"|'[^']*')", b"", attributes)
    return (
        b"<" + prefix + b"c" + attributes + b' t="s"><v>'
        + str(string_index).encode("ascii")
        + b"</v></" + prefix + b"c>"
    )


def _cell_column_index(reference: str) -> int:
    match = re.match(r"([A-Z]+)", reference)
    if match is None:
        raise ValueError(f"Некорректный адрес ячейки Excel: {reference}")
    column_index = 0
    for character in match.group(1):
        column_index = column_index * 26 + ord(character) - ord("A") + 1
    return column_index


def _repair_missing_cell_references(worksheet_xml: bytes) -> bytes:
    def repair_row(row_match: re.Match[bytes]) -> bytes:
        row_attributes = row_match.group("attrs")
        row_number_match = re.search(
            rb"(?:^|\s)r\s*=\s*([\"'])(\d+)\1",
            row_attributes,
        )
        if row_number_match is None:
            return row_match.group(0)
        row_number = row_number_match.group(2).decode("ascii")
        body = row_match.group("body")
        cells = list(CELL_ELEMENT_PATTERN.finditer(body))
        missing_references: dict[int, str] = {}
        index = 0
        while index < len(cells):
            match = cells[index]
            attributes = match.group("attrs") or match.group("self_attrs") or b""
            if re.search(rb"(?:^|\s)r\s*=", attributes):
                index += 1
                continue

            run_start = index
            while index < len(cells):
                current_attributes = (
                    cells[index].group("attrs")
                    or cells[index].group("self_attrs")
                    or b""
                )
                if re.search(rb"(?:^|\s)r\s*=", current_attributes):
                    break
                index += 1
            run_length = index - run_start
            if run_start == 0 or index == len(cells):
                continue

            previous_attributes = (
                cells[run_start - 1].group("attrs")
                or cells[run_start - 1].group("self_attrs")
                or b""
            )
            next_attributes = cells[index].group("attrs") or cells[index].group("self_attrs") or b""
            previous_match = re.search(
                rb"(?:^|\s)r\s*=\s*([\"'])([A-Z]+)\d+\1",
                previous_attributes,
            )
            next_match = re.search(
                rb"(?:^|\s)r\s*=\s*([\"'])([A-Z]+)\d+\1",
                next_attributes,
            )
            if previous_match is None or next_match is None:
                continue
            previous_column = _cell_column_index(previous_match.group(2).decode("ascii"))
            next_column = _cell_column_index(next_match.group(2).decode("ascii"))
            if next_column - previous_column != run_length + 1:
                continue

            for offset in range(run_length):
                column_number = previous_column + offset + 1
                column_name = ""
                while column_number:
                    column_number, remainder = divmod(column_number - 1, 26)
                    column_name = chr(ord("A") + remainder) + column_name
                missing_references[cells[run_start + offset].start()] = (
                    f"{column_name}{row_number}"
                )

        if not missing_references:
            return row_match.group(0)

        def add_reference(cell_match: re.Match[bytes]) -> bytes:
            reference = missing_references.get(cell_match.start())
            if reference is None:
                return cell_match.group(0)
            prefix = cell_match.group("prefix")
            if prefix is None:
                prefix = cell_match.group("self_prefix")
            if prefix is None:
                raise ValueError("Не удалось определить тип ячейки Excel.")
            opening_tag_end = len(b"<" + prefix + b"c")
            return (
                cell_match.group(0)[:opening_tag_end]
                + b' r="'
                + reference.encode("ascii")
                + b'"'
                + cell_match.group(0)[opening_tag_end:]
            )

        repaired_body = CELL_ELEMENT_PATTERN.sub(add_reference, body)
        matched_row = row_match.group(0)
        body_start = row_match.start("body") - row_match.start(0)
        body_end = row_match.end("body") - row_match.start(0)
        return matched_row[:body_start] + repaired_body + matched_row[body_end:]

    return ROW_ELEMENT_PATTERN.sub(repair_row, worksheet_xml)


def _patch_worksheet_xml(
    worksheet_xml: bytes,
    cell_values: dict[str, str],
    sheet_name: str,
    shared_string_indices: dict[str, int] | None = None,
) -> bytes:
    pending = dict(cell_values)

    def replace_cell(match: re.Match[bytes]) -> bytes:
        prefix = match.group("prefix")
        attributes = match.group("attrs")
        if prefix is None:
            prefix = match.group("self_prefix")
            attributes = match.group("self_attrs")
        if prefix is None or attributes is None:
            raise ValueError(f"Не удалось разобрать ячейку на листе «{sheet_name}».")

        reference_match = re.search(rb"(?:^|\s)r\s*=\s*([\"'])(.*?)\1", attributes)
        if reference_match is None:
            return match.group(0)
        reference = reference_match.group(2).decode("ascii")
        value = pending.pop(reference, None)
        if value is None:
            return match.group(0)
        if shared_string_indices is not None:
            return _shared_string_cell(prefix, attributes, shared_string_indices[value])
        return _inline_string_cell(prefix, attributes, value)

    patched_xml = CELL_ELEMENT_PATTERN.sub(replace_cell, worksheet_xml)
    if pending:
        additions_by_row: dict[int, list[tuple[int, str, str]]] = {}
        for reference, value in pending.items():
            address_match = re.fullmatch(r"([A-Z]+)(\d+)", reference)
            if address_match is None:
                raise ValueError(f"Некорректный адрес ячейки Excel: {reference}")
            row_index = int(address_match.group(2))
            additions_by_row.setdefault(row_index, []).append(
                (_cell_column_index(reference), reference, value)
            )

        inserted_rows: set[int] = set()

        def insert_missing_cells(row_match: re.Match[bytes]) -> bytes:
            attributes = row_match.group("attrs")
            body = row_match.group("body")
            row_number_match = re.search(rb"(?:^|\s)r\s*=\s*([\"'])(\d+)\1", attributes)
            if row_number_match is None:
                return row_match.group(0)
            row_index = int(row_number_match.group(2))
            additions = additions_by_row.get(row_index)
            if not additions:
                return row_match.group(0)

            for column_index, reference, value in sorted(additions):
                insertion_offset = len(body)
                for cell_match in CELL_ELEMENT_PATTERN.finditer(body):
                    cell_attributes = cell_match.group("attrs")
                    if cell_attributes is None:
                        cell_attributes = cell_match.group("self_attrs")
                    if cell_attributes is None:
                        continue
                    cell_reference_match = re.search(
                        rb"(?:^|\s)r\s*=\s*([\"'])(.*?)\1",
                        cell_attributes,
                    )
                    if cell_reference_match is None:
                        continue
                    cell_reference = cell_reference_match.group(2).decode("ascii")
                    if _cell_column_index(cell_reference) > column_index:
                        insertion_offset = cell_match.start()
                        break
                if shared_string_indices is None:
                    cell_bytes = _inline_string_cell(
                        b"",
                        b' r="' + reference.encode("ascii") + b'"',
                        value,
                    )
                else:
                    cell_bytes = _shared_string_cell(
                        b"",
                        b' r="' + reference.encode("ascii") + b'"',
                        shared_string_indices[value],
                    )
                body = body[:insertion_offset] + cell_bytes + body[insertion_offset:]
            inserted_rows.add(row_index)
            matched_row = row_match.group(0)
            body_start = row_match.start("body") - row_match.start(0)
            body_end = row_match.end("body") - row_match.start(0)
            return (
                matched_row[:body_start]
                + body
                + matched_row[body_end:]
            )

        patched_xml = ROW_ELEMENT_PATTERN.sub(insert_missing_cells, patched_xml)
        missing_rows = set(additions_by_row) - inserted_rows
        if missing_rows:
            missing = ", ".join(
                reference
                for row_index in sorted(missing_rows)
                for _, reference, _ in additions_by_row[row_index]
            )
            raise ValueError(
                f"Не удалось найти строки ячеек {missing} на листе «{sheet_name}»."
            )
    return patched_xml


def _update_shared_strings(
    shared_strings_xml: bytes,
    new_strings: Iterable[str],
    total_references: int,
) -> tuple[bytes, dict[str, int]]:
    root = ET.fromstring(shared_strings_xml)
    string_items = root.findall(f"{{{SPREADSHEET_NS}}}si")
    indices: dict[str, int] = {}
    for index, item in enumerate(string_items):
        value = "".join(item.itertext())
        indices.setdefault(value, index)

    to_append: list[str] = []
    for value in new_strings:
        if value not in indices:
            indices[value] = len(string_items) + len(to_append)
            to_append.append(value)

    start_tag_match = re.search(rb"<(?P<prefix>(?:[A-Za-z_][\w.-]*:)?)sst\b[^>]*>", shared_strings_xml)
    closing_tag_match = re.search(rb"</(?P<prefix>(?:[A-Za-z_][\w.-]*:)?)sst\s*>", shared_strings_xml)
    if start_tag_match is None or closing_tag_match is None:
        raise ValueError("Не удалось разобрать общую таблицу строк книги Excel.")

    start_tag = start_tag_match.group(0)
    attributes = {
        b"count": str(total_references).encode("ascii"),
        b"uniqueCount": str(len(string_items) + len(to_append)).encode("ascii"),
    }
    for name, value in attributes.items():
        attribute_pattern = re.compile(
            rb"(\s+" + name + rb"\s*=\s*)(?:\"[^\"]*\"|'[^']*')"
        )
        replacement = lambda match, value=value: match.group(1) + b'"' + value + b'"'
        if attribute_pattern.search(start_tag):
            start_tag = attribute_pattern.sub(replacement, start_tag, count=1)
        else:
            start_tag = start_tag[:-1] + b" " + name + b'="' + value + b'">'

    prefix = start_tag_match.group("prefix")
    appended_xml: list[bytes] = []
    for value in to_append:
        text = escape(value).encode("utf-8")
        preserve_space = value != value.strip()
        space_attribute = b' xml:space="preserve"' if preserve_space else b""
        appended_xml.append(
            b"<" + prefix + b"si><" + prefix + b"t" + space_attribute + b">"
            + text
            + b"</" + prefix + b"t></" + prefix + b"si>"
        )

    content_start = start_tag_match.end()
    updated_xml = (
        shared_strings_xml[:start_tag_match.start()]
        + start_tag
        + shared_strings_xml[content_start:closing_tag_match.start()]
        + b"".join(appended_xml)
        + shared_strings_xml[closing_tag_match.start():]
    )
    return updated_xml, indices


def _count_shared_string_references(worksheet_xml: bytes) -> int:
    root = ET.fromstring(worksheet_xml)
    return sum(
        1
        for cell in root.iter(f"{{{SPREADSHEET_NS}}}c")
        if cell.get("t") == "s"
    )


def _write_updated_archive(
    workbook_path: Path,
    output_path: Path,
    updates: dict[str, dict[str, str]],
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temporary_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            prefix=f".{output_path.stem}_",
            suffix=output_path.suffix,
            dir=output_path.parent,
            delete=False,
        ) as temporary_file:
            temporary_path = Path(temporary_file.name)

        with zipfile.ZipFile(workbook_path, "r") as source:
            sheet_paths = _worksheet_paths(source)
            path_updates: dict[str, tuple[str, dict[str, str]]] = {}
            for sheet_name, values in updates.items():
                worksheet_path = sheet_paths.get(sheet_name)
                if worksheet_path is None:
                    raise ValueError(f"В архиве Excel не найден лист «{sheet_name}».")
                path_updates[worksheet_path] = (sheet_name, values)

            contents = {
                entry.filename: source.read(entry.filename)
                for entry in source.infolist()
            }
            shared_string_indices: dict[str, int] | None = None
            shared_strings_path = "xl/sharedStrings.xml"
            if shared_strings_path in contents:
                all_strings = [
                    value
                    for values in updates.values()
                    for value in values.values()
                ]
                shared_strings_xml, shared_string_indices = _update_shared_strings(
                    contents[shared_strings_path],
                    all_strings,
                    total_references=0,
                )
                contents[shared_strings_path] = shared_strings_xml

            for worksheet_path, (sheet_name, values) in path_updates.items():
                worksheet_xml = _repair_missing_cell_references(
                    contents[worksheet_path]
                )
                contents[worksheet_path] = _patch_worksheet_xml(
                    worksheet_xml,
                    values,
                    sheet_name,
                    shared_string_indices,
                )

            if shared_string_indices is not None:
                total_references = sum(
                    _count_shared_string_references(contents[worksheet_path])
                    for worksheet_path in sheet_paths.values()
                )
                contents[shared_strings_path], _ = _update_shared_strings(
                    contents[shared_strings_path],
                    (),
                    total_references,
                )

            with zipfile.ZipFile(temporary_path, "w") as result:
                result.comment = source.comment
                for entry in source.infolist():
                    result.writestr(entry, contents[entry.filename])

        os.replace(temporary_path, output_path)
        temporary_path = None
    finally:
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)


def update_workbook(
    selected_reports_paths: Path | Iterable[Path],
    workbook_path: Path,
    output_path: Path,
    skip_existing_comments: bool = False,
    overwrite_source: bool = False,
) -> UpdateResult:
    try:
        from openpyxl import load_workbook
        from openpyxl.cell.cell import MergedCell
    except ImportError as error:
        raise RuntimeError(
            "Не установлена библиотека openpyxl. Установите зависимости из requirements.txt."
        ) from error

    workbook_path = workbook_path.expanduser().resolve()
    output_path = (
        workbook_path
        if overwrite_source
        else output_path.expanduser().resolve()
    )
    if not workbook_path.is_file():
        raise ValueError(f"Файл Excel не найден: {workbook_path}")
    if workbook_path.suffix.casefold() not in SUPPORTED_EXCEL_SUFFIXES:
        raise ValueError("Поддерживаются файлы Excel форматов .xlsx и .xlsm.")
    if output_path.suffix.casefold() != workbook_path.suffix.casefold():
        raise ValueError("Расширение выходного файла должно совпадать с исходным файлом.")
    if output_path == workbook_path and not overwrite_source:
        raise ValueError("Выберите новый файл для результата, чтобы сохранить исходную книгу.")

    (
        reports,
        vehicle_folders_found,
        report_folders_found,
        report_records,
    ) = _read_folder_reports(selected_reports_paths)
    if not reports:
        raise ValueError("В выбранных папках не найдено текстовых отчетов.")

    workbook = load_workbook(
        workbook_path,
        keep_vba=workbook_path.suffix.casefold() == ".xlsm",
    )
    comments_written = 0
    rows_without_report = 0
    rows_without_interpretation = 0
    rows_skipped_merged_cells = 0
    rows_skipped_existing_comments = 0
    worksheets_updated = 0
    sheets_with_headers = 0
    updates: dict[str, dict[str, str]] = {}
    workbook_plate_keys: set[str] = set()
    merged_cell_rows: list[tuple[str, str, int, str, str]] = []
    existing_comment_rows: list[tuple[str, str, int, str, str]] = []

    for sheet in workbook.worksheets:
        headers = find_header_columns(sheet)
        if headers is None:
            continue
        sheets_with_headers += 1
        header_row, plate_column, comment_column = headers
        sheet_written = False
        for row_index in range(header_row + 1, sheet.max_row + 1):
            plate = sheet.cell(row=row_index, column=plate_column).value
            if plate is None or not str(plate).strip():
                continue
            workbook_plate_keys.add(normalize_key(plate))

            note = reports.get(normalize_key(plate))
            if note is None:
                rows_without_report += 1
                continue
            if not note:
                rows_without_interpretation += 1
                note = "Гео НЕТ, ДУТ НЕТ."

            comment_cell = sheet.cell(row=row_index, column=comment_column)
            if isinstance(comment_cell, MergedCell):
                rows_skipped_merged_cells += 1
                merged_cell_rows.append(
                    (
                        str(plate),
                        sheet.title,
                        row_index,
                        comment_cell.coordinate,
                        note,
                    )
                )
                continue
            if (
                skip_existing_comments
                and comment_cell.value is not None
                and str(comment_cell.value).strip()
            ):
                rows_skipped_existing_comments += 1
                existing_comment_rows.append(
                    (
                        str(plate),
                        sheet.title,
                        row_index,
                        comment_cell.coordinate,
                        str(comment_cell.value),
                    )
                )
                continue
            updates.setdefault(sheet.title, {})[comment_cell.coordinate] = note
            comments_written += 1
            sheet_written = True
        if sheet_written:
            worksheets_updated += 1

    if not sheets_with_headers:
        raise ValueError(
            "В книге не найдена строка заголовков, содержащая столбцы «Госномер» и «Комментарий»."
        )

    _write_updated_archive(workbook_path, output_path, updates)
    duplicate_counts: dict[str, int] = {}
    for record in report_records:
        duplicate_counts[record.plate_key] = duplicate_counts.get(record.plate_key, 0) + 1
    duplicate_plate_keys = {
        plate_key
        for plate_key, count in duplicate_counts.items()
        if count > 1
    }
    extract_path = _write_reconciliation_extract(
        output_path,
        report_records,
        workbook_plate_keys,
        duplicate_plate_keys,
        merged_cell_rows,
        existing_comment_rows,
        reserved_paths=(workbook_path,),
    )
    return UpdateResult(
        comments_written=comments_written,
        rows_without_report=rows_without_report,
        rows_without_interpretation=rows_without_interpretation,
        rows_skipped_merged_cells=rows_skipped_merged_cells,
        rows_skipped_existing_comments=rows_skipped_existing_comments,
        worksheets_updated=worksheets_updated,
        vehicle_folders_found=vehicle_folders_found,
        report_folders_found=report_folders_found,
        unique_report_plates=len(reports),
        report_plates_not_in_workbook=len(set(reports) - workbook_plate_keys),
        extract_path=extract_path,
    )


class CommentsFromReportsApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Заполнение комментариев из отчетов")
        self.minsize(760, 430)
        self.reports_paths: list[Path] = []
        self.workbook_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.overwrite_source = tk.BooleanVar(value=False)
        self._separate_output_path: str | None = None
        self._build_ui()

    def _build_ui(self) -> None:
        frame = ttk.Frame(self, padding=16)
        frame.pack(fill=tk.BOTH, expand=True)
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(1, weight=1)

        ttk.Label(
            frame,
            text="Источники отчетов (папки и ZIP-архивы):",
        ).grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=(0, 4))
        self.reports_list = tk.Listbox(
            frame,
            height=6,
            selectmode=tk.EXTENDED,
            exportselection=False,
        )
        self.reports_list.grid(row=1, column=0, columnspan=2, sticky=tk.NSEW, pady=(0, 6))
        folder_buttons = ttk.Frame(frame)
        folder_buttons.grid(row=1, column=2, sticky=tk.N, padx=(8, 0))
        ttk.Button(
            folder_buttons,
            text="Добавить папки...",
            command=self._choose_reports_folders,
        ).pack(fill=tk.X, pady=(0, 4))
        ttk.Button(
            folder_buttons,
            text="Добавить ZIP...",
            command=self._choose_reports_archives,
        ).pack(fill=tk.X, pady=(0, 4))
        ttk.Button(
            folder_buttons,
            text="Удалить выбранные",
            command=self._remove_reports_folder,
        ).pack(fill=tk.X)

        self._add_path_row(
            frame,
            2,
            "Файл Excel:",
            self.workbook_path,
            self._choose_workbook,
        )
        self.output_entry, self.output_button = self._add_path_row(
            frame,
            3,
            "Сохранить результат:",
            self.output_path,
            self._choose_output,
        )

        ttk.Label(
            frame,
            text="В исходной книге найдутся столбцы «Госномер» и «Комментарий»; "
            "комментарий будет сформирован по статусам БК и ДУТ; раздел «ПРИМЕЧАНИЕ» не копируется. "
            "Для повторяющихся госномеров используется самый новый отчет.",
            wraplength=660,
        ).grid(row=6, column=0, columnspan=3, sticky=tk.W, pady=(10, 12))
        self.skip_existing_comments = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            frame,
            text="Не перезаписывать заполненные ячейки «Комментарий»",
            variable=self.skip_existing_comments,
        ).grid(row=4, column=0, columnspan=3, sticky=tk.W, pady=(4, 0))
        ttk.Checkbutton(
            frame,
            text="Перезаписать исходный файл (сохранить под исходным именем)",
            variable=self.overwrite_source,
            command=self._toggle_source_overwrite,
        ).grid(row=5, column=0, columnspan=3, sticky=tk.W, pady=(4, 0))
        ttk.Button(frame, text="Заполнить комментарии", command=self._process).grid(
            row=7, column=0, columnspan=3, sticky=tk.EW
        )

    @staticmethod
    def _add_path_row(
        parent: ttk.Frame,
        row: int,
        label: str,
        variable: tk.StringVar,
        command,
    ) -> tuple[ttk.Entry, ttk.Button]:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky=tk.W, pady=5)
        entry = ttk.Entry(parent, textvariable=variable)
        entry.grid(
            row=row, column=1, sticky=tk.EW, padx=8, pady=5
        )
        button = ttk.Button(parent, text="Выбрать...", command=command)
        button.grid(
            row=row, column=2, pady=5
        )
        return entry, button

    def _toggle_source_overwrite(self) -> None:
        workbook_path = self.workbook_path.get().strip()
        if self.overwrite_source.get():
            self._separate_output_path = self.output_path.get().strip() or None
            if workbook_path:
                self.output_path.set(workbook_path)
            self.output_entry.state(["disabled"])
            self.output_button.state(["disabled"])
            return

        self.output_entry.state(["!disabled"])
        self.output_button.state(["!disabled"])
        if self._separate_output_path:
            self.output_path.set(self._separate_output_path)
        elif workbook_path:
            source = Path(workbook_path)
            self.output_path.set(
                str(source.with_name(f"{source.stem}_с_комментариями{source.suffix}"))
            )
        self._separate_output_path = None

    def _choose_reports_folders(self) -> None:
        if platform.system() != "Darwin":
            self._choose_multiple_reports_folders_custom()
            return

        script = """
try
    set chosenFolders to choose folder with prompt "Выберите папки с отчетами" with multiple selections allowed
on error number -128
    return ""
end try
if class of chosenFolders is not list then set chosenFolders to {chosenFolders}
set AppleScript's text item delimiters to (ASCII character 30)
set folderPaths to {}
repeat with chosenFolder in chosenFolders
    set end of folderPaths to POSIX path of (chosenFolder as alias)
end repeat
return folderPaths as text
"""
        try:
            result = subprocess.run(
                ["osascript", "-e", script],
                capture_output=True,
                text=True,
                check=False,
            )
        except OSError as error:
            messagebox.showerror(
                "Не удалось открыть выбор папок",
                f"Системная панель выбора папок недоступна:\n{error}",
            )
            return

        if result.returncode != 0:
            messagebox.showerror(
                "Не удалось выбрать папки",
                result.stderr.strip() or "Системная панель выбора папок завершилась с ошибкой.",
            )
            return

        folders = [Path(path) for path in result.stdout.strip().split("\x1e") if path]
        self._add_reports_folders(folders)

    def _choose_reports_archives(self) -> None:
        archives = filedialog.askopenfilenames(
            title="Выберите ZIP-архивы с отчетами",
            filetypes=[("ZIP-архивы", "*.zip")],
        )
        if archives:
            self._add_reports_folders([Path(path) for path in archives])

    def _choose_reports_folder(self) -> None:
        self._choose_reports_folders()

    def _choose_multiple_reports_folders(self) -> None:
        self._choose_reports_folders()

    def _choose_multiple_reports_folders_custom(self) -> None:
        dialog = tk.Toplevel(self)
        dialog.title("Выбор нескольких папок с отчетами")
        dialog.transient(self)
        dialog.geometry("700x460")
        dialog.minsize(480, 320)
        dialog.columnconfigure(0, weight=1)
        dialog.rowconfigure(2, weight=1)

        initial_path = self.reports_paths[-1] if self.reports_paths else Path.home()
        current_path = initial_path
        path_variable = tk.StringVar(value=str(current_path))

        navigation = ttk.Frame(dialog, padding=(12, 12, 12, 6))
        navigation.grid(row=0, column=0, sticky=tk.EW)
        navigation.columnconfigure(1, weight=1)
        ttk.Button(
            navigation,
            text="Вверх",
            command=lambda: navigate(current_path.parent),
        ).grid(row=0, column=0, padx=(0, 6))
        path_entry = ttk.Entry(navigation, textvariable=path_variable)
        path_entry.grid(row=0, column=1, sticky=tk.EW, padx=(0, 6))
        ttk.Button(
            navigation,
            text="Перейти",
            command=lambda: navigate(Path(path_variable.get())),
        ).grid(row=0, column=2)

        ttk.Label(
            dialog,
            text="Выберите одну или несколько папок (Shift/Cmd/Ctrl), "
            "или дважды щелкните по папке, чтобы открыть ее:",
            wraplength=660,
        ).grid(row=1, column=0, sticky=tk.W, padx=12, pady=(0, 6))

        list_frame = ttk.Frame(dialog, padding=(12, 0, 12, 8))
        list_frame.grid(row=2, column=0, sticky=tk.NSEW)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        folders_list = tk.Listbox(
            list_frame,
            selectmode=tk.EXTENDED,
            exportselection=False,
        )
        folders_list.grid(row=0, column=0, sticky=tk.NSEW)
        scrollbar = ttk.Scrollbar(
            list_frame,
            orient=tk.VERTICAL,
            command=folders_list.yview,
        )
        scrollbar.grid(row=0, column=1, sticky=tk.NS)
        folders_list.configure(yscrollcommand=scrollbar.set)

        def refresh() -> None:
            folders_list.delete(0, tk.END)
            path_variable.set(str(current_path))
            try:
                children = sorted(
                    (child for child in current_path.iterdir() if child.is_dir()),
                    key=lambda child: child.name.casefold(),
                )
            except OSError as error:
                messagebox.showerror(
                    "Не удалось прочитать папку",
                    f"Не удалось прочитать содержимое «{current_path}»:\n{error}",
                    parent=dialog,
                )
                return
            for child in children:
                folders_list.insert(tk.END, child.name)

        def navigate(path: Path) -> None:
            nonlocal current_path
            target = path.expanduser().resolve()
            if not target.is_dir():
                messagebox.showwarning(
                    "Папка не найдена",
                    f"Указанный путь не является папкой:\n{target}",
                    parent=dialog,
                )
                return
            current_path = target
            refresh()

        def selected_paths() -> list[Path]:
            return [
                current_path / folders_list.get(index)
                for index in folders_list.curselection()
            ]

        def open_selected_folder(_event=None) -> None:
            selection = selected_paths()
            if len(selection) == 1:
                navigate(selection[0])

        folders_list.bind("<Double-Button-1>", open_selected_folder)
        path_entry.bind("<Return>", lambda _event: navigate(Path(path_variable.get())))

        actions = ttk.Frame(dialog, padding=(12, 0, 12, 12))
        actions.grid(row=3, column=0, sticky=tk.EW)
        ttk.Button(
            actions,
            text="Добавить выбранные",
            command=lambda: self._add_reports_folders(selected_paths()),
        ).pack(side=tk.LEFT)
        ttk.Button(
            actions,
            text="Добавить текущую папку",
            command=lambda: self._add_reports_folders([current_path]),
        ).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(actions, text="Готово", command=dialog.destroy).pack(side=tk.RIGHT)

        refresh()
        dialog.grab_set()
        path_entry.focus_set()

    def _add_reports_folders(self, paths: Iterable[Path]) -> None:
        rejected: list[str] = []
        for path in paths:
            resolved_path = path.expanduser().resolve()
            if resolved_path in self.reports_paths:
                continue
            try:
                find_reports_folder(resolved_path)
            except (ValueError, OSError, zipfile.BadZipFile):
                rejected.append(str(resolved_path))
                continue
            self.reports_paths.append(resolved_path)
            self.reports_list.insert(tk.END, str(resolved_path))

        if rejected:
            messagebox.showwarning(
                "Неподходящие источники",
                "В этих папках или ZIP-архивах не найдены отчеты или папки ТС:\n"
                + "\n".join(rejected),
            )

    def _remove_reports_folder(self) -> None:
        selection = self.reports_list.curselection()
        if not selection:
            messagebox.showwarning(
                "Не выбрана папка",
                "Выберите одну или несколько папок в списке для удаления.",
            )
            return
        for index in reversed(selection):
            del self.reports_paths[index]
            self.reports_list.delete(index)

    def _choose_workbook(self) -> None:
        path = filedialog.askopenfilename(
            title="Выберите книгу Excel",
            filetypes=[("Книги Excel", "*.xlsx *.xlsm")],
        )
        if path:
            previous_workbook = self.workbook_path.get().strip()
            self.workbook_path.set(path)
            if self.overwrite_source.get():
                self.output_path.set(path)
            elif (
                not self.output_path.get()
                or self.output_path.get() == previous_workbook
            ):
                source = Path(path)
                self.output_path.set(str(source.with_name(f"{source.stem}_с_комментариями{source.suffix}")))

    def _choose_output(self) -> None:
        workbook = Path(self.workbook_path.get()) if self.workbook_path.get() else None
        suffix = workbook.suffix if workbook else ".xlsx"
        initial_name = (
            f"{workbook.stem}_с_комментариями{suffix}"
            if workbook
            else f"Результат{suffix}"
        )
        path = filedialog.asksaveasfilename(
            title="Сохранить заполненную книгу",
            defaultextension=suffix,
            initialfile=initial_name,
            filetypes=[("Книги Excel", "*.xlsx *.xlsm")],
        )
        if path:
            self.output_path.set(path)

    def _process(self) -> None:
        if not self.reports_paths or not self.workbook_path.get().strip():
            messagebox.showwarning(
                "Не выбраны данные",
                "Добавьте папку с отчетами и выберите файл Excel.",
            )
            return
        if not self.overwrite_source.get() and not self.output_path.get().strip():
            messagebox.showwarning("Не указан результат", "Выберите имя и место для нового файла.")
            return

        try:
            workbook_path = Path(self.workbook_path.get())
            output_path = (
                workbook_path
                if self.overwrite_source.get()
                else Path(self.output_path.get())
            )
            result = update_workbook(
                self.reports_paths,
                workbook_path,
                output_path,
                skip_existing_comments=self.skip_existing_comments.get(),
                overwrite_source=self.overwrite_source.get(),
            )
        except Exception as error:
            messagebox.showerror("Ошибка обработки", str(error))
            return

        messagebox.showinfo(
            "Готово",
            f"Папок ТС просмотрено: {result.vehicle_folders_found}\n"
            f"Папок с распознанным отчетом: {result.report_folders_found}\n"
            f"Повторяющихся госномеров в папках: "
            f"{result.report_folders_found - result.unique_report_plates}\n"
            f"Уникальных госномеров в отчетах: {result.unique_report_plates}\n"
            f"Госномеров из отчетов нет в Excel: {result.report_plates_not_in_workbook}\n"
            f"Строк Excel с найденным отчетом: "
            f"{result.comments_written + result.rows_skipped_merged_cells}\n"
            f"Записано комментариев: {result.comments_written}\n"
            f"Пропущено объединенных ячеек: {result.rows_skipped_merged_cells}\n"
            f"Сохранено существующих комментариев: "
            f"{result.rows_skipped_existing_comments}\n"
            f"Обработано листов: {result.worksheets_updated}\n"
            f"Строк без папки/отчета: {result.rows_without_report}\n"
            f"Отчетов без данных для комментария: {result.rows_without_interpretation}\n\n"
            f"Файл сохранен:\n{output_path.expanduser()}\n\n"
            f"Выписка сохранена:\n{result.extract_path}",
        )


def main() -> int:
    try:
        app = CommentsFromReportsApp()
        app.mainloop()
    except tk.TclError as error:
        print(f"Не удалось запустить графический интерфейс: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
