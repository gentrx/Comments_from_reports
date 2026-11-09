import os
import re
import sys
import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from main import (
    ReportFolderRecord,
    _count_shared_string_references,
    _read_folder_reports,
    _repair_missing_cell_references,
    _write_reconciliation_extract,
    interpret_report,
    read_folder_reports,
    update_workbook,
)


def active_sheet(workbook: Workbook) -> Worksheet:
    sheet = workbook.active
    if sheet is None:
        raise AssertionError("Test workbook has no active worksheet.")
    return sheet


class ExtractNoteTests(unittest.TestCase):
    def test_interprets_compact_geo_and_dut_statuses(self):
        self.assertEqual(
            interpret_report("Гео +\nДУТ -\n"),
            "Гео есть, ДУТ нет.",
        )
        self.assertEqual(
            interpret_report("Гео -\nДУТ +\n"),
            "Гео НЕТ, ДУТ есть.",
        )
        self.assertEqual(
            interpret_report("Гео -\nДУТ -\n"),
            "Гео НЕТ, ДУТ НЕТ.",
        )

    def test_interprets_report_statuses_without_copying_note(self):
        report = (
            "----/13.10.2026/----\n"
            "ДУТ\n"
            "Кол-во: 0\n"
            "УДКМ\n"
            "Кол-во: 0\n"
            "БК\n"
            "Данные: есть\n"
            "ПРИМЕЧАНИЕ\n"
            "Дут отсутствует, данные по гео от 16.09\n"
        )
        self.assertEqual(
            interpret_report(report),
            "Гео есть, ДУТ нет.",
        )

    def test_interprets_geo_and_dut_presence(self):
        report = "ДУТ\nКол-во: 2\nБК\nДанные: есть\nПРИМЕЧАНИЕ\n"
        self.assertEqual(interpret_report(report), "Гео есть, ДУТ есть.")

    def test_does_not_copy_note_even_when_it_mentions_geo_and_dut(self):
        report = (
            "ДУТ\n"
            "Кол-во: 1\n"
            "ДУТ 1\n"
            "Данные: есть\n"
            "БК\n"
            "Данные: есть\n"
            "ПРИМЕЧАНИЕ\n"
            "Данные дут и гео от 10.10\n"
        )
        self.assertEqual(
            interpret_report(report),
            "Гео есть, ДУТ есть.",
        )

    def test_formats_missing_geo_and_dut_like_workbook_examples(self):
        report = "ДУТ\nКол-во: 0\nБК\nДанные: нет\n"
        self.assertEqual(interpret_report(report), "Гео НЕТ, ДУТ НЕТ.")

    def test_reads_bk_data_marked_netu_without_copying_note(self):
        report = (
            "----/08.10.2026/----\n"
            "ДУТ\n"
            "Кол-во: 0\n"
            "УДКМ\n"
            "Кол-во: 0\n"
            "БК\n"
            "Данные: нету\n"
            "ПРИМЕЧАНИЕ\n"
            "гео -, дут -\n"
        )
        self.assertEqual(
            interpret_report(report),
            "Гео НЕТ, ДУТ НЕТ.",
        )

    def test_missing_bk_status_defaults_to_geo_no_and_dut_status_is_kept(self):
        report = (
            "----/13.10.2026/----\n"
            "ДУТ\n"
            "Кол-во: 1\n"
            "ДУТ 1\n"
            "Модель:\n"
            "Серийный номер:\n"
            "Данные: нету\n"
            "УДКМ\n"
            "Кол-во: 0\n"
            "ПРИМЕЧАНИЕ\n"
            "не открывается. в ремонте\n"
        )
        self.assertEqual(
            interpret_report(report),
            "Гео НЕТ, ДУТ НЕТ.",
        )

    def test_dut_presence_depends_on_sensor_data_not_sensor_count(self):
        report = (
            "ДУТ\n"
            "Кол-во: 1\n"
            "ДУТ 1\n"
            "Данные: нету\n"
            "БК\n"
            "Данные: есть\n"
        )
        self.assertEqual(interpret_report(report), "Гео есть, ДУТ нет.")

    def test_dut_is_present_if_any_sensor_has_data(self):
        report = (
            "ДУТ\n"
            "Кол-во: 2\n"
            "ДУТ 1\n"
            "Данные: нету\n"
            "ДУТ 2\n"
            "Данные: есть\n"
            "БК\n"
            "Данные: есть\n"
        )
        self.assertEqual(interpret_report(report), "Гео есть, ДУТ есть.")

    def test_does_not_copy_location_from_note(self):
        report = "ДУТ\nКол-во: 1\nБК\nДанные: есть\nПРИМЕЧАНИЕ\nТС на адресе: Манчестерская 1\n"
        self.assertEqual(interpret_report(report), "Гео есть, ДУТ есть.")

    def test_does_not_copy_special_category_from_note(self):
        report = "ПРИМЕЧАНИЕ\nТС есть в базе, но нет адреса или топлива\n"
        self.assertEqual(interpret_report(report), "Гео НЕТ, ДУТ НЕТ.")

    def test_does_not_append_additional_note_details(self):
        report = (
            "ДУТ\nКол-во: 2\nБК\nДанные: есть\n"
            "ПРИМЕЧАНИЕ\nТС не введено в эксплуатацию, БНСО смонтирован\n"
        )
        self.assertEqual(
            interpret_report(report),
            "Гео есть, ДУТ есть.",
        )

    def test_defaults_bk_and_dut_to_no_if_report_has_no_status_data(self):
        self.assertEqual(
            interpret_report("Дата: 21.10.2026\nПРИМЕЧАНИЕ\nОтчет не заполнен\n"),
            "Гео НЕТ, ДУТ НЕТ.",
        )

    def test_empty_bk_and_dut_sections_default_to_no_data(self):
        self.assertEqual(
            interpret_report("БК\n\nДУТ\n\nПРИМЕЧАНИЕ\n"),
            "Гео НЕТ, ДУТ НЕТ.",
        )


class UpdateWorkbookTests(unittest.TestCase):
    def test_repairs_missing_cell_reference_when_neighbors_are_contiguous(self):
        worksheet_xml = (
            b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            b'<sheetData><row r="46"><c r="B46" t="s"><v>0</v></c>'
            b'<c t="inlineStr"><is><t>comment</t></is></c>'
            b'<c r="D46" t="s"><v>1</v></c></row></sheetData></worksheet>'
        )

        repaired_xml = _repair_missing_cell_references(worksheet_xml)
        cells = ET.fromstring(repaired_xml).findall(
            ".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c"
        )

        self.assertEqual([cell.get("r") for cell in cells], ["B46", "C46", "D46"])

    def test_counts_shared_string_cells_without_cell_references(self):
        worksheet_xml = (
            b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            b"<sheetData><row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c>"
            b'<c t="s"><v>1</v></c></row></sheetData></worksheet>'
        )

        self.assertEqual(_count_shared_string_references(worksheet_xml), 2)

    def test_reads_reports_from_multiple_vehicle_folder_roots(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            first_root = root / "# 2026-10-08..09"
            second_root = root / "Еще отчеты"
            first_vehicle = first_root / "car111"
            second_vehicle = second_root / "car222"
            first_vehicle.mkdir(parents=True)
            second_vehicle.mkdir(parents=True)
            (first_vehicle / "Отчет_car111.txt").write_text(
                "ПРИМЕЧАНИЕ\nНе ТС\n",
                encoding="utf-8",
            )
            (second_vehicle / "Отчет_car222.txt").write_text(
                "ДУТ\nКол-во: 1\nДУТ 1\nДанные: есть\nБК\nДанные: есть\n",
                encoding="utf-8",
            )

            reports = read_folder_reports([first_root, second_root])

            self.assertEqual(
                reports,
                {"car111": "Гео НЕТ, ДУТ НЕТ.", "car222": "Гео есть, ДУТ есть."},
            )

    def test_reads_and_applies_reports_from_zip_archive(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            archive_path = root / "reports.zip"
            with zipfile.ZipFile(archive_path, "w") as archive:
                archive.writestr(
                    "Отчеты/car123/Отчет_car123.txt",
                    "БК\nДанные: есть\nДУТ\nКол-во: 1\n",
                )
                archive.writestr("data.txt", "metadata")

            reports, vehicle_count, report_count, records = _read_folder_reports(
                archive_path
            )
            self.assertEqual(reports, {"car123": "Гео есть, ДУТ есть."})
            self.assertEqual(vehicle_count, 1)
            self.assertEqual(report_count, 1)
            self.assertIn("reports.zip!/Отчеты/car123/Отчет_car123.txt", str(records[0].report_path))

            workbook_path = root / "Исходный.xlsx"
            workbook = Workbook()
            sheet = active_sheet(workbook)
            sheet.append(["Госномер", "Комментарий"])
            sheet.append(["CAR123", None])
            workbook.save(workbook_path)

            result = update_workbook(
                archive_path,
                workbook_path,
                root / "Результат.xlsx",
            )

            updated = load_workbook(root / "Результат.xlsx")
            self.assertEqual(
                active_sheet(updated)["B2"].value,
                "Гео есть, ДУТ есть.",
            )
            updated.close()
            extract = load_workbook(result.extract_path, read_only=True)
            self.assertIn(
                "reports.zip!/Отчеты/car123/Отчет_car123.txt",
                extract["Распознанные отчеты"]["C2"].value,
            )
            extract.close()

    def test_rejects_zip_archive_without_reports(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            archive_path = Path(temporary_directory) / "empty.zip"
            with zipfile.ZipFile(archive_path, "w") as archive:
                archive.writestr("data.txt", "metadata")

            with self.assertRaisesRegex(ValueError, "нет папки «Отчеты»"):
                _read_folder_reports(archive_path)

    def test_reads_compact_status_reports_named_after_vehicle_folder(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            vehicle_folder = root / "B315ta98"
            vehicle_folder.mkdir()
            (vehicle_folder / "B315ta98.txt").write_text(
                "Гео +\nДУТ -\n",
                encoding="utf-8",
            )

            reports = read_folder_reports(root)

            self.assertEqual(reports, {"b315ta98": "Гео есть, ДУТ нет."})

    def test_uses_newest_report_when_plate_is_in_multiple_roots(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            old_report = root / "old" / "car111" / "Отчет_old.txt"
            new_report = root / "new" / "car111" / "Отчет_new.txt"
            old_report.parent.mkdir(parents=True)
            new_report.parent.mkdir(parents=True)
            old_report.write_text("БК\nДанные: нет\nДУТ\nКол-во: 0\n", encoding="utf-8")
            new_report.write_text("БК\nДанные: есть\nДУТ\nКол-во: 1\n", encoding="utf-8")
            old_report.touch()
            new_report.touch()
            os.utime(old_report, (100, 100))
            os.utime(new_report, (200, 200))

            reports = read_folder_reports([old_report.parents[1], new_report.parents[1]])

            self.assertEqual(reports["car111"], "Гео есть, ДУТ есть.")

    def test_counts_duplicate_report_folders_as_one_unique_plate(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            first = root / "first" / "car111"
            second = root / "second" / "car111"
            first.mkdir(parents=True)
            second.mkdir(parents=True)
            (first / "Отчет_car111.txt").write_text("БК\nДанные: нет\n", encoding="utf-8")
            (second / "Отчет_car111.txt").write_text("БК\nДанные: есть\n", encoding="utf-8")

            reports, vehicle_folder_count, report_folder_count, _ = _read_folder_reports(
                [first.parent, second.parent]
            )

            self.assertEqual(vehicle_folder_count, 2)
            self.assertEqual(report_folder_count, 2)
            self.assertEqual(len(reports), 1)

    def test_writes_reconciliation_extract_for_requested_categories(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            records = [
                ReportFolderRecord(
                    "car111",
                    "car111",
                    root / "first" / "car111",
                    root / "first" / "car111" / "Отчет_car111.txt",
                    "Гео есть, ДУТ есть.",
                ),
                ReportFolderRecord(
                    "car111",
                    "car111",
                    root / "second" / "car111",
                    root / "second" / "car111" / "Отчет_car111.txt",
                    "Гео НЕТ, ДУТ НЕТ.",
                ),
                ReportFolderRecord(
                    "car222",
                    "car222",
                    root / "car222",
                    root / "car222" / "car222.txt",
                    "Гео есть, ДУТ нет.",
                ),
            ]
            extract_path = _write_reconciliation_extract(
                root / "Результат.xlsx",
                records,
                workbook_plate_keys={"car111"},
                duplicate_plate_keys={"car111"},
                merged_cell_rows=[("car111", "Лист1", 5, "C5", "Гео есть, ДУТ есть.")],
                existing_comment_rows=[("car111", "Лист1", 6, "C6", "Ручной текст")],
            )

            extract = load_workbook(extract_path, read_only=True)
            self.assertEqual(extract["Распознанные отчеты"].max_row, 4)
            self.assertEqual(extract["Отсутствуют в Excel"].max_row, 2)
            self.assertEqual(extract["Отсутствуют в Excel"]["A2"].value, "car222")
            self.assertEqual(extract["Дубли папок"].max_row, 3)
            self.assertEqual(extract["Дубли папок"]["B2"].value, 2)
            self.assertEqual(extract["Объединенные ячейки"]["A2"].value, "car111")
            self.assertEqual(
                extract["Существующие комментарии"]["E2"].value,
                "Ручной текст",
            )
            extract.close()

    def test_writes_report_interpretation_for_matching_plate(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            reports = root / "Отчеты"
            vehicle_folder = reports / "b309tt78"
            vehicle_folder.mkdir(parents=True)
            (vehicle_folder / "Отчет_b309tt78.txt").write_text(
                "ПРИМЕЧАНИЕ\n----------\nНет данных по ДУТ",
                encoding="utf-8",
            )

            workbook_path = root / "Исходный.xlsx"
            workbook = Workbook()
            sheet = active_sheet(workbook)
            sheet.append(["Госномер", "Модель ТС", "Комментарий"])
            sheet.append(["B309TT78", "Грузовик", None])
            sheet.append(["other123", "Автобус", "Старый комментарий"])
            other_sheet = workbook.create_sheet("Другой лист")
            other_sheet.append(["Госномер", "Комментарий"])
            other_sheet.append(["b309tt78", None])
            formula_sheet = workbook.create_sheet("Арсенальная")
            formula_sheet["B3"] = '=IFERROR(INDEX(Результат!B$2:B$685,1),"")'
            workbook.save(workbook_path)
            with zipfile.ZipFile(workbook_path, "r") as source:
                entries = {
                    item.filename: (item, source.read(item.filename))
                    for item in source.infolist()
                }
            shared_strings_path = "xl/sharedStrings.xml"
            entries[shared_strings_path] = (
                zipfile.ZipInfo(shared_strings_path),
                b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                b'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                b'count="0" uniqueCount="1"><si><t>Unused</t></si></sst>',
            )
            relationships_info, relationships_xml = entries["xl/_rels/workbook.xml.rels"]
            relationships_xml = relationships_xml.replace(
                b"</Relationships>",
                b'<Relationship Id="rIdSharedStrings" '
                b'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" '
                b'Target="sharedStrings.xml"/></Relationships>',
            )
            entries["xl/_rels/workbook.xml.rels"] = (relationships_info, relationships_xml)
            content_types_info, content_types_xml = entries["[Content_Types].xml"]
            content_types_xml = content_types_xml.replace(
                b"</Types>",
                b'<Override PartName="/xl/sharedStrings.xml" '
                b'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
                b"</Types>",
            )
            entries["[Content_Types].xml"] = (content_types_info, content_types_xml)
            sheet_info, sheet_xml = entries["xl/worksheets/sheet3.xml"]
            self.assertIn(b'<c r="B3"><f>', sheet_xml)
            sheet_xml, replacements = re.subn(
                rb"<v\s*/>",
                b"<v>42</v>",
                sheet_xml,
                count=1,
            )
            self.assertEqual(replacements, 1)
            entries["xl/worksheets/sheet3.xml"] = (sheet_info, sheet_xml)
            with zipfile.ZipFile(workbook_path, "w") as destination:
                for entry, content in entries.values():
                    destination.writestr(entry, content)

            output_path = root / "Результат.xlsx"
            result = update_workbook(reports, workbook_path, output_path)

            self.assertEqual(result.comments_written, 2)
            self.assertEqual(result.rows_without_report, 1)
            self.assertEqual(result.rows_without_interpretation, 0)
            self.assertEqual(result.rows_skipped_merged_cells, 0)
            self.assertEqual(result.worksheets_updated, 2)
            self.assertEqual(result.vehicle_folders_found, 1)
            self.assertEqual(result.report_folders_found, 1)
            self.assertEqual(result.unique_report_plates, 1)
            self.assertEqual(result.report_plates_not_in_workbook, 0)
            self.assertTrue(result.extract_path.is_file())
            self.assertNotEqual(result.extract_path, workbook_path.resolve())
            extract = load_workbook(result.extract_path, read_only=True)
            self.assertEqual(
                extract.sheetnames,
                [
                    "Распознанные отчеты",
                    "Отсутствуют в Excel",
                    "Дубли папок",
                    "Объединенные ячейки",
                    "Существующие комментарии",
                ],
            )
            self.assertEqual(extract["Распознанные отчеты"]["A2"].value, "b309tt78")
            extract.close()

            updated = load_workbook(output_path)
            updated_sheet = active_sheet(updated)
            self.assertEqual(updated_sheet["C2"].value, "Гео НЕТ, ДУТ НЕТ.")
            self.assertEqual(updated_sheet["C3"].value, "Старый комментарий")
            self.assertEqual(updated["Другой лист"]["B2"].value, "Гео НЕТ, ДУТ НЕТ.")
            self.assertEqual(
                updated["Арсенальная"]["B3"].value,
                '=IFERROR(INDEX(Результат!B$2:B$685,1),"")',
            )
            values_only = load_workbook(output_path, data_only=True)
            self.assertEqual(values_only["Арсенальная"]["B3"].value, 42)
            with zipfile.ZipFile(workbook_path, "r") as original_archive:
                original_workbook_xml = original_archive.read("xl/workbook.xml")
            with zipfile.ZipFile(output_path, "r") as updated_archive:
                self.assertEqual(
                    updated_archive.read("xl/workbook.xml"),
                    original_workbook_xml,
                )
                updated_strings = ET.fromstring(updated_archive.read("xl/sharedStrings.xml"))
                for worksheet_path in (
                    "xl/worksheets/sheet1.xml",
                    "xl/worksheets/sheet2.xml",
                ):
                    worksheet_root = ET.fromstring(updated_archive.read(worksheet_path))
                    cells = worksheet_root.iter(
                        "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c"
                    )
                    self.assertTrue(all(cell.get("r") for cell in cells))
            string_items = updated_strings.findall(
                "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si"
            )
            self.assertEqual(updated_strings.get("count"), "2")
            self.assertEqual(updated_strings.get("uniqueCount"), str(len(string_items)))
            self.assertEqual(
                [
                    "".join(string_item.itertext())
                    for string_item in string_items
                ],
                ["Unused", "Гео НЕТ, ДУТ НЕТ."],
            )
            self.assertEqual(
                updated_sheet["C2"].value,
                "Гео НЕТ, ДУТ НЕТ.",
            )
            original = load_workbook(workbook_path)
            self.assertIsNone(active_sheet(original)["C2"].value)

    def test_skip_existing_comments_preserves_values_and_fills_empty_cells(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            reports = root / "Отчеты"
            vehicle_folder = reports / "car123"
            vehicle_folder.mkdir(parents=True)
            (vehicle_folder / "Отчет_car123.txt").write_text(
                "БК\nДанные: есть\nДУТ\nКол-во: 1\n",
                encoding="utf-8",
            )
            workbook_path = root / "Исходный.xlsx"
            workbook = Workbook()
            sheet = active_sheet(workbook)
            sheet.append(["Госномер", "Комментарий"])
            sheet.append(["car123", "Не менять"])
            sheet.append(["car123", None])
            workbook.save(workbook_path)

            result = update_workbook(
                reports,
                workbook_path,
                root / "Результат.xlsx",
                skip_existing_comments=True,
            )

            updated = load_workbook(root / "Результат.xlsx")
            updated_sheet = active_sheet(updated)
            self.assertEqual(updated_sheet["B2"].value, "Не менять")
            self.assertEqual(updated_sheet["B3"].value, "Гео есть, ДУТ есть.")
            self.assertEqual(result.comments_written, 1)
            self.assertEqual(result.rows_skipped_existing_comments, 1)
            updated.close()
            extract = load_workbook(result.extract_path, read_only=True)
            self.assertEqual(
                extract["Существующие комментарии"]["E2"].value,
                "Не менять",
            )
            extract.close()
            overwrite_result = update_workbook(
                reports,
                workbook_path,
                root / "Перезапись.xlsx",
            )
            overwrite = load_workbook(root / "Перезапись.xlsx")
            self.assertEqual(
                active_sheet(overwrite)["B2"].value,
                "Гео есть, ДУТ есть.",
            )
            self.assertEqual(overwrite_result.rows_skipped_existing_comments, 0)
            overwrite.close()

    def test_overwrite_source_keeps_original_path_and_name(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            reports = root / "Отчеты"
            vehicle_folder = reports / "car123"
            vehicle_folder.mkdir(parents=True)
            (vehicle_folder / "Отчет_car123.txt").write_text(
                "БК\nДанные: есть\nДУТ\nКол-во: 1\n",
                encoding="utf-8",
            )
            workbook_path = root / "ИсходнаяКнига.xlsx"
            workbook = Workbook()
            sheet = active_sheet(workbook)
            sheet.append(["Госномер", "Комментарий"])
            sheet.append(["car123", None])
            workbook.save(workbook_path)

            with self.assertRaisesRegex(ValueError, "новый файл"):
                update_workbook(reports, workbook_path, workbook_path)

            result = update_workbook(
                reports,
                workbook_path,
                root / "ДругоеИмя.xlsx",
                overwrite_source=True,
            )

            self.assertEqual(result.extract_path.name, "ИсходнаяКнига_выписка.xlsx")
            self.assertFalse((root / "ДругоеИмя.xlsx").exists())
            updated = load_workbook(workbook_path)
            self.assertEqual(
                active_sheet(updated)["B2"].value,
                "Гео есть, ДУТ есть.",
            )
            updated.close()

    def test_skips_merged_comment_cells(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            vehicle_folder = root / "Отчеты" / "car123"
            vehicle_folder.mkdir(parents=True)
            (vehicle_folder / "Отчет_car123.txt").write_text(
                "ПРИМЕЧАНИЕ\nНет данных",
                encoding="utf-8",
            )

            workbook_path = root / "Исходный.xlsx"
            workbook = Workbook()
            sheet = active_sheet(workbook)
            sheet.append(["Госномер", "Модель", "Комментарий"])
            sheet.append(["car123", "Объединенная ячейка", None])
            sheet.merge_cells("B2:C2")
            workbook.save(workbook_path)

            result = update_workbook(
                root / "Отчеты",
                workbook_path,
                root / "Результат.xlsx",
            )
            self.assertEqual(result.comments_written, 0)
            self.assertEqual(result.rows_skipped_merged_cells, 1)
            extract = load_workbook(result.extract_path, read_only=True)
            self.assertEqual(extract["Объединенные ячейки"]["A2"].value, "car123")
            extract.close()


if __name__ == "__main__":
    unittest.main()
