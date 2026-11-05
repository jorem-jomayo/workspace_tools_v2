import csv
import json
import sys
from datetime import date, datetime, time
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


SOURCE_COLUMNS = (42, 3, 4, 5, 28, 33, 37, 27)
PREVIEW_LIMIT = 100


def export_value(value):
    if isinstance(value, datetime):
        return value.strftime("%m/%d/%Y")
    if isinstance(value, date):
        return value.strftime("%m/%d/%Y")
    if isinstance(value, time):
        return value.strftime("%H:%M:%S")
    return value


def json_value(value):
    value = export_value(value)
    if value is None:
        return ""
    if isinstance(value, (str, int, float, bool)):
        return value
    return str(value)


def bank_names(source_path):
    workbook = load_workbook(source_path, read_only=True, data_only=True)
    try:
        sheet_name = next(
            (name for name in workbook.sheetnames if name.strip().upper() == "RESULT"),
            None,
        )
        if sheet_name is None:
            raise ValueError("A sheet named RESULT was not found in this workbook.")
        values = {}
        for row in workbook[sheet_name].iter_rows(
            min_row=2,
            min_col=33,
            max_col=33,
            values_only=True,
        ):
            bank = str(row[0]).strip() if row[0] is not None else ""
            if bank:
                values.setdefault(bank.casefold(), bank)
        return sorted(values.values(), key=str.casefold)
    finally:
        workbook.close()


def process(source_path, output_dir, bank_filters=None):
    if isinstance(bank_filters, dict):
        campaign = bank_filters.get("campaign", "pif")
        selected_bank = bank_filters.get("bank", "")
        if campaign not in ("all", "pif") or not isinstance(selected_bank, str):
            raise ValueError("Expected a valid campaign and bank selection.")
        selected_bank = selected_bank.strip().casefold()
        selected_banks = set()
    else:
        campaign = "all" if bank_filters else "pif"
        selected_bank = ""
        selected_banks = {
            bank.strip().casefold()
            for bank in (bank_filters or [])
            if bank.strip()
        }
    workbook = load_workbook(source_path, read_only=True, data_only=True)
    sheet_name = next(
        (name for name in workbook.sheetnames if name.strip().upper() == "RESULT"),
        None,
    )
    if sheet_name is None:
        raise ValueError("A sheet named RESULT was not found in this workbook.")

    source_sheet = workbook[sheet_name]
    output_book = Workbook(write_only=True)
    output_sheet = output_book.create_sheet("RESULT")
    font = Font(name="Calibri", size=9)
    column_widths = [10] * len(SOURCE_COLUMNS)
    preview_rows = []
    row_count = 0
    matched_data_rows = 0

    output_path = Path(output_dir)
    date_stamp = date.today().strftime("%m-%d-%Y")
    workbook_path = output_path / f"FIELD RESULT {date_stamp}.xlsx"
    csv_path = output_path / "fieldman.csv"

    with csv_path.open("w", encoding="utf-8-sig", newline="") as csv_file:
        csv_writer = csv.writer(csv_file, lineterminator="\r\n")
        csv_writer.writerow(["Ref Code", "Field Name"])

        for source_row in source_sheet.iter_rows(
            min_col=1,
            max_col=max(SOURCE_COLUMNS),
            values_only=True,
        ):
            is_header = row_count == 0
            bank = source_row[32] if len(source_row) > 32 else ""
            normalized_bank = str(bank or "").strip().casefold()
            if (
                not is_header
                and (
                    (campaign == "pif" and "pif" not in normalized_bank)
                    or (selected_bank and normalized_bank != selected_bank)
                    or (selected_banks and normalized_bank not in selected_banks)
                )
            ):
                continue

            selected = [
                export_value(source_row[column - 1])
                if column <= len(source_row)
                else ""
                for column in SOURCE_COLUMNS
            ]
            row_count += 1
            if not is_header:
                matched_data_rows += 1
            if row_count <= PREVIEW_LIMIT:
                preview_rows.append([json_value(value) for value in selected])

            output_cells = []
            for index, value in enumerate(selected):
                column_widths[index] = max(
                    column_widths[index],
                    min(len(str(value)) + 3 if value is not None else 10, 80),
                )
                cell = WriteOnlyCell(output_sheet, value=value)
                cell.font = font
                output_cells.append(cell)
            output_sheet.append(output_cells)

            if row_count > 1:
                reference_code = source_row[42 - 1] if len(source_row) >= 42 else ""
                field_name = source_row[27 - 1] if len(source_row) >= 27 else ""
                csv_writer.writerow(
                    [
                        "" if reference_code is None else reference_code,
                        "" if field_name is None else field_name,
                    ]
                )

    if row_count == 0:
        workbook.close()
        raise ValueError("The RESULT sheet is empty or contains no readable rows.")
    if (selected_banks or selected_bank or campaign == "pif") and matched_data_rows == 0:
        workbook.close()
        raise ValueError("No data rows were found for the selected campaign or bank in column AG.")

    for index, width in enumerate(column_widths, start=1):
        output_sheet.column_dimensions[get_column_letter(index)].width = width
    output_book.save(workbook_path)
    workbook.close()

    return {
        "campaign": campaign,
        "rows": row_count,
        "previewRows": preview_rows,
        "workbookName": workbook_path.name,
    }


def main():
    if len(sys.argv) == 3 and sys.argv[1] == "--list-banks":
        print(json.dumps(bank_names(Path(sys.argv[2])), ensure_ascii=False))
        return
    if len(sys.argv) not in (3, 4):
        raise ValueError("Expected an input workbook and output directory.")
    bank_filters = (
        json.loads(Path(sys.argv[3]).read_text(encoding="utf-8"))
        if len(sys.argv) == 4
        else []
    )
    if isinstance(bank_filters, list):
        if any(not isinstance(bank, str) or len(bank) > 256 for bank in bank_filters):
            raise ValueError("Expected a list of bank names.")
    elif (
        not isinstance(bank_filters, dict)
        or bank_filters.get("campaign", "pif") not in ("all", "pif")
        or not isinstance(bank_filters.get("bank", ""), str)
        or len(bank_filters.get("bank", "")) > 256
    ):
        raise ValueError("Expected valid Field Result filters.")
    result = process(Path(sys.argv[1]), Path(sys.argv[2]), bank_filters)
    print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        print(str(error), file=sys.stderr)
        sys.exit(1)
