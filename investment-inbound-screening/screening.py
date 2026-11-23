"""Extract application rows and append Codex's assessments to an Excel copy."""

import argparse
import json
from pathlib import Path
from shutil import copyfileobj
from tempfile import NamedTemporaryFile

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


COLUMNS = {
    "fit": "Thesis fit",
    "recommendation": "Recommendation",
    "reasoning": "Reasoning",
    "follow_up": "Missing information / key follow-up questions",
}
OUTCOMES = {
    "CALL": {"Strong"},
    "MORE INFO": {"Potential", "Unknown"},
    "RELEVANT FOR MARKET MAP": {"Potential"},
    "PASS": {"Poor"},
}


def present(cell):
    return cell.hyperlink is not None or (
        cell.value is not None and str(cell.value).strip() != ""
    )


def read_rows(sheet):
    headers = [(cell.column, str(cell.value).strip()) for cell in sheet[1] if present(cell)]
    names = [" ".join(label.split()).casefold() for _, label in headers]
    if "company name" not in names or len(names) != len(set(names)):
        raise ValueError("Row 1 must contain unique headers including Company Name.")
    output_names = {label.casefold() for label in COLUMNS.values()}
    if output_names.intersection(names):
        raise ValueError("Use the original application workbook.")

    rows = []
    for cells in sheet.iter_rows(min_row=2):
        values = {}
        for column, label in headers:
            cell = cells[column - 1]
            if present(cell):
                if cell.hyperlink:
                    values[label] = {
                        "value": cell.value,
                        "link": cell.hyperlink.target or cell.hyperlink.location,
                    }
                else:
                    values[label] = cell.value
        if values:
            rows.append({"row": cells[0].row, "application": values})
    return rows


def validate_results(results, rows):
    if not isinstance(results, list):
        raise ValueError("Results must be a JSON array.")
    expected = {item["row"] for item in rows}
    by_row = {}
    for item in results:
        if not isinstance(item, dict) or type(item.get("row")) is not int:
            raise ValueError("Each result needs an integer Excel row number.")
        row = item["row"]
        if row not in expected or row in by_row:
            raise ValueError(f"Unexpected or duplicate result for Excel row {row}.")
        if any(
            not isinstance(item.get(field), str)
            or not item[field].strip()
            or len(item[field]) > 32767
            for field in COLUMNS
        ):
            raise ValueError(f"Excel row {row} needs text in all four output fields.")
        if item["fit"] not in OUTCOMES.get(item["recommendation"], set()):
            raise ValueError(f"Invalid fit and recommendation for Excel row {row}.")
        by_row[row] = item
    missing = expected - by_row.keys()
    if missing:
        raise ValueError(f"Missing results for Excel rows {sorted(missing)}.")
    return by_row


def write_results(sheet, by_row):
    start = max(cell.column for cell in sheet[1] if present(cell)) + 1
    for offset, (field, label) in enumerate(COLUMNS.items()):
        column = start + offset
        header = sheet.cell(1, column, label)
        header.fill = PatternFill("solid", fgColor="19364A")
        header.font = Font(bold=True, color="FFFFFF")
        header.alignment = Alignment(wrap_text=True, vertical="center")
        sheet.column_dimensions[get_column_letter(column)].width = (16, 32, 64, 53)[offset]
        for row, item in by_row.items():
            cell = sheet.cell(row, column, item[field])
            cell.data_type = "s"
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            sheet.row_dimensions[row].height = max(sheet.row_dimensions[row].height or 0, 66)
    sheet.row_dimensions[1].height = max(sheet.row_dimensions[1].height or 0, 38)
    if not sheet.freeze_panes:
        sheet.freeze_panes = "A2"
    if not sheet.tables:
        last = get_column_letter(start + len(COLUMNS) - 1)
        sheet.auto_filter.ref = f"A1:{last}{max(by_row, default=1)}"


def save_copy(workbook, output):
    with NamedTemporaryFile(dir=output.parent, suffix=".xlsx", delete=False) as temp:
        temporary = Path(temp.name)
    try:
        workbook.save(temporary)
        with temporary.open("rb") as source, output.open("xb") as target:
            copyfileobj(source, target)
    finally:
        temporary.unlink()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    for name in ("extract", "export"):
        command = commands.add_parser(name)
        command.add_argument("input", type=Path)
        if name == "export":
            command.add_argument("results", type=Path)
        command.add_argument("output", type=Path)
        command.add_argument("--sheet", help="Application sheet name")
    args = parser.parse_args()

    try:
        suffix = ".json" if args.command == "extract" else ".xlsx"
        if args.input.suffix.lower() != ".xlsx" or args.output.suffix.lower() != suffix:
            raise ValueError(f"Use an .xlsx input and a {suffix} output.")
        if args.output.exists():
            raise ValueError(f"Output already exists. Choose a new path: {args.output}")
        workbook = load_workbook(args.input)
        try:
            sheet = workbook[args.sheet] if args.sheet else workbook.worksheets[0]
            rows = read_rows(sheet)
            if args.command == "extract":
                with args.output.open("x", encoding="utf-8") as output:
                    json.dump(rows, output, ensure_ascii=False, indent=2, default=str)
            else:
                results = json.loads(args.results.read_text(encoding="utf-8-sig"))
                write_results(sheet, validate_results(results, rows))
                save_copy(workbook, args.output)
        finally:
            workbook.close()
    except (OSError, ValueError, KeyError) as error:
        parser.exit(1, f"Error: {error}\n")
    print(f"{len(rows)} application rows. Saved {args.output}")


if __name__ == "__main__":
    main()
