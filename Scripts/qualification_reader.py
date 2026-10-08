import json
from pathlib import Path

from openpyxl import load_workbook


PROJECT_ROOT = Path(__file__).resolve().parent.parent
SETTINGS_PATH = PROJECT_ROOT / "Config" / "settings.json"

EXPECTED_QUALIFICATIONS = {
    "DUMOD0",
    "DUMOD1",
    "CUMOD0",
    "CUMOD1",
    "CUMOD2",
    "CUMOD3",
    "UCMOD0",
    "UCMOD1",
    "UCMOD2",
    "UCMOD3",
    "UCMOD4",
}


def load_settings():
    if not SETTINGS_PATH.exists():
        raise FileNotFoundError(
            f"Settings file not found: {SETTINGS_PATH}"
        )

    with SETTINGS_PATH.open(
        mode="r",
        encoding="utf-8",
    ) as settings_file:
        return json.load(settings_file)


def get_qualification_folder(settings):
    qualification_folder = Path(
        settings["qualification_folder"]
    )

    if not qualification_folder.exists():
        raise FileNotFoundError(
            "Qualification folder not found:\n"
            f"{qualification_folder}"
        )

    return qualification_folder


def find_qualification_files(qualification_folder):
    supported_extensions = {
        ".xlsx",
        ".xlsm",
    }

    return sorted(
        file
        for file in qualification_folder.iterdir()
        if file.is_file()
        and file.suffix.lower() in supported_extensions
        and not file.name.startswith("~$")
    )


def identify_qualification_code(file_path):
    filename_upper = file_path.stem.upper()

    matches = [
        qualification_code
        for qualification_code in EXPECTED_QUALIFICATIONS
        if qualification_code in filename_upper
    ]

    if len(matches) == 1:
        return matches[0]

    if not matches:
        return "UNRECOGNIZED"

    return "MULTIPLE MATCHES"


def inspect_workbook(file_path):
    workbook = load_workbook(
        file_path,
        read_only=True,
        data_only=True,
    )

    print("=" * 80)
    print(f"File: {file_path.name}")
    print(f"Worksheets: {workbook.sheetnames}")

    for worksheet in workbook.worksheets:
        worksheet.reset_dimensions()

        print()
        print(f"Worksheet: {worksheet.title}")
        print(
            f"Dimensions: "
            f"{worksheet.max_row} rows x "
            f"{worksheet.max_column} columns"
        )

        populated_rows_shown = 0

        for row_number, row in enumerate(
            worksheet.iter_rows(values_only=True),
            start=1,
        ):
            populated_values = [
                f"C{column_number}={value!r}"
                for column_number, value in enumerate(
                    row,
                    start=1,
                )
                if value is not None
            ]

            if not populated_values:
                continue

            print(
                f"Row {row_number}: "
                + " | ".join(populated_values)
            )

            populated_rows_shown += 1

            if populated_rows_shown >= 6:
                break

    workbook.close()
    print()


def inspect_qualification_files():
    settings = load_settings()
    qualification_folder = get_qualification_folder(settings)
    files = find_qualification_files(qualification_folder)

    print(f"Qualification folder: {qualification_folder}")
    print(f"Workbook count: {len(files)}")
    print()

    recognized_codes = set()

    for file_path in files:
        qualification_code = identify_qualification_code(
            file_path
        )

        if qualification_code not in EXPECTED_QUALIFICATIONS:
            print(
                f"Skipped: {file_path.name} "
                f"({qualification_code})"
            )
            continue

        recognized_codes.add(qualification_code)
        print(f"Qualification code: {qualification_code}")
        inspect_workbook(file_path)

    missing_codes = EXPECTED_QUALIFICATIONS - recognized_codes

    print("=" * 80)
    print(f"Recognized qualifications: {len(recognized_codes)}")

    if missing_codes:
        print(
            "Missing qualification codes: "
            + ", ".join(sorted(missing_codes))
        )
    else:
        print("Missing qualification codes: None")


if __name__ == "__main__":
    inspect_qualification_files()