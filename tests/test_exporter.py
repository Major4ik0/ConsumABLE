import os
import openpyxl
from core.exporter import export_report_to_excel


def test_export_report_creates_three_sheets_with_formulas(tmp_path):
    out_file = str(tmp_path / "test_report.xlsx")

    cadets = [
        {"id": 1, "rank": "курсант", "fio": "Иванов И.И.", "short_fio": "Иванов И.И.", "gender": "male", "group_num": "101", "course": 1},
        {"id": 2, "rank": "курсант", "fio": "Смирнова А.С.", "short_fio": "Смирнова А.С.", "gender": "female", "group_num": "101", "course": 1},
    ]
    absences = [
        {"id": 1, "cadet_id": 1, "rank": "курсант", "short_fio": "Иванов И.И.", "gender": "male", "group_num": "101", "type_name": "Дежурный", "category": "Наряд", "note": "КПП"}
    ]

    export_report_to_excel(
        output_path=out_file,
        report_date="2026-11-04",
        cadets=cadets,
        absences=absences,
        unit_title="1 курс 1 факультета",
        chief_title="Начальник курса",
        chief_name="майор В.Медков",
        active_categories=["Наряд", "Госпиталь", "Лазарет"]
    )

    assert os.path.exists(out_file)

    wb = openpyxl.load_workbook(out_file, data_only=False)
    # Check that all 3 sheets exist
    assert "Сводный расход" in wb.sheetnames
    assert "Мужской состав" in wb.sheetnames
    assert "Женский состав" in wb.sheetnames

    ws_total = wb["Сводный расход"]

    # Verify that cells contain real Excel formulas (e.g. SUM, difference), not static baked numbers!
    formula_cells = []
    for row in ws_total.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str) and cell.value.startswith("="):
                formula_cells.append(cell.value)

    assert len(formula_cells) > 0
    assert any("SUM" in f for f in formula_cells)
