import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from datetime import datetime
from typing import List, Dict, Any, Optional

from core.models import normalize_gender


def style_range(ws, cell_range, border=None, font=None, alignment=None, fill=None):
    """Applies styles to a rectangular range of cells."""
    for row in ws[cell_range]:
        for cell in row:
            if border: cell.border = border
            if font: cell.font = font
            if alignment: cell.alignment = alignment
            if fill: cell.fill = fill


def build_sheet_content(
    ws,
    sheet_title_suffix: str,
    report_date: str,
    unit_title: str,
    chief_title: str,
    chief_name: str,
    cadets: List[Dict[str, Any]],
    absences: List[Dict[str, Any]],
    categories: List[str]
):
    """Fills a single worksheet with structured строевая записка using real Excel formulas."""
    bold_font = Font(name="Times New Roman", size=12, bold=True)
    norm_font = Font(name="Times New Roman", size=12)
    title_font = Font(name="Times New Roman", size=12, bold=True)
    header_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    thin = Side(border_style="thin", color="000000")
    cell_border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # 1. Заголовки (строки 2-4)
    ws.merge_cells("B2:L2")
    ws["B2"] = f"РАЗВЕРНУТАЯ СТРОЕВАЯ ЗАПИСКА {sheet_title_suffix}".strip()
    style_range(ws, "B2:L2", font=title_font, alignment=center)

    ws.merge_cells("B3:L3")
    ws["B3"] = unit_title
    style_range(ws, "B3:L3", font=norm_font, alignment=center)

    ws.merge_cells("B4:L4")
    try:
        dt_val = datetime.strptime(report_date, "%Y-%m-%d").strftime("%d.%m.%Y г.")
    except Exception:
        dt_val = report_date
    ws["B4"] = dt_val
    style_range(ws, "B4:L4", font=norm_font, alignment=center)

    ws.row_dimensions[4].height = 17.0

    group_order_map = {}
    for c in cadets:
        grp = c.get('group_num', '1')
        if grp not in group_order_map:
            group_order_map[grp] = c.get('sort_order', 0)
    groups = sorted(group_order_map.keys(), key=lambda g: group_order_map[g])

    start_row = 5
    ws.row_dimensions[start_row].height = 18.0

    headers = ["№ п/п", "Подразд.", "П/С", "Н/Л"] + categories
    # Header cells start at column 2 (B)
    for idx, h in enumerate(headers, start=2):
        cell = ws.cell(row=start_row, column=idx, value=h)
        cell.font = norm_font
        cell.fill = header_fill
        cell.border = cell_border
        cell.alignment = center

    # Column "ВСЕГО" merged on 2 columns (e.g. K and L)
    v_col1 = 2 + len(headers)
    v_col2 = v_col1 + 1
    v_col1_letter = get_column_letter(v_col1)
    v_col2_letter = get_column_letter(v_col2)

    ws.merge_cells(start_row=start_row, start_column=v_col1, end_row=start_row, end_column=v_col2)
    ws.cell(row=start_row, column=v_col1, value="ВСЕГО")
    style_range(ws, f"{v_col1_letter}{start_row}:{v_col2_letter}{start_row}",
                border=cell_border, font=norm_font, fill=header_fill, alignment=center)

    cat_start_col = 6  # Column F (№=B, Подразд=C, П/С=D, Н/Л=E? Wait! Let's check headers:)
    # headers = ["№ п/п", "Подразд.", "П/С", "Н/Л"] + categories
    # B: col 2 = "№ п/п"
    # C: col 3 = "Подразд."
    # D: col 4 = "П/С"
    # E: col 5 = "Н/Л"
    # F: col 6 .. col 6+len(categories)-1 = categories
    # v_col1 = 2 + len(headers) = 2 + 4 + len(categories) = 6 + len(categories)
    first_cat_col = 6
    last_cat_col = first_cat_col + len(categories) - 1
    first_cat_letter = get_column_letter(first_cat_col)
    last_cat_letter = get_column_letter(last_cat_col)

    curr_row = start_row + 1
    first_grp_row = curr_row

    for idx, grp in enumerate(groups, 1):
        ws.row_dimensions[curr_row].height = 17.0
        grp_cadets = [c for c in cadets if str(c.get('group_num')) == str(grp)]
        ps = len(grp_cadets)
        grp_abs = [a for a in absences if str(a.get('group_num')) == str(grp)]

        row_counts = {}
        for cat in categories:
            cnt = sum(1 for a in grp_abs if a.get('category') == cat)
            row_counts[cat] = cnt

        # № п/п
        ws.cell(row=curr_row, column=2, value=idx).alignment = center
        # Подразд.
        ws.cell(row=curr_row, column=3, value=grp).alignment = center
        # П/С
        ws.cell(row=curr_row, column=4, value=ps).alignment = center
        # Н/Л (Formula: =D{curr_row}-v_col1{curr_row})
        formula_nl = f"=D{curr_row}-{v_col1_letter}{curr_row}"
        ws.cell(row=curr_row, column=5, value=formula_nl).alignment = center

        # Categories
        for c_offset, cat in enumerate(categories):
            col_num = first_cat_col + c_offset
            c_val = row_counts[cat] if row_counts[cat] > 0 else ""
            cell = ws.cell(row=curr_row, column=col_num, value=c_val)
            cell.alignment = center

        # Formula for ВСЕГО: =SUM(first_cat_letter{curr_row}:last_cat_letter{curr_row})
        formula_row_total = f"=SUM({first_cat_letter}{curr_row}:{last_cat_letter}{curr_row})"
        ws.merge_cells(start_row=curr_row, start_column=v_col1, end_row=curr_row, end_column=v_col2)
        ws.cell(row=curr_row, column=v_col1, value=formula_row_total)

        style_range(ws, f"B{curr_row}:{v_col2_letter}{curr_row}", border=cell_border, font=norm_font, alignment=center)
        curr_row += 1

    last_grp_row = curr_row - 1 if groups else curr_row

    # Additional sub-rows for "Мужчины" and "Женщины" if multiple genders exist on the sheet
    has_males = any(normalize_gender(c.get('gender')) == 'male' for c in cadets)
    has_females = any(normalize_gender(c.get('gender')) == 'female' for c in cadets)

    if has_males and has_females:
        for g_label, g_norm, num_idx in [
            ("Мужчины", "male", len(groups) + 1),
            ("Женщины", "female", len(groups) + 2)
        ]:
            ws.row_dimensions[curr_row].height = 17.0
            g_cadets = [c for c in cadets if normalize_gender(c.get('gender')) == g_norm]
            g_ps = len(g_cadets)
            g_abs = [a for a in absences if normalize_gender(a.get('gender')) == g_norm]
            row_counts = {cat: sum(1 for a in g_abs if a.get('category') == cat) for cat in categories}

            ws.cell(row=curr_row, column=2, value=num_idx).alignment = center
            ws.cell(row=curr_row, column=3, value=g_label).alignment = center
            ws.cell(row=curr_row, column=4, value=g_ps).alignment = center
            ws.cell(row=curr_row, column=5, value=f"=D{curr_row}-{v_col1_letter}{curr_row}").alignment = center

            for c_offset, cat in enumerate(categories):
                col_num = first_cat_col + c_offset
                c_val = row_counts[cat] if row_counts[cat] > 0 else ""
                ws.cell(row=curr_row, column=col_num, value=c_val).alignment = center

            formula_total = f"=SUM({first_cat_letter}{curr_row}:{last_cat_letter}{curr_row})"
            ws.merge_cells(start_row=curr_row, start_column=v_col1, end_row=curr_row, end_column=v_col2)
            ws.cell(row=curr_row, column=v_col1, value=formula_total)

            style_range(ws, f"B{curr_row}:{v_col2_letter}{curr_row}", border=cell_border, font=norm_font, alignment=center)
            curr_row += 1

    # Строка «Всего» with dynamic formulas
    ws.row_dimensions[curr_row].height = 18.0
    ws.cell(row=curr_row, column=2, value=len(groups) + 3 if (has_males and has_females) else len(groups) + 1)
    ws.cell(row=curr_row, column=3, value="Всего")

    if groups:
        ws.cell(row=curr_row, column=4, value=f"=SUM(D{first_grp_row}:D{last_grp_row})")
        ws.cell(row=curr_row, column=5, value=f"=SUM(E{first_grp_row}:E{last_grp_row})")
        for c_offset in range(len(categories)):
            col_letter = get_column_letter(first_cat_col + c_offset)
            ws.cell(row=curr_row, column=first_cat_col + c_offset, value=f"=SUM({col_letter}{first_grp_row}:{col_letter}{last_grp_row})")
        ws.merge_cells(start_row=curr_row, start_column=v_col1, end_row=curr_row, end_column=v_col2)
        ws.cell(row=curr_row, column=v_col1, value=f"=SUM({first_cat_letter}{curr_row}:{last_cat_letter}{curr_row})")
    else:
        ws.cell(row=curr_row, column=4, value=0)
        ws.cell(row=curr_row, column=5, value=0)
        ws.merge_cells(start_row=curr_row, start_column=v_col1, end_row=curr_row, end_column=v_col2)
        ws.cell(row=curr_row, column=v_col1, value=0)

    style_range(ws, f"B{curr_row}:{v_col2_letter}{curr_row}", border=cell_border, font=bold_font, alignment=center)
    curr_row += 2

    # 3. Список отсутствующих
    last_col_letter = v_col2_letter
    ws.merge_cells(f"B{curr_row}:{last_col_letter}{curr_row}")
    ws[f"B{curr_row}"] = "Отсутствующие"
    style_range(ws, f"B{curr_row}:{last_col_letter}{curr_row}", font=norm_font, alignment=center)
    ws.row_dimensions[curr_row].height = 18.0
    curr_row += 1

    # Заголовки таблицы отсутствующих
    ws.row_dimensions[curr_row].height = 18.0
    ws.cell(row=curr_row, column=2, value="№ п/п").alignment = center
    ws.cell(row=curr_row, column=3, value="В/звание").alignment = center
    ws.merge_cells(start_row=curr_row, start_column=4, end_row=curr_row, end_column=6)
    ws.cell(row=curr_row, column=4, value="Фамилия и инициалы").alignment = center
    ws.merge_cells(start_row=curr_row, start_column=7, end_row=curr_row, end_column=8)
    ws.cell(row=curr_row, column=7, value="Причина отсутствия").alignment = center
    ws.cell(row=curr_row, column=9, value="Подразд.").alignment = center
    ws.merge_cells(start_row=curr_row, start_column=10, end_row=curr_row, end_column=v_col2)
    ws.cell(row=curr_row, column=10, value="Примечание").alignment = center

    style_range(ws, f"B{curr_row}:{last_col_letter}{curr_row}", border=cell_border, font=norm_font, fill=header_fill, alignment=center)
    curr_row += 1

    # Строки отсутствующих
    for idx, a in enumerate(absences, 1):
        ws.row_dimensions[curr_row].height = 18.0
        ws.cell(row=curr_row, column=2, value=idx)
        ws.cell(row=curr_row, column=3, value=a.get('rank', 'курсант'))
        ws.merge_cells(start_row=curr_row, start_column=4, end_row=curr_row, end_column=6)
        ws.cell(row=curr_row, column=4, value=a.get('short_fio') or a.get('fio', ''))
        ws.merge_cells(start_row=curr_row, start_column=7, end_row=curr_row, end_column=8)
        ws.cell(row=curr_row, column=7, value=a.get('type_name', ''))
        ws.cell(row=curr_row, column=9, value=a.get('group_num', ''))

        ws.merge_cells(start_row=curr_row, start_column=10, end_row=curr_row, end_column=v_col2)
        note_text = str(a.get('note') or "").strip()
        ws.cell(row=curr_row, column=10, value=note_text)
        style_range(ws, f"B{curr_row}:{last_col_letter}{curr_row}", border=cell_border, font=norm_font, alignment=center)
        curr_row += 1

    # 4. Подписи начальника
    curr_row += 1
    ws.merge_cells(start_row=curr_row, start_column=2, end_row=curr_row, end_column=v_col2)
    ws.cell(row=curr_row, column=2, value=chief_title.strip())
    style_range(ws, f"B{curr_row}:{last_col_letter}{curr_row}", font=norm_font, alignment=center)

    curr_row += 1
    ws.merge_cells(start_row=curr_row, start_column=2, end_row=curr_row, end_column=v_col2)
    ws.cell(row=curr_row, column=2, value=chief_name)
    style_range(ws, f"B{curr_row}:{last_col_letter}{curr_row}", font=norm_font, alignment=center)

    # 5. Ширина колонок
    col_widths = {
        'A': 3.0, 'B': 7.0, 'C': 11.5, 'D': 8.5, 'E': 9.0, 'F': 13.0,
        'G': 12.5, 'H': 11.5, 'I': 15.0, 'J': 15.0, 'K': 9.0, 'L': 13.0
    }
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

    # 6. Параметры страницы
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    ws.page_margins.left = 0.2
    ws.page_margins.right = 0.2
    ws.page_margins.top = 0.3
    ws.page_margins.bottom = 0.3
    ws.page_margins.header = 0.0
    ws.page_margins.footer = 0.0

    ws.print_options.horizontalCentered = True
    ws.print_area = f"B2:{last_col_letter}{curr_row}"


def export_report_to_excel(
    output_path: str,
    report_date: str,
    cadets: List[Dict[str, Any]],
    absences: List[Dict[str, Any]],
    unit_title: str,
    chief_title: str,
    chief_name: str,
    active_categories: Optional[List[str]] = None
):
    """Generates the multi-slice Excel report with:
    - Sheet 1: «Сводный расход» (Overall summary with formulas and sub-rows)
    - Sheet 2: «Мужской состав» (Male personnel slice with formulas)
    - Sheet 3: «Женский состав» (Female personnel slice with formulas)
    """
    wb = openpyxl.Workbook()

    standard_categories = ["Наряд", "Госпиталь", "Лазарет", "Командировка", "Похороны"]
    if active_categories:
        ordered_cats = [c for c in standard_categories if c in active_categories]
        other_cats = [c for c in active_categories if c not in standard_categories]
        categories = ordered_cats + other_cats
    else:
        categories = standard_categories

    # Sheet 1: Сводный расход
    ws_total = wb.active
    ws_total.title = "Сводный расход"
    build_sheet_content(
        ws_total,
        sheet_title_suffix="",
        report_date=report_date,
        unit_title=unit_title,
        chief_title=chief_title,
        chief_name=chief_name,
        cadets=cadets,
        absences=absences,
        categories=categories
    )

    # Sheet 2: Мужской состав
    ws_male = wb.create_sheet(title="Мужской состав")
    male_cadets = [c for c in cadets if normalize_gender(c.get('gender')) == 'male']
    male_absences = [a for a in absences if normalize_gender(a.get('gender')) == 'male']
    build_sheet_content(
        ws_male,
        sheet_title_suffix="(Мужской состав)",
        report_date=report_date,
        unit_title=unit_title,
        chief_title=chief_title,
        chief_name=chief_name,
        cadets=male_cadets,
        absences=male_absences,
        categories=categories
    )

    # Sheet 3: Женский состав
    ws_female = wb.create_sheet(title="Женский состав")
    female_cadets = [c for c in cadets if normalize_gender(c.get('gender')) == 'female']
    female_absences = [a for a in absences if normalize_gender(a.get('gender')) == 'female']
    build_sheet_content(
        ws_female,
        sheet_title_suffix="(Женский состав)",
        report_date=report_date,
        unit_title=unit_title,
        chief_title=chief_title,
        chief_name=chief_name,
        cadets=female_cadets,
        absences=female_absences,
        categories=categories
    )

    wb.save(output_path)
