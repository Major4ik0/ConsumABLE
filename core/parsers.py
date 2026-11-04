import io
import os
import re
import uuid
from typing import List, Tuple, Optional, Union, BinaryIO, Any
import openpyxl

from core.models import StagingCadetDTO, normalize_gender, gender_to_ru

# Endings specification
MALE_ENDINGS = ("ич", "ыч", "оглы", "уулу")
FEMALE_ENDINGS = ("на", "вна", "чна", "кызы", "гызы")


def make_short_fio(full_name: str) -> str:
    """Converts full name 'Иванов Иван Иванович' to short form 'Иванов И.И.'."""
    name = full_name.strip()
    m = re.match(r'^([А-ЯЁа-яё\-]+)\s+([А-ЯЁа-яё])\.\s*([А-ЯЁа-яё])\.$', name)
    if m:
        return f"{m.group(1)} {m.group(2).upper()}.{m.group(3).upper()}."
    parts = name.split()
    if len(parts) >= 3:
        i1 = parts[1][0].upper()
        i2 = parts[2][0].upper()
        return f"{parts[0]} {i1}.{i2}."
    elif len(parts) == 2:
        if '.' in parts[1]:
            initials = ''.join([c for c in parts[1] if c.isalpha()]).upper()
            if len(initials) >= 2:
                return f"{parts[0]} {initials[0]}.{initials[1]}."
            elif len(initials) == 1:
                return f"{parts[0]} {initials[0]}."
        return f"{parts[0]} {parts[1][0].upper()}."
    return name


def detect_gender(patronymic: str) -> str:
    """Core gender detection by patronymic according to specification:
    - Endings ('ич', 'ыч', 'оглы') -> 'male'
    - Endings ('на', 'вна', 'чна', 'кызы') -> 'female'
    - If absent or unrecognized -> 'male'
    """
    gender, _ = detect_gender_with_flag(patronymic)
    return gender


def detect_gender_with_flag(patronymic_or_fio: str, explicit_val: Optional[str] = None) -> Tuple[str, bool]:
    """Detects gender and returns (gender, needs_manual_review).
    If explicit_val is provided, uses it directly.
    Otherwise checks patronymic endings. If unrecognized, defaults to 'male' with needs_manual_review=True.
    """
    if explicit_val is not None:
        v = str(explicit_val).strip().lower()
        if v in ("ж", "f", "жен", "женский", "female"):
            return "female", False
        if v in ("м", "m", "муж", "мужской", "male"):
            return "male", False

    text = str(patronymic_or_fio or "").strip()
    if not text:
        return "male", True

    parts = text.split()
    # If full name is passed (e.g. "Иванов Иван Иванович"), patronymic is parts[2]
    if len(parts) >= 3:
        patronymic = parts[2].lower()
    else:
        patronymic = parts[-1].lower()

    # Check female endings
    for ending in FEMALE_ENDINGS:
        if patronymic.endswith(ending):
            return "female", False

    # Check male endings
    for ending in MALE_ENDINGS:
        if patronymic.endswith(ending):
            return "male", False

    # Default for military structures: 'male' with manual review flag
    return "male", True


def extract_course(group_val: Any, explicit_course: Any = None) -> int:
    """Attempts to determine course (1..5) from explicit column or group number (e.g. '221' -> 2)."""
    if explicit_course is not None:
        try:
            c = int(str(explicit_course).strip())
            if 1 <= c <= 6:
                return c
        except (ValueError, TypeError):
            pass

    grp_str = str(group_val or "").strip()
    # Common academic pattern: 3-4 digit group where 1st digit or 2nd is course: e.g. 1280 (1st or 2nd)?
    # Look for patterns like "1 курс", "2 курс" or first digit in group number:
    m_course_word = re.search(r'([1-5])\s*курс', grp_str, re.IGNORECASE)
    if m_course_word:
        return int(m_course_word.group(1))

    digits = "".join([ch for ch in grp_str if ch.isdigit()])
    if len(digits) >= 1:
        first_digit = int(digits[0])
        if 1 <= first_digit <= 5:
            return first_digit

    return 1


def parse_cadets_sheet(ws, default_group: str = "", seen_groups: Optional[List[str]] = None) -> List[StagingCadetDTO]:
    """Parses an openpyxl worksheet into StagingCadetDTO objects without database side-effects."""
    if seen_groups is None:
        seen_groups = []

    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []

    header_idx = -1
    col_map = {}

    for r_i in range(min(15, len(rows))):
        row = [str(c or "").strip().lower() for c in rows[r_i]]
        has_fio = any(k in c for c in row for k in ['фио', 'ф.и.о.', 'фамили', 'курсант', 'слушател', 'обучающ'])
        has_rank = any(k in c for c in row for k in ['звание', 'в/зв', 'в/звание', 'воинск'])
        has_grp = any(k in c for c in row for k in ['групп', 'взвод', 'подразд', 'отделен', 'к/о', 'класс', 'курс'])

        if (has_fio and (has_rank or has_grp)) or (has_fio and any('№' in c or 'п/п' in c for c in row)):
            header_idx = r_i
            for c_i, cell_text in enumerate(row):
                if any(k in cell_text for k in ['в/зв', 'звание', 'воинск']):
                    col_map['rank'] = c_i
                elif cell_text in ['фио', 'ф.и.о.', 'фамилия имя отчество', 'фамилия, имя, отчество', 'фамилия имя']:
                    col_map['fio'] = c_i
                elif 'сокр' in cell_text and 'фио' in cell_text:
                    col_map['short_fio'] = c_i
                elif any(k in cell_text for k in ['фио', 'ф.и.о.', 'фамилия']) and 'fio' not in col_map:
                    col_map['fio'] = c_i
                elif 'фамил' in cell_text and 'fio' not in col_map:
                    col_map['surname'] = c_i
                elif cell_text == 'имя' or cell_text.startswith('имя '):
                    col_map['name'] = c_i
                elif 'отчеств' in cell_text:
                    col_map['patronymic'] = c_i
                elif 'курс' in cell_text:
                    col_map['course'] = c_i
                elif any(k in cell_text for k in ['групп', 'взвод', 'подразд', 'отделен', 'к/о', 'класс']):
                    col_map['group'] = c_i
                elif any(k in cell_text for k in ['пол', 'м/ж']):
                    col_map['gender'] = c_i
            break

    start_row = header_idx + 1 if header_idx != -1 else 0

    if 'fio' not in col_map and 'surname' not in col_map:
        col_scores = {}
        for r in rows[start_row:start_row + 15]:
            for c_i, cell in enumerate(r):
                val = str(cell or "").strip()
                if not val:
                    continue
                words = val.split()
                if len(words) in [2, 3] and any(ord('а') <= ord(ch.lower()) <= ord('я') for ch in val):
                    col_scores[c_i] = col_scores.get(c_i, 0) + 1
        if col_scores:
            best_fio_col = max(col_scores, key=col_scores.get)
            col_map['fio'] = best_fio_col
            if best_fio_col > 0:
                col_map['rank'] = best_fio_col - 1
            if best_fio_col + 1 < len(rows[start_row]):
                col_map['group'] = best_fio_col + 1
        else:
            col_map = {'rank': 0, 'fio': 1, 'group': 2}

    staging_cadets: List[StagingCadetDTO] = []
    current_group = default_group or "1"

    for r_i in range(start_row, len(rows)):
        row = rows[r_i]
        if not row or all(c is None or str(c).strip() == "" for c in row):
            continue

        fio = ""
        if 'fio' in col_map and col_map['fio'] < len(row) and row[col_map['fio']] is not None:
            fio = str(row[col_map['fio']]).strip()
        elif 'surname' in col_map:
            sur = str(row[col_map['surname']] or "").strip() if col_map['surname'] < len(row) else ""
            nm = str(row[col_map.get('name', 999)] or "").strip() if col_map.get('name', 999) < len(row) else ""
            pat = str(row[col_map.get('patronymic', 999)] or "").strip() if col_map.get('patronymic', 999) < len(row) else ""
            fio = " ".join(filter(None, [sur, nm, pat])).strip()

        fio_lower = fio.lower()
        if not fio or fio_lower in ['фио', 'ф.и.о.', 'фамилия', 'итого', 'всего']:
            continue

        if len(fio.split()) == 1 and ('взвод' in fio_lower or 'группа' in fio_lower or 'отделение' in fio_lower):
            current_group = fio
            continue

        rank = ""
        if 'rank' in col_map and col_map['rank'] < len(row) and row[col_map['rank']] is not None:
            rank = str(row[col_map['rank']]).strip()
            if rank.isdigit():
                rank = "курсант"
        if not rank:
            rank = "курсант"

        group_raw = None
        if 'group' in col_map and col_map['group'] < len(row):
            group_raw = row[col_map['group']]

        if group_raw is not None and str(group_raw).strip():
            if isinstance(group_raw, (int, float)):
                group_num = str(int(group_raw))
            else:
                group_num = str(group_raw).strip()
        else:
            group_num = current_group

        if group_num not in seen_groups:
            seen_groups.append(group_num)
        sort_order = seen_groups.index(group_num)

        short_fio = ""
        if 'short_fio' in col_map and col_map['short_fio'] < len(row) and row[col_map['short_fio']]:
            short_fio = str(row[col_map['short_fio']]).strip()
        if not short_fio:
            short_fio = make_short_fio(fio)

        # Detect gender with review flag
        raw_gender = None
        if 'gender' in col_map and col_map['gender'] < len(row):
            raw_gender = row[col_map['gender']]

        gender, needs_review = detect_gender_with_flag(fio, raw_gender)

        # Course detection
        explicit_course = None
        if 'course' in col_map and col_map['course'] < len(row):
            explicit_course = row[col_map['course']]
        course = extract_course(group_num, explicit_course)

        warnings = []
        status = "active"
        if course > 5:
            status = "graduated"
            warnings.append(f"Курс {course} превышает 5 (выпускник)")
        if needs_review:
            warnings.append("Пол определен эвристически (проверьте)")

        staging_cadets.append(StagingCadetDTO(
            temp_id=str(uuid.uuid4())[:8],
            rank=rank,
            fio=fio,
            short_fio=short_fio,
            group_num=group_num,
            course=course,
            gender=gender,
            status=status,
            needs_manual_review=needs_review,
            warnings=warnings,
            sort_order=sort_order
        ))

    return staging_cadets


def parse_roster_excel(file_source: Union[str, bytes, BinaryIO]) -> List[StagingCadetDTO]:
    """Main Excel parser entry point. Reads .xlsx from path or in-memory bytes without modifying the DB."""
    if isinstance(file_source, str):
        if not os.path.exists(file_source):
            raise FileNotFoundError(f"Файл не найден: {file_source}")
        ext = os.path.splitext(file_source)[1].lower()
        if ext == ".xls":
            raise ValueError(
                "Файл имеет устаревший формат .xls (Excel 97-2003).\n"
                "Пожалуйста, откройте его в Excel и сохраните как «Книга Excel (.xlsx)»."
            )
        wb = openpyxl.load_workbook(file_source, data_only=True)
    elif isinstance(file_source, bytes):
        wb = openpyxl.load_workbook(io.BytesIO(file_source), data_only=True)
    else:
        wb = openpyxl.load_workbook(file_source, data_only=True)

    priority_names = ['база', 'личный состав', 'курсанты', 'список', 'список личного состава', 'ведомость']
    target_sheets = []

    for sheet_name in wb.sheetnames:
        clean_name = sheet_name.strip().lower()
        if any(p in clean_name for p in priority_names):
            target_sheets.append(wb[sheet_name])
            break

    if not target_sheets:
        for sname in wb.sheetnames:
            ws = wb[sname]
            if ws.max_row and ws.max_row > 1:
                target_sheets.append(ws)

    if not target_sheets:
        target_sheets = [wb.active]

    all_cadets: List[StagingCadetDTO] = []
    seen_groups: List[str] = []

    for ws in target_sheets:
        default_grp = ws.title.strip() if len(target_sheets) > 1 else ""
        cadets = parse_cadets_sheet(ws, default_group=default_grp, seen_groups=seen_groups)
        all_cadets.extend(cadets)

    return all_cadets


# =====================================================================
# STAGING WIZARD TRANSFORMATION FUNCTIONS
# =====================================================================

def increment_course_all(cadets: List[StagingCadetDTO], max_course: int = 5) -> List[StagingCadetDTO]:
    """Increments course by 1 for all active records.
    Records that reach > max_course are transitioned to 'graduated' with a warning badge.
    """
    updated: List[StagingCadetDTO] = []
    for c in cadets:
        cadet = c.model_copy(deep=True)
        if cadet.status == "active":
            new_course = cadet.course + 1
            cadet.course = new_course
            if new_course > max_course:
                cadet.status = "graduated"
                warn = f"Выпускник (курс {new_course} > {max_course})"
                if warn not in cadet.warnings:
                    cadet.warnings.append(warn)
        updated.append(cadet)
    return updated


def bulk_set_course(
    cadets: List[StagingCadetDTO],
    target_course: int,
    selected_temp_ids: Optional[List[str]] = None,
    max_course: int = 5
) -> List[StagingCadetDTO]:
    """Sets target course (1..5) for selected temp_ids or for all if selected_temp_ids is None."""
    target_ids = set(selected_temp_ids) if selected_temp_ids else None
    updated: List[StagingCadetDTO] = []

    for c in cadets:
        cadet = c.model_copy(deep=True)
        if target_ids is None or cadet.temp_id in target_ids:
            cadet.course = target_course
            if target_course > max_course:
                cadet.status = "graduated"
                warn = f"Курс {target_course} превышает {max_course} (выпускник)"
                if warn not in cadet.warnings:
                    cadet.warnings.append(warn)
            else:
                cadet.status = "active"
                # remove graduated warning if course was lowered
                cadet.warnings = [w for w in cadet.warnings if "выпускник" not in w.lower()]
        updated.append(cadet)
    return updated


def remove_graduates(cadets: List[StagingCadetDTO], max_course: int = 5) -> List[StagingCadetDTO]:
    """Filters out any cadets with status == 'graduated' or course > max_course."""
    return [c for c in cadets if c.status != "graduated" and c.course <= max_course]


def toggle_gender(cadet: StagingCadetDTO) -> StagingCadetDTO:
    """Toggles gender between 'male' and 'female' and clears manual review flag."""
    c = cadet.model_copy(deep=True)
    c.gender = "female" if c.gender == "male" else "male"
    c.needs_manual_review = False
    c.warnings = [w for w in c.warnings if "пол определен" not in w.lower()]
    return c
