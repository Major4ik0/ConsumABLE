import os
from typing import List, Dict, Any, Optional

from core.parsers import (
    make_short_fio,
    detect_gender as core_detect_gender,
    detect_gender_with_flag,
    parse_roster_excel,
    parse_cadets_sheet,
    parse_group_code,
    change_group_course,
    get_course_from_group
)
from core.models import gender_to_ru, normalize_gender
from core.exporter import export_report_to_excel, style_range
from core.balance import calculate_slice_metrics, compute_tri_slice_balance


def detect_gender(fio: str, explicit_val=None) -> str:
    """Wrapper that returns Russian 'М' or 'Ж' for backwards compatibility."""
    g, _ = detect_gender_with_flag(fio, explicit_val)
    return gender_to_ru(g)


def import_cadets_from_excel(file_path: str) -> List[Dict[str, Any]]:
    """Legacy compatibility loader: parses file via core parser and returns list of dicts."""
    staging = parse_roster_excel(file_path)
    result = []
    for c in staging:
        result.append({
            'rank': c.rank,
            'fio': c.fio,
            'short_fio': c.short_fio,
            'group_num': c.group_num,
            'course': c.course,
            'gender': gender_to_ru(c.gender),
            'status': c.status,
            'sort_order': c.sort_order,
            'needs_manual_review': c.needs_manual_review,
            'warnings': c.warnings
        })
    return result
