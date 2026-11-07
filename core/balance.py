from typing import List, Dict, Any, Literal, Optional
from core.models import RosterSliceBalance, TriSliceBalance, normalize_gender


def calculate_slice_metrics(
    cadets: List[Dict[str, Any]],
    absences: List[Dict[str, Any]],
    slice_name: Literal["total", "male", "female"],
    categories: Optional[List[str]] = None
) -> RosterSliceBalance:
    """Calculates balance metrics for a given slice ('total', 'male', 'female').
    Strict mathematical condition: по_списку == налицо + сумма(все_категории_отсутствующих).
    """
    if slice_name == "male":
        filtered_cadets = [c for c in cadets if normalize_gender(c.get("gender", "male")) == "male"]
        filtered_absences = [a for a in absences if normalize_gender(a.get("gender", "male")) == "male"]
    elif slice_name == "female":
        filtered_cadets = [c for c in cadets if normalize_gender(c.get("gender", "male")) == "female"]
        filtered_absences = [a for a in absences if normalize_gender(a.get("gender", "male")) == "female"]
    else:
        filtered_cadets = list(cadets)
        filtered_absences = list(absences)

    listed = len(filtered_cadets)
    absent = len(filtered_absences)
    present = listed - absent

    # Categorize absences
    cat_counts: Dict[str, int] = {}
    if categories:
        for cat in categories:
            cat_counts[cat] = 0

    for a in filtered_absences:
        cat = str(a.get("category") or "Другое").strip()
        cat_counts[cat] = cat_counts.get(cat, 0) + 1

    return RosterSliceBalance(
        slice_name=slice_name,
        listed_count=listed,
        present_count=present,
        absent_count=absent,
        categories_breakdown=cat_counts
    )


def compute_tri_slice_balance(
    cadets: List[Dict[str, Any]],
    absences: List[Dict[str, Any]],
    categories: Optional[List[str]] = None
) -> TriSliceBalance:
    """Computes all 3 slices (Total, Male, Female) and verifies mathematical identity:
    Общий == Мужчины + Женщины for listed, present, absent, and each absence category.
    """
    total_slice = calculate_slice_metrics(cadets, absences, "total", categories=categories)
    male_slice = calculate_slice_metrics(cadets, absences, "male", categories=categories)
    female_slice = calculate_slice_metrics(cadets, absences, "female", categories=categories)

    return TriSliceBalance(
        total=total_slice,
        male=male_slice,
        female=female_slice
    )


def compute_gender_distribution(
    cadets: List[Dict[str, Any]],
    absences: List[Dict[str, Any]],
    categories: Optional[List[str]] = None
) -> Dict[str, Dict[str, Any]]:
    """Computes comprehensive gender distribution stats for:
    - Общее количество людей и общий расход (total)
    - Количество мужчин и расход на мужчин (male)
    - Количество женщин и расход на женщин (female)
    """
    bal = compute_tri_slice_balance(cadets, absences, categories=categories)
    return {
        "total": {
            "listed": bal.total.listed_count,
            "present": bal.total.present_count,
            "absent": bal.total.absent_count,
            "categories": bal.total.categories_breakdown
        },
        "male": {
            "listed": bal.male.listed_count,
            "present": bal.male.present_count,
            "absent": bal.male.absent_count,
            "categories": bal.male.categories_breakdown
        },
        "female": {
            "listed": bal.female.listed_count,
            "present": bal.female.present_count,
            "absent": bal.female.absent_count,
            "categories": bal.female.categories_breakdown
        },
        "is_balanced": bal.is_sum_balanced,
        "error_message": bal.error_message
    }
