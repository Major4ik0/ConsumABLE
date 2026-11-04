from typing import Literal, Optional, List, Dict, Any, Union
from pydantic import BaseModel, Field, model_validator

GenderType = Literal["male", "female"]
CadetStatusType = Literal["active", "graduated", "expelled"]

def normalize_gender(val: Any) -> GenderType:
    """Normalizes various gender representations ('М', 'Ж', 'male', 'female', etc.) to 'male' or 'female'."""
    if not val:
        return "male"
    s = str(val).strip().lower()
    if s in ("female", "ж", "жен", "женский", "f"):
        return "female"
    return "male"

def gender_to_ru(gender: Union[GenderType, str]) -> str:
    """Converts normalized gender to Russian single letter ('М' or 'Ж')."""
    return "Ж" if normalize_gender(gender) == "female" else "М"


class CadetBase(BaseModel):
    rank: str = "курсант"
    fio: str
    short_fio: Optional[str] = None
    group_num: str
    course: int = Field(default=1, ge=1, le=6)
    gender: GenderType = "male"
    status: CadetStatusType = "active"
    sort_order: int = 0

    @model_validator(mode="before")
    @classmethod
    def preprocess_data(cls, data: Any) -> Any:
        if isinstance(data, dict):
            if "gender" in data:
                data["gender"] = normalize_gender(data["gender"])
            if "course" in data and data["course"] is not None:
                try:
                    data["course"] = int(data["course"])
                except (ValueError, TypeError):
                    data["course"] = 1
        return data


class CadetRecord(CadetBase):
    id: Optional[int] = None


class StagingCadetDTO(BaseModel):
    temp_id: str
    id: Optional[int] = None
    rank: str = "курсант"
    fio: str
    short_fio: str
    group_num: str
    course: int = 1
    gender: GenderType = "male"
    status: CadetStatusType = "active"
    needs_manual_review: bool = False
    warnings: List[str] = Field(default_factory=list)
    sort_order: int = 0

    @model_validator(mode="before")
    @classmethod
    def preprocess_staging(cls, data: Any) -> Any:
        if isinstance(data, dict):
            if "gender" in data:
                data["gender"] = normalize_gender(data["gender"])
            if "course" in data and data["course"] is not None:
                try:
                    data["course"] = int(data["course"])
                except (ValueError, TypeError):
                    data["course"] = 1
        return data


class RosterSliceBalance(BaseModel):
    """Represents a single attendance slice (Total, Male, or Female).
    Validates: по_списку == налицо + сумма(все_категории_отсутствующих)
    """
    slice_name: Literal["total", "male", "female"]
    listed_count: int = Field(description="По списку (П/С)")
    present_count: int = Field(description="Налицо / В строю (Н/Л)")
    absent_count: int = Field(description="В расходе (Всего отсутствующих)")
    categories_breakdown: Dict[str, int] = Field(default_factory=dict)
    is_balanced: bool = True
    discrepancy: int = 0
    error_message: Optional[str] = None

    @model_validator(mode="after")
    def validate_slice_balance(self) -> "RosterSliceBalance":
        categories_sum = sum(self.categories_breakdown.values())
        if self.absent_count != categories_sum:
            self.is_balanced = False
            self.discrepancy = self.absent_count - categories_sum
            self.error_message = (
                f"Несовпадение суммы категорий в срезе '{self.slice_name}': "
                f"Всего в расходе ({self.absent_count}) != сумма категорий ({categories_sum})"
            )
            return self

        # Rule: по_списку == налицо + сумма(все_категории_отсутствующих)
        expected_listed = self.present_count + categories_sum
        diff = self.listed_count - expected_listed

        if diff != 0:
            self.is_balanced = False
            self.discrepancy = diff
            self.error_message = (
                f"Сбой баланса в срезе '{self.slice_name}': "
                f"По списку ({self.listed_count}) != Налицо ({self.present_count}) + "
                f"Отсутствующие ({categories_sum}). Нестыковка: {diff}"
            )
        else:
            self.is_balanced = True
            self.discrepancy = 0
            self.error_message = None

        return self


class TriSliceBalance(BaseModel):
    """Tri-slice attendance balance: Total (Общий), Male (Мужской состав), Female (Женский состав).
    Verifies that Общий == Мужчины + Женщины across all metrics.
    """
    total: RosterSliceBalance
    male: RosterSliceBalance
    female: RosterSliceBalance
    is_sum_balanced: bool = True
    error_message: Optional[str] = None

    @model_validator(mode="after")
    def validate_tri_slices_consistency(self) -> "TriSliceBalance":
        m = self.male
        f = self.female
        t = self.total

        errors = []
        if t.listed_count != (m.listed_count + f.listed_count):
            errors.append(
                f"По списку: Общий ({t.listed_count}) != М ({m.listed_count}) + Ж ({f.listed_count})"
            )
        if t.present_count != (m.present_count + f.present_count):
            errors.append(
                f"Налицо: Общий ({t.present_count}) != М ({m.present_count}) + Ж ({f.present_count})"
            )
        if t.absent_count != (m.absent_count + f.absent_count):
            errors.append(
                f"В расходе: Общий ({t.absent_count}) != М ({m.absent_count}) + Ж ({f.absent_count})"
            )

        all_categories = set(t.categories_breakdown.keys()) | set(m.categories_breakdown.keys()) | set(f.categories_breakdown.keys())
        for cat in all_categories:
            cat_t = t.categories_breakdown.get(cat, 0)
            cat_m = m.categories_breakdown.get(cat, 0)
            cat_f = f.categories_breakdown.get(cat, 0)
            if cat_t != (cat_m + cat_f):
                errors.append(
                    f"Категория '{cat}': Общий ({cat_t}) != М ({cat_m}) + Ж ({cat_f})"
                )

        if not t.is_balanced:
            errors.append(f"Общий срез: {t.error_message}")
        if not m.is_balanced:
            errors.append(f"Мужской срез: {m.error_message}")
        if not f.is_balanced:
            errors.append(f"Женский срез: {f.error_message}")

        if errors:
            self.is_sum_balanced = False
            self.error_message = "; ".join(errors)
        else:
            self.is_sum_balanced = True
            self.error_message = None

        return self
