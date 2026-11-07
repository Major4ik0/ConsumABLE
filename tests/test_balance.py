import pytest
from pydantic import ValidationError
from core.models import RosterSliceBalance, TriSliceBalance
from core.balance import calculate_slice_metrics, compute_tri_slice_balance, compute_gender_distribution


class TestSliceBalanceMath:
    def test_valid_slice_balance(self):
        # 10 listed = 7 present + (2 наряд + 1 лазарет)
        slice_b = RosterSliceBalance(
            slice_name="total",
            listed_count=10,
            present_count=7,
            absent_count=3,
            categories_breakdown={"Наряд": 2, "Лазарет": 1}
        )
        assert slice_b.is_balanced is True
        assert slice_b.discrepancy == 0
        assert slice_b.error_message is None

    def test_invalid_slice_balance_detected(self):
        # 10 listed != 6 present + 3 absences (diff = 1)
        slice_b = RosterSliceBalance(
            slice_name="male",
            listed_count=10,
            present_count=6,
            absent_count=3,
            categories_breakdown={"Наряд": 3}
        )
        assert slice_b.is_balanced is False
        assert slice_b.discrepancy == 1
        assert "Сбой баланса" in slice_b.error_message

    def test_categories_sum_mismatch(self):
        # absent_count = 4, but category breakdown adds up to 3
        slice_b = RosterSliceBalance(
            slice_name="female",
            listed_count=10,
            present_count=6,
            absent_count=4,
            categories_breakdown={"Наряд": 2, "Госпиталь": 1}
        )
        assert slice_b.is_balanced is False
        assert "Несовпадение суммы категорий" in slice_b.error_message


class TestTriSliceBalanceIdentity:
    @pytest.fixture
    def sample_data(self):
        cadets = [
            {"id": 1, "fio": "Иванов И.И.", "gender": "male", "group_num": "101"},
            {"id": 2, "fio": "Петров П.П.", "gender": "male", "group_num": "101"},
            {"id": 3, "fio": "Сидоров С.С.", "gender": "male", "group_num": "101"},
            {"id": 4, "fio": "Смирнова А.А.", "gender": "female", "group_num": "101"},
            {"id": 5, "fio": "Ковалева Е.Е.", "gender": "female", "group_num": "101"},
        ]
        absences = [
            {"cadet_id": 1, "gender": "male", "category": "Наряд"},
            {"cadet_id": 4, "gender": "female", "category": "Госпиталь"},
        ]
        categories = ["Наряд", "Госпиталь", "Лазарет"]
        return cadets, absences, categories

    def test_tri_slice_balance_successful_equality(self, sample_data):
        cadets, absences, categories = sample_data
        tri_balance = compute_tri_slice_balance(cadets, absences, categories)

        assert tri_balance.is_sum_balanced is True
        assert tri_balance.error_message is None

        # Total = Male + Female checks
        assert tri_balance.total.listed_count == tri_balance.male.listed_count + tri_balance.female.listed_count
        assert tri_balance.total.present_count == tri_balance.male.present_count + tri_balance.female.present_count
        assert tri_balance.total.absent_count == tri_balance.male.absent_count + tri_balance.female.absent_count

        # Category checks
        assert tri_balance.total.categories_breakdown["Наряд"] == 1
        assert tri_balance.male.categories_breakdown["Наряд"] == 1
        assert tri_balance.female.categories_breakdown["Наряд"] == 0

        assert tri_balance.total.categories_breakdown["Госпиталь"] == 1
        assert tri_balance.male.categories_breakdown["Госпиталь"] == 0
        assert tri_balance.female.categories_breakdown["Госпиталь"] == 1

    def test_tri_slice_detects_deliberate_inconsistency(self):
        # Construct intentionally broken balances
        total = RosterSliceBalance(
            slice_name="total", listed_count=10, present_count=8, absent_count=2,
            categories_breakdown={"Наряд": 2}
        )
        male = RosterSliceBalance(
            slice_name="male", listed_count=6, present_count=5, absent_count=1,
            categories_breakdown={"Наряд": 1}
        )
        # Female has 3 listed, so 6 + 3 = 9 != 10
        female = RosterSliceBalance(
            slice_name="female", listed_count=3, present_count=2, absent_count=1,
            categories_breakdown={"Наряд": 1}
        )

        tri = TriSliceBalance(total=total, male=male, female=female)
        assert tri.is_sum_balanced is False
        assert "Общий (10) != М (6) + Ж (3)" in tri.error_message

    def test_compute_gender_distribution(self, sample_data):
        cadets, absences, categories = sample_data
        dist = compute_gender_distribution(cadets, absences, categories=categories)
        assert dist["total"]["listed"] == 5
        assert dist["male"]["listed"] == 3
        assert dist["female"]["listed"] == 2
        assert dist["total"]["absent"] == 2
        assert dist["male"]["absent"] == 1
        assert dist["female"]["absent"] == 1
        assert dist["is_balanced"] is True
