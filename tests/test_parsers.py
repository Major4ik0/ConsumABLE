import pytest
from core.parsers import (
    detect_gender,
    detect_gender_with_flag,
    extract_course,
    increment_course_all,
    bulk_set_course,
    remove_graduates,
    toggle_gender,
    make_short_fio
)
from core.models import StagingCadetDTO


class TestGenderDetection:
    def test_male_patronymics(self):
        # Окончания "ич", "ыч", "оглы"
        assert detect_gender("Иванович") == "male"
        assert detect_gender("Петрович") == "male"
        assert detect_gender("Павлыч") == "male"
        assert detect_gender("Ильич") == "male"
        assert detect_gender("Мамед оглы") == "male"
        assert detect_gender("Али уулу") == "male"

    def test_female_patronymics(self):
        # Окончания "на", "вна", "чна", "кызы"
        assert detect_gender("Ивановна") == "female"
        assert detect_gender("Петровна") == "female"
        assert detect_gender("Ильинична") == "female"
        assert detect_gender("Никитична") == "female"
        assert detect_gender("Айгуль кызы") == "female"
        assert detect_gender("Айнур гызы") == "female"

    def test_full_fio_detection(self):
        # Full FIO strings
        assert detect_gender("Смирнов Алексей Владимирович") == "male"
        assert detect_gender("Ковалева Екатерина Сергеевна") == "female"
        assert detect_gender("Кузнецова Анна Кузьминична") == "female"

    def test_unknown_or_missing_patronymic_defaults_to_male_with_review_flag(self):
        # If patronymic is missing or unrecognized -> male with needs_manual_review=True
        gender, needs_review = detect_gender_with_flag("Джон Смит")
        assert gender == "male"
        assert needs_review is True

        gender, needs_review = detect_gender_with_flag("")
        assert gender == "male"
        assert needs_review is True

    def test_explicit_gender_override(self):
        # Explicit column takes priority
        gender, needs_review = detect_gender_with_flag("Иванов Иван", explicit_val="Жен")
        assert gender == "female"
        assert needs_review is False

        gender, needs_review = detect_gender_with_flag("Петрова Анна", explicit_val="М")
        assert gender == "male"
        assert needs_review is False


class TestShortFio:
    def test_standard_fio(self):
        assert make_short_fio("Иванов Иван Иванович") == "Иванов И.И."
        assert make_short_fio("Медков Владимир") == "Медков В."
        assert make_short_fio("Сидоров А. Б.") == "Сидоров А.Б."


class TestCourseTransformations:
    @pytest.fixture
    def sample_cadets(self):
        return [
            StagingCadetDTO(temp_id="1", rank="ряд.", fio="Иванов И.И.", short_fio="Иванов И.И.", group_num="121", course=1, gender="male", status="active"),
            StagingCadetDTO(temp_id="2", rank="ряд.", fio="Петров П.П.", short_fio="Петров П.П.", group_num="221", course=2, gender="male", status="active"),
            StagingCadetDTO(temp_id="3", rank="серж.", fio="Сидорова С.С.", short_fio="Сидорова С.С.", group_num="521", course=5, gender="female", status="active"),
        ]

    def test_increment_course_all(self, sample_cadets):
        updated = increment_course_all(sample_cadets, max_course=5)
        # Course 1 -> 2 (active)
        assert updated[0].course == 2
        assert updated[0].status == "active"
        # Course 2 -> 3 (active)
        assert updated[1].course == 3
        assert updated[1].status == "active"
        # Course 5 -> 6 (graduated)
        assert updated[2].course == 6
        assert updated[2].status == "graduated"
        assert any("выпускник" in w.lower() for w in updated[2].warnings)

    def test_bulk_set_course(self, sample_cadets):
        # Set course for all to 3
        updated = bulk_set_course(sample_cadets, target_course=3)
        for c in updated:
            assert c.course == 3
            assert c.status == "active"

        # Set course only for selected
        selective = bulk_set_course(sample_cadets, target_course=4, selected_temp_ids=["1"])
        assert selective[0].course == 4
        assert selective[1].course == 2
        assert selective[2].course == 5

    def test_remove_graduates(self, sample_cadets):
        incremented = increment_course_all(sample_cadets, max_course=5)
        assert len(incremented) == 3
        active_only = remove_graduates(incremented, max_course=5)
        assert len(active_only) == 2
        assert all(c.status != "graduated" for c in active_only)

    def test_toggle_gender(self, sample_cadets):
        c = sample_cadets[0]
        assert c.gender == "male"
        toggled = toggle_gender(c)
        assert toggled.gender == "female"
        assert toggled.needs_manual_review is False
        toggled_back = toggle_gender(toggled)
        assert toggled_back.gender == "male"

    def test_extract_course(self):
        assert extract_course("121") == 1
        assert extract_course("221") == 2
        assert extract_course("3 курс") == 3
        assert extract_course("520", explicit_course=4) == 4
