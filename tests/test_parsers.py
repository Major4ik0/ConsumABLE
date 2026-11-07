import pytest
from core.parsers import (
    detect_gender,
    detect_gender_with_flag,
    extract_course,
    increment_course_all,
    bulk_set_course,
    remove_graduates,
    toggle_gender,
    make_short_fio,
    parse_group_code,
    change_group_course
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
        assert make_short_fio("Петров Владимир") == "Петров В."
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
        # 4-digit academic mask: 2nd digit is course
        assert extract_course("1280") == 2
        assert extract_course("1390") == 3
        assert extract_course("1412") == 4

    def test_parse_group_code(self):
        # 4-digit mask: faculty, course, subgroup
        assert parse_group_code("1280") == (1, 2, "80")
        assert parse_group_code("1212") == (1, 2, "12")
        assert parse_group_code("2390") == (2, 3, "90")
        # 3-digit fallback
        assert parse_group_code("221") == (None, 2, "21")
        # None / invalid
        assert parse_group_code(None) == (None, None, None)

    def test_change_group_course(self):
        # Substitutes 2nd digit of 4-digit mask
        assert change_group_course("1280", 3) == "1380"
        assert change_group_course("1212", 4) == "1412"
        assert change_group_course("1490", 1) == "1190"
        assert change_group_course("2380", "5") == "2580"
        # 3-digit fallback
        assert change_group_course("221", 3) == "321"
        assert change_group_course("", 2) == ""

    def test_get_course_from_group(self):
        from core.parsers import get_course_from_group
        assert get_course_from_group("1280") == "2"
        assert get_course_from_group("1212") == "2"
        assert get_course_from_group("1390") == "3"
        assert get_course_from_group("2180") == "1"
        assert get_course_from_group("221") == "?"
        assert get_course_from_group("") == "?"
        assert get_course_from_group(None) == "?"

    def test_update_header_text(self):
        from core.parsers import update_header_text

        # 1. From star placeholders
        unit = "* курс * факультета (автоматизированных систем управления)"
        chief = "Начальник * курса * факультета"
        assert update_header_text(unit, faculty=1, course=2) == "2 курс 1 факультета (автоматизированных систем управления)"
        assert update_header_text(chief, faculty=1, course=2) == "Начальник 2 курса 1 факультета"

        # 2. Incrementing (+1 курс)
        u_updated = update_header_text(unit, faculty=1, course=2)
        c_updated = update_header_text(chief, faculty=1, course=2)
        assert update_header_text(u_updated, course=3) == "3 курс 1 факультета (автоматизированных систем управления)"
        assert update_header_text(c_updated, course=3) == "Начальник 3 курса 1 факультета"

        # 3. Re-importing with another faculty/course
        assert update_header_text(u_updated, faculty=2, course=4) == "4 курс 2 факультета (автоматизированных систем управления)"
        assert update_header_text(c_updated, faculty=2, course=4) == "Начальник 4 курса 2 факультета"

    def test_parse_signature_from_excel(self, tmp_path):
        import openpyxl
        from core.parsers import parse_signature_from_excel

        test_file = str(tmp_path / "test_sig.xlsx")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Расход"
        ws["B10"] = "Начальник 3 курса 1 факультета"
        ws["B11"] = "майор                          И.Иванов"
        wb.save(test_file)

        title, rank, name = parse_signature_from_excel(test_file)
        assert title == "Начальник 3 курса 1 факультета"
        assert rank == "майор"
        assert name == "И.Иванов"
