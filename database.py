import sqlite3
from typing import List, Dict, Any, Optional
from core.models import normalize_gender, gender_to_ru

class Database:
    def __init__(self, db_path="app_data.db"):
        self.conn = sqlite3.connect(db_path, check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self.create_tables()

    def create_tables(self):
        with self.conn:
            # Настройки курса и подписей
            self.conn.execute("""
                CREATE TABLE IF NOT EXISTS app_settings (
                    key TEXT PRIMARY KEY,
                    value TEXT
                )
            """)
            # Личный состав
            self.conn.execute("""
                CREATE TABLE IF NOT EXISTS cadets (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    rank TEXT,
                    fio TEXT NOT NULL,
                    short_fio TEXT,
                    group_num TEXT NOT NULL,
                    course INTEGER DEFAULT 1,
                    gender TEXT DEFAULT 'М',
                    status TEXT DEFAULT 'active',
                    sort_order INTEGER DEFAULT 0,
                    is_active INTEGER DEFAULT 1
                )
            """)
            # Справочник причин и нарядов
            self.conn.execute("""
                CREATE TABLE IF NOT EXISTS absence_types (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL,
                    code TEXT,
                    category TEXT NOT NULL
                )
            """)
            # Отсутствующие на дату
            self.conn.execute("""
                CREATE TABLE IF NOT EXISTS daily_absences (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    report_date TEXT NOT NULL,
                    cadet_id INTEGER NOT NULL,
                    type_id INTEGER NOT NULL,
                    note TEXT,
                    FOREIGN KEY(cadet_id) REFERENCES cadets(id) ON DELETE CASCADE,
                    FOREIGN KEY(type_id) REFERENCES absence_types(id) ON DELETE RESTRICT,
                    UNIQUE(report_date, cadet_id)
                )
            """)

            # Migration checks for columns in existing databases
            cur = self.conn.cursor()
            cur.execute("PRAGMA table_info(cadets)")
            existing_cols = [r[1] for r in cur.fetchall()]
            if "course" not in existing_cols:
                self.conn.execute("ALTER TABLE cadets ADD COLUMN course INTEGER DEFAULT 1")
            if "status" not in existing_cols:
                self.conn.execute("ALTER TABLE cadets ADD COLUMN status TEXT DEFAULT 'active'")

            self._seed_defaults()
            self._ensure_default_types()

    def _seed_defaults(self):
        cur = self.conn.cursor()
        cur.execute("SELECT COUNT(*) FROM app_settings")
        if cur.fetchone()[0] == 0:
            defaults = [
                ("unit_title", "2 курс 1 факультета (автоматизированных систем управления)"),
                ("chief_title", "Начальник 2 курса 1 факультета"),
                ("chief_name", "майор                          В.Медков"),
            ]
            self.conn.executemany("INSERT INTO app_settings (key, value) VALUES (?, ?)", defaults)

        cur.execute("SELECT COUNT(*) FROM absence_types")
        if cur.fetchone()[0] == 0:
            types = [
                ("Дежурный по курсу", "Н", "Наряд"),
                ("Дневальный по курсу", "Н", "Наряд"),
                ("Патруль", "Н", "Наряд"),
                ("Увольнение", "У", "Увольнение"),
                ("Госпиталь", "Г", "Госпиталь"),
                ("Лазарет", "Л", "Лазарет"),
                ("Командировка", "К", "Командировка"),
            ]
            self.conn.executemany(
                "INSERT INTO absence_types (name, code, category) VALUES (?, ?, ?)",
                types
            )

    def _ensure_default_types(self):
        with self.conn:
            self.conn.execute("""
                UPDATE absence_types 
                SET name = 'Увольнение', code = 'УВ', category = 'Увольнение' 
                WHERE name IN ('Увальнение', 'увальнение', 'увольнение')
            """)
            cur = self.conn.cursor()
            cur.execute("SELECT COUNT(*) FROM absence_types WHERE name = 'Увольнение'")
            if cur.fetchone()[0] == 0:
                self.conn.execute("""
                    INSERT INTO absence_types (name, code, category) 
                    VALUES ('Увольнение', 'УВ', 'Увольнение')
                """)

    def get_setting(self, key, default=""):
        cur = self.conn.cursor()
        cur.execute("SELECT value FROM app_settings WHERE key = ?", (key,))
        row = cur.fetchone()
        return row[0] if row else default

    def set_setting(self, key, value):
        with self.conn:
            self.conn.execute("INSERT OR REPLACE INTO app_settings (key, value) VALUES (?, ?)", (key, value))

    def get_all_cadets(self, active_only: bool = True):
        cur = self.conn.cursor()
        if active_only:
            cur.execute("""
                SELECT id, rank, fio, short_fio, group_num, course, gender, status, sort_order 
                FROM cadets 
                WHERE is_active = 1 AND (status IS NULL OR status = 'active')
                ORDER BY sort_order ASC, group_num ASC, fio ASC
            """)
        else:
            cur.execute("""
                SELECT id, rank, fio, short_fio, group_num, course, gender, status, sort_order 
                FROM cadets 
                ORDER BY sort_order ASC, group_num ASC, fio ASC
            """)
        return [dict(r) for r in cur.fetchall()]

    def get_cadets_by_gender(self, gender_val: str, active_only: bool = True):
        all_cadets = self.get_all_cadets(active_only=active_only)
        target = normalize_gender(gender_val)
        return [c for c in all_cadets if normalize_gender(c.get('gender')) == target]

    def get_cadet_by_id(self, cadet_id):
        cur = self.conn.cursor()
        cur.execute("""
            SELECT id, rank, fio, short_fio, group_num, course, gender, status, sort_order 
            FROM cadets 
            WHERE id = ?
        """, (cadet_id,))
        row = cur.fetchone()
        return dict(row) if row else None

    def get_distinct_groups(self):
        cur = self.conn.cursor()
        cur.execute("""
            SELECT DISTINCT group_num 
            FROM cadets 
            WHERE is_active = 1 AND (status IS NULL OR status = 'active')
            ORDER BY sort_order ASC, group_num ASC
        """)
        return [r[0] for r in cur.fetchall()]

    def add_cadet(self, rank, fio, group_num, gender="М", short_fio=None, course=1, status="active"):
        if not short_fio:
            from core.parsers import make_short_fio
            short_fio = make_short_fio(fio)
        cur = self.conn.cursor()
        cur.execute("SELECT sort_order FROM cadets WHERE group_num = ? LIMIT 1", (group_num,))
        row = cur.fetchone()
        if row:
            sort_order = row[0]
        else:
            cur.execute("SELECT COALESCE(MAX(sort_order), 0) + 1 FROM cadets")
            sort_order = cur.fetchone()[0]

        ru_gender = gender_to_ru(gender)

        with self.conn:
            cur = self.conn.execute("""
                INSERT INTO cadets (rank, fio, short_fio, group_num, course, gender, status, sort_order, is_active)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, 1)
            """, (rank, fio, short_fio, group_num, course, ru_gender, status, sort_order))
            return cur.lastrowid

    def update_cadet(self, cadet_id, rank, fio, group_num, gender="М", short_fio=None, course=1, status="active"):
        if not short_fio:
            from core.parsers import make_short_fio
            short_fio = make_short_fio(fio)
        ru_gender = gender_to_ru(gender)
        with self.conn:
            self.conn.execute("""
                UPDATE cadets 
                SET rank = ?, fio = ?, short_fio = ?, group_num = ?, course = ?, gender = ?, status = ?
                WHERE id = ?
            """, (rank, fio, short_fio, group_num, course, ru_gender, status, cadet_id))

    def toggle_cadet_gender(self, cadet_id: int) -> str:
        cadet = self.get_cadet_by_id(cadet_id)
        if not cadet:
            return "М"
        current_g = cadet.get('gender', 'М')
        new_g = "Ж" if normalize_gender(current_g) == "male" else "М"
        with self.conn:
            self.conn.execute("UPDATE cadets SET gender = ? WHERE id = ?", (new_g, cadet_id))
        return new_g

    def update_cadet_course(self, cadet_id: int, new_course: int, update_group: bool = True):
        from core.parsers import change_group_course
        cadet = self.get_cadet_by_id(cadet_id)
        if not cadet:
            return
        new_group = cadet.get('group_num', '')
        if update_group:
            new_group = change_group_course(new_group, new_course)
        with self.conn:
            self.conn.execute(
                "UPDATE cadets SET course = ?, group_num = ? WHERE id = ?",
                (new_course, new_group, cadet_id)
            )

    def delete_cadet(self, cadet_id):
        with self.conn:
            self.conn.execute("DELETE FROM daily_absences WHERE cadet_id = ?", (cadet_id,))
            self.conn.execute("DELETE FROM cadets WHERE id = ?", (cadet_id,))

    def clear_and_import_cadets(self, cadets_list):
        from core.parsers import get_course_from_group, make_short_fio
        with self.conn:
            self.conn.execute("DELETE FROM daily_absences")
            self.conn.execute("DELETE FROM cadets")
            prepared = []
            for idx, c in enumerate(cadets_list):
                if isinstance(c, dict):
                    rank = c.get('rank', 'курсант')
                    fio = c['fio']
                    short_fio = c.get('short_fio') or make_short_fio(fio)
                    group_num = str(c['group_num'])
                    course = c.get('course')
                    if not course or course == 1:
                        cg = get_course_from_group(group_num)
                        course = int(cg) if cg != "?" else (course or 1)
                    gender = gender_to_ru(c.get('gender', 'М'))
                    status = c.get('status', 'active')
                    sort_order = c.get('sort_order', idx)
                    prepared.append((rank, fio, short_fio, group_num, course, gender, status, sort_order))
                elif hasattr(c, 'fio'):
                    # Pydantic or DTO object
                    grp_str = str(c.group_num)
                    crs = getattr(c, 'course', None)
                    if not crs or crs == 1:
                        cg = get_course_from_group(grp_str)
                        crs = int(cg) if cg != "?" else (crs or 1)
                    s_fio = getattr(c, 'short_fio', None) or make_short_fio(c.fio)
                    prepared.append((
                        c.rank, c.fio, s_fio, grp_str, crs,
                        gender_to_ru(c.gender), getattr(c, 'status', 'active'), getattr(c, 'sort_order', idx)
                    ))
                else:
                    sort_order = c[5] if len(c) > 5 else idx
                    grp_str = str(c[3])
                    course = c[6] if len(c) > 6 else None
                    if not course or course == 1:
                        cg = get_course_from_group(grp_str)
                        course = int(cg) if cg != "?" else (course or 1)
                    status = c[7] if len(c) > 7 else 'active'
                    s_fio = c[2] if len(c) > 2 and c[2] else make_short_fio(c[1])
                    prepared.append((c[0], c[1], s_fio, grp_str, course, gender_to_ru(c[4]), status, sort_order))

            self.conn.executemany(
                "INSERT INTO cadets (rank, fio, short_fio, group_num, course, gender, status, sort_order) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                prepared
            )

    def bulk_increment_course(self, max_course: int = 5, update_group_numbers: bool = True):
        from core.parsers import change_group_course
        with self.conn:
            cursor = self.conn.cursor()
            cursor.execute("SELECT id, group_num, course FROM cadets WHERE is_active = 1 AND status = 'active'")
            for cid, grp, crs in cursor.fetchall():
                if crs >= max_course:
                    cursor.execute("UPDATE cadets SET status = 'graduated' WHERE id = ?", (cid,))
                else:
                    new_crs = crs + 1
                    new_grp = change_group_course(grp, new_crs) if update_group_numbers else grp
                    cursor.execute("UPDATE cadets SET course = ?, group_num = ? WHERE id = ?", (new_crs, new_grp, cid))

    def bulk_set_course(self, course: int, cadet_ids: Optional[List[int]] = None, update_group_numbers: bool = True):
        from core.parsers import change_group_course
        with self.conn:
            cursor = self.conn.cursor()
            if cadet_ids:
                q_marks = ",".join("?" for _ in cadet_ids)
                cursor.execute(f"SELECT id, group_num FROM cadets WHERE id IN ({q_marks})", cadet_ids)
            else:
                cursor.execute("SELECT id, group_num FROM cadets WHERE is_active = 1")
            rows = cursor.fetchall()

            for cid, grp in rows:
                new_grp = change_group_course(grp, course) if update_group_numbers else grp
                cursor.execute(
                    "UPDATE cadets SET course = ?, group_num = ?, status = 'active' WHERE id = ?",
                    (course, new_grp, cid)
                )

    def get_absence_types(self):
        cur = self.conn.cursor()
        cur.execute("""
            SELECT id, name, code, category 
            FROM absence_types 
            ORDER BY 
                CASE category 
                    WHEN 'Наряд' THEN 1 
                    WHEN 'Увольнение' THEN 2 
                    WHEN 'Госпиталь' THEN 3 
                    WHEN 'Лазарет' THEN 4 
                    WHEN 'Командировка' THEN 5 
                    WHEN 'Похороны' THEN 6 
                    WHEN 'Отпуск' THEN 7 
                    ELSE 8 
                END, 
                name
        """)
        return [dict(r) for r in cur.fetchall()]

    def get_absence_type_by_id(self, type_id):
        cur = self.conn.cursor()
        cur.execute("SELECT id, name, code, category FROM absence_types WHERE id = ?", (type_id,))
        row = cur.fetchone()
        return dict(row) if row else None

    def get_distinct_categories(self):
        cur = self.conn.cursor()
        cur.execute("""
            SELECT DISTINCT category 
            FROM absence_types 
            ORDER BY 
                CASE category 
                    WHEN 'Наряд' THEN 1 
                    WHEN 'Увольнение' THEN 2 
                    WHEN 'Госпиталь' THEN 3 
                    WHEN 'Лазарет' THEN 4 
                    WHEN 'Командировка' THEN 5 
                    WHEN 'Похороны' THEN 6 
                    WHEN 'Отпуск' THEN 7 
                    ELSE 8 
                END, 
                category
        """)
        return [r[0] for r in cur.fetchall()]

    def add_absence_type(self, name, code, category):
        with self.conn:
            cur = self.conn.execute(
                "INSERT INTO absence_types (name, code, category) VALUES (?, ?, ?)",
                (name, code, category)
            )
            return cur.lastrowid

    def update_absence_type(self, type_id, name, code, category):
        with self.conn:
            self.conn.execute("""
                UPDATE absence_types 
                SET name = ?, code = ?, category = ? 
                WHERE id = ?
            """, (name, code, category, type_id))

    def delete_absence_type(self, type_id):
        with self.conn:
            self.conn.execute("DELETE FROM absence_types WHERE id = ?", (type_id,))

    def get_absences_for_date(self, target_date, gender_filter: Optional[str] = None):
        cur = self.conn.cursor()
        cur.execute("""
            SELECT da.id, c.id AS cadet_id, c.rank, c.short_fio, c.fio, c.group_num, c.course, c.gender, c.status,
                   t.id AS type_id, t.name AS type_name, t.code AS type_code, t.category, da.note
            FROM daily_absences da
            JOIN cadets c ON da.cadet_id = c.id
            JOIN absence_types t ON da.type_id = t.id
            WHERE da.report_date = ?
            ORDER BY c.sort_order ASC, c.group_num ASC, c.fio ASC
        """, (target_date,))
        rows = [dict(r) for r in cur.fetchall()]
        if gender_filter:
            norm_target = normalize_gender(gender_filter)
            return [r for r in rows if normalize_gender(r.get('gender')) == norm_target]
        return rows

    def set_absence(self, target_date, cadet_id, type_id, note=""):
        with self.conn:
            self.conn.execute("""
                INSERT OR REPLACE INTO daily_absences (report_date, cadet_id, type_id, note)
                VALUES (?, ?, ?, ?)
            """, (target_date, cadet_id, type_id, note))

    def remove_absence(self, absence_id):
        with self.conn:
            self.conn.execute("DELETE FROM daily_absences WHERE id = ?", (absence_id,))
