import sqlite3

class Database:
    def __init__(self, db_path="app_data.db"):
        self.conn = sqlite3.connect(db_path)
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
                    gender TEXT DEFAULT 'М',
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
        # Исправление опечатки и добавление Увольнения в существующие БД
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

    def get_all_cadets(self):
        cur = self.conn.cursor()
        cur.execute("""
            SELECT id, rank, fio, short_fio, group_num, gender, sort_order 
            FROM cadets 
            WHERE is_active = 1 
            ORDER BY sort_order ASC, group_num ASC, fio ASC
        """)
        return [dict(r) for r in cur.fetchall()]

    def get_cadet_by_id(self, cadet_id):
        cur = self.conn.cursor()
        cur.execute("""
            SELECT id, rank, fio, short_fio, group_num, gender, sort_order 
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
            WHERE is_active = 1 
            ORDER BY sort_order ASC, group_num ASC
        """)
        return [r[0] for r in cur.fetchall()]

    def add_cadet(self, rank, fio, group_num, gender="М", short_fio=None):
        if not short_fio:
            from excel_manager import make_short_fio
            short_fio = make_short_fio(fio)
        cur = self.conn.cursor()
        cur.execute("SELECT sort_order FROM cadets WHERE group_num = ? LIMIT 1", (group_num,))
        row = cur.fetchone()
        if row:
            sort_order = row[0]
        else:
            cur.execute("SELECT COALESCE(MAX(sort_order), 0) + 1 FROM cadets")
            sort_order = cur.fetchone()[0]

        with self.conn:
            cur = self.conn.execute("""
                INSERT INTO cadets (rank, fio, short_fio, group_num, gender, sort_order, is_active)
                VALUES (?, ?, ?, ?, ?, ?, 1)
            """, (rank, fio, short_fio, group_num, gender, sort_order))
            return cur.lastrowid

    def update_cadet(self, cadet_id, rank, fio, group_num, gender="М", short_fio=None):
        if not short_fio:
            from excel_manager import make_short_fio
            short_fio = make_short_fio(fio)
        with self.conn:
            self.conn.execute("""
                UPDATE cadets 
                SET rank = ?, fio = ?, short_fio = ?, group_num = ?, gender = ?
                WHERE id = ?
            """, (rank, fio, short_fio, group_num, gender, cadet_id))

    def delete_cadet(self, cadet_id):
        with self.conn:
            self.conn.execute("DELETE FROM daily_absences WHERE cadet_id = ?", (cadet_id,))
            self.conn.execute("DELETE FROM cadets WHERE id = ?", (cadet_id,))

    def clear_and_import_cadets(self, cadets_list):
        with self.conn:
            self.conn.execute("DELETE FROM daily_absences")
            self.conn.execute("DELETE FROM cadets")
            prepared = []
            for idx, c in enumerate(cadets_list):
                if isinstance(c, dict):
                    prepared.append((
                        c.get('rank', ''), c['fio'], c.get('short_fio', ''), 
                        c['group_num'], c.get('gender', 'М'), c.get('sort_order', idx)
                    ))
                else:
                    sort_order = c[5] if len(c) > 5 else idx
                    prepared.append((c[0], c[1], c[2], c[3], c[4], sort_order))

            self.conn.executemany(
                "INSERT INTO cadets (rank, fio, short_fio, group_num, gender, sort_order) VALUES (?, ?, ?, ?, ?, ?)",
                prepared
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
            self.conn.execute(
                "INSERT INTO absence_types (name, code, category) VALUES (?, ?, ?)",
                (name, code, category)
            )

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

    def get_absences_for_date(self, target_date):
        cur = self.conn.cursor()
        cur.execute("""
            SELECT da.id, c.id AS cadet_id, c.rank, c.short_fio, c.fio, c.group_num, c.gender, 
                   t.id AS type_id, t.name AS type_name, t.code AS type_code, t.category, da.note
            FROM daily_absences da
            JOIN cadets c ON da.cadet_id = c.id
            JOIN absence_types t ON da.type_id = t.id
            WHERE da.report_date = ?
            ORDER BY c.sort_order ASC, c.group_num ASC, c.fio ASC
        """, (target_date,))
        return [dict(r) for r in cur.fetchall()]

    def set_absence(self, target_date, cadet_id, type_id, note=""):
        with self.conn:
            self.conn.execute("""
                INSERT OR REPLACE INTO daily_absences (report_date, cadet_id, type_id, note)
                VALUES (?, ?, ?, ?)
            """, (target_date, cadet_id, type_id, note))

    def remove_absence(self, absence_id):
        with self.conn:
            self.conn.execute("DELETE FROM daily_absences WHERE id = ?", (absence_id,))
