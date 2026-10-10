import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QTableWidget, QTableWidgetItem, QComboBox,
    QLineEdit, QDateEdit, QFileDialog, QMessageBox, QTabWidget,
    QHeaderView, QFrame
)
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtGui import QFont

from database import Database
from excel_manager import import_cadets_from_excel, export_report_to_excel

MODERN_STYLE = """
QMainWindow { background-color: #F8FAFC; }
QTabWidget::pane { border: 1px solid #E2E8F0; background: #FFFFFF; border-radius: 8px; }
QTabBar::tab {
    background: #EDF2F7; color: #4A5568; padding: 10px 22px;
    font-size: 13px; font-weight: bold; border-top-left-radius: 6px;
    border-top-right-radius: 6px; margin-right: 4px;
}
QTabBar::tab:selected { background: #3182CE; color: white; }
QPushButton {
    background-color: #3182CE; color: white; font-weight: bold;
    font-size: 12px; border-radius: 6px; padding: 8px 16px; border: none;
}
QPushButton:hover { background-color: #2B6CB0; }
QPushButton#deleteBtn { background-color: #E53E3E; padding: 4px 10px; }
QPushButton#deleteBtn:hover { background-color: #C53030; }
QPushButton#exportBtn { background-color: #38A169; font-size: 14px; padding: 10px 22px; }
QPushButton#exportBtn:hover { background-color: #2F855A; }
QTableWidget {
    background-color: white; gridline-color: #E2E8F0;
    border: 1px solid #CBD5E0; border-radius: 6px; font-size: 13px;
}
QHeaderView::section {
    background-color: #EDF2F7; color: #2D3748; font-weight: bold;
    padding: 6px; border: 1px solid #E2E8F0;
}
QLineEdit, QComboBox, QDateEdit {
    padding: 6px 10px; border: 1px solid #CBD5E0; border-radius: 6px;
    background: white; font-size: 13px;
}
QFrame#cardFrame {
    background-color: white; border: 1px solid #E2E8F0;
    border-radius: 8px; padding: 10px;
}
"""

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Строевая записка курса — Автоматизированный расход")
        self.resize(1150, 780)
        self.db = Database()
        self.setStyleSheet(MODERN_STYLE)
        self.init_ui()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        tabs = QTabWidget()
        tabs.addTab(self.create_today_tab(), "📋 Расход на сегодня")
        tabs.addTab(self.create_cadets_tab(), "👥 Личный состав")
        tabs.addTab(self.create_settings_tab(), "⚙️ Справочник нарядов и шаблоны")
        main_layout.addWidget(tabs)

    def create_today_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)

        top_bar = QHBoxLayout()
        top_bar.addWidget(QLabel("<b>Дата расхода:</b>"))
        self.date_edit = QDateEdit(QDate.currentDate())
        self.date_edit.setCalendarPopup(True)
        self.date_edit.dateChanged.connect(self.load_daily_data)
        top_bar.addWidget(self.date_edit)

        self.stat_card = QFrame()
        self.stat_card.setObjectName("cardFrame")
        card_layout = QHBoxLayout(self.stat_card)
        self.lbl_stats = QLabel("По списку: 0 | Налицо: 0 | Отсутствуют: 0")
        self.lbl_stats.setStyleSheet("font-size: 14px;")
        card_layout.addWidget(self.lbl_stats)
        top_bar.addWidget(self.stat_card)

        top_bar.addStretch()

        btn_export = QPushButton("📤 Экспорт в Excel")
        btn_export.setObjectName("exportBtn")
        btn_export.clicked.connect(self.export_excel)
        top_bar.addWidget(btn_export)
        layout.addLayout(top_bar)

        # Панель добавления
        add_box = QFrame()
        add_box.setObjectName("cardFrame")
        add_layout = QHBoxLayout(add_box)

        self.cb_cadets = QComboBox()
        self.cb_cadets.setEditable(True)
        self.cb_cadets.setPlaceholderText("Фамилия или группа...")
        add_layout.addWidget(self.cb_cadets, 3)

        self.cb_types = QComboBox()
        add_layout.addWidget(self.cb_types, 2)

        self.edit_note = QLineEdit()
        self.edit_note.setPlaceholderText("Примечание (город, корпус, задача)...")
        add_layout.addWidget(self.edit_note, 3)

        btn_add = QPushButton("➕ Добавить в расход")
        btn_add.clicked.connect(self.add_absence_record)
        add_layout.addWidget(btn_add, 1)
        layout.addWidget(add_box)

        self.table_absences = QTableWidget()
        self.table_absences.setColumnCount(7)
        self.table_absences.setHorizontalHeaderLabels([
            "№", "Звание", "ФИО", "Группа", "Причина (наряд)", "Примечание", "Действие"
        ])
        self.table_absences.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table_absences.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table_absences.horizontalHeader().setSectionResizeMode(6, QHeaderView.ResizeToContents)
        layout.addWidget(self.table_absences)

        self.load_types_combo()
        self.load_cadets_combo()
        self.load_daily_data()
        return tab

    def create_cadets_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)

        top_panel = QHBoxLayout()
        btn_import = QPushButton("📥 Импорт базы курса из Excel")
        btn_import.clicked.connect(self.import_cadets_excel)
        top_panel.addWidget(btn_import)

        top_panel.addStretch()
        self.search_cadet = QLineEdit()
        self.search_cadet.setPlaceholderText("Поиск по фамилии или группе...")
        self.search_cadet.textChanged.connect(self.filter_cadets_table)
        top_panel.addWidget(self.search_cadet)
        layout.addLayout(top_panel)

        self.table_cadets = QTableWidget()
        self.table_cadets.setColumnCount(5)
        self.table_cadets.setHorizontalHeaderLabels(["ID", "Звание", "ФИО", "Группа", "Пол"])
        self.table_cadets.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.table_cadets)

        self.load_cadets_table()
        return tab

    def create_settings_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # Блок настроек шапки
        layout.addWidget(QLabel("<b>Шапка и подписи строевой записки:</b>"))
        head_box = QFrame()
        head_box.setObjectName("cardFrame")
        h_layout = QVBoxLayout(head_box)

        self.edit_unit = QLineEdit(self.db.get_setting("unit_title"))
        self.edit_unit.setPlaceholderText("Название курса / факультета")
        h_layout.addWidget(QLabel("Наименование подразделения:"))
        h_layout.addWidget(self.edit_unit)

        r_chief = QHBoxLayout()
        self.edit_chief_title = QLineEdit(self.db.get_setting("chief_title"))
        self.edit_chief_title.setPlaceholderText("Должность начальника курса")
        self.edit_chief_name = QLineEdit(self.db.get_setting("chief_name"))
        self.edit_chief_name.setPlaceholderText("Воинское звание и инициалы")
        r_chief.addWidget(self.edit_chief_title, 2)
        r_chief.addWidget(self.edit_chief_name, 2)
        
        btn_save_head = QPushButton("Сохранить шапку")
        btn_save_head.clicked.connect(self.save_header_settings)
        r_chief.addWidget(btn_save_head, 1)
        h_layout.addLayout(r_chief)
        layout.addWidget(head_box)

        # Блок нарядов
        layout.addWidget(QLabel("<b>Добавить новый наряд или причину отсутствия:</b>"))
        row1 = QHBoxLayout()
        self.new_type_name = QLineEdit()
        self.new_type_name.setPlaceholderText("Название (например, Дежурный по парку)")
        row1.addWidget(self.new_type_name, 3)

        self.new_type_code = QLineEdit()
        self.new_type_code.setPlaceholderText("Шифр (например, ДпП)")
        row1.addWidget(self.new_type_code, 1)

        self.new_type_cat = QComboBox()
        self.new_type_cat.setEditable(True)
        self.new_type_cat.addItems(["Наряд", "Госпиталь", "Лазарет", "Командировка", "Похороны", "Отпуск", "Карантин"])
        row1.addWidget(self.new_type_cat, 2)

        btn_save_type = QPushButton("Добавить причину")
        btn_save_type.clicked.connect(self.save_new_type)
        row1.addWidget(btn_save_type, 1)
        layout.addLayout(row1)

        self.table_types = QTableWidget()
        self.table_types.setColumnCount(4)
        self.table_types.setHorizontalHeaderLabels(["ID", "Название / Наряд", "Шифр", "Категория в сводке"])
        self.table_types.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.table_types)

        self.load_types_table()
        return tab

    # ---------------- Вспомогательные методы ----------------
    def get_selected_date_str(self):
        return self.date_edit.date().toString("yyyy-MM-dd")

    def load_types_combo(self):
        self.cb_types.clear()
        types = self.db.get_absence_types()
        for t in types:
            display = f"{t['name']} ({t['code']}) [{t['category']}]" if t['code'] else f"{t['name']} [{t['category']}]"
            self.cb_types.addItem(display, userData=t['id'])

    def load_cadets_combo(self):
        self.cb_cadets.clear()
        cadets = self.db.get_all_cadets()
        for c in cadets:
            self.cb_cadets.addItem(f"{c['short_fio']} | {c['rank']} (гр. {c['group_num']})", userData=c['id'])

    def load_daily_data(self):
        d_str = self.get_selected_date_str()
        absences = self.db.get_absences_for_date(d_str)
        cadets = self.db.get_all_cadets()

        total = len(cadets)
        abs_count = len(absences)
        present = total - abs_count
        self.lbl_stats.setText(
            f"<b>По списку:</b> {total}  |  "
            f"<b style='color:#38A169;'>Налицо:</b> {present}  |  "
            f"<b style='color:#E53E3E;'>Отсутствуют:</b> {abs_count}"
        )

        self.table_absences.setRowCount(0)
        for idx, a in enumerate(absences, 1):
            r = self.table_absences.rowCount()
            self.table_absences.insertRow(r)
            self.table_absences.setItem(r, 0, QTableWidgetItem(str(idx)))
            self.table_absences.setItem(r, 1, QTableWidgetItem(a['rank']))
            self.table_absences.setItem(r, 2, QTableWidgetItem(a['short_fio']))
            self.table_absences.setItem(r, 3, QTableWidgetItem(a['group_num']))
            self.table_absences.setItem(r, 4, QTableWidgetItem(f"{a['type_name']} ({a['category']})"))
            self.table_absences.setItem(r, 5, QTableWidgetItem(a['note'] or ""))

            btn_del = QPushButton("Удалить")
            btn_del.setObjectName("deleteBtn")
            abs_id = a['id']
            btn_del.clicked.connect(lambda _, aid=abs_id: self.delete_absence(aid))
            self.table_absences.setCellWidget(r, 6, btn_del)

    def add_absence_record(self):
        cadet_id = self.cb_cadets.currentData()
        type_id = self.cb_types.currentData()
        note = self.edit_note.text().strip()
        if not cadet_id or not type_id:
            QMessageBox.warning(self, "Внимание", "Выберите курсанта и наряд/причину!")
            return
        d_str = self.get_selected_date_str()
        self.db.set_absence(d_str, cadet_id, type_id, note)
        self.edit_note.clear()
        self.load_daily_data()

    def delete_absence(self, abs_id):
        self.db.remove_absence(abs_id)
        self.load_daily_data()

    def export_excel(self):
        d_str = self.get_selected_date_str()
        cadets = self.db.get_all_cadets()
        absences = self.db.get_absences_for_date(d_str)

        if not cadets:
            QMessageBox.warning(self, "Внимание", "Сначала загрузите личный состав курса!")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить строевую записку", f"Строевая_записка_{d_str}.xlsx", "Excel Files (*.xlsx)"
        )
        if file_path:
            try:
                # Получаем все уникальные категории из используемых типов отсутствия
                all_types = self.db.get_absence_types()
                active_cats = []
                for t in all_types:
                    if t['category'] not in active_cats:
                        active_cats.append(t['category'])

                export_report_to_excel(
                    output_path=file_path,
                    report_date=d_str,
                    cadets=cadets,
                    absences=absences,
                    unit_title=self.db.get_setting("unit_title"),
                    chief_title=self.db.get_setting("chief_title"),
                    chief_name=self.db.get_setting("chief_name"),
                    active_categories=active_cats
                )
                QMessageBox.information(self, "Успех", f"Строевая записка успешно создана:\n{file_path}")
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Не удалось сформировать файл:\n{e}")

    def import_cadets_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Выберите файл с базой курса", "", "Excel Files (*.xlsx *.xls)"
        )
        if file_path:
            try:
                cadets = import_cadets_from_excel(file_path)
                if not cadets:
                    QMessageBox.warning(self, "Внимание", "Курсанты в файле не найдены.")
                    return
                self.db.clear_and_import_cadets(cadets)
                self.load_cadets_table()
                self.load_cadets_combo()
                self.load_daily_data()
                QMessageBox.information(self, "Успех", f"Загружено курсантов: {len(cadets)} чел.")
            except Exception as e:
                QMessageBox.critical(self, "Ошибка импорта", f"Ошибка чтения файла:\n{e}")

    def load_cadets_table(self):
        cadets = self.db.get_all_cadets()
        self.table_cadets.setRowCount(0)
        for c in cadets:
            r = self.table_cadets.rowCount()
            self.table_cadets.insertRow(r)
            self.table_cadets.setItem(r, 0, QTableWidgetItem(str(c['id'])))
            self.table_cadets.setItem(r, 1, QTableWidgetItem(c['rank']))
            self.table_cadets.setItem(r, 2, QTableWidgetItem(c['fio']))
            self.table_cadets.setItem(r, 3, QTableWidgetItem(c['group_num']))
            self.table_cadets.setItem(r, 4, QTableWidgetItem(c['gender']))

    def filter_cadets_table(self, text):
        for r in range(self.table_cadets.rowCount()):
            fio = self.table_cadets.item(r, 2).text().lower()
            grp = self.table_cadets.item(r, 3).text().lower()
            match = (text.lower() in fio) or (text.lower() in grp)
            self.table_cadets.setRowHidden(r, not match)

    def save_header_settings(self):
        self.db.set_setting("unit_title", self.edit_unit.text().strip())
        self.db.set_setting("chief_title", self.edit_chief_title.text().strip())
        self.db.set_setting("chief_name", self.edit_chief_name.text().strip())
        QMessageBox.information(self, "Сохранено", "Настройки бланка успешно обновлены!")

    def save_new_type(self):
        name = self.new_type_name.text().strip()
        code = self.new_type_code.text().strip()
        cat = self.new_type_cat.currentText().strip()
        if not name or not cat:
            return
        self.db.add_absence_type(name, code, cat)
        self.new_type_name.clear()
        self.new_type_code.clear()
        self.load_types_table()
        self.load_types_combo()

    def load_types_table(self):
        types = self.db.get_absence_types()
        self.table_types.setRowCount(0)
        for t in types:
            r = self.table_types.rowCount()
            self.table_types.insertRow(r)
            self.table_types.setItem(r, 0, QTableWidgetItem(str(t['id'])))
            self.table_types.setItem(r, 1, QTableWidgetItem(t['name']))
            self.table_types.setItem(r, 2, QTableWidgetItem(t['code'] or ""))
            self.table_types.setItem(r, 3, QTableWidgetItem(t['category']))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
