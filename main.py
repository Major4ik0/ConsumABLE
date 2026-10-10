import sys
import os
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QTableWidget, QTableWidgetItem, QComboBox,
    QLineEdit, QDateEdit, QFileDialog, QMessageBox, QTabWidget,
    QHeaderView, QFrame
)
from PyQt5.QtCore import Qt, QDate, QByteArray, QSize
from PyQt5.QtGui import QIcon, QPixmap, QPainter
from PyQt5.QtSvg import QSvgRenderer

from database import Database[span_0](start_span)[span_0](end_span)
from excel_manager import import_cadets_from_excel, export_report_to_excel[span_1](start_span)[span_1](end_span)

# Векторные SVG-иконки
SVG_CLIPBOARD = '''<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2"/><rect x="8" y="2" width="8" height="4" rx="1" ry="1"/><path d="M9 12h6"/><path d="M9 16h6"/></svg>'''
SVG_USERS = '''<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>'''
SVG_SETTINGS = '''<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z"/></svg>'''
SVG_EXPORT = '''<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>'''
SVG_IMPORT = '''<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>'''
SVG_PLUS = '''<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>'''
SVG_TRASH = '''<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>'''
SVG_SAVE = '''<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>'''

def create_svg_icon(svg_str: str, color: str = "#2D3748", size: int = 20) -> QIcon:
    colored_svg = svg_str.replace('currentColor', color)
    renderer = QSvgRenderer(QByteArray(colored_svg.encode('utf-8')))
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.transparent)
    painter = QPainter(pixmap)
    renderer.render(painter)
    painter.end()
    return QIcon(pixmap)

MODERN_STYLE = """
QMainWindow { background-color: #F8FAFC; }
QTabWidget::pane { border: 1px solid #E2E8F0; background: #FFFFFF; border-radius: 8px; }
QTabBar::tab {
    background: #EDF2F7; color: #4A5568; padding: 10px 24px;
    font-size: 13px; font-weight: bold; border-top-left-radius: 6px;
    border-top-right-radius: 6px; margin-right: 4px;
}
QTabBar::tab:selected { background: #3182CE; color: white; }
QPushButton {
    background-color: #3182CE; color: white; font-weight: bold;
    font-size: 12px; border-radius: 6px; padding: 8px 16px; border: none;
}
QPushButton:hover { background-color: #2B6CB0; }
QPushButton#deleteBtn { background-color: #E53E3E; padding: 5px 10px; font-size: 11px; }
QPushButton#deleteBtn:hover { background-color: #C53030; }
QPushButton#exportBtn { background-color: #2B8A3E; font-size: 13px; padding: 9px 20px; }
QPushButton#exportBtn:hover { background-color: #237032; }
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
        self.setWindowTitle("ConsumABLE — Управление строевой запиской")
        self.resize(1150, 780)
        self.db = Database()[span_2](start_span)[span_2](end_span)

        # Установка иконки окна из icon.png
        icon_path = os.path.join(os.path.dirname(__file__), "icon.png")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        elif os.path.exists("icon.png"):
            self.setWindowIcon(QIcon("icon.png"))

        self.setStyleSheet(MODERN_STYLE)
        self.init_ui()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        tabs = QTabWidget()
        tabs.setIconSize(QSize(18, 18))

        icon_today = create_svg_icon(SVG_CLIPBOARD, "#4A5568")
        icon_cadets = create_svg_icon(SVG_USERS, "#4A5568")
        icon_settings = create_svg_icon(SVG_SETTINGS, "#4A5568")

        tabs.addTab(self.create_today_tab(), icon_today, "Расход на сегодня")
        tabs.addTab(self.create_cadets_tab(), icon_cadets, "Личный состав")
        tabs.addTab(self.create_settings_tab(), icon_settings, "Справочник и настройки")
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

        btn_export = QPushButton("Экспорт в Excel")
        btn_export.setObjectName("exportBtn")
        btn_export.setIcon(create_svg_icon(SVG_EXPORT, "#FFFFFF", 16))
        btn_export.clicked.connect(self.export_excel)
        top_bar.addWidget(btn_export)
        layout.addLayout(top_bar)

        # Панель быстрого добавления
        add_box = QFrame()
        add_box.setObjectName("cardFrame")
        add_layout = QHBoxLayout(add_box)

        self.cb_cadets = QComboBox()
        self.cb_cadets.setEditable(True)
        self.cb_cadets.setPlaceholderText("Фамилия или подразделение...")
        add_layout.addWidget(self.cb_cadets, 3)

        self.cb_types = QComboBox()
        add_layout.addWidget(self.cb_types, 2)

        self.edit_note = QLineEdit()
        self.edit_note.setPlaceholderText("Примечание (город, корпус, задача)...")
        add_layout.addWidget(self.edit_note, 3)

        btn_add = QPushButton("Добавить в расход")
        btn_add.setIcon(create_svg_icon(SVG_PLUS, "#FFFFFF", 14))
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
        btn_import = QPushButton("Импорт базы из Excel")
        btn_import.setIcon(create_svg_icon(SVG_IMPORT, "#FFFFFF", 16))
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

        layout.addWidget(QLabel("<b>Шапка и подписи строевой записки:</b>"))
        head_box = QFrame()
        head_box.setObjectName("cardFrame")
        h_layout = QVBoxLayout(head_box)

        self.edit_unit = QLineEdit(self.db.get_setting("unit_title"))[span_3](start_span)[span_3](end_span)
        self.edit_unit.setPlaceholderText("Название курса / факультета")
        h_layout.addWidget(QLabel("Наименование подразделения:"))
        h_layout.addWidget(self.edit_unit)

        r_chief = QHBoxLayout()
        self.edit_chief_title = QLineEdit(self.db.get_setting("chief_title"))[span_4](start_span)[span_4](end_span)
        self.edit_chief_title.setPlaceholderText("Должность начальника курса")
        self.edit_chief_name = QLineEdit(self.db.get_setting("chief_name"))[span_5](start_span)[span_5](end_span)
        self.edit_chief_name.setPlaceholderText("Воинское звание и инициалы")
        r_chief.addWidget(self.edit_chief_title, 2)
        r_chief.addWidget(self.edit_chief_name, 2)
        
        btn_save_head = QPushButton("Сохранить шапку")
        btn_save_head.setIcon(create_svg_icon(SVG_SAVE, "#FFFFFF", 14))
        btn_save_head.clicked.connect(self.save_header_settings)
        r_chief.addWidget(btn_save_head, 1)
        h_layout.addLayout(r_chief)
        layout.addWidget(head_box)

        layout.addWidget(QLabel("<b>Добавить новый наряд или причину отсутствия:</b>"))
        row1 = QHBoxLayout()
        self.new_type_name = QLineEdit()
        self.new_type_name.setPlaceholderText("Название (например, Дежурный по КПП)")
        row1.addWidget(self.new_type_name, 3)

        self.new_type_code = QLineEdit()
        self.new_type_code.setPlaceholderText("Шифр (например, ДпКПП)")
        row1.addWidget(self.new_type_code, 1)

        self.new_type_cat = QComboBox()
        self.new_type_cat.setEditable(True)
        self.new_type_cat.addItems(["Наряд", "Госпиталь", "Лазарет", "Командировка", "Похороны", "Отпуск", "Карантин"])
        row1.addWidget(self.new_type_cat, 2)

        btn_save_type = QPushButton("Добавить причину")
        btn_save_type.setIcon(create_svg_icon(SVG_PLUS, "#FFFFFF", 14))
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

    def get_selected_date_str(self):
        return self.date_edit.date().toString("yyyy-MM-dd")

    def load_types_combo(self):
        self.cb_types.clear()
        types = self.db.get_absence_types()[span_6](start_span)[span_6](end_span)
        for t in types:
            display = f"{t['name']} ({t['code']}) [{t['category']}]" if t['code'] else f"{t['name']} [{t['category']}]"
            self.cb_types.addItem(display, userData=t['id'])

    def load_cadets_combo(self):
        self.cb_cadets.clear()
        cadets = self.db.get_all_cadets()[span_7](start_span)[span_7](end_span)
        for c in cadets:
            self.cb_cadets.addItem(f"{c['short_fio']} | {c['rank']} (гр. {c['group_num']})", userData=c['id'])

    def load_daily_data(self):
        d_str = self.get_selected_date_str()
        absences = self.db.get_absences_for_date(d_str)[span_8](start_span)[span_8](end_span)
        cadets = self.db.get_all_cadets()[span_9](start_span)[span_9](end_span)

        total = len(cadets)
        abs_count = len(absences)
        present = total - abs_count
        self.lbl_stats.setText(
            f"<b>По списку:</b> {total}  |  "
            f"<b style='color:#2B8A3E;'>Налицо:</b> {present}  |  "
            f"<b style='color:#E53E3E;'>Отсутствуют:</b> {abs_count}"
        )

        self.table_absences.setRowCount(0)
        icon_del = create_svg_icon(SVG_TRASH, "#FFFFFF", 12)
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
            btn_del.setIcon(icon_del)
            abs_id = a['id']
            btn_del.clicked.connect(lambda _, aid=abs_id: self.delete_absence(aid))
            self.table_absences.setCellWidget(r, 6, btn_del)

    def add_absence_record(self):
        cadet_id = self.cb_cadets.currentData()
        type_id = self.cb_types.currentData()
        note = self.edit_note.text().strip()
        if not cadet_id or not type_id:
            QMessageBox.warning(self, "Внимание", "Выберите курсанта и причину отсутствия.")
            return
        d_str = self.get_selected_date_str()
        self.db.set_absence(d_str, cadet_id, type_id, note)[span_10](start_span)[span_10](end_span)
        self.edit_note.clear()
        self.load_daily_data()

    def delete_absence(self, abs_id):
        self.db.remove_absence(abs_id)[span_11](start_span)[span_11](end_span)
        self.load_daily_data()

    def export_excel(self):
        d_str = self.get_selected_date_str()
        cadets = self.db.get_all_cadets()[span_12](start_span)[span_12](end_span)
        absences = self.db.get_absences_for_date(d_str)[span_13](start_span)[span_13](end_span)

        if not cadets:
            QMessageBox.warning(self, "Внимание", "Сначала загрузите личный состав курса.")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить строевую записку", f"Строевая_записка_{d_str}.xlsx", "Excel Files (*.xlsx)"
        )
        if file_path:
            try:
                all_types = self.db.get_absence_types()[span_14](start_span)[span_14](end_span)
                active_cats = []
                for t in all_types:
                    if t['category'] not in active_cats:
                        active_cats.append(t['category'])

                export_report_to_excel(
                    output_path=file_path,
                    report_date=d_str,
                    cadets=cadets,
                    absences=absences,
                    unit_title=self.db.get_setting("unit_title"),[span_15](start_span)[span_15](end_span)
                    chief_title=self.db.get_setting("chief_title"),[span_16](start_span)[span_16](end_span)
                    chief_name=self.db.get_setting("chief_name"),[span_17](start_span)[span_17](end_span)
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
                cadets = import_cadets_from_excel(file_path)[span_18](start_span)[span_18](end_span)
                if not cadets:
                    QMessageBox.warning(self, "Внимание", "Курсанты в файле не найдены.")
                    return
                self.db.clear_and_import_cadets(cadets)[span_19](start_span)[span_19](end_span)
                self.load_cadets_table()
                self.load_cadets_combo()
                self.load_daily_data()
                QMessageBox.information(self, "Успех", f"Загружено курсантов: {len(cadets)} чел.")
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Ошибка чтения файла:\n{e}")

    def load_cadets_table(self):
        cadets = self.db.get_all_cadets()[span_20](start_span)[span_20](end_span)
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
        self.db.set_setting("unit_title", self.edit_unit.text().strip())[span_21](start_span)[span_21](end_span)
        self.db.set_setting("chief_title", self.edit_chief_title.text().strip())[span_22](start_span)[span_22](end_span)
        self.db.set_setting("chief_name", self.edit_chief_name.text().strip())[span_23](start_span)[span_23](end_span)
        QMessageBox.information(self, "Сохранено", "Настройки бланка успешно обновлены.")

    def save_new_type(self):
        name = self.new_type_name.text().strip()
        code = self.new_type_code.text().strip()
        cat = self.new_type_cat.currentText().strip()
        if not name or not cat:
            return
        self.db.add_absence_type(name, code, cat)[span_24](start_span)[span_24](end_span)
        self.new_type_name.clear()
        self.new_type_code.clear()
        self.load_types_table()
        self.load_types_combo()

    def load_types_table(self):
        types = self.db.get_absence_types()[span_25](start_span)[span_25](end_span)
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
