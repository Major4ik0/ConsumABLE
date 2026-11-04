import os
import sys
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QTableWidget, QTableWidgetItem, QComboBox,
    QLineEdit, QDateEdit, QFileDialog, QMessageBox, QTabWidget,
    QHeaderView, QFrame, QAbstractItemView, QDialog, QFormLayout, QDialogButtonBox
)
from PyQt5.QtCore import Qt, QDate, QByteArray, QSize
from PyQt5.QtGui import QIcon, QPixmap, QPainter, QColor
from PyQt5.QtSvg import QSvgRenderer

from database import Database
from core.parsers import parse_roster_excel, make_short_fio
from core.models import normalize_gender, gender_to_ru
from core.balance import compute_tri_slice_balance
from core.exporter import export_report_to_excel
from ui.dialogs.import_wizard import ImportReviewDialog

# SVG Icons
SVG_CLIPBOARD = '''<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2"/><rect x="8" y="2" width="8" height="4" rx="1" ry="1"/><path d="M9 12h6"/><path d="M9 16h6"/></svg>'''
SVG_USERS = '''<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>'''
SVG_SETTINGS = '''<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z"/></svg>'''
SVG_EXPORT = '''<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>'''
SVG_IMPORT = '''<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>'''
SVG_PLUS = '''<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>'''
SVG_TRASH = '''<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>'''
SVG_CALENDAR = '''<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>'''
SVG_EDIT = '''<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>'''

def create_svg_icon(svg_str: str, color: str = "#2563EB", size: int = 18) -> QIcon:
    colored_svg = svg_str.replace('currentColor', color)
    renderer = QSvgRenderer(QByteArray(colored_svg.encode('utf-8')))
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.transparent)
    painter = QPainter(pixmap)
    renderer.render(painter)
    painter.end()
    return QIcon(pixmap)

MODERN_STYLE = """
* {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
    color: #0F172A;
}

QMainWindow {
    background-color: #F8FAFC;
}

QDialog {
    background-color: #FFFFFF;
}

QTabWidget::pane {
    border: 1px solid #E2E8F0;
    background: #FFFFFF;
    border-radius: 12px;
    top: -1px;
}

QTabBar::tab {
    background: #F1F5F9;
    color: #64748B;
    padding: 10px 24px;
    font-size: 13px;
    font-weight: 600;
    border: 1px solid #E2E8F0;
    border-top-left-radius: 8px;
    border-top-right-radius: 8px;
    margin-right: 4px;
    min-height: 18px;
}

QTabBar::tab:hover {
    background: #E2E8F0;
    color: #1E293B;
}

QTabBar::tab:selected {
    background: #2563EB;
    color: #FFFFFF;
    border-color: #2563EB;
}

QFrame#cardFrame {
    background-color: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 10px;
    padding: 12px;
}

QLineEdit, QComboBox, QDateEdit {
    background-color: #FFFFFF;
    color: #0F172A;
    border: 1px solid #CBD5E1;
    border-radius: 8px;
    padding: 7px 12px;
    font-size: 13px;
    min-height: 22px;
}

QLineEdit:hover, QComboBox:hover, QDateEdit:hover {
    border-color: #94A3B8;
}

QLineEdit:focus, QComboBox:focus, QDateEdit:focus {
    border: 2px solid #2563EB;
    padding: 6px 11px;
}

QPushButton {
    background-color: #2563EB;
    color: #FFFFFF;
    font-weight: 600;
    font-size: 13px;
    border-radius: 8px;
    padding: 8px 18px;
    border: none;
}

QPushButton:hover {
    background-color: #1D4ED8;
}

QPushButton#exportBtn {
    background-color: #059669;
}
QPushButton#exportBtn:hover {
    background-color: #047857;
}

QPushButton#importBtn {
    background-color: #4F46E5;
}
QPushButton#importBtn:hover {
    background-color: #4338CA;
}

QPushButton#todayBtn {
    background-color: #F1F5F9;
    color: #334155;
    border: 1px solid #CBD5E1;
}

QPushButton#deleteBtn {
    background-color: #FEE2E2;
    color: #DC2626;
    border: 1px solid #FECACA;
    border-radius: 6px;
    padding: 4px 10px;
    font-size: 12px;
    font-weight: 600;
}
QPushButton#deleteBtn:hover {
    background-color: #DC2626;
    color: #FFFFFF;
}

QTableWidget {
    background-color: #FFFFFF;
    color: #0F172A;
    gridline-color: #F1F5F9;
    border: 1px solid #E2E8F0;
    border-radius: 10px;
    font-size: 13px;
    selection-background-color: #EFF6FF;
    selection-color: #1E40AF;
    alternate-background-color: #F8FAFC;
}

QHeaderView::section {
    background-color: #F8FAFC;
    color: #475569;
    font-weight: 700;
    font-size: 12px;
    border: none;
    border-bottom: 2px solid #E2E8F0;
    padding: 8px;
}
"""


class CadetDialog(QDialog):
    def __init__(self, parent=None, cadet=None, available_groups=None):
        super().__init__(parent)
        self.cadet = cadet
        self.setWindowTitle("Редактировать курсанта" if cadet else "Добавить курсанта")
        self.resize(460, 340)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(14)

        form = QFormLayout()
        form.setSpacing(10)

        self.edit_rank = QLineEdit(cadet['rank'] if cadet else "курсант")
        self.edit_fio = QLineEdit(cadet['fio'] if cadet else "")
        self.edit_fio.setPlaceholderText("Фамилия Имя Отчество")

        self.edit_group = QComboBox()
        self.edit_group.setEditable(True)
        if available_groups:
            self.edit_group.addItems([str(g) for g in available_groups])
        if cadet:
            self.edit_group.setCurrentText(str(cadet['group_num']))

        self.edit_course = QComboBox()
        self.edit_course.addItems(["1 курс", "2 курс", "3 курс", "4 курс", "5 курс"])
        curr_course = cadet.get('course', 1) if cadet else 1
        self.edit_course.setCurrentIndex(max(0, min(4, curr_course - 1)))

        self.edit_gender = QComboBox()
        self.edit_gender.addItems(["М (Мужской)", "Ж (Женский)"])
        if cadet:
            g = normalize_gender(cadet.get('gender', 'male'))
            self.edit_gender.setCurrentIndex(1 if g == 'female' else 0)

        form.addRow("Воинское звание:", self.edit_rank)
        form.addRow("Полное ФИО:", self.edit_fio)
        form.addRow("Подразделение (Группа):", self.edit_group)
        form.addRow("Курс обучения:", self.edit_course)
        form.addRow("Пол:", self.edit_gender)
        layout.addLayout(form)

        btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btn_box.button(QDialogButtonBox.Ok).setText("Сохранить")
        btn_box.button(QDialogButtonBox.Cancel).setText("Отмена")
        btn_box.accepted.connect(self.validate_and_accept)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)

    def validate_and_accept(self):
        fio = self.edit_fio.text().strip()
        grp = self.edit_group.currentText().strip()
        if not fio:
            QMessageBox.warning(self, "Внимание", "Пожалуйста, введите ФИО курсанта.")
            self.edit_fio.setFocus()
            return
        if not grp:
            QMessageBox.warning(self, "Внимание", "Пожалуйста, укажите группу курсанта.")
            self.edit_group.setFocus()
            return
        self.accept()

    def get_data(self):
        return {
            'rank': self.edit_rank.text().strip() or "курсант",
            'fio': self.edit_fio.text().strip(),
            'group_num': self.edit_group.currentText().strip(),
            'course': self.edit_course.currentIndex() + 1,
            'gender': "Ж" if self.edit_gender.currentIndex() == 1 else "М",
            'status': "active"
        }


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ConsumABLE — Управление строевой запиской")
        self.resize(1200, 840)
        self.db = Database()

        icon_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "icon.png")
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
        main_layout.setContentsMargins(14, 14, 14, 14)
        main_layout.setSpacing(10)

        self.tabs = QTabWidget()
        self.tabs.setIconSize(QSize(18, 18))

        icon_today = create_svg_icon(SVG_CLIPBOARD, "#2563EB")
        icon_cadets = create_svg_icon(SVG_USERS, "#2563EB")
        icon_settings = create_svg_icon(SVG_SETTINGS, "#2563EB")

        self.tabs.addTab(self.create_today_tab(), icon_today, "Расход на сегодня")
        self.tabs.addTab(self.create_cadets_tab(), icon_cadets, "Личный состав")
        self.tabs.addTab(self.create_settings_tab(), icon_settings, "Справочник и настройки")
        main_layout.addWidget(self.tabs)

    def create_today_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(12)

        # Top Bar
        top_bar = QHBoxLayout()
        top_bar.setSpacing(10)

        lbl_date = QLabel("Дата расхода:")
        lbl_date.setStyleSheet("font-weight: 700; font-size: 13px; color: #334155;")
        top_bar.addWidget(lbl_date)

        self.date_edit = QDateEdit(QDate.currentDate())
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDisplayFormat("dd.MM.yyyy")
        self.date_edit.setMinimumWidth(135)
        self.date_edit.setFixedHeight(36)
        self.date_edit.dateChanged.connect(self.load_daily_data)
        top_bar.addWidget(self.date_edit)

        btn_today = QPushButton("Сегодня")
        btn_today.setObjectName("todayBtn")
        btn_today.setFixedHeight(36)
        btn_today.setIcon(create_svg_icon(SVG_CALENDAR, "#475569", 14))
        btn_today.clicked.connect(lambda: self.date_edit.setDate(QDate.currentDate()))
        top_bar.addWidget(btn_today)

        top_bar.addStretch()

        btn_export = QPushButton("Экспорт в Excel (3 среза + формулы)")
        btn_export.setObjectName("exportBtn")
        btn_export.setFixedHeight(36)
        btn_export.setIcon(create_svg_icon(SVG_EXPORT, "#FFFFFF", 16))
        btn_export.clicked.connect(self.export_excel)
        top_bar.addWidget(btn_export)

        layout.addLayout(top_bar)

        # Быстрое добавление в расход
        add_box = QFrame()
        add_box.setObjectName("cardFrame")
        add_layout = QHBoxLayout(add_box)
        add_layout.setContentsMargins(10, 10, 10, 10)
        add_layout.setSpacing(10)

        self.cb_cadets = QComboBox()
        self.cb_cadets.setEditable(True)
        self.cb_cadets.setInsertPolicy(QComboBox.NoInsert)
        if self.cb_cadets.completer():
            self.cb_cadets.completer().setFilterMode(Qt.MatchContains)
        self.cb_cadets.setPlaceholderText("Фамилия курсанта...")
        if self.cb_cadets.lineEdit():
            self.cb_cadets.lineEdit().returnPressed.connect(self.on_cadet_return_pressed)
        add_layout.addWidget(self.cb_cadets, 4)

        self.cb_types = QComboBox()
        self.cb_types.setEditable(True)
        self.cb_types.setInsertPolicy(QComboBox.NoInsert)
        if self.cb_types.completer():
            self.cb_types.completer().setFilterMode(Qt.MatchContains)
        self.cb_types.setPlaceholderText("Причина (наряд, увольнение...)...")
        if self.cb_types.lineEdit():
            self.cb_types.lineEdit().returnPressed.connect(self.on_type_return_pressed)
        add_layout.addWidget(self.cb_types, 3)

        self.edit_note = QLineEdit()
        self.edit_note.setPlaceholderText("Примечание (город, корпус, задача)...")
        self.edit_note.returnPressed.connect(self.add_absence_record)
        add_layout.addWidget(self.edit_note, 3)

        btn_add = QPushButton("Добавить в расход")
        btn_add.setFixedHeight(36)
        btn_add.setIcon(create_svg_icon(SVG_PLUS, "#FFFFFF", 14))
        btn_add.clicked.connect(self.add_absence_record)
        add_layout.addWidget(btn_add, 2)

        layout.addWidget(add_box)

        # -------------------------------------------------------------
        # 3 НЕЗАВИСИМЫХ СРЕЗА РАСХОДА (QTabWidget)
        # -------------------------------------------------------------
        self.slice_tabs = QTabWidget()
        self.slice_tabs.setStyleSheet("""
            QTabBar::tab {
                padding: 8px 18px;
                font-size: 13px;
                font-weight: 700;
            }
        """)

        # Срез 1: Общий расход
        self.tab_slice_total = self.create_single_slice_tab("total")
        # Срез 2: Мужской состав
        self.tab_slice_male = self.create_single_slice_tab("male")
        # Срез 3: Женский состав
        self.tab_slice_female = self.create_single_slice_tab("female")

        self.slice_tabs.addTab(self.tab_slice_total, "👥 Общий расход")
        self.slice_tabs.addTab(self.tab_slice_male, "👨 Мужской состав")
        self.slice_tabs.addTab(self.tab_slice_female, "👩 Женский состав")
        layout.addWidget(self.slice_tabs)

        # Контрольная плашка математического баланса
        self.balance_bar = QLabel("")
        self.balance_bar.setStyleSheet(
            "background: #ECFDF5; color: #065F46; border: 1px solid #A7F3D0; "
            "border-radius: 8px; padding: 10px 14px; font-weight: 700; font-size: 12px;"
        )
        layout.addWidget(self.balance_bar)

        self.load_types_combo()
        self.load_cadets_combo()
        self.load_daily_data()
        return tab

    def create_single_slice_tab(self, slice_key: str) -> QWidget:
        w = QWidget()
        l = QVBoxLayout(w)
        l.setContentsMargins(8, 8, 8, 8)
        l.setSpacing(10)

        # KPI Badges
        kpi_layout = QHBoxLayout()
        kpi_layout.setSpacing(8)

        lbl_ps = QLabel("По списку: 0")
        lbl_ps.setStyleSheet("background: #F1F5F9; color: #1E293B; border: 1px solid #CBD5E1; border-radius: 6px; padding: 5px 12px; font-weight: 700;")
        lbl_nl = QLabel("В строю: 0")
        lbl_nl.setStyleSheet("background: #ECFDF5; color: #065F46; border: 1px solid #A7F3D0; border-radius: 6px; padding: 5px 12px; font-weight: 700;")
        lbl_abs = QLabel("В расходе: 0")
        lbl_abs.setStyleSheet("background: #FFF1F2; color: #9F1239; border: 1px solid #FECDD3; border-radius: 6px; padding: 5px 12px; font-weight: 700;")
        lbl_cats = QLabel("Категории: —")
        lbl_cats.setStyleSheet("background: #EFF6FF; color: #1E40AF; border: 1px solid #BFDBFE; border-radius: 6px; padding: 5px 12px; font-size: 12px;")

        kpi_layout.addWidget(lbl_ps)
        kpi_layout.addWidget(lbl_nl)
        kpi_layout.addWidget(lbl_abs)
        kpi_layout.addWidget(lbl_cats)
        kpi_layout.addStretch()
        l.addLayout(kpi_layout)

        # Table
        tbl = QTableWidget()
        tbl.setColumnCount(8)
        tbl.setHorizontalHeaderLabels([
            "№", "Звание", "ФИО", "Группа", "Курс", "Причина (категория)", "Примечание", "Действие"
        ])
        tbl.setAlternatingRowColors(True)
        tbl.verticalHeader().setVisible(False)
        tbl.setSelectionBehavior(QAbstractItemView.SelectRows)
        tbl.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        tbl.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        tbl.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        tbl.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        tbl.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        tbl.horizontalHeader().setSectionResizeMode(7, QHeaderView.ResizeToContents)
        l.addWidget(tbl)

        setattr(self, f"badge_ps_{slice_key}", lbl_ps)
        setattr(self, f"badge_nl_{slice_key}", lbl_nl)
        setattr(self, f"badge_abs_{slice_key}", lbl_abs)
        setattr(self, f"badge_cats_{slice_key}", lbl_cats)
        setattr(self, f"table_absences_{slice_key}", tbl)

        return w

    def create_cadets_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(12)

        top_panel = QHBoxLayout()
        top_panel.setSpacing(10)

        btn_add_cadet = QPushButton("Добавить курсанта")
        btn_add_cadet.setFixedHeight(36)
        btn_add_cadet.setIcon(create_svg_icon(SVG_PLUS, "#FFFFFF", 14))
        btn_add_cadet.clicked.connect(self.open_add_cadet_dialog)
        top_panel.addWidget(btn_add_cadet)

        btn_import = QPushButton("Импорт базы из Excel (Staging Wizard)")
        btn_import.setObjectName("importBtn")
        btn_import.setFixedHeight(36)
        btn_import.setIcon(create_svg_icon(SVG_IMPORT, "#FFFFFF", 16))
        btn_import.clicked.connect(self.import_cadets_excel)
        top_panel.addWidget(btn_import)

        # Фильтр по курсу
        self.combo_course_filter = QComboBox()
        self.combo_course_filter.setFixedHeight(36)
        self.combo_course_filter.addItems(["Все курсы", "1 курс", "2 курс", "3 курс", "4 курс", "5 курс"])
        self.combo_course_filter.currentIndexChanged.connect(self.load_cadets_table)
        top_panel.addWidget(self.combo_course_filter)

        # Фильтр по полу
        self.combo_gender_filter = QComboBox()
        self.combo_gender_filter.setFixedHeight(36)
        self.combo_gender_filter.addItems(["Все (М и Ж)", "Только мужчины", "Только женщины"])
        self.combo_gender_filter.currentIndexChanged.connect(self.load_cadets_table)
        top_panel.addWidget(self.combo_gender_filter)

        # Выпадающий фильтр по учебным группам
        self.combo_group_filter = QComboBox()
        self.combo_group_filter.setFixedHeight(36)
        self.combo_group_filter.currentIndexChanged.connect(self.load_cadets_table)
        top_panel.addWidget(self.combo_group_filter)

        self.edit_search = QLineEdit()
        self.edit_search.setFixedHeight(36)
        self.edit_search.setPlaceholderText("Поиск по фамилии или званию...")
        self.edit_search.textChanged.connect(self.load_cadets_table)
        top_panel.addWidget(self.edit_search)

        layout.addLayout(top_panel)

        self.table_cadets = QTableWidget()
        self.table_cadets.setColumnCount(8)
        self.table_cadets.setHorizontalHeaderLabels([
            "№", "Звание", "ФИО", "Краткое ФИО", "Группа", "Курс", "Пол", "Действия"
        ])
        self.table_cadets.setAlternatingRowColors(True)
        self.table_cadets.verticalHeader().setVisible(False)
        self.table_cadets.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table_cadets.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table_cadets.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table_cadets.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table_cadets.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.table_cadets.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)
        self.table_cadets.horizontalHeader().setSectionResizeMode(6, QHeaderView.ResizeToContents)
        self.table_cadets.horizontalHeader().setSectionResizeMode(7, QHeaderView.ResizeToContents)

        layout.addWidget(self.table_cadets)

        self.update_group_filter_options()
        self.load_cadets_table()
        return tab

    def create_settings_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(14)

        card_signatures = QFrame()
        card_signatures.setObjectName("cardFrame")
        sig_layout = QVBoxLayout(card_signatures)

        lbl_sig_title = QLabel("Шапка и подписи строевой записки")
        lbl_sig_title.setStyleSheet("font-size: 14px; font-weight: 700; color: #1E293B;")
        sig_layout.addWidget(lbl_sig_title)

        form_sig = QFormLayout()
        self.edit_unit = QLineEdit(self.db.get_setting("unit_title", "3 курс 1 факультета"))
        self.edit_chief_title = QLineEdit(self.db.get_setting("chief_title", "Начальник 3 курса 1 факультета"))
        self.edit_chief_name = QLineEdit(self.db.get_setting("chief_name", "майор                          В.Медков"))

        form_sig.addRow("Наименование подразделения:", self.edit_unit)
        form_sig.addRow("Должность начальника:", self.edit_chief_title)
        form_sig.addRow("Воинское звание и фамилия:", self.edit_chief_name)
        sig_layout.addLayout(form_sig)

        btn_save_sig = QPushButton("Сохранить параметры")
        btn_save_sig.setIcon(create_svg_icon(SVG_SAVE, "#FFFFFF", 14))
        btn_save_sig.clicked.connect(self.save_signatures)
        sig_layout.addWidget(btn_save_sig, alignment=Qt.AlignRight)

        layout.addWidget(card_signatures)
        layout.addStretch()
        return tab

    def save_signatures(self):
        self.db.set_setting("unit_title", self.edit_unit.text().strip())
        self.db.set_setting("chief_title", self.edit_chief_title.text().strip())
        self.db.set_setting("chief_name", self.edit_chief_name.text().strip())
        QMessageBox.information(self, "Успех", "Параметры строевой записки сохранены.")

    def get_selected_date_str(self):
        return self.date_edit.date().toString("yyyy-MM-dd")

    def load_types_combo(self):
        self.cb_types.clear()
        types = self.db.get_absence_types()
        for t in types:
            self.cb_types.addItem(f"{t['name']} ({t['category']})", userData=t['id'])

    def load_cadets_combo(self):
        self.cb_cadets.clear()
        cadets = self.db.get_all_cadets()
        for c in cadets:
            self.cb_cadets.addItem(f"{c['short_fio']} | {c['rank']} (гр. {c['group_num']}, {c.get('course', 1)} к.)", userData=c['id'])

    def load_daily_data(self):
        d_str = self.get_selected_date_str()
        absences = self.db.get_absences_for_date(d_str)
        cadets = self.db.get_all_cadets()

        # Compute 3-slice balance
        categories = self.db.get_distinct_categories()
        balance = compute_tri_slice_balance(cadets, absences, categories)

        slices_data = [
            ("total", balance.total, absences),
            ("male", balance.male, [a for a in absences if normalize_gender(a.get('gender')) == 'male']),
            ("female", balance.female, [a for a in absences if normalize_gender(a.get('gender')) == 'female'])
        ]

        icon_del = create_svg_icon(SVG_TRASH, "#DC2626", 12)

        for key, s_bal, s_absences in slices_data:
            badge_ps = getattr(self, f"badge_ps_{key}", None)
            badge_nl = getattr(self, f"badge_nl_{key}", None)
            badge_abs = getattr(self, f"badge_abs_{key}", None)
            badge_cats = getattr(self, f"badge_cats_{key}", None)
            tbl = getattr(self, f"table_absences_{key}", None)

            if badge_ps: badge_ps.setText(f"📋 По списку: {s_bal.listed_count}")
            if badge_nl: badge_nl.setText(f"✅ В строю: {s_bal.present_count}")
            if badge_abs: badge_abs.setText(f"⚠️ В расходе: {s_bal.absent_count}")

            if badge_cats:
                breakdown = [f"{k}: {v}" for k, v in s_bal.categories_breakdown.items() if v > 0]
                badge_cats.setText(" | ".join(breakdown) if breakdown else "Все в строю")

            if tbl:
                tbl.setRowCount(0)
                for idx, a in enumerate(s_absences, 1):
                    r = tbl.rowCount()
                    tbl.insertRow(r)
                    tbl.setRowHeight(r, 36)

                    tbl.setItem(r, 0, QTableWidgetItem(str(idx)))
                    tbl.setItem(r, 1, QTableWidgetItem(a.get('rank', 'курсант')))
                    tbl.setItem(r, 2, QTableWidgetItem(a.get('short_fio') or a.get('fio', '')))
                    tbl.setItem(r, 3, QTableWidgetItem(str(a.get('group_num', ''))))
                    tbl.setItem(r, 4, QTableWidgetItem(str(a.get('course', 1))))
                    tbl.setItem(r, 5, QTableWidgetItem(f"{a.get('type_name', '')} ({a.get('category', '')})"))
                    tbl.setItem(r, 6, QTableWidgetItem(a.get('note') or ""))

                    btn_del = QPushButton("Удалить")
                    btn_del.setObjectName("deleteBtn")
                    btn_del.setIcon(icon_del)
                    abs_id = a['id']
                    btn_del.clicked.connect(lambda _, aid=abs_id: self.delete_absence(aid))
                    tbl.setCellWidget(r, 7, btn_del)

        # Контроль баланса
        if balance.is_sum_balanced:
            self.balance_bar.setStyleSheet(
                "background: #ECFDF5; color: #065F46; border: 1px solid #A7F3D0; "
                "border-radius: 8px; padding: 10px 14px; font-weight: 700; font-size: 12px;"
            )
            self.balance_bar.setText(
                f"⚖️ Сквозной баланс строя подтвержден: "
                f"Общий [П/С: {balance.total.listed_count}, В строю: {balance.total.present_count}, Расход: {balance.total.absent_count}] "
                f"= Мужчины [{balance.male.listed_count}/{balance.male.present_count}/{balance.male.absent_count}] "
                f"+ Женщины [{balance.female.listed_count}/{balance.female.present_count}/{balance.female.absent_count}] ✓"
            )
        else:
            self.balance_bar.setStyleSheet(
                "background: #FEF2F2; color: #991B1B; border: 1px solid #FECACA; "
                "border-radius: 8px; padding: 10px 14px; font-weight: 700; font-size: 12px;"
            )
            self.balance_bar.setText(f"❌ НАРУШЕНИЕ БАЛАНСА: {balance.error_message}")

    def on_cadet_return_pressed(self):
        text = self.cb_cadets.currentText().strip().lower()
        if text:
            for i in range(self.cb_cadets.count()):
                if text in self.cb_cadets.itemText(i).lower():
                    self.cb_cadets.setCurrentIndex(i)
                    break
        self.cb_types.setFocus()

    def on_type_return_pressed(self):
        text = self.cb_types.currentText().strip().lower()
        if text:
            for i in range(self.cb_types.count()):
                if text in self.cb_types.itemText(i).lower():
                    self.cb_types.setCurrentIndex(i)
                    break
        self.edit_note.setFocus()

    def add_absence_record(self):
        cadet_id = self.cb_cadets.currentData()
        type_id = self.cb_types.currentData()
        note = self.edit_note.text().strip()

        if not cadet_id:
            text = self.cb_cadets.currentText().strip().lower()
            if text:
                for i in range(self.cb_cadets.count()):
                    if text in self.cb_cadets.itemText(i).lower():
                        cadet_id = self.cb_cadets.itemData(i)
                        break

        if not type_id:
            text = self.cb_types.currentText().strip().lower()
            if text:
                for i in range(self.cb_types.count()):
                    if text in self.cb_types.itemText(i).lower():
                        type_id = self.cb_types.itemData(i)
                        break

        if not type_id and self.cb_types.currentText().strip():
            raw_text = self.cb_types.currentText().strip()
            low = raw_text.lower()
            cat = "Наряд"
            if "увольн" in low or "отпуск" in low:
                cat = "Увольнение"
            elif "госпитал" in low:
                cat = "Госпиталь"
            elif "лазарет" in low or "больнич" in low or "медпункт" in low:
                cat = "Лазарет"
            elif "командир" in low:
                cat = "Командировка"
            elif "похорон" in low:
                cat = "Похороны"
            type_id = self.db.add_absence_type(raw_text, raw_text[:2].upper(), cat)
            self.load_types_combo()

        if not cadet_id or not type_id:
            QMessageBox.warning(self, "Внимание", "Пожалуйста, введите фамилию курсанта и причину отсутствия.")
            return

        d_str = self.get_selected_date_str()
        self.db.set_absence(d_str, cadet_id, type_id, note)
        self.edit_note.clear()
        self.load_daily_data()
        self.cb_cadets.setFocus()

    def delete_absence(self, abs_id):
        self.db.remove_absence(abs_id)
        self.load_daily_data()

    def export_excel(self):
        d_str = self.get_selected_date_str()
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить строевую записку", f"Строевая_записка_{d_str}.xlsx", "Книга Excel (*.xlsx)"
        )
        if file_path:
            try:
                cadets = self.db.get_all_cadets()
                absences = self.db.get_absences_for_date(d_str)
                export_report_to_excel(
                    output_path=file_path,
                    report_date=d_str,
                    cadets=cadets,
                    absences=absences,
                    unit_title=self.db.get_setting("unit_title", "3 курс 1 факультета"),
                    chief_title=self.db.get_setting("chief_title", "Начальник 3 курса 1 факультета"),
                    chief_name=self.db.get_setting("chief_name", "майор                          В.Медков"),
                    active_categories=self.db.get_distinct_categories()
                )
                QMessageBox.information(
                    self, "Успех",
                    f"Строевая записка успешно сформирована с 3 срезами и рабочими формулами:\n{file_path}"
                )
            except Exception as e:
                QMessageBox.critical(self, "Ошибка экспорта", f"Не удалось сформировать файл:\n{e}")

    def import_cadets_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Выберите файл с базой курса", "", "Книга Excel (*.xlsx *.xlsm);;Все файлы (*.*)"
        )
        if not file_path:
            return

        try:
            staging_cadets = parse_roster_excel(file_path)
            if not staging_cadets:
                QMessageBox.warning(
                    self, "Внимание",
                    "В файле не найдены данные курсантов.\nУбедитесь, что таблица содержит ФИО и группы."
                )
                return

            # Открываем Staging Wizard
            wizard = ImportReviewDialog(self, staging_cadets=staging_cadets)
            if wizard.exec_() == QDialog.Accepted:
                final_cadets = wizard.get_staged_cadets()
                self.db.clear_and_import_cadets(final_cadets)
                self.update_group_filter_options()
                self.load_cadets_table()
                self.load_cadets_combo()
                self.load_daily_data()
                QMessageBox.information(
                    self, "Успех импорта",
                    f"База личного состава успешно обновлена!\nЗагружено курсантов: {len(final_cadets)} чел."
                )
        except Exception as e:
            QMessageBox.critical(self, "Ошибка чтения файла", f"{e}")

    def update_group_filter_options(self):
        current = self.combo_group_filter.currentText()
        self.combo_group_filter.blockSignals(True)
        self.combo_group_filter.clear()
        self.combo_group_filter.addItem("Все подразделения")
        for g in self.db.get_distinct_groups():
            self.combo_group_filter.addItem(str(g))
        idx = self.combo_group_filter.findText(current)
        if idx >= 0:
            self.combo_group_filter.setCurrentIndex(idx)
        self.combo_group_filter.blockSignals(False)

    def load_cadets_table(self):
        cadets = self.db.get_all_cadets(active_only=False)

        # Filter by course
        course_idx = self.combo_course_filter.currentIndex()
        if course_idx > 0:
            cadets = [c for c in cadets if c.get('course') == course_idx]

        # Filter by gender
        gender_idx = self.combo_gender_filter.currentIndex()
        if gender_idx == 1:
            cadets = [c for c in cadets if normalize_gender(c.get('gender')) == 'male']
        elif gender_idx == 2:
            cadets = [c for c in cadets if normalize_gender(c.get('gender')) == 'female']

        # Filter by group
        grp = self.combo_group_filter.currentText()
        if grp and grp != "Все подразделения":
            cadets = [c for c in cadets if str(c.get('group_num')) == grp]

        # Filter by search
        search_txt = self.edit_search.text().strip().lower()
        if search_txt:
            cadets = [c for c in cadets if search_txt in c.get('fio', '').lower() or search_txt in c.get('rank', '').lower()]

        self.table_cadets.setRowCount(0)
        icon_edit = create_svg_icon(SVG_EDIT, "#2563EB", 12)
        icon_del = create_svg_icon(SVG_TRASH, "#DC2626", 12)

        for idx, c in enumerate(cadets, 1):
            r = self.table_cadets.rowCount()
            self.table_cadets.insertRow(r)
            self.table_cadets.setRowHeight(r, 36)

            self.table_cadets.setItem(r, 0, QTableWidgetItem(str(idx)))
            self.table_cadets.setItem(r, 1, QTableWidgetItem(c.get('rank', 'курсант')))
            self.table_cadets.setItem(r, 2, QTableWidgetItem(c.get('fio', '')))
            self.table_cadets.setItem(r, 3, QTableWidgetItem(c.get('short_fio', '')))
            self.table_cadets.setItem(r, 4, QTableWidgetItem(str(c.get('group_num', ''))))
            self.table_cadets.setItem(r, 5, QTableWidgetItem(str(c.get('course', 1))))

            g_ru = gender_to_ru(c.get('gender', 'male'))
            g_item = QTableWidgetItem(g_ru)
            g_item.setTextAlignment(Qt.AlignCenter)
            g_item.setForeground(QColor("#BE185D" if g_ru == "Ж" else "#1D4ED8"))
            self.table_cadets.setItem(r, 6, g_item)

            act_widget = QWidget()
            act_layout = QHBoxLayout(act_widget)
            act_layout.setContentsMargins(4, 2, 4, 2)
            act_layout.setSpacing(6)

            btn_edit = QPushButton("Ред.")
            btn_edit.setStyleSheet("background-color: #EFF6FF; color: #2563EB; border: 1px solid #BFDBFE; border-radius: 6px; padding: 4px 8px; font-size: 12px;")
            btn_edit.setIcon(icon_edit)
            cid = c['id']
            btn_edit.clicked.connect(lambda _, cadet_id=cid: self.open_edit_cadet_dialog(cadet_id))
            act_layout.addWidget(btn_edit)

            btn_del = QPushButton("Удал.")
            btn_del.setObjectName("deleteBtn")
            btn_del.setIcon(icon_del)
            btn_del.clicked.connect(lambda _, cadet_id=cid: self.delete_cadet_action(cadet_id))
            act_layout.addWidget(btn_del)

            self.table_cadets.setCellWidget(r, 7, act_widget)

    def open_add_cadet_dialog(self):
        groups = self.db.get_distinct_groups()
        dlg = CadetDialog(self, available_groups=groups)
        if dlg.exec_() == QDialog.Accepted:
            data = dlg.get_data()
            self.db.add_cadet(
                rank=data['rank'],
                fio=data['fio'],
                group_num=data['group_num'],
                course=data['course'],
                gender=data['gender']
            )
            self.update_group_filter_options()
            self.load_cadets_table()
            self.load_cadets_combo()
            self.load_daily_data()

    def open_edit_cadet_dialog(self, cadet_id):
        cadet = self.db.get_cadet_by_id(cadet_id)
        if not cadet:
            return
        groups = self.db.get_distinct_groups()
        dlg = CadetDialog(self, cadet=cadet, available_groups=groups)
        if dlg.exec_() == QDialog.Accepted:
            data = dlg.get_data()
            self.db.update_cadet(
                cadet_id=cadet_id,
                rank=data['rank'],
                fio=data['fio'],
                group_num=data['group_num'],
                course=data['course'],
                gender=data['gender']
            )
            self.update_group_filter_options()
            self.load_cadets_table()
            self.load_cadets_combo()
            self.load_daily_data()

    def delete_cadet_action(self, cadet_id):
        c = self.db.get_cadet_by_id(cadet_id)
        if not c:
            return
        ans = QMessageBox.question(
            self, "Подтверждение удаления",
            f"Вы действительно хотите удалить курсанта:\n{c['fio']} (гр. {c['group_num']})?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if ans == QMessageBox.Yes:
            self.db.delete_cadet(cadet_id)
            self.update_group_filter_options()
            self.load_cadets_table()
            self.load_cadets_combo()
            self.load_daily_data()
