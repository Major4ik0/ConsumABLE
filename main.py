import sys
import os
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QTableWidget, QTableWidgetItem, QComboBox,
    QLineEdit, QDateEdit, QFileDialog, QMessageBox, QTabWidget,
    QHeaderView, QFrame, QAbstractItemView, QDialog, QFormLayout, QDialogButtonBox
)
from PyQt5.QtCore import Qt, QDate, QByteArray, QSize
from PyQt5.QtGui import QIcon, QPixmap, QPainter, QPalette, QColor
from PyQt5.QtSvg import QSvgRenderer

from database import Database
from excel_manager import import_cadets_from_excel, export_report_to_excel, make_short_fio, detect_gender

# Векторные SVG-иконки
SVG_CLIPBOARD = '''<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2"/><rect x="8" y="2" width="8" height="4" rx="1" ry="1"/><path d="M9 12h6"/><path d="M9 16h6"/></svg>'''
SVG_USERS = '''<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>'''
SVG_SETTINGS = '''<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z"/></svg>'''
SVG_EXPORT = '''<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>'''
SVG_IMPORT = '''<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>'''
SVG_PLUS = '''<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>'''
SVG_TRASH = '''<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>'''
SVG_SAVE = '''<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>'''
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

QDateEdit {
    min-width: 130px;
    font-weight: 600;
}

QDateEdit::drop-down, QComboBox::drop-down {
    subcontrol-origin: padding;
    subcontrol-position: top right;
    width: 28px;
    border-left: 1px solid #E2E8F0;
    border-top-right-radius: 8px;
    border-bottom-right-radius: 8px;
    background: #F8FAFC;
}

QComboBox QAbstractItemView {
    background-color: #FFFFFF;
    color: #0F172A;
    border: 1px solid #CBD5E1;
    border-radius: 8px;
    selection-background-color: #EFF6FF;
    selection-color: #1E40AF;
    padding: 4px;
}

QPushButton {
    background-color: #2563EB;
    color: #FFFFFF;
    font-weight: 600;
    font-size: 13px;
    border-radius: 8px;
    padding: 8px 18px;
    border: none;
    min-height: 22px;
}

QPushButton:hover {
    background-color: #1D4ED8;
}

QPushButton:pressed {
    background-color: #1E40AF;
}

QPushButton#exportBtn {
    background-color: #059669;
    font-size: 13px;
    padding: 9px 20px;
}

QPushButton#exportBtn:hover {
    background-color: #047857;
}

QPushButton#importBtn {
    background-color: #4F46E5;
    font-size: 13px;
    padding: 9px 20px;
}

QPushButton#importBtn:hover {
    background-color: #4338CA;
}

QPushButton#todayBtn {
    background-color: #F1F5F9;
    color: #334155;
    border: 1px solid #CBD5E1;
    padding: 6px 14px;
    font-size: 12px;
    font-weight: 600;
}

QPushButton#todayBtn:hover {
    background-color: #E2E8F0;
    color: #0F172A;
}

QPushButton#editBtn {
    background-color: #EFF6FF;
    color: #2563EB;
    border: 1px solid #BFDBFE;
    border-radius: 6px;
    padding: 4px 10px;
    font-size: 12px;
    font-weight: 600;
}

QPushButton#editBtn:hover {
    background-color: #2563EB;
    color: #FFFFFF;
    border-color: #2563EB;
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
    border-color: #DC2626;
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

QTableWidget::item {
    padding: 6px 10px;
    border-bottom: 1px solid #F1F5F9;
}

QHeaderView::section {
    background-color: #F8FAFC;
    color: #475569;
    font-weight: 700;
    font-size: 12px;
    padding: 9px 10px;
    border: none;
    border-bottom: 2px solid #E2E8F0;
    border-right: 1px solid #F1F5F9;
}

QScrollBar:vertical {
    border: none;
    background: #F1F5F9;
    width: 8px;
    border-radius: 4px;
    margin: 0px;
}

QScrollBar::handle:vertical {
    background: #CBD5E1;
    min-height: 25px;
    border-radius: 4px;
}

QScrollBar::handle:vertical:hover {
    background: #94A3B8;
}

QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}

QCalendarWidget QWidget#qt_calendar_navigationbar {
    background-color: #F8FAFC;
    border-bottom: 1px solid #E2E8F0;
}

QCalendarWidget QToolButton {
    color: #0F172A;
    background-color: transparent;
    border: none;
    font-weight: 600;
    font-size: 13px;
    padding: 4px 8px;
}

QCalendarWidget QToolButton:hover {
    background-color: #E2E8F0;
    border-radius: 4px;
}

QCalendarWidget QTableView {
    background-color: #FFFFFF;
    color: #0F172A;
    selection-background-color: #2563EB;
    selection-color: #FFFFFF;
    alternate-background-color: #FFFFFF;
}
"""

class CadetDialog(QDialog):
    def __init__(self, parent=None, cadet_data=None, available_groups=None):
        super().__init__(parent)
        self.cadet_data = cadet_data
        self.setWindowTitle("Редактировать курсанта" if cadet_data else "Добавить курсанта")
        self.setFixedWidth(460)
        self.init_ui(available_groups or [])

    def init_ui(self, available_groups):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(14)

        form = QFormLayout()
        form.setSpacing(10)

        self.edit_rank = QComboBox()
        self.edit_rank.setEditable(True)
        self.edit_rank.addItems(["курсант", "ряд.", "ефр.", "мл. с-т.", "с-т.", "ст. с-т.", "пр-к."])

        self.edit_fio = QLineEdit()
        self.edit_fio.setPlaceholderText("Например: Иванов Иван Иванович")

        self.edit_short_fio = QLineEdit()
        self.edit_short_fio.setPlaceholderText("Например: Иванов И.И.")

        self.edit_group = QComboBox()
        self.edit_group.setEditable(True)
        if available_groups:
            self.edit_group.addItems([str(g) for g in available_groups])
        self.edit_group.setPlaceholderText("Например: 1280")

        self.combo_gender = QComboBox()
        self.combo_gender.addItems(["М", "Ж"])

        # Автоматическая подсказка инициалов при вводе ФИО
        self.edit_fio.textChanged.connect(self._on_fio_changed)

        form.addRow("Воинское звание:", self.edit_rank)
        form.addRow("ФИО полностью:", self.edit_fio)
        form.addRow("Фамилия и инициалы:", self.edit_short_fio)
        form.addRow("Группа / Взвод:", self.edit_group)
        form.addRow("Пол:", self.combo_gender)
        layout.addLayout(form)

        # Заполнение текущими данными при редактировании
        if self.cadet_data:
            self.edit_rank.setCurrentText(str(self.cadet_data.get('rank', 'курсант')))
            self.edit_fio.setText(str(self.cadet_data.get('fio', '')))
            self.edit_short_fio.setText(str(self.cadet_data.get('short_fio', '')))
            self.edit_group.setCurrentText(str(self.cadet_data.get('group_num', '')))
            gender_val = str(self.cadet_data.get('gender', 'М')).upper()
            self.combo_gender.setCurrentText("Ж" if gender_val == "Ж" else "М")

        btn_box = QHBoxLayout()
        btn_box.addStretch()

        btn_cancel = QPushButton("Отмена")
        btn_cancel.setObjectName("todayBtn")
        btn_cancel.clicked.connect(self.reject)
        btn_box.addWidget(btn_cancel)

        btn_save = QPushButton("Сохранить")
        btn_save.clicked.connect(self._on_save)
        btn_box.addWidget(btn_save)

        layout.addLayout(btn_box)

    def _on_fio_changed(self, text):
        if not self.cadet_data or not self.edit_short_fio.text():
            self.edit_short_fio.setText(make_short_fio(text))
            g = detect_gender(text)
            self.combo_gender.setCurrentText(g)

    def _on_save(self):
        fio = self.edit_fio.text().strip()
        grp = self.edit_group.currentText().strip()
        if not fio:
            QMessageBox.warning(self, "Внимание", "Пожалуйста, укажите ФИО курсанта.")
            self.edit_fio.setFocus()
            return
        if not grp:
            QMessageBox.warning(self, "Внимание", "Пожалуйста, укажите группу или подразделение.")
            self.edit_group.setFocus()
            return
        self.accept()

    def get_data(self):
        fio = self.edit_fio.text().strip()
        short_fio = self.edit_short_fio.text().strip() or make_short_fio(fio)
        return {
            'rank': self.edit_rank.currentText().strip() or "курсант",
            'fio': fio,
            'short_fio': short_fio,
            'group_num': self.edit_group.currentText().strip(),
            'gender': self.combo_gender.currentText().strip()
        }

class AbsenceTypeDialog(QDialog):
    def __init__(self, parent=None, type_data=None, available_categories=None):
        super().__init__(parent)
        self.type_data = type_data
        self.setWindowTitle("Редактировать причину отсутствия" if type_data else "Добавить причину")
        self.setFixedWidth(460)
        self.init_ui(available_categories or [])

    def init_ui(self, available_categories):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(14)

        form = QFormLayout()
        form.setSpacing(10)

        self.edit_name = QLineEdit()
        self.edit_name.setPlaceholderText("Например: Дежурный по КПП или Увольнение")

        self.edit_code = QLineEdit()
        self.edit_code.setPlaceholderText("Например: ДпКПП или УВ")

        self.combo_cat = QComboBox()
        self.combo_cat.setEditable(True)
        default_cats = ["Наряд", "Увольнение", "Госпиталь", "Лазарет", "Командировка", "Отпуск", "Похороны", "Карантин"]
        all_cats = list(dict.fromkeys(default_cats + available_categories))
        self.combo_cat.addItems(all_cats)
        self.combo_cat.setPlaceholderText("Выберите или введите категорию в сводке...")

        form.addRow("Название / Наряд:", self.edit_name)
        form.addRow("Шифр (сокращение):", self.edit_code)
        form.addRow("Категория в сводке:", self.combo_cat)
        layout.addLayout(form)

        if self.type_data:
            self.edit_name.setText(str(self.type_data.get('name', '')))
            self.edit_code.setText(str(self.type_data.get('code', '')))
            self.combo_cat.setCurrentText(str(self.type_data.get('category', 'Наряд')))

        btn_box = QHBoxLayout()
        btn_box.addStretch()

        btn_cancel = QPushButton("Отмена")
        btn_cancel.setObjectName("todayBtn")
        btn_cancel.clicked.connect(self.reject)
        btn_box.addWidget(btn_cancel)

        btn_save = QPushButton("Сохранить")
        btn_save.clicked.connect(self._on_save)
        btn_box.addWidget(btn_save)

        layout.addLayout(btn_box)

    def _on_save(self):
        name = self.edit_name.text().strip()
        cat = self.combo_cat.currentText().strip()
        if not name:
            QMessageBox.warning(self, "Внимание", "Пожалуйста, укажите название причины.")
            self.edit_name.setFocus()
            return
        if not cat:
            QMessageBox.warning(self, "Внимание", "Пожалуйста, укажите категорию в сводке.")
            self.combo_cat.setFocus()
            return
        self.accept()

    def get_data(self):
        return {
            'name': self.edit_name.text().strip(),
            'code': self.edit_code.text().strip(),
            'category': self.combo_cat.currentText().strip()
        }

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ConsumABLE — Управление строевой запиской")
        self.resize(1180, 800)
        self.db = Database()

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

        top_bar.addSpacing(10)

        # KPI-карточки сводки
        self.badge_total = QLabel("По списку: 0")
        self.badge_total.setStyleSheet(
            "background: #F1F5F9; color: #1E293B; border: 1px solid #CBD5E1; "
            "border-radius: 8px; padding: 6px 14px; font-weight: 700; font-size: 13px;"
        )
        self.badge_present = QLabel("В строю: 0")
        self.badge_present.setStyleSheet(
            "background: #ECFDF5; color: #065F46; border: 1px solid #A7F3D0; "
            "border-radius: 8px; padding: 6px 14px; font-weight: 700; font-size: 13px;"
        )
        self.badge_absent = QLabel("В расходе: 0")
        self.badge_absent.setStyleSheet(
            "background: #FFF1F2; color: #9F1239; border: 1px solid #FECDD3; "
            "border-radius: 8px; padding: 6px 14px; font-weight: 700; font-size: 13px;"
        )
        top_bar.addWidget(self.badge_total)
        top_bar.addWidget(self.badge_present)
        top_bar.addWidget(self.badge_absent)

        top_bar.addStretch()

        btn_export = QPushButton("Экспорт в Excel")
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
        self.cb_cadets.setPlaceholderText("Выберите курсанта...")
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

        # Таблица отсутствующих
        self.table_absences = QTableWidget()
        self.table_absences.setColumnCount(7)
        self.table_absences.setHorizontalHeaderLabels([
            "№", "Звание", "ФИО", "Группа", "Причина (наряд)", "Примечание", "Действие"
        ])
        self.table_absences.setAlternatingRowColors(True)
        self.table_absences.verticalHeader().setVisible(False)
        self.table_absences.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table_absences.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table_absences.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table_absences.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table_absences.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.table_absences.horizontalHeader().setSectionResizeMode(6, QHeaderView.ResizeToContents)

        layout.addWidget(self.table_absences)

        self.load_types_combo()
        self.load_cadets_combo()
        self.load_daily_data()
        return tab

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

        btn_import = QPushButton("Импорт базы из Excel")
        btn_import.setObjectName("importBtn")
        btn_import.setFixedHeight(36)
        btn_import.setIcon(create_svg_icon(SVG_IMPORT, "#FFFFFF", 16))
        btn_import.clicked.connect(self.import_cadets_excel)
        top_panel.addWidget(btn_import)

        # Выпадающий фильтр по учебным группам
        self.combo_group_filter = QComboBox()
        self.combo_group_filter.setFixedHeight(36)
        self.combo_group_filter.setMinimumWidth(150)
        self.combo_group_filter.addItem("Все группы")
        self.combo_group_filter.currentTextChanged.connect(self.filter_cadets_table)
        top_panel.addWidget(self.combo_group_filter)

        self.lbl_cadet_count = QLabel("Всего курсантов: 0")
        self.lbl_cadet_count.setStyleSheet(
            "background: #F1F5F9; color: #475569; border: 1px solid #CBD5E1; "
            "border-radius: 8px; padding: 6px 14px; font-weight: 600; font-size: 13px;"
        )
        top_panel.addWidget(self.lbl_cadet_count)

        top_panel.addStretch()

        self.search_cadet = QLineEdit()
        self.search_cadet.setPlaceholderText("🔍 Поиск по фамилии, званию или группе...")
        self.search_cadet.setMinimumWidth(260)
        self.search_cadet.setFixedHeight(36)
        self.search_cadet.textChanged.connect(self.filter_cadets_table)
        top_panel.addWidget(self.search_cadet)

        layout.addLayout(top_panel)

        self.table_cadets = QTableWidget()
        self.table_cadets.setColumnCount(7)
        self.table_cadets.setHorizontalHeaderLabels([
            "ID", "Воинское звание", "ФИО (полное)", "Инициалы", "Группа / Взвод", "Пол", "Действия"
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

        # Двойной клик по строке открывает окно редактирования
        self.table_cadets.cellDoubleClicked.connect(self._on_cadet_cell_double_clicked)

        layout.addWidget(self.table_cadets)

        self.load_cadets_table()
        return tab

    def create_settings_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(14)

        lbl_head = QLabel("Шапка и подписи строевой записки:")
        lbl_head.setStyleSheet("font-weight: 700; font-size: 14px; color: #1E293B;")
        layout.addWidget(lbl_head)

        head_box = QFrame()
        head_box.setObjectName("cardFrame")
        h_layout = QVBoxLayout(head_box)
        h_layout.setSpacing(10)

        h_layout.addWidget(QLabel("Наименование подразделения:"))
        self.edit_unit = QLineEdit(self.db.get_setting("unit_title"))
        self.edit_unit.setPlaceholderText("Например: 2 курс 1 факультета (АСУ)")
        h_layout.addWidget(self.edit_unit)

        r_chief = QHBoxLayout()
        r_chief.setSpacing(10)

        v1 = QVBoxLayout()
        v1.addWidget(QLabel("Должность начальника:"))
        self.edit_chief_title = QLineEdit(self.db.get_setting("chief_title"))
        self.edit_chief_title.setPlaceholderText("Например: Начальник 2 курса 1 факультета")
        v1.addWidget(self.edit_chief_title)
        r_chief.addLayout(v1, 3)

        v2 = QVBoxLayout()
        v2.addWidget(QLabel("Воинское звание и фамилия:"))
        self.edit_chief_name = QLineEdit(self.db.get_setting("chief_name"))
        self.edit_chief_name.setPlaceholderText("Например: майор В. Медков")
        v2.addWidget(self.edit_chief_name)
        r_chief.addLayout(v2, 3)

        btn_save_head = QPushButton("Сохранить реквизиты")
        btn_save_head.setFixedHeight(36)
        btn_save_head.setIcon(create_svg_icon(SVG_SAVE, "#FFFFFF", 14))
        btn_save_head.clicked.connect(self.save_header_settings)

        v3 = QVBoxLayout()
        v3.addWidget(QLabel(" "))
        v3.addWidget(btn_save_head)
        r_chief.addLayout(v3, 2)

        h_layout.addLayout(r_chief)
        layout.addWidget(head_box)

        lbl_types = QLabel("Справочник причин отсутствия и нарядов:")
        lbl_types.setStyleSheet("font-weight: 700; font-size: 14px; color: #1E293B;")
        layout.addWidget(lbl_types)

        add_type_box = QFrame()
        add_type_box.setObjectName("cardFrame")
        row1 = QHBoxLayout(add_type_box)
        row1.setSpacing(10)

        self.new_type_name = QLineEdit()
        self.new_type_name.setPlaceholderText("Название (например, Дежурный по КПП)")
        row1.addWidget(self.new_type_name, 3)

        self.new_type_code = QLineEdit()
        self.new_type_code.setPlaceholderText("Шифр (ДпКПП)")
        row1.addWidget(self.new_type_code, 1)

        self.new_type_cat = QComboBox()
        self.new_type_cat.setEditable(True)
        self.new_type_cat.addItems([
            "Наряд", "Увольнение", "Госпиталь", "Лазарет", "Командировка", "Отпуск", "Похороны", "Карантин"
        ])
        row1.addWidget(self.new_type_cat, 2)

        btn_save_type = QPushButton("Добавить причину")
        btn_save_type.setFixedHeight(36)
        btn_save_type.setIcon(create_svg_icon(SVG_PLUS, "#FFFFFF", 14))
        btn_save_type.clicked.connect(self.save_new_type)
        row1.addWidget(btn_save_type, 2)

        layout.addWidget(add_type_box)

        self.table_types = QTableWidget()
        self.table_types.setColumnCount(5)
        self.table_types.setHorizontalHeaderLabels(["ID", "Название / Наряд", "Шифр", "Категория в сводке", "Действия"])
        self.table_types.setAlternatingRowColors(True)
        self.table_types.verticalHeader().setVisible(False)
        self.table_types.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table_types.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table_types.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table_types.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.table_types.cellDoubleClicked.connect(self._on_type_cell_double_clicked)

        layout.addWidget(self.table_types)

        self.load_types_table()
        return tab

    def get_selected_date_str(self):
        return self.date_edit.date().toString("yyyy-MM-dd")

    def load_types_combo(self):
        self.cb_types.clear()
        types = self.db.get_absence_types()
        for t in types:
            if t['category'] == t['name']:
                display = f"{t['name']} ({t['code']})" if t['code'] else t['name']
            else:
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
        present = max(0, total - abs_count)

        self.badge_total.setText(f"📋 По списку: {total}")
        self.badge_present.setText(f"✅ В строю: {present}")
        self.badge_absent.setText(f"⚠️ В расходе: {abs_count}")

        self.table_absences.setRowCount(0)
        icon_del = create_svg_icon(SVG_TRASH, "#DC2626", 12)

        for idx, a in enumerate(absences, 1):
            r = self.table_absences.rowCount()
            self.table_absences.insertRow(r)
            self.table_absences.setRowHeight(r, 38)

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

        # Поиск совпадения по введенному тексту, если не было явного клика
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

        # Если введена новая причина отсутствия, которой еще нет в списке
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
            self.load_absence_types()

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
        cadets = self.db.get_all_cadets()
        absences = self.db.get_absences_for_date(d_str)

        if not cadets:
            QMessageBox.warning(self, "Внимание", "Сначала загрузите личный состав курса во вкладке «Личный состав».")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить строевую записку", f"Строевая_записка_{d_str}.xlsx", "Книга Excel (*.xlsx)"
        )

        if file_path:
            try:
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
                QMessageBox.information(self, "Успех", f"Строевая записка успешно сформирована:\n{file_path}")
            except Exception as e:
                QMessageBox.critical(self, "Ошибка формирования", f"Не удалось создать файл:\n{e}")

    def import_cadets_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Выберите файл с базой курса", "", "Книга Excel (*.xlsx *.xlsm);;Все файлы (*.*)"
        )

        if file_path:
            try:
                cadets = import_cadets_from_excel(file_path)
                if not cadets:
                    QMessageBox.warning(
                        self, "Внимание",
                        "В файле не найдены данные курсантов.\n\n"
                        "Убедитесь, что таблица содержит строки с ФИО и номером группы/взвода."
                    )
                    return

                self.db.clear_and_import_cadets(cadets)
                self.load_cadets_table()
                self.load_cadets_combo()
                self.load_daily_data()
                QMessageBox.information(
                    self, "Успех",
                    f"Успешно импортирован личный состав:\nЗагружено {len(cadets)} чел."
                )
            except Exception as e:
                QMessageBox.critical(self, "Ошибка чтения файла", f"{e}")

    def open_add_cadet_dialog(self):
        groups = self.db.get_distinct_groups()
        dlg = CadetDialog(self, available_groups=groups)
        if dlg.exec_() == QDialog.Accepted:
            data = dlg.get_data()
            self.db.add_cadet(
                rank=data['rank'],
                fio=data['fio'],
                group_num=data['group_num'],
                gender=data['gender'],
                short_fio=data['short_fio']
            )
            self.load_cadets_table()
            self.load_cadets_combo()
            self.load_daily_data()
            QMessageBox.information(self, "Успех", f"Курсант {data['fio']} успешно добавлен в список.")

    def open_edit_cadet_dialog(self, cadet_id):
        cadet = self.db.get_cadet_by_id(cadet_id)
        if not cadet:
            QMessageBox.warning(self, "Ошибка", "Курсант не найден в базе данных.")
            return

        groups = self.db.get_distinct_groups()
        dlg = CadetDialog(self, cadet_data=cadet, available_groups=groups)
        if dlg.exec_() == QDialog.Accepted:
            data = dlg.get_data()
            self.db.update_cadet(
                cadet_id=cadet_id,
                rank=data['rank'],
                fio=data['fio'],
                group_num=data['group_num'],
                gender=data['gender'],
                short_fio=data['short_fio']
            )
            self.load_cadets_table()
            self.load_cadets_combo()
            self.load_daily_data()

    def confirm_delete_cadet(self, cadet_id, rank, fio, group_num):
        msg = (
            f"Вы действительно хотите удалить курсанта:\n\n"
            f"{rank} {fio} (гр. {group_num})?\n\n"
            f"Курсант будет удален из списков личного состава и расходов "
            f"(например, в связи с переводом или отчислением)."
        )
        reply = QMessageBox.question(
            self, "Подтверждение удаления курсанта", msg,
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            self.db.delete_cadet(cadet_id)
            self.load_cadets_table()
            self.load_cadets_combo()
            self.load_daily_data()

    def _on_cadet_cell_double_clicked(self, row, column):
        id_item = self.table_cadets.item(row, 0)
        if id_item:
            try:
                c_id = int(id_item.text())
                self.open_edit_cadet_dialog(c_id)
            except ValueError:
                pass

    def load_cadets_table(self):
        cadets = self.db.get_all_cadets()
        self.table_cadets.setRowCount(0)
        self.lbl_cadet_count.setText(f"Всего курсантов: {len(cadets)} чел.")

        # Обновление выпадающего списка групп
        current_filter = self.combo_group_filter.currentText()
        self.combo_group_filter.blockSignals(True)
        self.combo_group_filter.clear()
        self.combo_group_filter.addItem("Все группы")
        groups = self.db.get_distinct_groups()
        for g in groups:
            self.combo_group_filter.addItem(str(g))
        # Восстановление выбранного фильтра если он есть
        idx = self.combo_group_filter.findText(current_filter)
        if idx >= 0:
            self.combo_group_filter.setCurrentIndex(idx)
        self.combo_group_filter.blockSignals(False)

        icon_edit = create_svg_icon(SVG_EDIT, "#2563EB", 12)
        icon_del = create_svg_icon(SVG_TRASH, "#DC2626", 12)

        for c in cadets:
            r = self.table_cadets.rowCount()
            self.table_cadets.insertRow(r)
            self.table_cadets.setRowHeight(r, 38)

            c_id = c['id']
            c_rank = c['rank'] or ""
            c_fio = c['fio']
            c_short = c.get('short_fio') or make_short_fio(c_fio)
            c_grp = str(c['group_num'])
            c_gender = c.get('gender', 'М')

            self.table_cadets.setItem(r, 0, QTableWidgetItem(str(c_id)))
            self.table_cadets.setItem(r, 1, QTableWidgetItem(c_rank))
            self.table_cadets.setItem(r, 2, QTableWidgetItem(c_fio))
            self.table_cadets.setItem(r, 3, QTableWidgetItem(c_short))
            self.table_cadets.setItem(r, 4, QTableWidgetItem(c_grp))
            self.table_cadets.setItem(r, 5, QTableWidgetItem(c_gender))

            # Панель действий: Изменить и Удалить
            action_widget = QWidget()
            action_layout = QHBoxLayout(action_widget)
            action_layout.setContentsMargins(4, 2, 4, 2)
            action_layout.setSpacing(6)

            btn_edit = QPushButton("Изменить")
            btn_edit.setObjectName("editBtn")
            btn_edit.setIcon(icon_edit)
            btn_edit.clicked.connect(lambda _, cid=c_id: self.open_edit_cadet_dialog(cid))
            action_layout.addWidget(btn_edit)

            btn_del = QPushButton("Удалить")
            btn_del.setObjectName("deleteBtn")
            btn_del.setIcon(icon_del)
            btn_del.clicked.connect(lambda _, cid=c_id, cr=c_rank, cf=c_fio, cg=c_grp: self.confirm_delete_cadet(cid, cr, cf, cg))
            action_layout.addWidget(btn_del)

            self.table_cadets.setCellWidget(r, 6, action_widget)

        self.filter_cadets_table()

    def filter_cadets_table(self):
        query = self.search_cadet.text().lower().strip()
        selected_group = self.combo_group_filter.currentText()
        visible_count = 0
        total = self.table_cadets.rowCount()

        for r in range(total):
            rank = self.table_cadets.item(r, 1).text().lower() if self.table_cadets.item(r, 1) else ""
            fio = self.table_cadets.item(r, 2).text().lower() if self.table_cadets.item(r, 2) else ""
            short_fio = self.table_cadets.item(r, 3).text().lower() if self.table_cadets.item(r, 3) else ""
            grp = self.table_cadets.item(r, 4).text() if self.table_cadets.item(r, 4) else ""

            match_text = (not query) or (query in fio) or (query in grp.lower()) or (query in rank) or (query in short_fio)
            match_group = (selected_group == "Все группы") or (grp == selected_group)

            match = match_text and match_group
            self.table_cadets.setRowHidden(r, not match)
            if match:
                visible_count += 1

        if query or selected_group != "Все группы":
            self.lbl_cadet_count.setText(f"Найдено: {visible_count} из {total}")
        else:
            self.lbl_cadet_count.setText(f"Всего курсантов: {total} чел.")

    def save_header_settings(self):
        self.db.set_setting("unit_title", self.edit_unit.text().strip())
        self.db.set_setting("chief_title", self.edit_chief_title.text().strip())
        self.db.set_setting("chief_name", self.edit_chief_name.text().strip())
        QMessageBox.information(self, "Сохранено", "Реквизиты и подписи бланка успешно сохранены.")

    def save_new_type(self):
        name = self.new_type_name.text().strip()
        code = self.new_type_code.text().strip()
        cat = self.new_type_cat.currentText().strip()

        if not name or not cat:
            QMessageBox.warning(self, "Внимание", "Заполните название причины/наряда.")
            return

        self.db.add_absence_type(name, code, cat)
        self.new_type_name.clear()
        self.new_type_code.clear()
        self.load_types_table()
        self.load_types_combo()

    def _on_type_cell_double_clicked(self, row, column):
        id_item = self.table_types.item(row, 0)
        if id_item:
            try:
                t_id = int(id_item.text())
                self.open_edit_type_dialog(t_id)
            except ValueError:
                pass

    def open_edit_type_dialog(self, type_id):
        t_data = self.db.get_absence_type_by_id(type_id)
        if not t_data:
            QMessageBox.warning(self, "Ошибка", "Причина отсутствия не найдена.")
            return

        categories = self.db.get_distinct_categories()
        dlg = AbsenceTypeDialog(self, type_data=t_data, available_categories=categories)
        if dlg.exec_() == QDialog.Accepted:
            data = dlg.get_data()
            self.db.update_absence_type(type_id, data['name'], data['code'], data['category'])
            self.load_types_table()
            self.load_types_combo()
            self.load_daily_data()
            QMessageBox.information(
                self, "Успех",
                f"Причина «{data['name']}» успешно обновлена.\n"
                f"Категория в сводке: {data['category']}"
            )

    def confirm_delete_type(self, type_id, name):
        reply = QMessageBox.question(
            self, "Подтверждение удаления",
            f"Удалить причину отсутствия «{name}» из справочника?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            self.db.delete_absence_type(type_id)
            self.load_types_table()
            self.load_types_combo()
            self.load_daily_data()

    def load_types_table(self):
        types = self.db.get_absence_types()
        self.table_types.setRowCount(0)

        # Синхронизация списка категорий в панели добавления
        current_cat = self.new_type_cat.currentText()
        distinct_cats = self.db.get_distinct_categories()
        default_cats = ["Наряд", "Увольнение", "Госпиталь", "Лазарет", "Командировка", "Отпуск", "Похороны", "Карантин"]
        all_cats = list(dict.fromkeys(default_cats + distinct_cats))
        self.new_type_cat.blockSignals(True)
        self.new_type_cat.clear()
        self.new_type_cat.addItems(all_cats)
        if current_cat:
            self.new_type_cat.setCurrentText(current_cat)
        self.new_type_cat.blockSignals(False)

        icon_edit = create_svg_icon(SVG_EDIT, "#2563EB", 12)
        icon_del = create_svg_icon(SVG_TRASH, "#DC2626", 12)

        for t in types:
            r = self.table_types.rowCount()
            self.table_types.insertRow(r)
            self.table_types.setRowHeight(r, 38)

            t_id = t['id']
            t_name = t['name']
            t_code = t['code'] or ""
            t_cat = t['category']

            self.table_types.setItem(r, 0, QTableWidgetItem(str(t_id)))
            self.table_types.setItem(r, 1, QTableWidgetItem(t_name))
            self.table_types.setItem(r, 2, QTableWidgetItem(t_code))
            self.table_types.setItem(r, 3, QTableWidgetItem(t_cat))

            # Панель действий: Изменить и Удалить
            action_widget = QWidget()
            action_layout = QHBoxLayout(action_widget)
            action_layout.setContentsMargins(4, 2, 4, 2)
            action_layout.setSpacing(6)

            btn_edit = QPushButton("Изменить")
            btn_edit.setObjectName("editBtn")
            btn_edit.setIcon(icon_edit)
            btn_edit.clicked.connect(lambda _, tid=t_id: self.open_edit_type_dialog(tid))
            action_layout.addWidget(btn_edit)

            btn_del = QPushButton("Удалить")
            btn_del.setObjectName("deleteBtn")
            btn_del.setIcon(icon_del)
            btn_del.clicked.connect(lambda _, tid=t_id, tn=t_name: self.confirm_delete_type(tid, tn))
            action_layout.addWidget(btn_del)

            self.table_types.setCellWidget(r, 4, action_widget)

if __name__ == "__main__":
    app = QApplication(sys.argv)

    palette = app.palette()
    palette.setColor(QPalette.Window, QColor("#F8FAFC"))
    palette.setColor(QPalette.WindowText, QColor("#0F172A"))
    palette.setColor(QPalette.Base, QColor("#FFFFFF"))
    palette.setColor(QPalette.AlternateBase, QColor("#F8FAFC"))
    palette.setColor(QPalette.ToolTipBase, QColor("#FFFFFF"))
    palette.setColor(QPalette.ToolTipText, QColor("#0F172A"))
    palette.setColor(QPalette.Text, QColor("#0F172A"))
    palette.setColor(QPalette.Button, QColor("#F8FAFC"))
    palette.setColor(QPalette.ButtonText, QColor("#0F172A"))
    palette.setColor(QPalette.BrightText, QColor("#EF4444"))
    palette.setColor(QPalette.Highlight, QColor("#2563EB"))
    palette.setColor(QPalette.HighlightedText, QColor("#FFFFFF"))
    app.setPalette(palette)

    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
