import os
import sys
import re
from typing import Optional, Tuple
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QTableWidget, QTableWidgetItem, QComboBox,
    QLineEdit, QDateEdit, QFileDialog, QMessageBox, QTabWidget,
    QHeaderView, QFrame, QAbstractItemView, QDialog, QFormLayout, QDialogButtonBox,
    QButtonGroup, QListView
)
from PyQt5.QtCore import Qt, QDate, QByteArray, QSize
from PyQt5.QtGui import QIcon, QPixmap, QPainter, QColor
from PyQt5.QtSvg import QSvgRenderer

from database import Database
from core.parsers import (
    parse_roster_excel, make_short_fio, change_group_course,
    parse_group_code, get_course_from_group, update_header_text,
    parse_signature_from_excel
)
from core.models import normalize_gender, gender_to_ru
from core.balance import compute_tri_slice_balance
from core.exporter import export_report_to_excel
from ui.dialogs.import_wizard import ImportReviewDialog
from ui.icons import (
    SVG_CLIPBOARD, SVG_USERS, SVG_SETTINGS, SVG_EXPORT, SVG_IMPORT,
    SVG_PLUS, SVG_TRASH, SVG_CALENDAR, SVG_EDIT, SVG_SAVE, SVG_MALE,
    SVG_FEMALE, SVG_GRADUATE, SVG_FAST_FORWARD, SVG_CHECK, SVG_CROSS,
    SVG_SCALE, SVG_ALERT, SVG_SEARCH, SVG_FILTER, SVG_REFRESH, SVG_INFO,
    SVG_CODE, create_svg_icon
)

POPUP_STYLE = """
QListView {
    background-color: #ffffff;
    color: #1e293b;
    selection-background-color: #2563eb;
    selection-color: #ffffff;
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    outline: none;
}
QListView::item {
    min-height: 28px;
    padding: 5px 8px;
    background-color: #ffffff;
    color: #1e293b;
}
QListView::item:selected,
QListView::item:hover {
    background-color: #2563eb;
    color: #ffffff;
}
"""

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

QComboBox {
    background-color: #FFFFFF;
    color: #1E293B;
    border: 1px solid #CBD5E1;
    border-radius: 6px;
    padding: 5px 10px;
}

QComboBox::drop-down {
    border: none;
}

QComboBox QAbstractItemView,
QComboBox QListView,
QListView {
    background-color: #FFFFFF;
    color: #1E293B;
    selection-background-color: #2563EB;
    selection-color: #FFFFFF;
    border: 1px solid #CBD5E1;
    outline: none;
}

QComboBox QAbstractItemView::item,
QComboBox QListView::item,
QListView::item {
    min-height: 28px;
    padding: 4px 8px;
    background-color: #FFFFFF;
    color: #1E293B;
}

QComboBox QAbstractItemView::item:selected,
QComboBox QListView::item:selected,
QListView::item:selected {
    background-color: #2563EB;
    color: #FFFFFF;
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

QPushButton#bulkCourseBtn {
    background-color: #0284C7;
}
QPushButton#bulkCourseBtn:hover {
    background-color: #0369A1;
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
        self.edit_group.setView(QListView())
        self.edit_group.view().setStyleSheet(POPUP_STYLE)
        self.edit_group.setEditable(True)
        if available_groups:
            self.edit_group.addItems([str(g) for g in available_groups])
        if cadet:
            self.edit_group.setCurrentText(str(cadet['group_num']))

        self.edit_course = QComboBox()
        self.edit_course.setView(QListView())
        self.edit_course.view().setStyleSheet(POPUP_STYLE)
        self.edit_course.addItems(["1 курс", "2 курс", "3 курс", "4 курс", "5 курс"])
        curr_course = cadet.get('course', 1) if cadet else 1
        self.edit_course.setCurrentIndex(max(0, min(4, curr_course - 1)))

        self.edit_gender = QComboBox()
        self.edit_gender.setView(QListView())
        self.edit_gender.view().setStyleSheet(POPUP_STYLE)
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


class BulkCourseDialog(QDialog):
    """Диалог массового перевода курса для всех учебных групп."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Массовый перевод курса")
        self.setFixedSize(400, 220)
        self.setStyleSheet("""
            QDialog {
                background-color: #FFFFFF;
            }
            QLabel {
                color: #0F172A;
                font-size: 13px;
            }
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(22, 20, 22, 20)
        layout.setSpacing(14)

        title_lbl = QLabel("Массовый перевод курса")
        title_lbl.setStyleSheet("font-weight: 700; font-size: 16px; color: #0F172A;")
        layout.addWidget(title_lbl)

        desc_lbl = QLabel("У всех 4-значных групп (например, 1280) вторая цифра будет заменена на выбранный курс (станет 1380).")
        desc_lbl.setWordWrap(True)
        desc_lbl.setStyleSheet("color: #64748B; font-size: 12px; line-height: 1.4;")
        layout.addWidget(desc_lbl)

        form_layout = QHBoxLayout()
        form_lbl = QLabel("Установить курс:")
        form_lbl.setStyleSheet("font-weight: 600; color: #334155;")
        self.cb_course = QComboBox()
        self.cb_course.setView(QListView())
        self.cb_course.view().setStyleSheet(POPUP_STYLE)
        self.cb_course.addItems(["1 курс", "2 курс", "3 курс", "4 курс", "5 курс"])
        self.cb_course.setCurrentIndex(2)  # 3 курс по умолчанию
        self.cb_course.setFixedHeight(34)
        form_layout.addWidget(form_lbl)
        form_layout.addWidget(self.cb_course, 1)
        layout.addLayout(form_layout)

        layout.addStretch()

        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        self.btn_cancel = QPushButton("Отмена")
        self.btn_cancel.setStyleSheet("""
            QPushButton {
                background-color: #F1F5F9;
                color: #334155;
                font-weight: 600;
                font-size: 13px;
                border-radius: 6px;
                padding: 7px 16px;
                border: 1px solid #CBD5E1;
            }
            QPushButton:hover {
                background-color: #E2E8F0;
            }
        """)
        self.btn_cancel.clicked.connect(self.reject)
        btn_layout.addWidget(self.btn_cancel)

        self.btn_apply = QPushButton("Применить ко всем")
        self.btn_apply.setStyleSheet("""
            QPushButton {
                background-color: #2563EB;
                color: #FFFFFF;
                font-weight: 600;
                font-size: 13px;
                border-radius: 6px;
                padding: 7px 16px;
                border: none;
            }
            QPushButton:hover {
                background-color: #1D4ED8;
            }
        """)
        self.btn_apply.setIcon(create_svg_icon(SVG_CHECK, "#FFFFFF", 14))
        self.btn_apply.clicked.connect(self.accept)
        btn_layout.addWidget(self.btn_apply)

        layout.addLayout(btn_layout)

    def get_selected_course(self) -> int:
        return self.cb_course.currentIndex() + 1


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ConsumABLE v1.3.4 — Управление строевой запиской")
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

        # Нижняя строка состояния (StatusBar) с информацией об авторе
        status_bar = self.statusBar()
        status_bar.setStyleSheet("""
            QStatusBar {
                background-color: #F8FAFC;
                border-top: 1px solid #E2E8F0;
                color: #64748B;
                padding: 4px 10px;
                font-size: 11px;
            }
            QStatusBar QLabel {
                color: #64748B;
                font-size: 11px;
            }
        """)

        status_left = QLabel("ConsumABLE v1.3.4 • Система автоматизации строевого учета")
        status_bar.addWidget(status_left)

        author_widget = QWidget()
        author_layout = QHBoxLayout(author_widget)
        author_layout.setContentsMargins(0, 0, 0, 0)
        author_layout.setSpacing(6)

        author_icon = QLabel()
        author_icon.setPixmap(create_svg_icon(SVG_CODE, "#64748B", 13).pixmap(13, 13))
        author_label = QLabel("Создал: Журавлев Алексей")
        author_label.setStyleSheet("color: #64748B; font-size: 11px; font-weight: 500;")

        author_layout.addWidget(author_icon)
        author_layout.addWidget(author_label)
        status_bar.addPermanentWidget(author_widget)

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
        self.cb_cadets.setView(QListView())
        self.cb_cadets.view().setStyleSheet(POPUP_STYLE)
        self.cb_cadets.setEditable(True)
        self.cb_cadets.setInsertPolicy(QComboBox.NoInsert)
        if self.cb_cadets.completer():
            self.cb_cadets.completer().setFilterMode(Qt.MatchContains)
            comp_popup = QListView()
            comp_popup.setStyleSheet(POPUP_STYLE)
            self.cb_cadets.completer().setPopup(comp_popup)
        self.cb_cadets.setPlaceholderText("Фамилия курсанта...")
        if self.cb_cadets.lineEdit():
            self.cb_cadets.lineEdit().returnPressed.connect(self.on_cadet_return_pressed)
        add_layout.addWidget(self.cb_cadets, 4)

        self.cb_types = QComboBox()
        self.cb_types.setView(QListView())
        self.cb_types.view().setStyleSheet(POPUP_STYLE)
        self.cb_types.setEditable(True)
        self.cb_types.setInsertPolicy(QComboBox.NoInsert)
        if self.cb_types.completer():
            self.cb_types.completer().setFilterMode(Qt.MatchContains)
            comp_popup_types = QListView()
            comp_popup_types.setStyleSheet(POPUP_STYLE)
            self.cb_types.completer().setPopup(comp_popup_types)
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
        # 3 ИНФОРМАЦИОННЫЕ КАРТОЧКИ СО СТАТИСТИКОЙ (Всего / Мужчины / Девушки)
        # -------------------------------------------------------------
        stats_frame = QFrame()
        stats_frame.setObjectName("cardFrame")
        stats_layout = QHBoxLayout(stats_frame)
        stats_layout.setContentsMargins(10, 10, 10, 10)
        stats_layout.setSpacing(12)

        def create_stat_card(title: str, icon_svg: str, icon_color: str, card_id: str):
            card = QFrame()
            card.setStyleSheet("""
                QFrame {
                    background: #FFFFFF;
                    border: 1px solid #E2E8F0;
                    border-radius: 8px;
                    padding: 8px 12px;
                }
            """)
            c_layout = QVBoxLayout(card)
            c_layout.setContentsMargins(4, 4, 4, 4)
            c_layout.setSpacing(4)

            header = QHBoxLayout()
            lbl_icon = QLabel()
            lbl_icon.setPixmap(create_svg_icon(icon_svg, icon_color, 16).pixmap(16, 16))
            lbl_title = QLabel(title)
            lbl_title.setStyleSheet(f"font-weight: 700; font-size: 13px; color: {icon_color};")
            header.addWidget(lbl_icon)
            header.addWidget(lbl_title)
            header.addStretch()
            c_layout.addLayout(header)

            val_lbl = QLabel("По списку: 0 чел. | В строю: 0 чел. | В расходе: 0 чел.")
            val_lbl.setStyleSheet("font-size: 12px; font-weight: 600; color: #334155;")
            c_layout.addWidget(val_lbl)

            setattr(self, f"card_stat_{card_id}", val_lbl)
            return card

        stats_layout.addWidget(create_stat_card("Общий состав (Всего)", SVG_USERS, "#2563EB", "total"))
        stats_layout.addWidget(create_stat_card("Мужской состав (Парни)", SVG_MALE, "#1D4ED8", "male"))
        stats_layout.addWidget(create_stat_card("Женский состав (Девушки)", SVG_FEMALE, "#DB2777", "female"))
        layout.addWidget(stats_frame)

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

        self.slice_tabs.addTab(self.tab_slice_total, create_svg_icon(SVG_USERS, "#2563EB", 16), "Общий расход")
        self.slice_tabs.addTab(self.tab_slice_male, create_svg_icon(SVG_MALE, "#1D4ED8", 16), "Мужской состав")
        self.slice_tabs.addTab(self.tab_slice_female, create_svg_icon(SVG_FEMALE, "#DB2777", 16), "Женский состав")
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

        btn_bulk_course = QPushButton("Сменить курс всех групп")
        btn_bulk_course.setObjectName("bulkCourseBtn")
        btn_bulk_course.setFixedHeight(36)
        btn_bulk_course.setIcon(create_svg_icon(SVG_REFRESH, "#FFFFFF", 14))
        btn_bulk_course.clicked.connect(self.open_bulk_course_dialog)
        top_panel.addWidget(btn_bulk_course)

        btn_inc_course = QPushButton("+1 курс")
        btn_inc_course.setObjectName("incCourseBtn")
        btn_inc_course.setFixedHeight(36)
        btn_inc_course.setToolTip("Повысить курс всех учебных групп на +1 (1280 -> 1380)")
        btn_inc_course.setIcon(create_svg_icon(SVG_FAST_FORWARD, "#FFFFFF", 14))
        btn_inc_course.clicked.connect(self.increment_all_courses)
        top_panel.addWidget(btn_inc_course)

        # Фильтр по курсу
        self.combo_course_filter = QComboBox()
        self.combo_course_filter.setView(QListView())
        self.combo_course_filter.view().setStyleSheet(POPUP_STYLE)
        self.combo_course_filter.setFixedHeight(36)
        self.combo_course_filter.addItems(["Все курсы", "1 курс", "2 курс", "3 курс", "4 курс", "5 курс"])
        self.combo_course_filter.currentIndexChanged.connect(self.load_cadets_table)
        top_panel.addWidget(self.combo_course_filter)

        # Фильтр по полу
        self.combo_gender_filter = QComboBox()
        self.combo_gender_filter.setView(QListView())
        self.combo_gender_filter.view().setStyleSheet(POPUP_STYLE)
        self.combo_gender_filter.setFixedHeight(36)
        self.combo_gender_filter.addItems(["Все (М и Ж)", "Только мужчины", "Только женщины"])
        self.combo_gender_filter.currentIndexChanged.connect(self.load_cadets_table)
        top_panel.addWidget(self.combo_gender_filter)

        # Выпадающий фильтр по учебным группам
        self.combo_group_filter = QComboBox()
        self.combo_group_filter.setView(QListView())
        self.combo_group_filter.view().setStyleSheet(POPUP_STYLE)
        self.combo_group_filter.setFixedHeight(36)
        self.combo_group_filter.currentIndexChanged.connect(self.load_cadets_table)
        top_panel.addWidget(self.combo_group_filter)

        self.edit_search = QLineEdit()
        self.edit_search.setFixedHeight(36)
        self.edit_search.setPlaceholderText("Поиск по фамилии или званию...")
        self.edit_search.textChanged.connect(self.load_cadets_table)
        top_panel.addWidget(self.edit_search)

        layout.addLayout(top_panel)

        # Сегментированный переключатель курса: [ Все ] [ 1 ] [ 2 ] [ 3 ] [ 4 ] [ 5 ]
        course_bar = QFrame()
        course_bar.setObjectName("cardFrame")
        course_bar_layout = QHBoxLayout(course_bar)
        course_bar_layout.setContentsMargins(8, 6, 8, 6)
        course_bar_layout.setSpacing(6)

        lbl_c_filter = QLabel("Курс:")
        lbl_c_filter.setStyleSheet("font-weight: 700; font-size: 13px; color: #334155; margin-right: 4px;")
        course_bar_layout.addWidget(lbl_c_filter)

        self.course_btn_group = QButtonGroup(self)
        self.course_btn_group.setExclusive(True)
        self.selected_course_filter = 0

        courses = [("Все курсы", 0), ("1 курс", 1), ("2 курс", 2), ("3 курс", 3), ("4 курс", 4), ("5 курс", 5)]
        for lbl_text, c_val in courses:
            btn = QPushButton(lbl_text)
            btn.setCheckable(True)
            if c_val == 0:
                btn.setChecked(True)
            btn.setStyleSheet("""
                QPushButton {
                    border: 1px solid #CBD5E1;
                    background: #FFFFFF;
                    color: #334155;
                    border-radius: 6px;
                    padding: 5px 14px;
                    font-size: 12px;
                    font-weight: 600;
                }
                QPushButton:checked {
                    background: #2563EB;
                    color: #FFFFFF;
                    border-color: #1D4ED8;
                }
                QPushButton:hover:!checked {
                    background: #F1F5F9;
                }
            """)
            self.course_btn_group.addButton(btn, c_val)
            course_bar_layout.addWidget(btn)

        self.course_btn_group.buttonClicked[int].connect(self.on_course_segmented_clicked)
        course_bar_layout.addStretch()
        layout.addWidget(course_bar)

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
        self.edit_unit = QLineEdit(self.db.get_setting("unit_title", "* курс * факультета (автоматизированных систем управления)"))
        self.edit_chief_title = QLineEdit(self.db.get_setting("chief_title", "Начальник * курса * факультета"))

        self.combo_commander_rank = QComboBox()
        self.combo_commander_rank.setView(QListView())
        self.combo_commander_rank.view().setStyleSheet(POPUP_STYLE)
        self.combo_commander_rank.addItems([
            "", "майор", "капитан", "подполковник", "старший лейтенант", "лейтенант",
            "полковник", "младший лейтенант", "старший прапорщик", "прапорщик",
            "старшина", "старший сержант", "сержант", "младший сержант", "ефрейтор", "рядовой"
        ])
        curr_rank = self.db.get_setting("commander_rank", "")
        idx_r = self.combo_commander_rank.findText(curr_rank)
        if idx_r >= 0:
            self.combo_commander_rank.setCurrentIndex(idx_r)
        else:
            self.combo_commander_rank.setCurrentIndex(0)

        self.edit_commander_name = QLineEdit(self.db.get_setting("commander_name", ""))
        self.edit_commander_name.setPlaceholderText("Фамилия И.О.")

        form_sig.addRow("Наименование подразделения:", self.edit_unit)
        form_sig.addRow("Должность начальника:", self.edit_chief_title)
        form_sig.addRow("В/звание начальника:", self.combo_commander_rank)
        form_sig.addRow("Фамилия и инициалы:", self.edit_commander_name)
        sig_layout.addLayout(form_sig)

        btn_save_sig = QPushButton("Сохранить параметры")
        btn_save_sig.setIcon(create_svg_icon(SVG_SAVE, "#FFFFFF", 14))
        btn_save_sig.clicked.connect(self.save_signatures)
        sig_layout.addWidget(btn_save_sig, alignment=Qt.AlignRight)

        layout.addWidget(card_signatures)

        # Информационный блок / Карточка «О программе»
        card_about = QFrame()
        card_about.setObjectName("cardFrame")
        about_layout = QVBoxLayout(card_about)
        about_layout.setSpacing(10)

        about_header = QHBoxLayout()
        about_header.setSpacing(8)
        lbl_info_icon = QLabel()
        lbl_info_icon.setPixmap(create_svg_icon(SVG_INFO, "#2563EB", 18).pixmap(18, 18))
        lbl_about_title = QLabel("О программе")
        lbl_about_title.setStyleSheet("font-size: 14px; font-weight: 700; color: #1E293B;")
        about_header.addWidget(lbl_info_icon)
        about_header.addWidget(lbl_about_title)
        about_header.addStretch()
        about_layout.addLayout(about_header)

        lbl_app_name = QLabel("ConsumABLE v1.3.4")
        lbl_app_name.setStyleSheet("font-size: 14px; font-weight: 700; color: #2563EB;")
        about_layout.addWidget(lbl_app_name)

        lbl_author = QLabel("Автор / Создатель: Журавлев Алексей")
        lbl_author.setStyleSheet("font-size: 13px; font-weight: 600; color: #1E293B;")
        about_layout.addWidget(lbl_author)

        lbl_purpose = QLabel("Назначение: Автоматизация учета и расхода личного состава с формированием строевых записок.")
        lbl_purpose.setStyleSheet("font-size: 12px; color: #64748B;")
        lbl_purpose.setWordWrap(True)
        about_layout.addWidget(lbl_purpose)

        layout.addWidget(card_about)
        layout.addStretch()
        return tab

    def get_chief_full_name(self) -> str:
        rank = self.db.get_setting("commander_rank", "").strip()
        fio = self.db.get_setting("commander_name", "").strip()
        if rank and fio:
            return f"{rank}                          {fio}"
        elif rank:
            return rank
        elif fio:
            return fio
        return ""

    def save_signatures(self):
        self.db.set_setting("unit_title", self.edit_unit.text().strip())
        self.db.set_setting("chief_title", self.edit_chief_title.text().strip())
        rank = self.combo_commander_rank.currentText().strip()
        fio = self.edit_commander_name.text().strip()
        self.db.set_setting("commander_rank", rank)
        self.db.set_setting("commander_name", fio)
        full_chief = f"{rank}                          {fio}" if (rank and fio) else (rank or fio)
        self.db.set_setting("chief_name", full_chief)
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
            cg = get_course_from_group(c['group_num'])
            course_display = cg if cg != "?" else str(c.get('course', 1))
            self.cb_cadets.addItem(f"{c['short_fio']} | {c['rank']} (гр. {c['group_num']}, {course_display} к.)", userData=c['id'])

    def update_header_settings(self, faculty: Optional[int] = None, course: Optional[int] = None):
        """Updates unitTitle and chiefTitle settings and UI fields with real course/faculty numbers."""
        unit_title = self.db.get_setting("unit_title", "* курс * факультета (автоматизированных систем управления)")
        chief_title = self.db.get_setting("chief_title", "Начальник * курса * факультета")

        new_unit = update_header_text(unit_title, faculty=faculty, course=course)
        new_chief = update_header_text(chief_title, faculty=faculty, course=course)

        self.db.set_setting("unit_title", new_unit)
        self.db.set_setting("chief_title", new_chief)

        if hasattr(self, "edit_unit") and self.edit_unit is not None:
            self.edit_unit.setText(new_unit)
        if hasattr(self, "edit_chief_title") and self.edit_chief_title is not None:
            self.edit_chief_title.setText(new_chief)

    def open_bulk_course_dialog(self):
        dlg = BulkCourseDialog(self)
        if dlg.exec_() == QDialog.Accepted:
            target_course = dlg.get_selected_course()
            self.db.bulk_set_course(target_course, update_group_numbers=True)
            self.update_header_settings(course=target_course)
            self.update_group_filter_options()
            self.load_cadets_table()
            self.load_cadets_combo()
            self.load_daily_data()
            QMessageBox.information(
                self,
                "Успех",
                f"Курс всех групп успешно изменен на {target_course} курс.\nНомера групп обновлены (например, 1280 -> 1{target_course}80)."
            )

    def increment_all_courses(self):
        reply = QMessageBox.question(
            self,
            "Подтверждение",
            "Повысить курс всех учебных групп на +1?\n\nВ номерах групп 1X** цифра курса увеличится (например, 1280 -> 1380), курсанты старше 5 курса будут переведены в статус 'выпущен'.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes
        )
        if reply == QMessageBox.Yes:
            self.db.bulk_increment_course(max_course=5, update_group_numbers=True)
            cursor = self.db.conn.cursor()
            cursor.execute("SELECT course FROM cadets WHERE is_active = 1 AND status = 'active' LIMIT 1")
            row = cursor.fetchone()
            if row and row[0]:
                self.update_header_settings(course=row[0])
            self.update_group_filter_options()
            self.load_cadets_table()
            self.load_cadets_combo()
            self.load_daily_data()
            QMessageBox.information(
                self,
                "Успех",
                "Курс всех учебных групп успешно повышен на +1.\nНомера групп автоматически обновлены."
            )

    def on_course_segmented_clicked(self, course_id: int):
        self.selected_course_filter = course_id
        if hasattr(self, "combo_course_filter"):
            self.combo_course_filter.blockSignals(True)
            self.combo_course_filter.setCurrentIndex(course_id)
            self.combo_course_filter.blockSignals(False)
        self.load_cadets_table()

    def toggle_cadet_gender_action(self, cadet_id: int):
        self.db.toggle_cadet_gender(cadet_id)
        self.load_cadets_table()
        self.load_daily_data()

    def load_daily_data(self):
        d_str = self.get_selected_date_str()
        absences = self.db.get_absences_for_date(d_str)
        cadets = self.db.get_all_cadets()

        # Compute 3-slice balance
        categories = self.db.get_distinct_categories()
        balance = compute_tri_slice_balance(cadets, absences, categories)

        # Update 3 information cards (Всего / Мужчины / Девушки)
        if hasattr(self, "card_stat_total"):
            self.card_stat_total.setText(
                f"По списку: {balance.total.listed_count} чел. | В строю: {balance.total.present_count} чел. | В расходе: {balance.total.absent_count} чел."
            )
        if hasattr(self, "card_stat_male"):
            self.card_stat_male.setText(
                f"По списку: {balance.male.listed_count} чел. | В строю: {balance.male.present_count} чел. | В расходе: {balance.male.absent_count} чел."
            )
        if hasattr(self, "card_stat_female"):
            self.card_stat_female.setText(
                f"По списку: {balance.female.listed_count} чел. | В строю: {balance.female.present_count} чел. | В расходе: {balance.female.absent_count} чел."
            )

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

            if badge_ps: badge_ps.setText(f"По списку: {s_bal.listed_count}")
            if badge_nl: badge_nl.setText(f"В строю: {s_bal.present_count}")
            if badge_abs: badge_abs.setText(f"В расходе: {s_bal.absent_count}")

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
                f"Сквозной баланс строя подтвержден: "
                f"Общий [П/С: {balance.total.listed_count}, В строю: {balance.total.present_count}, Расход: {balance.total.absent_count}] "
                f"= Мужчины [{balance.male.listed_count}/{balance.male.present_count}/{balance.male.absent_count}] "
                f"+ Женщины [{balance.female.listed_count}/{balance.female.present_count}/{balance.female.absent_count}]"
            )
        else:
            self.balance_bar.setStyleSheet(
                "background: #FEF2F2; color: #991B1B; border: 1px solid #FECACA; "
                "border-radius: 8px; padding: 10px 14px; font-weight: 700; font-size: 12px;"
            )
            self.balance_bar.setText(f"НАРУШЕНИЕ БАЛАНСА: {balance.error_message}")

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
                    unit_title=self.db.get_setting("unit_title", "* курс * факультета (автоматизированных систем управления)"),
                    chief_title=self.db.get_setting("chief_title", "Начальник * курса * факультета"),
                    chief_name=self.get_chief_full_name(),
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
            # Поиск подписи должностного лица (начальника курса) в Excel
            sig_title, sig_rank, sig_name = parse_signature_from_excel(file_path)
            if sig_rank:
                self.db.set_setting("commander_rank", sig_rank)
                if hasattr(self, "combo_commander_rank"):
                    idx_r = self.combo_commander_rank.findText(sig_rank)
                    if idx_r >= 0:
                        self.combo_commander_rank.setCurrentIndex(idx_r)
            else:
                self.db.set_setting("commander_rank", "")
                if hasattr(self, "combo_commander_rank"):
                    self.combo_commander_rank.setCurrentIndex(0)

            if sig_name:
                self.db.set_setting("commander_name", sig_name)
                if hasattr(self, "edit_commander_name"):
                    self.edit_commander_name.setText(sig_name)
            else:
                self.db.set_setting("commander_name", "")
                if hasattr(self, "edit_commander_name"):
                    self.edit_commander_name.clear()

            if sig_title:
                self.db.set_setting("chief_title", sig_title)
                if hasattr(self, "edit_chief_title"):
                    self.edit_chief_title.setText(sig_title)

            rk = self.db.get_setting("commander_rank", "").strip()
            nm = self.db.get_setting("commander_name", "").strip()
            if rk and nm:
                self.db.set_setting("chief_name", f"{rk}                          {nm}")
            elif rk:
                self.db.set_setting("chief_name", rk)
            elif nm:
                self.db.set_setting("chief_name", nm)
            else:
                self.db.set_setting("chief_name", "")

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

                # Автоподстановка курса и факультета в шапку строевой записки
                detected_faculty = None
                detected_course = None
                for c in final_cadets:
                    grp = getattr(c, 'group_num', None) or (c.get('group_num') if isinstance(c, dict) else None)
                    fac, crs, _ = parse_group_code(grp)
                    if fac and not detected_faculty:
                        detected_faculty = fac
                    if crs and not detected_course:
                        detected_course = crs
                    if detected_faculty and detected_course:
                        break

                if detected_faculty or detected_course:
                    self.update_header_settings(faculty=detected_faculty, course=detected_course)

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
        c_filter = getattr(self, "selected_course_filter", 0)
        if c_filter > 0:
            cadets = [
                c for c in cadets
                if (get_course_from_group(c.get('group_num')) == str(c_filter) if get_course_from_group(c.get('group_num')) != "?" else c.get('course') == c_filter)
            ]
        else:
            course_idx = self.combo_course_filter.currentIndex()
            if course_idx > 0:
                cadets = [
                    c for c in cadets
                    if (get_course_from_group(c.get('group_num')) == str(course_idx) if get_course_from_group(c.get('group_num')) != "?" else c.get('course') == course_idx)
                ]

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
            cg = get_course_from_group(c.get('group_num', ''))
            course_val = cg if cg != "?" else str(c.get('course', 1))
            self.table_cadets.setItem(r, 4, QTableWidgetItem(str(c.get('group_num', ''))))
            self.table_cadets.setItem(r, 5, QTableWidgetItem(f"{course_val} курс"))

            g_ru = gender_to_ru(c.get('gender', 'male'))
            btn_gender = QPushButton(f" {g_ru}")
            cid = c['id']
            if g_ru == "Ж":
                btn_gender.setIcon(create_svg_icon(SVG_FEMALE, "#BE185D", 14))
                btn_gender.setStyleSheet("""
                    QPushButton {
                        background: #FDF2F8;
                        color: #BE185D;
                        border: 1px solid #FBCFE8;
                        border-radius: 4px;
                        font-weight: 700;
                        padding: 3px 6px;
                        font-size: 12px;
                    }
                    QPushButton:hover { background: #FCE7F3; }
                """)
            else:
                btn_gender.setIcon(create_svg_icon(SVG_MALE, "#1D4ED8", 14))
                btn_gender.setStyleSheet("""
                    QPushButton {
                        background: #EFF6FF;
                        color: #1D4ED8;
                        border: 1px solid #BFDBFE;
                        border-radius: 4px;
                        font-weight: 700;
                        padding: 3px 6px;
                        font-size: 12px;
                    }
                    QPushButton:hover { background: #DBEAFE; }
                """)
            btn_gender.setToolTip("Нажмите для быстрой смены пола курсанта (М ↔ Ж)")
            btn_gender.clicked.connect(lambda _, cadet_id=cid: self.toggle_cadet_gender_action(cadet_id))
            self.table_cadets.setCellWidget(r, 6, btn_gender)

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
