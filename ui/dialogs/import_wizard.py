from typing import List, Optional
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QTableView, QToolBar,
    QPushButton, QLabel, QComboBox, QSpinBox, QMessageBox,
    QStyledItemDelegate, QHeaderView, QWidget, QFrame
)
from PyQt5.QtCore import Qt, QAbstractTableModel, QModelIndex, QVariant
from PyQt5.QtGui import QColor, QBrush, QFont

from core.models import StagingCadetDTO, gender_to_ru, normalize_gender
from core.parsers import (
    increment_course_all,
    bulk_set_course,
    remove_graduates,
    toggle_gender
)


class CourseDelegate(QStyledItemDelegate):
    def createEditor(self, parent, option, index):
        editor = QSpinBox(parent)
        editor.setRange(1, 6)
        editor.setStyleSheet("font-size: 13px; padding: 2px;")
        return editor

    def setEditorData(self, editor, index):
        value = index.model().data(index, Qt.EditRole)
        try:
            editor.setValue(int(value))
        except (ValueError, TypeError):
            editor.setValue(1)

    def setModelData(self, editor, index, model):
        model.setData(index, editor.value(), Qt.EditRole)

    def updateEditorGeometry(self, editor, option, index):
        editor.setGeometry(option.rect)


class GenderDelegate(QStyledItemDelegate):
    def createEditor(self, parent, option, index):
        editor = QComboBox(parent)
        editor.addItems(["М (Мужской)", "Ж (Женский)"])
        editor.setStyleSheet("font-size: 13px;")
        return editor

    def setEditorData(self, editor, index):
        val = str(index.model().data(index, Qt.DisplayRole)).strip()
        if "Ж" in val.upper() or "F" in val.upper():
            editor.setCurrentIndex(1)
        else:
            editor.setCurrentIndex(0)

    def setModelData(self, editor, index, model):
        chosen = "female" if editor.currentIndex() == 1 else "male"
        model.setData(index, chosen, Qt.EditRole)

    def updateEditorGeometry(self, editor, option, index):
        editor.setGeometry(option.rect)


class CadetStagingTableModel(QAbstractTableModel):
    COLUMNS = [
        "№", "Звание", "ФИО", "Группа", "Курс", "Пол", "Статус", "Проверка / Предупреждения"
    ]

    def __init__(self, cadets: List[StagingCadetDTO], parent=None):
        super().__init__(parent)
        self.cadets: List[StagingCadetDTO] = [c.model_copy(deep=True) for c in cadets]

    def rowCount(self, parent=QModelIndex()):
        return len(self.cadets)

    def columnCount(self, parent=QModelIndex()):
        return len(self.COLUMNS)

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self.COLUMNS[section]
        return QVariant()

    def flags(self, index):
        if not index.isValid():
            return Qt.NoItemFlags
        base_flags = Qt.ItemIsEnabled | Qt.ItemIsSelectable
        # Course (4) and Gender (5) are inline editable
        if index.column() in (4, 5):
            return base_flags | Qt.ItemIsEditable
        return base_flags

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid() or not (0 <= index.row() < len(self.cadets)):
            return QVariant()

        c = self.cadets[index.row()]
        col = index.column()

        if role in (Qt.DisplayRole, Qt.EditRole):
            if col == 0:
                return str(index.row() + 1)
            elif col == 1:
                return c.rank
            elif col == 2:
                return c.fio
            elif col == 3:
                return c.group_num
            elif col == 4:
                return c.course
            elif col == 5:
                return "Ж" if normalize_gender(c.gender) == "female" else "М"
            elif col == 6:
                if c.status == "graduated":
                    return "Выпускник"
                elif c.status == "expelled":
                    return "Отчислен"
                return "В строю"
            elif col == 7:
                return "; ".join(c.warnings) if c.warnings else "OK"

        elif role == Qt.TextAlignmentRole:
            if col in (0, 4, 5, 6):
                return Qt.AlignCenter
            return Qt.AlignLeft | Qt.AlignVCenter

        elif role == Qt.ForegroundRole:
            if col == 5:
                # Color code gender
                if normalize_gender(c.gender) == "female":
                    return QBrush(QColor("#BE185D"))  # Pink/Magenta
                return QBrush(QColor("#1D4ED8"))  # Blue
            elif col == 6 and c.status == "graduated":
                return QBrush(QColor("#B45309"))  # Amber
            elif col == 7:
                if c.warnings:
                    return QBrush(QColor("#DC2626"))  # Red warning
                return QBrush(QColor("#16A34A"))  # Green OK

        elif role == Qt.BackgroundRole:
            if c.status == "graduated":
                return QBrush(QColor("#FEF3C7"))  # Light amber highlight
            elif c.needs_manual_review:
                return QBrush(QColor("#FFFBEB"))  # Subtle yellow

        return QVariant()

    def setData(self, index, value, role=Qt.EditRole):
        if not index.isValid() or role != Qt.EditRole:
            return False

        c = self.cadets[index.row()]
        col = index.column()

        if col == 4:
            try:
                new_course = int(value)
                c.course = new_course
                if new_course > 5:
                    c.status = "graduated"
                    warn = f"Курс {new_course} > 5 (выпускник)"
                    if warn not in c.warnings:
                        c.warnings.append(warn)
                else:
                    c.status = "active"
                    c.warnings = [w for w in c.warnings if "выпускник" not in w.lower()]
                self.dataChanged.emit(index, self.index(index.row(), 7))
                return True
            except (ValueError, TypeError):
                return False

        elif col == 5:
            norm = normalize_gender(value)
            c.gender = norm
            c.needs_manual_review = False
            c.warnings = [w for w in c.warnings if "пол определен" not in w.lower()]
            self.dataChanged.emit(index, self.index(index.row(), 7))
            return True

        return False

    def refresh_data(self, new_cadets: List[StagingCadetDTO]):
        self.beginResetModel()
        self.cadets = new_cadets
        self.endResetModel()


class ImportReviewDialog(QDialog):
    """Interactive Staging Wizard dialog for Excel list import with:
    - Quick course increment (+1 курс всем)
    - Bulk course override (1..5)
    - Remove graduates
    - Inline course and gender editing
    """
    def __init__(self, parent=None, staging_cadets: Optional[List[StagingCadetDTO]] = None):
        super().__init__(parent)
        self.setWindowTitle("Интерактивный предпросмотр импорта (Staging Wizard)")
        self.resize(1100, 650)
        self.setMinimumSize(850, 500)

        self.initial_cadets: List[StagingCadetDTO] = staging_cadets or []
        self.model = CadetStagingTableModel(self.initial_cadets, self)

        self.init_ui()
        self.update_summary_label()

    def init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(14, 14, 14, 14)
        main_layout.setSpacing(12)

        # Header Info Card
        header_card = QFrame()
        header_card.setStyleSheet(
            "background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 10px;"
        )
        h_layout = QHBoxLayout(header_card)
        h_layout.setContentsMargins(6, 6, 6, 6)

        title_lbl = QLabel("<b>Мастер предварительного просмотра и подготовки личного состава</b>")
        title_lbl.setStyleSheet("font-size: 13px; color: #1E293B;")
        h_layout.addWidget(title_lbl)
        h_layout.addStretch()

        self.summary_badge = QLabel("")
        self.summary_badge.setStyleSheet(
            "background: #EFF6FF; color: #1E40AF; border: 1px solid #BFDBFE; "
            "border-radius: 6px; padding: 4px 10px; font-weight: 600; font-size: 12px;"
        )
        h_layout.addWidget(self.summary_badge)
        main_layout.addWidget(header_card)

        # Toolbar
        toolbar = QToolBar()
        toolbar.setStyleSheet(
            "QToolBar { background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 8px; padding: 6px; gap: 8px; }"
        )
        toolbar.setMovable(False)

        # 1. Кнопка +1 курс всем
        btn_inc = QPushButton("⏩ +1 курс всем (Перевод на след. год)")
        btn_inc.setStyleSheet(
            "background-color: #2563EB; color: white; font-weight: 600; padding: 6px 14px; border-radius: 6px;"
        )
        btn_inc.clicked.connect(self.on_increment_course_all)
        toolbar.addWidget(btn_inc)

        toolbar.addSeparator()

        # 2. Селектор и кнопка массовой установки курса
        lbl_bulk = QLabel("  Назначить курс: ")
        lbl_bulk.setStyleSheet("font-size: 12px; font-weight: 600; color: #475569;")
        toolbar.addWidget(lbl_bulk)

        self.combo_target_course = QComboBox()
        self.combo_target_course.addItems(["1 курс", "2 курс", "3 курс", "4 курс", "5 курс"])
        self.combo_target_course.setCurrentIndex(0)
        self.combo_target_course.setStyleSheet("padding: 4px 8px; font-size: 12px; min-width: 90px;")
        toolbar.addWidget(self.combo_target_course)

        btn_apply_course = QPushButton("Применить")
        btn_apply_course.setStyleSheet(
            "background-color: #F1F5F9; color: #334155; border: 1px solid #CBD5E1; "
            "font-weight: 600; padding: 6px 12px; border-radius: 6px;"
        )
        btn_apply_course.clicked.connect(self.on_bulk_set_course)
        toolbar.addWidget(btn_apply_course)

        toolbar.addSeparator()

        # 3. Удалить выпускников
        btn_remove_grad = QPushButton("🎓 Исключить выпускников")
        btn_remove_grad.setStyleSheet(
            "background-color: #FEF3C7; color: #92400E; border: 1px solid #FDE68A; "
            "font-weight: 600; padding: 6px 12px; border-radius: 6px;"
        )
        btn_remove_grad.clicked.connect(self.on_remove_graduates)
        toolbar.addWidget(btn_remove_grad)

        # 4. Сменить пол выделенным
        btn_toggle_gender = QPushButton("⚧ Переключить пол выделенным")
        btn_toggle_gender.setStyleSheet(
            "background-color: #F3E8FF; color: #6B21A8; border: 1px solid #E9D5FF; "
            "font-weight: 600; padding: 6px 12px; border-radius: 6px;"
        )
        btn_toggle_gender.clicked.connect(self.on_toggle_selected_gender)
        toolbar.addWidget(btn_toggle_gender)

        main_layout.addWidget(toolbar)

        # Table View
        self.table_view = QTableView()
        self.table_view.setModel(self.model)
        self.table_view.setAlternatingRowColors(True)
        self.table_view.setItemDelegateForColumn(4, CourseDelegate(self.table_view))
        self.table_view.setItemDelegateForColumn(5, GenderDelegate(self.table_view))
        self.table_view.setSelectionBehavior(QTableView.SelectRows)
        self.table_view.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table_view.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table_view.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table_view.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.table_view.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.table_view.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)
        self.table_view.horizontalHeader().setSectionResizeMode(6, QHeaderView.ResizeToContents)
        main_layout.addWidget(self.table_view)

        # Bottom Action Bar
        bottom_bar = QHBoxLayout()
        bottom_bar.setSpacing(10)

        hint_lbl = QLabel("💡 <i>Кликните дважды на ячейку «Курс» или «Пол» для мгновенного изменения</i>")
        hint_lbl.setStyleSheet("color: #64748B; font-size: 12px;")
        bottom_bar.addWidget(hint_lbl)
        bottom_bar.addStretch()

        btn_cancel = QPushButton("Отмена")
        btn_cancel.setStyleSheet("padding: 8px 18px; font-weight: 600; color: #475569;")
        btn_cancel.clicked.connect(self.reject)
        bottom_bar.addWidget(btn_cancel)

        btn_confirm = QPushButton("Подтвердить и внести в расход")
        btn_confirm.setStyleSheet(
            "background-color: #059669; color: white; font-weight: 700; padding: 8px 22px; border-radius: 6px;"
        )
        btn_confirm.clicked.connect(self.on_confirm)
        bottom_bar.addWidget(btn_confirm)

        main_layout.addLayout(bottom_bar)

    def update_summary_label(self):
        total = len(self.model.cadets)
        graduated = sum(1 for c in self.model.cadets if c.status == "graduated" or c.course > 5)
        males = sum(1 for c in self.model.cadets if normalize_gender(c.gender) == "male")
        females = sum(1 for c in self.model.cadets if normalize_gender(c.gender) == "female")
        needs_review = sum(1 for c in self.model.cadets if c.needs_manual_review)

        self.summary_badge.setText(
            f"Всего: {total} чел. | 👨 М: {males} | 👩 Ж: {females}"
            + (f" | ⚠️ Выпускники: {graduated}" if graduated else "")
            + (f" | 🔍 Проверить пол: {needs_review}" if needs_review else "")
        )

    def on_increment_course_all(self):
        self.model.refresh_data(increment_course_all(self.model.cadets))
        self.update_summary_label()
        QMessageBox.information(
            self, "Перевод на следующий курс",
            "Курс всех активных курсантов увеличен на +1.\n"
            "Курсанты, превысившие 5-й курс, переведены в статус «Выпускник»."
        )

    def on_bulk_set_course(self):
        target_course = self.combo_target_course.currentIndex() + 1
        selected_indexes = self.table_view.selectionModel().selectedRows()

        if selected_indexes:
            selected_ids = [self.model.cadets[idx.row()].temp_id for idx in selected_indexes]
            self.model.refresh_data(bulk_set_course(self.model.cadets, target_course, selected_ids))
            msg = f"Курс {target_course} назначен для выбранных ({len(selected_ids)} чел.)."
        else:
            self.model.refresh_data(bulk_set_course(self.model.cadets, target_course, None))
            msg = f"Курс {target_course} назначен для ВСЕХ курсантов."

        self.update_summary_label()
        QMessageBox.information(self, "Курс обновлен", msg)

    def on_remove_graduates(self):
        count_before = len(self.model.cadets)
        filtered = remove_graduates(self.model.cadets)
        removed = count_before - len(filtered)
        self.model.refresh_data(filtered)
        self.update_summary_label()
        QMessageBox.information(self, "Исключение выпускников", f"Исключено выпускников из списка: {removed} чел.")

    def on_toggle_selected_gender(self):
        selected_indexes = self.table_view.selectionModel().selectedRows()
        if not selected_indexes:
            QMessageBox.information(self, "Информация", "Выделите строки в таблице для смены пола.")
            return

        for idx in selected_indexes:
            row = idx.row()
            self.model.cadets[row] = toggle_gender(self.model.cadets[row])

        self.model.refresh_data(self.model.cadets)
        self.update_summary_label()

    def on_confirm(self):
        graduates = [c for c in self.model.cadets if c.status == "graduated" or c.course > 5]
        if graduates:
            ans = QMessageBox.question(
                self, "Внимание: в списке есть выпускники",
                f"В списке обнаружено {len(graduates)} выпускников (курс > 5).\n"
                f"Исключить их перед внесением в расход?",
                QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel,
                QMessageBox.Yes
            )
            if ans == QMessageBox.Cancel:
                return
            elif ans == QMessageBox.Yes:
                self.model.refresh_data(remove_graduates(self.model.cadets))

        self.accept()

    def get_staged_cadets(self) -> List[StagingCadetDTO]:
        return self.model.cadets
