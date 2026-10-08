from __future__ import annotations

import json
import math
import subprocess
import sys
import uuid
from pathlib import Path
from typing import Callable

from PySide6.QtCore import QDate, QPoint, QPointF, QRect, QThread, QTimer, QUrl, Qt, Signal
from PySide6.QtGui import QColor, QDesktopServices, QFont, QFontDatabase, QFontMetrics, QIcon, QPainter, QPen, QPixmap
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QComboBox,
    QDateEdit,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSpinBox,
    QSplitter,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QToolButton,
    QToolTip,
    QVBoxLayout,
    QWidget,
)

from .config_store import ConfigStore
from .engine import COMPARE_MODES, DataEngine, compare_periods, is_session_header, list_chart_headers
from .models import AnalysisResult, Record, SourceConfig
from .source_reader import (
    SourceReader,
    excel_column,
    normalize_column_schema,
    schema_field_names,
    split_names,
)
from .version import APP_VERSION, RELEASES_URL, download_release_installer, fetch_latest_release, is_newer


APP_TITLE = "表数通"
QUERY_SOURCES = ["extract", "aggregate", "direct"]
RANGE_LABELS = ["当月", "最近7天", "最近2天", "自定义"]
UI_RANGE_ALIASES = {"week": "month", "month": "month", "days7": "days7", "days2": "days2", "custom": "custom"}
CHIP_VISIBLE = 4
DELTA_UP = QColor("#15803d")
DELTA_DOWN = QColor("#dc2626")
DELTA_FLAT = QColor("#64748b")
PIE_COLORS = [
    QColor("#087fbb"), QColor("#f59e0b"), QColor("#10b981"), QColor("#8b5cf6"),
    QColor("#ef4444"), QColor("#14b8a6"), QColor("#f97316"), QColor("#6366f1"),
    QColor("#84cc16"),
]
COLUMN_LETTERS = [excel_column(index) for index in range(52)]
DEFAULT_QUERY_RESULT_FIELDS = ["输入电话号码", "专页ID", "姓名", "评论贴文", "号码", "日期", "修正格式"]
PHONE_HEADERS = {"号码", "手机号", "手机号码", "电话", "联系电话", "phone", "number"}
PHONE_TABLE_NAME_HINTS = ("号码表", "号码数据")
PAGE_ID_HEADERS = {"专页id", "pageid"}


def is_phone_data_table(fields: list[str] | None = None, source_name: str = "") -> bool:
    name = str(source_name or "")
    if any(hint in name for hint in PHONE_TABLE_NAME_HINTS):
        return True
    folded = {str(item).strip().casefold() for item in (fields or [])}
    has_phone = any(item.casefold() in folded for item in PHONE_HEADERS)
    has_page = any(item.casefold() in folded for item in PAGE_ID_HEADERS)
    return has_phone and has_page


def query_result_headers(
    store: ConfigStore,
    source: str,
    query_field: str,
    result_fields: list[str],
    source_name: str = "",
) -> list[str]:
    table_fields = list(result_fields or [])
    if is_phone_data_table(table_fields, source_name):
        return DEFAULT_QUERY_RESULT_FIELDS.copy()
    return table_fields or [query_field]


class ColumnMapRow(QFrame):
    changed = Signal()
    move_requested = Signal(object, int)
    remove_requested = Signal(object)

    def __init__(self, name: str = "", column: str = "A", enabled: bool = True, parent=None):
        super().__init__(parent)
        self.setObjectName("columnMapRow")
        layout = QHBoxLayout(self)
        layout.setContentsMargins(10, 8, 10, 8)
        layout.setSpacing(8)
        self.enabled_box = QCheckBox()
        self.enabled_box.setChecked(enabled)
        self.name_edit = QLineEdit(name)
        self.name_edit.setPlaceholderText("表头名称，例如：贴文ID")
        self.column_box = QComboBox()
        self.column_box.setEditable(True)
        self.column_box.setInsertPolicy(QComboBox.NoInsert)
        self.column_box.addItems(COLUMN_LETTERS)
        self.column_box.setCurrentText((column or "A").upper())
        self.column_box.setFixedWidth(72)
        self.column_box.setObjectName("columnLetter")
        up = QToolButton()
        up.setText("↑")
        down = QToolButton()
        down.setText("↓")
        delete = QToolButton()
        delete.setText("×")
        delete.setToolTip("删除这一行")
        layout.addWidget(self.enabled_box)
        layout.addWidget(self.name_edit, 1)
        layout.addWidget(self.column_box)
        layout.addWidget(up)
        layout.addWidget(down)
        layout.addWidget(delete)
        self.enabled_box.toggled.connect(self.changed)
        self.name_edit.textChanged.connect(self.changed)
        self.column_box.currentTextChanged.connect(self.changed)
        up.clicked.connect(lambda: self.move_requested.emit(self, -1))
        down.clicked.connect(lambda: self.move_requested.emit(self, 1))
        delete.clicked.connect(lambda: self.remove_requested.emit(self))

    def value(self) -> dict[str, object]:
        return {
            "name": self.name_edit.text().strip(),
            "column": self.column_box.currentText().strip().upper() or "A",
            "enabled": self.enabled_box.isChecked(),
        }


class ColumnMapWidget(QWidget):
    changed = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(8)
        header = QHBoxLayout()
        enabled_label = QLabel("选用")
        enabled_label.setFixedWidth(36)
        name_label = QLabel("表头名称")
        column_label = QLabel("对应列")
        column_label.setFixedWidth(72)
        header.addWidget(enabled_label)
        header.addWidget(name_label, 1)
        header.addWidget(column_label)
        header.addSpacing(92)
        outer.addLayout(header)
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setMinimumHeight(180)
        self.scroll.setMaximumHeight(280)
        self.list_host = QWidget()
        self.rows_layout = QVBoxLayout(self.list_host)
        self.rows_layout.setContentsMargins(0, 0, 0, 0)
        self.rows_layout.setSpacing(6)
        self.rows_layout.addStretch()
        self.scroll.setWidget(self.list_host)
        outer.addWidget(self.scroll)
        actions = QHBoxLayout()
        add = QPushButton("添加字段")
        add.clicked.connect(lambda: self.add_row())
        actions.addWidget(add)
        actions.addStretch()
        outer.addLayout(actions)
        self._rows: list[ColumnMapRow] = []

    def add_row(self, name: str = "", column: str = "", enabled: bool = True) -> None:
        letter = column or excel_column(len(self._rows))
        row = ColumnMapRow(name, letter, enabled, self.list_host)
        row.changed.connect(self.changed)
        row.move_requested.connect(self.move_row)
        row.remove_requested.connect(self.remove_row)
        self.rows_layout.insertWidget(self.rows_layout.count() - 1, row)
        self._rows.append(row)
        self.changed.emit()

    def move_row(self, row: ColumnMapRow, step: int) -> None:
        index = self._rows.index(row)
        target = index + step
        if target < 0 or target >= len(self._rows):
            return
        self._rows[index], self._rows[target] = self._rows[target], self._rows[index]
        self.rows_layout.removeWidget(row)
        self.rows_layout.insertWidget(target, row)
        self.changed.emit()

    def remove_row(self, row: ColumnMapRow) -> None:
        if row in self._rows:
            self._rows.remove(row)
        self.rows_layout.removeWidget(row)
        row.deleteLater()
        self.changed.emit()

    def set_schema(self, schema: list[object] | None) -> None:
        for row in list(self._rows):
            self.remove_row(row)
        entries = normalize_column_schema(schema)
        if not entries:
            self.add_row("", "A", True)
            return
        for entry in entries:
            self.add_row(str(entry.get("name") or ""), str(entry.get("column") or "A"), bool(entry.get("enabled", True)))

    def schema(self) -> list[dict[str, object]]:
        return [row.value() for row in self._rows]

    def setEnabled(self, enabled: bool) -> None:
        super().setEnabled(enabled)
        self.scroll.setEnabled(enabled)


class TaskThread(QThread):
    succeeded = Signal(object)
    failed = Signal(str)

    def __init__(self, job: Callable[[], object], parent=None):
        super().__init__(parent)
        self.job = job

    def run(self) -> None:
        try:
            self.succeeded.emit(self.job())
        except Exception as exc:
            self.failed.emit(str(exc))


class SourceDialog(QDialog):
    def __init__(self, store: ConfigStore, source: SourceConfig | None = None, parent=None):
        super().__init__(parent)
        self.store = store
        self.source = source
        self.setWindowTitle("编辑数据源" if source else "添加数据源")
        self.setMinimumWidth(760)
        self.setMinimumHeight(640)
        form = QFormLayout(self)
        self.name = QLineEdit(source.name if source else "")
        self.url = QLineEdit(source.url if source else "")
        self.url.setPlaceholderText("Google 表格链接；也支持本地 .xlsx 文件")
        self.include = QTextEdit("\n".join(source.include_sheets) if source else "")
        self.include.setMaximumHeight(85)
        self.include.setPlaceholderText("留空＝遍历全部；多个名称每行一个")
        self.exclude = QTextEdit("\n".join(source.exclude_sheets) if source else "")
        self.exclude.setMaximumHeight(85)
        self.exclude.setPlaceholderText("多个名称每行一个；排除规则优先")
        self.header_row = QSpinBox()
        self.header_row.setRange(1, 100)
        self.header_row.setValue(source.header_row if source else 1)
        self.credential = QLineEdit(source.credential_path if source else "")
        self.credential.setPlaceholderText("留空则使用设置里的服务账号池（轮询）")
        credential_button = QPushButton("选择…")
        credential_button.clicked.connect(self.choose_credential)
        credential_row = QHBoxLayout()
        credential_row.addWidget(self.credential, 1)
        credential_row.addWidget(credential_button)
        self.enabled = QCheckBox("启用此数据源")
        self.enabled.setChecked(source.enabled if source else True)
        form.addRow("数据源名称*", self.name)
        form.addRow("表格链接/文件*", self.url)
        form.addRow("指定 Sheet", self.include)
        form.addRow("排除 Sheet", self.exclude)
        form.addRow("表头所在行", self.header_row)
        form.addRow("服务账号 JSON", credential_row)
        form.addRow("", self.enabled)
        self.use_own_schema = QCheckBox("使用独立列结构（此表可指定自己的表头和对应列）")
        self.use_own_schema.setChecked(bool(source.column_schema_enabled) if source else False)
        self.schema_editor = ColumnMapWidget()
        self.schema_editor.set_schema(source.column_schema if source else [])
        form.addRow("", self.use_own_schema)
        form.addRow("字段分配", self.schema_editor)
        hint = QLabel("勾选要用的字段，填写表头名称，再选择对应的表格列（A、B、C…）。可用上下箭头调整顺序。独立列结构只作用于当前数据源。")
        hint.setObjectName("muted")
        hint.setWordWrap(True)
        form.addRow("", hint)
        self.use_own_schema.toggled.connect(self.schema_editor.setEnabled)
        self.schema_editor.setEnabled(self.use_own_schema.isChecked())
        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.validate_and_accept)
        buttons.rejected.connect(self.reject)
        form.addRow(buttons)

    def choose_credential(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "选择服务账号 JSON", "", "JSON 文件 (*.json)")
        if path:
            self.credential.setText(path)

    def validate_and_accept(self) -> None:
        if not self.name.text().strip() or not self.url.text().strip():
            QMessageBox.warning(self, "缺少信息", "请填写数据源名称和表格链接。")
            return
        if self.use_own_schema.isChecked() and not schema_field_names(self.schema_editor.schema()):
            QMessageBox.warning(self, "缺少列结构", "启用独立列结构时，请至少勾选并填写一个表头名称。")
            return
        self.accept()

    def value(self) -> SourceConfig:
        return SourceConfig(
            id=self.source.id if self.source else uuid.uuid4().hex,
            name=self.name.text().strip(),
            url=self.url.text().strip(),
            include_sheets=split_names(self.include.toPlainText()),
            exclude_sheets=split_names(self.exclude.toPlainText()),
            header_row=self.header_row.value(),
            enabled=self.enabled.isChecked(),
            credential_path=self.credential.text().strip(),
            column_schema_enabled=self.use_own_schema.isChecked(),
            column_schema=self.schema_editor.schema(),
        )


class MainWindow(QMainWindow):
    def __init__(self, store: ConfigStore, icon_path: Path | None = None):
        super().__init__()
        self.store = store
        self.tasks: list[TaskThread] = []
        self.icon_path = icon_path
        self._restoring_settings = True
        self.setWindowTitle(f"{APP_TITLE} · 多表数据查询与提取")
        self.resize(1220, 780)
        self.setMinimumSize(980, 650)
        if icon_path and icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))
        self._build_ui()
        self.refresh_sources()
        self.refresh_field_controls()
        self.refresh_logs()
        self._restoring_settings = False
        QTimer.singleShot(1500, self.check_updates_silent)

    def _build_ui(self) -> None:
        central = QWidget()
        outer = QHBoxLayout(central)
        outer.setContentsMargins(0, 0, 0, 0)
        sidebar = QFrame()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(210)
        sidebar_layout = QVBoxLayout(sidebar)
        sidebar_layout.setContentsMargins(20, 24, 20, 20)
        brand_row = QHBoxLayout()
        if self.icon_path and self.icon_path.exists():
            logo = QLabel()
            logo.setPixmap(QPixmap(str(self.icon_path)).scaled(46, 46, Qt.KeepAspectRatio, Qt.SmoothTransformation))
            brand_row.addWidget(logo)
        brand_text = QVBoxLayout()
        brand = QLabel(APP_TITLE)
        brand.setObjectName("brand")
        subtitle = QLabel("多表数据工作台")
        subtitle.setObjectName("sidebarMuted")
        brand_text.addWidget(brand)
        brand_text.addWidget(subtitle)
        brand_row.addLayout(brand_text, 1)
        self.nav = QListWidget()
        self.nav.setObjectName("navigation")
        self.nav.addItems(["数据源", "汇总同步", "数据查询", "数据分析", "时间提取", "字段配置", "运行日志", "设置"])
        self.nav.setCurrentRow(0)
        sidebar_layout.addLayout(brand_row)
        sidebar_layout.addSpacing(22)
        sidebar_layout.addWidget(self.nav, 1)
        self.version_label = QLabel(f"v{APP_VERSION}")
        self.version_label.setObjectName("sidebarMuted")
        self.sidebar_update_button = QPushButton("检查并安装更新")
        self.sidebar_update_button.clicked.connect(self.check_updates)
        sidebar_layout.addWidget(self.version_label)
        sidebar_layout.addWidget(self.sidebar_update_button)
        self.pages = QStackedWidget()
        self.pages.addWidget(self._sources_page())
        self.pages.addWidget(self._sync_page())
        self.pages.addWidget(self._query_page())
        self.pages.addWidget(self._analysis_page())
        self.pages.addWidget(self._extract_page())
        self.pages.addWidget(self._fields_page())
        self.pages.addWidget(self._logs_page())
        self.pages.addWidget(self._settings_page())
        self.nav.currentRowChanged.connect(self.pages.setCurrentIndex)
        outer.addWidget(sidebar)
        outer.addWidget(self.pages, 1)
        self.setCentralWidget(central)
        self.statusBar().showMessage("就绪")

    def _page(self, title: str, description: str) -> tuple[QWidget, QVBoxLayout]:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(30, 26, 30, 24)
        title_label = QLabel(title)
        title_label.setObjectName("pageTitle")
        desc_label = QLabel(description)
        desc_label.setObjectName("muted")
        desc_label.setWordWrap(True)
        layout.addWidget(title_label)
        layout.addWidget(desc_label)
        layout.addSpacing(14)
        return page, layout

    def _sources_page(self) -> QWidget:
        page, layout = self._page("数据源", "配置一个或多个 Google 表格链接；每个数据源都可以有自己的列数和表头。")
        actions = QHBoxLayout()
        add = QPushButton("＋ 添加数据源")
        add.setObjectName("primary")
        edit = QPushButton("编辑")
        remove = QPushButton("删除")
        scan = QPushButton("扫描 Sheet")
        add.clicked.connect(self.add_source)
        edit.clicked.connect(self.edit_source)
        remove.clicked.connect(self.delete_source)
        scan.clicked.connect(self.scan_source)
        actions.addWidget(add)
        actions.addWidget(edit)
        actions.addWidget(remove)
        actions.addWidget(scan)
        actions.addStretch()
        self.source_table = QTableWidget(0, 8)
        self.source_table.setHorizontalHeaderLabels(["启用", "名称", "表格链接", "指定 Sheet", "排除 Sheet", "表头行", "列结构", "认证"])
        self.source_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.source_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.source_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.source_table.verticalHeader().setVisible(False)
        self.source_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.source_table.doubleClicked.connect(self.edit_source)
        rule = QLabel("读取规则：指定名称留空时遍历全部；填写后仅抓取指定项；排除项始终优先。不同表格列数不一样时，在编辑数据源里勾选「独立列结构」。")
        rule.setObjectName("infoBox")
        layout.addLayout(actions)
        layout.addWidget(self.source_table, 1)
        layout.addWidget(rule)
        return page

    def _sync_page(self) -> QWidget:
        page, layout = self._page("汇总同步", "勾选要汇总的数据源，并可指定写入的 Google 表格或本地文件。同步同时会更新查询缓存。")
        card = QFrame()
        card.setObjectName("card")
        form = QFormLayout(card)
        self.sync_source_list = QListWidget()
        self.sync_source_list.setSelectionMode(QAbstractItemView.NoSelection)
        self.sync_source_list.setMinimumHeight(140)
        pick_row = QHBoxLayout()
        select_all = QPushButton("全选")
        select_none = QPushButton("全不选")
        select_all.clicked.connect(lambda: self.set_sync_sources_checked(True))
        select_none.clicked.connect(lambda: self.set_sync_sources_checked(False))
        pick_row.addWidget(select_all)
        pick_row.addWidget(select_none)
        pick_row.addStretch()
        self.write_aggregate = QCheckBox("同时写入本地汇总数据库")
        self.write_aggregate.setChecked(bool(self.store.get("write_aggregate", True)))
        self.write_aggregate.toggled.connect(lambda checked: self.store.set("write_aggregate", checked))
        self.max_rows = QSpinBox()
        self.max_rows.setRange(1000, 10_000_000)
        self.max_rows.setSingleStep(10000)
        self.max_rows.setValue(int(self.store.get("max_rows_per_db", 500000)))
        self.max_rows.valueChanged.connect(lambda value: self.store.set("max_rows_per_db", value))
        self.sync_google_url = QLineEdit(str(self.store.get("sync_google_url", "") or ""))
        self.sync_google_url.setPlaceholderText("可选：https://docs.google.com/spreadsheets/d/...")
        self.sync_google_sheet = QLineEdit(str(self.store.get("sync_google_sheet", "汇总结果") or "汇总结果"))
        self.sync_local_xlsx = QLineEdit(str(self.store.get("sync_local_xlsx", "") or ""))
        self.sync_local_xlsx.setPlaceholderText("可选：本地 Excel 路径")
        choose = QPushButton("选择…")
        choose.clicked.connect(self.choose_sync_xlsx)
        xlsx_row = QHBoxLayout()
        xlsx_row.addWidget(self.sync_local_xlsx, 1)
        xlsx_row.addWidget(choose)
        form.addRow("选择数据源", self.sync_source_list)
        form.addRow("", pick_row)
        form.addRow("本地汇总库", self.write_aggregate)
        form.addRow("单库最大行数", self.max_rows)
        form.addRow("写入 Google 表格", self.sync_google_url)
        form.addRow("目标工作表", self.sync_google_sheet)
        form.addRow("写入本地 Excel", xlsx_row)
        hint = QLabel("至少勾选一个数据源。Google 表格和本地 Excel 都是可选写入目标；不填则只更新缓存/汇总库。")
        hint.setObjectName("muted")
        hint.setWordWrap(True)
        form.addRow("", hint)
        self.sync_button = QPushButton("开始同步并更新缓存")
        self.sync_button.setObjectName("primary")
        self.sync_button.clicked.connect(self.run_sync)
        self.sync_status = QTextEdit()
        self.sync_status.setReadOnly(True)
        self.sync_status.setPlaceholderText("运行明细会显示在这里，并同步写入运行日志。")
        layout.addWidget(card)
        layout.addWidget(self.sync_button, 0, Qt.AlignLeft)
        layout.addWidget(self.sync_status, 1)
        return page

    def _query_page(self) -> QWidget:
        page, layout = self._page(
            "数据查询",
            "查询走本地库。点「同步本地库」会下载表格最新数据并替换本地库。只有号码数据表才复制号码和修正格式。",
        )
        bar = QHBoxLayout()
        self.query_field = QComboBox()
        self.query_field.setEditable(True)
        self.query_field.setInsertPolicy(QComboBox.NoInsert)
        self.query_field.setMinimumWidth(140)
        self.query_value = QTextEdit()
        self.query_value.setMaximumHeight(78)
        self.query_value.setPlaceholderText("每行一个查询值，也支持逗号分隔批量粘贴")
        self.query_mode = QComboBox()
        self.query_mode.addItems(["查询提取表", "查询汇总数据库", "查询数据源"])
        saved_source = str(self.store.get("query_source", "extract"))
        if saved_source in QUERY_SOURCES:
            self.query_mode.setCurrentIndex(QUERY_SOURCES.index(saved_source))
        self.query_table_label = QLabel("数据表")
        self.query_source_pick = QComboBox()
        self.query_source_pick.setMinimumWidth(140)
        self.query_fuzzy = QCheckBox("模糊匹配")
        self.query_fuzzy.setChecked(bool(self.store.get("query_fuzzy", False)))
        self.query_fuzzy.setToolTip("勾选后可用姓名、编号或部分文字查询；查询来源时也会匹配修正格式")
        self.query_date_enabled = QCheckBox("限制日期范围")
        self.query_date_enabled.setChecked(bool(self.store.get("query_date_enabled", False)))
        self.query_date_field = QComboBox()
        self.query_date_field.setEditable(True)
        self.query_date_field.setMinimumWidth(120)
        self.query_start_date = QDateEdit(QDate.currentDate().addMonths(-1))
        self.query_end_date = QDateEdit(QDate.currentDate())
        for widget, key in (
            (self.query_start_date, "query_start_date"),
            (self.query_end_date, "query_end_date"),
        ):
            widget.setCalendarPopup(True)
            widget.setDisplayFormat("yyyy-MM-dd")
            saved = QDate.fromString(str(self.store.get(key, "") or ""), "yyyy-MM-dd")
            if saved.isValid():
                widget.setDate(saved)
        self.query_mode.currentIndexChanged.connect(self.on_query_mode_changed)
        self.query_source_pick.currentIndexChanged.connect(self.on_query_table_changed)
        self.query_fuzzy.toggled.connect(self.persist_workspace_settings)
        self.query_date_enabled.toggled.connect(self.update_query_date_controls)
        self.query_date_enabled.toggled.connect(self.persist_workspace_settings)
        self.query_date_field.currentTextChanged.connect(self.persist_workspace_settings)
        self.query_start_date.dateChanged.connect(self.persist_workspace_settings)
        self.query_end_date.dateChanged.connect(self.persist_workspace_settings)
        self.query_field.currentTextChanged.connect(self.persist_workspace_settings)
        button = QPushButton("查询")
        button.setObjectName("primary")
        button.clicked.connect(self.run_query)
        self.refresh_cache_button = QPushButton("同步本地库")
        self.refresh_cache_button.setToolTip("从表格下载最新数据并替换本地库。内容完全相同则跳过写入。")
        self.refresh_cache_button.clicked.connect(self.run_refresh_cache)
        self.copy_query_button = QPushButton("一键复制结果")
        self.copy_query_button.setEnabled(False)
        self.copy_query_button.clicked.connect(self.copy_query_results)
        bar.addWidget(QLabel("表格来源"))
        bar.addWidget(self.query_mode)
        bar.addWidget(self.query_table_label)
        bar.addWidget(self.query_source_pick)
        bar.addWidget(QLabel("查询字段"))
        bar.addWidget(self.query_field)
        bar.addWidget(self.query_value, 1)
        bar.addWidget(self.query_fuzzy)
        bar.addWidget(self.refresh_cache_button)
        bar.addWidget(button)
        bar.addWidget(self.copy_query_button)
        date_bar = QHBoxLayout()
        date_bar.addWidget(self.query_date_enabled)
        date_bar.addWidget(QLabel("日期字段"))
        date_bar.addWidget(self.query_date_field)
        date_bar.addWidget(QLabel("日期范围"))
        date_bar.addWidget(self.query_start_date)
        date_bar.addWidget(QLabel("至"))
        date_bar.addWidget(self.query_end_date)
        date_bar.addWidget(QLabel("排除关键词"))
        self.query_exclude = QLineEdit(str(self.store.get("query_exclude_keywords", "") or ""))
        self.query_exclude.setPlaceholderText("多个用逗号分隔，命中任一则排除")
        self.query_exclude.setMinimumWidth(180)
        self.query_exclude.editingFinished.connect(self.persist_workspace_settings)
        date_bar.addWidget(self.query_exclude, 1)
        date_bar.addStretch()
        self.query_table = QTableWidget()
        self.query_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.query_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.query_table.verticalHeader().setVisible(False)
        self.query_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        hint = QLabel("其他表即使查询号码字段，也按该表表头显示。只有号码数据表才显示修正格式，并提供“复制号码和修正格式”。")
        hint.setObjectName("muted")
        hint.setWordWrap(True)
        self.query_result_summary = QLabel("查询结果：0 条")
        self.query_result_summary.setObjectName("muted")
        self.query_result_summary.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.query_cache_status = QLabel("缓存：尚未读取")
        self.query_cache_status.setObjectName("muted")
        layout.addLayout(bar)
        layout.addLayout(date_bar)
        layout.addWidget(hint)
        layout.addWidget(self.query_cache_status)
        layout.addWidget(self.query_result_summary)
        layout.addWidget(self.query_table, 1)
        self.update_query_date_controls()
        return page

    def _analysis_page(self) -> QWidget:
        page, layout = self._page(
            "数据分析",
            "查询和分析走本地库。同步会下载表格最新数据并替换本地库。场记只数含 D 的条数。",
        )
        self._analysis_result = None
        self._analysis_chart_headers: list[str] = []
        self._analysis_header_selected: list[str] = list(self.store.get("analysis_stat_headers", []) or [])
        bar = QHBoxLayout()
        bar.setSpacing(8)
        self.analysis_source_pick = QComboBox()
        self.analysis_source_pick.setMinimumWidth(140)
        self.analysis_date_field = QComboBox()
        self.analysis_date_field.setMinimumWidth(130)
        self.analysis_name_field = QComboBox()
        self.analysis_name_field.setMinimumWidth(130)
        self.analysis_team = QComboBox()
        self.analysis_team.setEditable(True)
        self.analysis_team.setInsertPolicy(QComboBox.NoInsert)
        self.analysis_team.setMinimumWidth(120)
        self.analysis_names = QLineEdit(str(self.store.get("analysis_names", "") or ""))
        self.analysis_names.setPlaceholderText("留空=整个队别")
        self.analysis_exclude = QLineEdit(str(self.store.get("analysis_exclude_keywords", "") or ""))
        self.analysis_exclude.setPlaceholderText("排除关键词，逗号分隔")
        self.analysis_cache_status = QLabel("本地库：尚未同步")
        self.analysis_cache_status.setObjectName("muted")
        self.analysis_sync_button = QPushButton("同步本地库")
        self.analysis_sync_button.setToolTip("从表格下载最新数据并替换本地库。内容完全相同则跳过写入。")
        self.analysis_sync_button.clicked.connect(self.run_analysis_sync)
        self.analysis_button = QPushButton("开始分析")
        self.analysis_button.setObjectName("primary")
        self.analysis_button.clicked.connect(self.run_analysis)
        bar.addWidget(QLabel("数据源"))
        bar.addWidget(self.analysis_source_pick)
        bar.addWidget(QLabel("日期"))
        bar.addWidget(self.analysis_date_field)
        bar.addWidget(QLabel("名字列"))
        bar.addWidget(self.analysis_name_field)
        bar.addWidget(QLabel("队别"))
        bar.addWidget(self.analysis_team)
        bar.addWidget(QLabel("名字"))
        bar.addWidget(self.analysis_names, 1)
        bar.addWidget(self.analysis_exclude, 1)
        bar.addWidget(self.analysis_sync_button)
        bar.addWidget(self.analysis_button)

        compare_bar = QHBoxLayout()
        compare_bar.setSpacing(8)
        self.analysis_range = QComboBox()
        self.analysis_range.addItems(RANGE_LABELS)
        saved_mode = UI_RANGE_ALIASES.get(str(self.store.get("analysis_compare_mode", "month") or "month"), "month")
        if saved_mode in COMPARE_MODES:
            self.analysis_range.setCurrentIndex(COMPARE_MODES.index(saved_mode))
        self.analysis_compare_enabled = QCheckBox("对比")
        self.analysis_compare_enabled.setChecked(bool(self.store.get("analysis_compare_enabled", False)))
        self.analysis_reference_label = QLabel("截止")
        self.analysis_reference = QDateEdit(QDate.currentDate())
        self.analysis_reference.setCalendarPopup(True)
        self.analysis_reference.setDisplayFormat("yyyy-MM-dd")
        saved_ref = QDate.fromString(str(self.store.get("analysis_reference_date", "") or ""), "yyyy-MM-dd")
        if saved_ref.isValid():
            self.analysis_reference.setDate(saved_ref)
        self.analysis_current_label = QLabel("本期")
        self.analysis_current_to = QLabel("至")
        self.analysis_previous_label = QLabel("对比期")
        self.analysis_previous_to = QLabel("至")
        self.analysis_current_start = QDateEdit(QDate.currentDate().addDays(-6))
        self.analysis_current_end = QDateEdit(QDate.currentDate())
        self.analysis_previous_start = QDateEdit(QDate.currentDate().addDays(-13))
        self.analysis_previous_end = QDateEdit(QDate.currentDate().addDays(-7))
        for widget, key in (
            (self.analysis_current_start, "analysis_current_start"),
            (self.analysis_current_end, "analysis_current_end"),
            (self.analysis_previous_start, "analysis_previous_start"),
            (self.analysis_previous_end, "analysis_previous_end"),
        ):
            widget.setCalendarPopup(True)
            widget.setDisplayFormat("yyyy-MM-dd")
            saved = QDate.fromString(str(self.store.get(key, "") or ""), "yyyy-MM-dd")
            if saved.isValid():
                widget.setDate(saved)
        self.analysis_chart_line = QPushButton("曲线")
        self.analysis_chart_pie = QPushButton("饼图")
        for button in (self.analysis_chart_line, self.analysis_chart_pie):
            button.setCheckable(True)
            button.setObjectName("chartToggle")
            button.setFixedWidth(56)
        chart_mode = str(self.store.get("analysis_chart_mode", "line") or "line")
        self.analysis_chart_line.setChecked(chart_mode != "pie")
        self.analysis_chart_pie.setChecked(chart_mode == "pie")
        self.analysis_chart_line.clicked.connect(lambda: self.set_analysis_chart_mode("line"))
        self.analysis_chart_pie.clicked.connect(lambda: self.set_analysis_chart_mode("pie"))
        compare_bar.addWidget(QLabel("时间"))
        compare_bar.addWidget(self.analysis_range)
        compare_bar.addWidget(self.analysis_reference_label)
        compare_bar.addWidget(self.analysis_reference)
        compare_bar.addWidget(self.analysis_current_label)
        compare_bar.addWidget(self.analysis_current_start)
        compare_bar.addWidget(self.analysis_current_to)
        compare_bar.addWidget(self.analysis_current_end)
        compare_bar.addWidget(self.analysis_compare_enabled)
        compare_bar.addWidget(self.analysis_previous_label)
        compare_bar.addWidget(self.analysis_previous_start)
        compare_bar.addWidget(self.analysis_previous_to)
        compare_bar.addWidget(self.analysis_previous_end)
        compare_bar.addStretch()
        compare_bar.addWidget(self.analysis_chart_line)
        compare_bar.addWidget(self.analysis_chart_pie)

        self.analysis_source_pick.currentIndexChanged.connect(self.on_analysis_table_changed)
        self.analysis_date_field.currentTextChanged.connect(self.persist_workspace_settings)
        self.analysis_name_field.currentTextChanged.connect(self.persist_workspace_settings)
        self.analysis_team.currentTextChanged.connect(self.persist_workspace_settings)
        self.analysis_range.currentIndexChanged.connect(self.sync_analysis_periods)
        self.analysis_range.currentIndexChanged.connect(self.persist_workspace_settings)
        self.analysis_compare_enabled.toggled.connect(self.sync_analysis_periods)
        self.analysis_compare_enabled.toggled.connect(self.persist_workspace_settings)
        self.analysis_reference.dateChanged.connect(self.sync_analysis_periods)
        self.analysis_reference.dateChanged.connect(self.persist_workspace_settings)
        for widget in (
            self.analysis_current_start, self.analysis_current_end,
            self.analysis_previous_start, self.analysis_previous_end,
        ):
            widget.dateChanged.connect(self.persist_workspace_settings)
        self.analysis_names.editingFinished.connect(self.persist_workspace_settings)
        self.analysis_exclude.editingFinished.connect(self.persist_workspace_settings)

        stats = QHBoxLayout()
        stats.setSpacing(8)
        self.analysis_stat_current = self._stat_card("本期")
        self.analysis_stat_previous = self._stat_card("对比期")
        stats.addWidget(self.analysis_stat_current)
        stats.addWidget(self.analysis_stat_previous)

        self.analysis_summary = QLabel("选数据源后开始分析。勾选加友途径等表头，曲线会按渠道分开。")
        self.analysis_summary.setObjectName("muted")
        self.analysis_summary.setWordWrap(True)

        chip_bar = QHBoxLayout()
        chip_bar.setSpacing(6)
        chip_label = QLabel("曲线数据")
        chip_label.setObjectName("sectionTitle")
        self.analysis_count_check = QCheckBox("记录数")
        self.analysis_count_check.setObjectName("chipCheck")
        self.analysis_count_check.setChecked(bool(self.store.get("analysis_show_count", True)))
        self.analysis_count_check.toggled.connect(self.on_analysis_header_checks)
        self.analysis_chip_host = QWidget()
        self.analysis_chip_layout = QHBoxLayout(self.analysis_chip_host)
        self.analysis_chip_layout.setContentsMargins(0, 0, 0, 0)
        self.analysis_chip_layout.setSpacing(6)
        self.analysis_more_headers = QPushButton("更多")
        self.analysis_more_headers.setObjectName("chartToggle")
        self.analysis_more_headers.setFixedHeight(26)
        self.analysis_more_headers.clicked.connect(self.open_analysis_header_more)
        chip_bar.addWidget(chip_label)
        chip_bar.addWidget(self.analysis_count_check)
        chip_bar.addWidget(self.analysis_chip_host, 1)
        chip_bar.addWidget(self.analysis_more_headers)
        chip_bar.addStretch()

        self.analysis_chart = LineChartWidget()
        self.analysis_chart.set_mode(chart_mode)

        tables = QSplitter(Qt.Horizontal)
        self.analysis_daily_table = QTableWidget()
        self.analysis_people_table = QTableWidget()
        for table in (self.analysis_daily_table, self.analysis_people_table):
            table.setEditTriggers(QAbstractItemView.NoEditTriggers)
            table.setSelectionBehavior(QAbstractItemView.SelectRows)
            table.verticalHeader().setVisible(False)
            table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
            table.setAlternatingRowColors(True)
        self.analysis_daily_wrap = self._labeled_table("每日记录", self.analysis_daily_table)
        self.analysis_people_wrap = self._labeled_table("人员明细", self.analysis_people_table)
        tables.addWidget(self.analysis_daily_wrap)
        tables.addWidget(self.analysis_people_wrap)
        tables.setStretchFactor(0, 3)
        tables.setStretchFactor(1, 2)
        tables.setMinimumHeight(180)

        layout.addLayout(bar)
        layout.addWidget(self.analysis_cache_status)
        layout.addLayout(compare_bar)
        layout.addLayout(stats)
        layout.addWidget(self.analysis_summary)
        layout.addLayout(chip_bar)
        layout.addWidget(self.analysis_chart, 2)
        layout.addWidget(tables, 3)
        self.sync_analysis_periods()
        return page

    def _labeled_table(self, title: str, table: QTableWidget) -> QWidget:
        wrap = QWidget()
        box = QVBoxLayout(wrap)
        box.setContentsMargins(0, 0, 0, 0)
        box.setSpacing(4)
        label = QLabel(title)
        label.setObjectName("sectionTitle")
        box.addWidget(label)
        box.addWidget(table, 1)
        wrap.title_label = label
        return wrap

    def _stat_card(self, title: str) -> QFrame:
        card = QFrame()
        card.setObjectName("statCard")
        box = QVBoxLayout(card)
        box.setContentsMargins(12, 8, 12, 8)
        heading = QLabel(title)
        heading.setObjectName("muted")
        value = QLabel("—")
        value.setObjectName("statValue")
        value.setWordWrap(True)
        extra = QLabel("")
        extra.setObjectName("muted")
        extra.setWordWrap(True)
        box.addWidget(heading)
        box.addWidget(value)
        box.addWidget(extra)
        card.title_label = heading
        card.value_label = value
        card.extra_label = extra
        return card

    def _extract_page(self) -> QWidget:
        page, layout = self._page("时间提取", "按日期范围提取数据并写入独立 Excel；历史提取记录会自动排重。")
        card = QFrame()
        card.setObjectName("card")
        form = QFormLayout(card)
        self.extract_date_field = QComboBox()
        self.extract_date_field.currentTextChanged.connect(self.persist_workspace_settings)
        self.start_date = QDateEdit(QDate.currentDate().addMonths(-1))
        self.end_date = QDateEdit(QDate.currentDate())
        for widget in (self.start_date, self.end_date):
            widget.setCalendarPopup(True)
            widget.setDisplayFormat("yyyy-MM-dd")
        date_row = QHBoxLayout()
        date_row.addWidget(self.start_date)
        date_row.addWidget(QLabel("至"))
        date_row.addWidget(self.end_date)
        date_row.addStretch()
        self.dedup_fields = QLineEdit(str(self.store.get("extract_dedup_fields", "号码,日期") or "号码,日期"))
        self.dedup_fields.setPlaceholderText("多个字段用逗号分隔；留空时使用源位置和行指纹")
        self.extract_signature_enabled = QCheckBox("添加签字栏")
        self.extract_signature_enabled.setChecked(bool(self.store.get("extract_signature_enabled", False)))
        self.extract_signature_header = QLineEdit(str(self.store.get("extract_signature_header", "签字") or "签字"))
        self.extract_signature_header.setPlaceholderText("例如：签字")
        self.extract_signature_value = QLineEdit(str(self.store.get("extract_signature_value", "") or ""))
        self.extract_signature_value.setPlaceholderText("提取人姓名")
        self.extract_signature_column = QComboBox()
        self.extract_signature_column.setEditable(True)
        self.extract_signature_column.setInsertPolicy(QComboBox.NoInsert)
        self.extract_signature_column.addItems(["", *COLUMN_LETTERS])
        self.extract_signature_column.setCurrentText(str(self.store.get("extract_signature_column", "") or ""))
        signature_row = QHBoxLayout()
        signature_row.addWidget(self.extract_signature_enabled)
        signature_row.addWidget(QLabel("表头"))
        signature_row.addWidget(self.extract_signature_header)
        signature_row.addWidget(QLabel("内容"))
        signature_row.addWidget(self.extract_signature_value)
        signature_row.addWidget(QLabel("列"))
        signature_row.addWidget(self.extract_signature_column)
        self.extract_mode = QComboBox()
        self.extract_mode.addItems(["从汇总数据库提取", "直接从数据源提取"])
        self.extract_mode.setCurrentIndex(1 if str(self.store.get("extract_mode", "aggregate")) == "direct" else 0)
        self.extract_source_label = QLabel("指定数据源")
        self.extract_source_pick = QComboBox()
        self.extract_source_pick.setMinimumWidth(220)
        self.extract_source_pick.currentIndexChanged.connect(self.on_extract_source_changed)
        self.output_type = QComboBox()
        self.output_type.addItems(["本地 Excel 文件", "Google 表格链接"])
        saved_type = str(self.store.get("extract_destination_type", "local"))
        self.output_type.setCurrentIndex(1 if saved_type == "google" else 0)
        self.output_sheet_name = QLineEdit(str(self.store.get("google_output_sheet", "提取结果") or "提取结果"))
        self.output_sheet_name.setPlaceholderText("例如：查询结果")
        self.google_output_url = QLineEdit(str(self.store.get("google_output_url", "") or ""))
        self.google_output_url.setPlaceholderText("https://docs.google.com/spreadsheets/d/...")
        saved_path = str(self.store.get("extract_output_path", "") or "").strip()
        self.output_path = QLineEdit(saved_path or str(Path.home() / "Desktop" / "提取结果.xlsx"))
        choose = QPushButton("选择…")
        choose.clicked.connect(self.choose_output)
        output_row = QHBoxLayout()
        output_row.addWidget(self.output_path, 1)
        output_row.addWidget(choose)
        form.addRow("日期字段", self.extract_date_field)
        form.addRow("日期范围", date_row)
        form.addRow("去重字段", self.dedup_fields)
        form.addRow("签字栏", signature_row)
        form.addRow("数据来源", self.extract_mode)
        form.addRow(self.extract_source_label, self.extract_source_pick)
        form.addRow("输出目标", self.output_type)
        form.addRow("目标表格链接", self.google_output_url)
        form.addRow("目标工作表", self.output_sheet_name)
        form.addRow("输出文件", output_row)
        self.extract_schema_enabled = QCheckBox("提取表使用独立字段分配（表头和列可以与数据源不同）")
        self.extract_schema_enabled.setChecked(bool(self.store.get("extract_column_schema_enabled", False)))
        self.extract_schema_editor = ColumnMapWidget()
        self.extract_schema_editor.set_schema(self.store.get("extract_column_schema", []) or [])
        self.extract_schema_editor.setEnabled(self.extract_schema_enabled.isChecked())
        self.extract_schema_enabled.toggled.connect(self.extract_schema_editor.setEnabled)
        form.addRow("", self.extract_schema_enabled)
        form.addRow("提取表字段", self.extract_schema_editor)
        self.output_type.currentIndexChanged.connect(self.update_output_destination)
        self.extract_mode.currentIndexChanged.connect(self.on_extract_mode_changed)
        self.dedup_fields.editingFinished.connect(self.persist_workspace_settings)
        self.extract_signature_enabled.toggled.connect(self.persist_workspace_settings)
        self.extract_signature_header.editingFinished.connect(self.persist_workspace_settings)
        self.extract_signature_value.editingFinished.connect(self.persist_workspace_settings)
        self.extract_signature_column.currentTextChanged.connect(self.persist_workspace_settings)
        self.google_output_url.editingFinished.connect(self.persist_workspace_settings)
        self.output_sheet_name.editingFinished.connect(self.persist_workspace_settings)
        self.output_path.editingFinished.connect(self.persist_workspace_settings)
        self.extract_schema_enabled.toggled.connect(self.persist_workspace_settings)
        self.local_output_widgets = [self.output_path, choose]
        self.refresh_extract_source_picker()
        self.update_extract_source_visibility()
        self.update_output_destination()
        self.extract_button = QPushButton("开始提取")
        self.extract_button.setObjectName("primary")
        self.extract_button.clicked.connect(self.run_extract)
        self.clear_extract_cache_button = QPushButton("清除本地排重缓存")
        self.clear_extract_cache_button.clicked.connect(self.clear_extract_cache)
        extract_actions = QHBoxLayout()
        extract_actions.addWidget(self.extract_button)
        extract_actions.addWidget(self.clear_extract_cache_button)
        extract_actions.addStretch()
        self.extract_status = QLabel("等待运行")
        self.extract_status.setObjectName("muted")
        layout.addWidget(card)
        layout.addLayout(extract_actions)
        layout.addWidget(self.extract_status)
        layout.addStretch()
        return page

    def _fields_page(self) -> QWidget:
        page, layout = self._page("列与字段配置", "给每个字段勾选、起名、指定对应列。某个数据源列不一样时，到该数据源里单独分配。")
        schema_card = QFrame()
        schema_card.setObjectName("card")
        schema_layout = QVBoxLayout(schema_card)
        schema_actions = QHBoxLayout()
        self.schema_enabled = QCheckBox("启用全局字段分配")
        self.schema_enabled.setChecked(bool(self.store.get("column_schema_enabled", False)))
        example = QPushButton("填入示例")
        example.clicked.connect(self.fill_example_schema)
        schema_actions.addWidget(self.schema_enabled)
        schema_actions.addWidget(example)
        schema_actions.addStretch()
        self.schema_editor = ColumnMapWidget()
        schema_layout.addLayout(schema_actions)
        schema_layout.addWidget(self.schema_editor)
        schema_hint = QLabel("勾选字段、填写表头名称、选择对应列（A/B/C…），可用上下箭头调整顺序。已单独配置的数据源不会使用这里的分配。")
        schema_hint.setObjectName("muted")
        schema_hint.setWordWrap(True)
        schema_layout.addWidget(schema_hint)
        layout.addWidget(schema_card)
        layout.addWidget(QLabel("标准字段与表头别名"))
        self.fields_table = QTableWidget(0, 2)
        self.fields_table.setHorizontalHeaderLabels(["标准字段", "可匹配的表头别名（逗号分隔）"])
        self.fields_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        actions = QHBoxLayout()
        add = QPushButton("添加字段")
        remove = QPushButton("删除选中")
        save = QPushButton("保存列与字段配置")
        save.setObjectName("primary")
        add.clicked.connect(self.add_field_row)
        remove.clicked.connect(self.remove_field_row)
        save.clicked.connect(self.save_fields)
        actions.addWidget(add)
        actions.addWidget(remove)
        actions.addStretch()
        actions.addWidget(save)
        layout.addWidget(self.fields_table, 1)
        layout.addLayout(actions)
        self.schema_enabled.toggled.connect(self.update_schema_enabled)
        self.load_fields_table()
        return page

    def _logs_page(self) -> QWidget:
        page, layout = self._page("运行日志", "查看读取、匹配、写入、去重和错误的详细记录。")
        actions = QHBoxLayout()
        refresh = QPushButton("刷新")
        clear = QPushButton("清空日志")
        refresh.clicked.connect(self.refresh_logs)
        clear.clicked.connect(self.clear_logs)
        actions.addWidget(refresh)
        actions.addWidget(clear)
        actions.addStretch()
        self.log_table = QTableWidget(0, 5)
        self.log_table.setHorizontalHeaderLabels(["时间", "级别", "操作", "消息", "详情"])
        self.log_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.log_table.verticalHeader().setVisible(False)
        self.log_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        layout.addLayout(actions)
        layout.addWidget(self.log_table, 1)
        return page

    def _settings_page(self) -> QWidget:
        page, layout = self._page("设置", "配置所有数据源共用的默认排除项、认证文件，以及软件更新。")
        card = QFrame()
        card.setObjectName("card")
        form = QFormLayout(card)
        self.global_excludes = QTextEdit("\n".join(self.store.get("global_excludes", [])))
        self.global_excludes.setMaximumHeight(130)
        self.credential_list = QListWidget()
        self.credential_list.setMaximumHeight(140)
        self._reload_credential_list()
        add_credential = QPushButton("添加…")
        remove_credential = QPushButton("移除")
        add_credential.clicked.connect(self.add_default_credentials)
        remove_credential.clicked.connect(self.remove_default_credential)
        credential_buttons = QVBoxLayout()
        credential_buttons.addWidget(add_credential)
        credential_buttons.addWidget(remove_credential)
        credential_buttons.addStretch()
        credential_row = QHBoxLayout()
        credential_row.addWidget(self.credential_list, 1)
        credential_row.addLayout(credential_buttons)
        credential_hint = QLabel("可添加多个 JSON。读取和写入会按顺序轮询，遇到 429 自动换下一个。请把表格共享给每一个服务账号邮箱。")
        credential_hint.setObjectName("muted")
        credential_hint.setWordWrap(True)
        form.addRow("全局排除 Sheet", self.global_excludes)
        form.addRow("服务账号 JSON", credential_row)
        form.addRow("", credential_hint)
        save = QPushButton("保存设置")
        save.setObjectName("primary")
        save.clicked.connect(self.save_settings)
        data_path = QLabel(str(self.store.data_dir))
        data_path.setTextInteractionFlags(Qt.TextSelectableByMouse)
        form.addRow("本地数据目录", data_path)
        version_row = QHBoxLayout()
        version_row.addWidget(QLabel(f"当前版本 v{APP_VERSION}"))
        self.update_button = QPushButton("检查并安装更新")
        self.update_button.clicked.connect(self.check_updates)
        version_row.addWidget(self.update_button)
        version_row.addStretch()
        form.addRow("软件版本", version_row)
        config_row = QHBoxLayout()
        export_config = QPushButton("导出配置")
        import_config = QPushButton("导入配置")
        export_config.clicked.connect(self.export_config_file)
        import_config.clicked.connect(self.import_config_file)
        config_row.addWidget(export_config)
        config_row.addWidget(import_config)
        config_row.addStretch()
        form.addRow("配置迁移", config_row)
        layout.addWidget(card)
        layout.addWidget(save, 0, Qt.AlignLeft)
        layout.addStretch()
        return page

    def selected_source(self) -> SourceConfig | None:
        row = self.source_table.currentRow()
        sources = self.store.load_sources()
        return sources[row] if 0 <= row < len(sources) else None

    def refresh_sources(self) -> None:
        sources = self.store.load_sources()
        self.source_table.setRowCount(len(sources))
        for row, source in enumerate(sources):
            if source.column_schema_enabled and source.column_schema:
                schema_label = f"独立{len(schema_field_names(source.column_schema))}列"
            else:
                schema_label = "自动/全局"
            values = [
                "是" if source.enabled else "否", source.name, source.url,
                "、".join(source.include_sheets) or "全部", "、".join(source.exclude_sheets) or "—",
                str(source.header_row), schema_label, self._source_auth_label(source),
            ]
            for column, value in enumerate(values):
                self.source_table.setItem(row, column, QTableWidgetItem(value))
        self.refresh_query_source_picker()
        self.refresh_extract_source_picker()
        self.refresh_analysis_source_picker()
        self.refresh_sync_source_list()
        if not getattr(self, "_restoring_settings", False):
            self.refresh_query_fields()
            self.refresh_analysis_fields()

    def add_source(self) -> None:
        dialog = SourceDialog(self.store, parent=self)
        if dialog.exec():
            value = dialog.value()
            self.store.save_source(value)
            self.refresh_sources()

    def edit_source(self) -> None:
        source = self.selected_source()
        if not source:
            QMessageBox.information(self, "请选择", "请先选择一个数据源。")
            return
        dialog = SourceDialog(self.store, source, self)
        if dialog.exec():
            self.store.save_source(dialog.value())
            self.refresh_sources()

    def delete_source(self) -> None:
        source = self.selected_source()
        if not source:
            return
        if QMessageBox.question(self, "确认删除", f"确定删除数据源“{source.name}”吗？") == QMessageBox.Yes:
            self.store.delete_source(source.id)
            self.refresh_sources()

    def scan_source(self) -> None:
        source = self.selected_source()
        if not source:
            QMessageBox.information(self, "请选择", "请先选择一个数据源。")
            return
        reader = SourceReader(self.store.get("field_aliases", {}), self.store.get("global_excludes", []))
        self.run_task(lambda: reader.list_sheets(source), self.scan_finished, "正在扫描 Sheet…")

    def scan_finished(self, names: list[str]) -> None:
        QMessageBox.information(self, "Sheet 列表", "\n".join(names) if names else "没有发现子 Sheet。")

    def refresh_sync_source_list(self) -> None:
        if not hasattr(self, "sync_source_list"):
            return
        saved = {str(item) for item in (self.store.get("sync_source_ids", []) or [])}
        self.sync_source_list.clear()
        for source in self.store.load_sources():
            if not source.enabled:
                continue
            item = QListWidgetItem(source.name)
            item.setData(Qt.UserRole, source.id)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked if (not saved or source.id in saved) else Qt.Unchecked)
            self.sync_source_list.addItem(item)

    def set_sync_sources_checked(self, checked: bool) -> None:
        state = Qt.Checked if checked else Qt.Unchecked
        for row in range(self.sync_source_list.count()):
            self.sync_source_list.item(row).setCheckState(state)

    def selected_sync_source_ids(self) -> list[str]:
        if not hasattr(self, "sync_source_list"):
            return []
        return [
            str(self.sync_source_list.item(row).data(Qt.UserRole))
            for row in range(self.sync_source_list.count())
            if self.sync_source_list.item(row).checkState() == Qt.Checked
        ]

    def choose_sync_xlsx(self) -> None:
        path, _ = QFileDialog.getSaveFileName(self, "保存汇总结果", self.sync_local_xlsx.text(), "Excel 工作簿 (*.xlsx)")
        if path:
            if not path.lower().endswith(".xlsx"):
                path += ".xlsx"
            self.sync_local_xlsx.setText(path)

    def run_sync(self) -> None:
        ids = self.selected_sync_source_ids()
        if not ids:
            QMessageBox.warning(self, "请选择", "请至少勾选一个要汇总的数据源。")
            return
        self.persist_workspace_settings()
        self.sync_status.clear()
        self.sync_button.setEnabled(False)
        engine = DataEngine(self.store)
        self.run_task(
            lambda: engine.sync(
                ids,
                self.write_aggregate.isChecked(),
                self.sync_google_url.text().strip(),
                self.sync_google_sheet.text().strip() or "汇总结果",
                self.sync_local_xlsx.text().strip(),
            ),
            self.sync_finished,
            "正在同步并更新缓存…",
            self.sync_button,
        )

    def sync_finished(self, result: dict[str, int]) -> None:
        self.refresh_logs()
        recent = list(reversed(self.store.read_logs(100)))
        self.sync_status.setPlainText("\n".join(
            f"[{row['created_at']}] {row['level']} · {row['message']}" for row in recent
            if row["operation"] in {"汇总同步", "读取数据源", "刷新缓存"}
        ))
        self.update_query_cache_status()
        sources = result.get("sources", 0)
        QMessageBox.information(self, "同步完成", f"已处理 {sources} 个数据源，共 {result.get('rows', 0)} 行，查询缓存已更新。")

    def persist_workspace_settings(self) -> None:
        if getattr(self, "_restoring_settings", False):
            return
        if hasattr(self, "query_mode"):
            self.store.set("query_source", QUERY_SOURCES[self.query_mode.currentIndex()])
            self.store.set("query_fuzzy", self.query_fuzzy.isChecked())
            self.store.set("query_date_enabled", self.query_date_enabled.isChecked())
            self.store.set("query_date_field", self.query_date_field.currentText().strip() or "日期")
            self.store.set("query_start_date", self.query_start_date.date().toString("yyyy-MM-dd"))
            self.store.set("query_end_date", self.query_end_date.date().toString("yyyy-MM-dd"))
            field = self.query_field.currentText().strip()
            if field:
                self.store.set("query_field", field)
                saved_map = dict(self.store.get("query_fields_by_mode", {}) or {})
                saved_map[QUERY_SOURCES[self.query_mode.currentIndex()]] = field
                self.store.set("query_fields_by_mode", saved_map)
            if hasattr(self, "query_source_pick"):
                self.store.set("query_direct_source_id", self.query_source_pick.currentData() or "")
            if hasattr(self, "query_exclude"):
                self.store.set("query_exclude_keywords", self.query_exclude.text().strip())
        if hasattr(self, "analysis_source_pick"):
            self.store.set("analysis_source", "direct")
            self.store.set("analysis_direct_source_id", self.analysis_source_pick.currentData() or "")
            team = self.analysis_team.currentText().strip()
            if team.startswith("全部"):
                team = ""
            self.store.set("analysis_team", team)
            self.store.set("analysis_date_field", self.analysis_date_field.currentText().strip())
            self.store.set("analysis_name_field", self.analysis_name_field.currentText().strip())
            self.store.set("analysis_names", self.analysis_names.text().strip())
            self.store.set("analysis_chart_mode", "pie" if self.analysis_chart_pie.isChecked() else "line")
            self.store.set("analysis_stat_headers", self.selected_analysis_headers())
            self.store.set("analysis_show_count", self.analysis_count_check.isChecked() if hasattr(self, "analysis_count_check") else True)
            self.store.set("analysis_exclude_keywords", self.analysis_exclude.text().strip())
            self.store.set("analysis_compare_mode", COMPARE_MODES[self.analysis_range.currentIndex()])
            self.store.set("analysis_compare_enabled", self.analysis_compare_enabled.isChecked())
            self.store.set("analysis_reference_date", self.analysis_reference.date().toString("yyyy-MM-dd"))
            self.store.set("analysis_current_start", self.analysis_current_start.date().toString("yyyy-MM-dd"))
            self.store.set("analysis_current_end", self.analysis_current_end.date().toString("yyyy-MM-dd"))
            self.store.set("analysis_previous_start", self.analysis_previous_start.date().toString("yyyy-MM-dd"))
            self.store.set("analysis_previous_end", self.analysis_previous_end.date().toString("yyyy-MM-dd"))
        if hasattr(self, "output_type"):
            self.store.set("extract_destination_type", "google" if self.output_type.currentIndex() == 1 else "local")
            self.store.set("google_output_url", self.google_output_url.text().strip())
            self.store.set("google_output_sheet", self.output_sheet_name.text().strip() or "提取结果")
            self.store.set("extract_output_path", self.output_path.text().strip())
            self.store.set("extract_mode", "direct" if self.extract_mode.currentIndex() == 1 else "aggregate")
            if hasattr(self, "extract_source_pick"):
                self.store.set("extract_direct_source_id", self.extract_source_pick.currentData() or "")
            self.store.set("extract_dedup_fields", self.dedup_fields.text().strip() or "号码,日期")
            if hasattr(self, "extract_signature_enabled"):
                self.store.set("extract_signature_enabled", self.extract_signature_enabled.isChecked())
                self.store.set("extract_signature_header", self.extract_signature_header.text().strip() or "签字")
                self.store.set("extract_signature_value", self.extract_signature_value.text().strip())
                self.store.set("extract_signature_column", self.extract_signature_column.currentText().strip().upper())
            date_field = self.extract_date_field.currentText().strip()
            if date_field:
                self.store.set("extract_date_field", date_field)
            if hasattr(self, "extract_schema_enabled"):
                self.store.set("extract_column_schema_enabled", self.extract_schema_enabled.isChecked())
                self.store.set("extract_column_schema", self.extract_schema_editor.schema())
        if hasattr(self, "sync_google_url"):
            self.store.set("sync_google_url", self.sync_google_url.text().strip())
            self.store.set("sync_google_sheet", self.sync_google_sheet.text().strip() or "汇总结果")
            self.store.set("sync_local_xlsx", self.sync_local_xlsx.text().strip())
            self.store.set("sync_source_ids", self.selected_sync_source_ids())

    def closeEvent(self, event) -> None:
        self.persist_workspace_settings()
        super().closeEvent(event)

    def extract_query_target(self) -> tuple[str, str]:
        sheet_name = str(self.store.get("google_output_sheet", "提取结果") or "提取结果")
        if hasattr(self, "output_sheet_name"):
            sheet_name = self.output_sheet_name.text().strip() or sheet_name
        if hasattr(self, "output_type"):
            if self.output_type.currentIndex() == 1:
                return self.google_output_url.text().strip(), sheet_name
            return self.output_path.text().strip(), sheet_name
        if str(self.store.get("extract_destination_type", "local")) == "google":
            return str(self.store.get("google_output_url", "") or ""), sheet_name
        return str(self.store.get("extract_output_path", "") or ""), sheet_name

    def current_query_mode(self) -> str:
        return QUERY_SOURCES[self.query_mode.currentIndex()]

    def current_query_source_id(self) -> str:
        if self.current_query_mode() != "direct" or not hasattr(self, "query_source_pick"):
            return ""
        return str(self.query_source_pick.currentData() or "")

    def current_query_source_name(self) -> str:
        if self.current_query_mode() != "direct" or not hasattr(self, "query_source_pick"):
            return ""
        if not self.query_source_pick.currentData():
            return ""
        return self.query_source_pick.currentText().strip()

    def current_table_is_phone(self, fields: list[str] | None = None) -> bool:
        return is_phone_data_table(fields if fields is not None else self._current_result_fields(), self.current_query_source_name())

    def _current_result_fields(self) -> list[str]:
        extract_target, extract_sheet = self.extract_query_target()
        return DataEngine(self.store).list_query_fields(
            self.current_query_mode(), extract_target, extract_sheet, self.current_query_source_id(),
        )

    def update_copy_button(self) -> None:
        if not hasattr(self, "copy_query_button"):
            return
        if self.current_table_is_phone():
            self.copy_query_button.setText("一键复制号码和修正格式")
        else:
            self.copy_query_button.setText("一键复制结果")

    def update_query_cache_status(self) -> None:
        if not hasattr(self, "query_cache_status"):
            return
        engine = DataEngine(self.store)
        mode = self.current_query_mode()
        if mode == "aggregate":
            self.query_cache_status.setText("缓存：汇总库为本地数据，查询较快")
            return
        cache_id = "extract-table" if mode == "extract" else self.current_query_source_id()
        if mode == "direct" and not cache_id:
            self.query_cache_status.setText("本地库：查询时使用各数据源已同步的数据")
            return
        if cache_id and engine.cache.has(cache_id):
            self.query_cache_status.setText(
                f"本地库：{engine.cache.row_count(cache_id)} 行，更新于 {engine.cache.updated_at(cache_id)}"
            )
        else:
            self.query_cache_status.setText("本地库：尚未同步。请先点「同步本地库」从表格下载。")

    def on_query_mode_changed(self) -> None:
        if getattr(self, "_restoring_settings", False):
            return
        self.update_query_source_visibility()
        self.refresh_query_fields()
        self.persist_workspace_settings()

    def on_query_table_changed(self) -> None:
        if getattr(self, "_restoring_settings", False):
            return
        self.refresh_query_fields()
        self.persist_workspace_settings()

    def update_query_source_visibility(self) -> None:
        if not hasattr(self, "query_source_pick"):
            return
        direct = self.current_query_mode() == "direct"
        self.query_table_label.setVisible(direct)
        self.query_source_pick.setVisible(direct)

    def current_extract_source_id(self) -> str:
        if not hasattr(self, "extract_source_pick"):
            return ""
        if not hasattr(self, "extract_mode") or self.extract_mode.currentIndex() != 1:
            return ""
        return str(self.extract_source_pick.currentData() or "")

    def update_extract_source_visibility(self) -> None:
        if not hasattr(self, "extract_source_pick"):
            return
        direct = self.extract_mode.currentIndex() == 1
        self.extract_source_label.setVisible(direct)
        self.extract_source_pick.setVisible(direct)

    def refresh_extract_source_picker(self) -> None:
        if not hasattr(self, "extract_source_pick"):
            return
        restoring = self._restoring_settings
        self._restoring_settings = True
        saved = str(self.store.get("extract_direct_source_id", "") or "")
        current = self.extract_source_pick.currentData()
        self.extract_source_pick.clear()
        self.extract_source_pick.addItem("全部数据源", "")
        for source in self.store.load_sources():
            if source.enabled:
                self.extract_source_pick.addItem(source.name, source.id)
        target = current if current not in (None, "") else saved
        index = self.extract_source_pick.findData(target)
        self.extract_source_pick.setCurrentIndex(index if index >= 0 else 0)
        self._restoring_settings = restoring

    def update_query_date_controls(self) -> None:
        enabled = self.query_date_enabled.isChecked()
        for widget in (self.query_date_field, self.query_start_date, self.query_end_date):
            widget.setEnabled(enabled)

    def refresh_query_source_picker(self) -> None:
        if not hasattr(self, "query_source_pick"):
            return
        restoring = self._restoring_settings
        self._restoring_settings = True
        saved = str(self.store.get("query_direct_source_id", "") or "")
        current = self.query_source_pick.currentData()
        self.query_source_pick.clear()
        self.query_source_pick.addItem("全部数据源", "")
        for source in self.store.load_sources():
            if source.enabled:
                self.query_source_pick.addItem(source.name, source.id)
        target = current if current not in (None, "") else saved
        index = self.query_source_pick.findData(target)
        self.query_source_pick.setCurrentIndex(index if index >= 0 else 0)
        self._restoring_settings = restoring
        self.update_query_source_visibility()

    def refresh_query_fields(self) -> None:
        if not hasattr(self, "query_field"):
            return
        mode = self.current_query_mode()
        extract_target, extract_sheet = self.extract_query_target()
        fields = DataEngine(self.store).list_query_fields(
            mode, extract_target, extract_sheet, self.current_query_source_id(),
        )
        restoring = self._restoring_settings
        self._restoring_settings = True
        current = self.query_field.currentText().strip()
        saved_map = self.store.get("query_fields_by_mode", {}) or {}
        saved = str(saved_map.get(mode) or self.store.get("query_field", "") or "").strip()
        self.query_field.clear()
        self.query_field.addItems(fields)
        if saved in fields:
            pick = saved
        elif current in fields:
            pick = current
        elif "号码" in fields:
            pick = "号码"
        elif "手机号码" in fields:
            pick = "手机号码"
        elif fields:
            pick = fields[0]
        else:
            pick = saved or current
        if pick:
            if pick not in fields:
                self.query_field.insertItem(0, pick)
            self.query_field.setCurrentText(pick)
        current_date_field = self.query_date_field.currentText().strip()
        saved_date_field = str(self.store.get("query_date_field", "日期") or "日期").strip()
        self.query_date_field.clear()
        self.query_date_field.addItems(fields)
        date_pick = ""
        for candidate in (current_date_field, saved_date_field):
            if candidate in fields:
                date_pick = candidate
                break
        if not date_pick:
            date_pick = DataEngine(self.store).suggest_field(fields, "日期")
        if date_pick:
            self.query_date_field.setCurrentText(date_pick)
        self._restoring_settings = restoring
        self.update_copy_button()
        self.update_query_cache_status()

    def current_analysis_source_id(self) -> str:
        if not hasattr(self, "analysis_source_pick"):
            return ""
        return str(self.analysis_source_pick.currentData() or "")

    def current_analysis_team(self) -> str:
        if not hasattr(self, "analysis_team"):
            return ""
        text = self.analysis_team.currentText().strip()
        if not text or text.startswith("全部"):
            return ""
        return text

    def on_analysis_table_changed(self) -> None:
        if getattr(self, "_restoring_settings", False):
            return
        self.refresh_analysis_fields()
        self.persist_workspace_settings()

    def refresh_analysis_source_picker(self) -> None:
        if not hasattr(self, "analysis_source_pick"):
            return
        restoring = self._restoring_settings
        self._restoring_settings = True
        saved = str(self.store.get("analysis_direct_source_id", "") or "")
        current = self.analysis_source_pick.currentData()
        sources = [source for source in self.store.load_sources() if source.enabled]
        self.analysis_source_pick.clear()
        self.analysis_source_pick.addItem("请选择数据源", "")
        for source in sources:
            self.analysis_source_pick.addItem(source.name, source.id)
        target = current if current not in (None, "") else saved
        index = self.analysis_source_pick.findData(target)
        if index < 0 and len(sources) == 1:
            index = 1
        self.analysis_source_pick.setCurrentIndex(index if index >= 0 else 0)
        self._restoring_settings = restoring
        self.refresh_analysis_fields()

    def refresh_analysis_fields(self) -> None:
        if not hasattr(self, "analysis_date_field"):
            return
        source_id = self.current_analysis_source_id()
        engine = DataEngine(self.store)
        headers = engine.source_headers(source_id) if source_id else []
        date_headers = engine.list_date_headers(headers)
        name_headers = engine.list_name_headers(headers)
        date_field, team_field, name_field = engine.detect_analysis_fields(headers)
        restoring = self._restoring_settings
        self._restoring_settings = True
        saved_date = self.analysis_date_field.currentText().strip() or str(self.store.get("analysis_date_field", "") or "")
        self.analysis_date_field.clear()
        self.analysis_date_field.addItems(date_headers)
        if saved_date in date_headers:
            self.analysis_date_field.setCurrentText(saved_date)
        elif date_field:
            self.analysis_date_field.setCurrentText(date_field)
        saved_name_field = self.analysis_name_field.currentText().strip() or str(self.store.get("analysis_name_field", "") or "")
        self.analysis_name_field.clear()
        self.analysis_name_field.addItems(name_headers)
        if saved_name_field in name_headers:
            self.analysis_name_field.setCurrentText(saved_name_field)
        elif name_field:
            self.analysis_name_field.setCurrentText(name_field)
        saved_team = self.current_analysis_team() or str(self.store.get("analysis_team", "") or "")
        self.analysis_team.clear()
        self.analysis_team.addItem("全部队别")
        teams = engine.list_analysis_teams(source_id, team_field) if source_id and team_field else []
        self.analysis_team.addItems(teams)
        if saved_team:
            index = self.analysis_team.findText(saved_team)
            if index >= 0:
                self.analysis_team.setCurrentIndex(index)
            else:
                self.analysis_team.setEditText(saved_team)
        else:
            self.analysis_team.setCurrentIndex(0)
        self._restoring_settings = restoring
        self.rebuild_analysis_header_checks(headers)
        self.update_analysis_cache_status()

    def current_analysis_range_mode(self) -> str:
        if not hasattr(self, "analysis_range"):
            return "month"
        index = self.analysis_range.currentIndex()
        if 0 <= index < len(COMPARE_MODES):
            return COMPARE_MODES[index]
        return "month"

    def analysis_compare_on(self) -> bool:
        return bool(hasattr(self, "analysis_compare_enabled") and self.analysis_compare_enabled.isChecked())

    def sync_analysis_periods(self) -> None:
        if not hasattr(self, "analysis_range"):
            return
        mode = self.current_analysis_range_mode()
        custom = mode == "custom"
        compare = self.analysis_compare_on()
        self.analysis_reference_label.setVisible(not custom)
        self.analysis_reference.setVisible(not custom)
        for widget in (
            self.analysis_current_label, self.analysis_current_start,
            self.analysis_current_to, self.analysis_current_end,
        ):
            widget.setVisible(custom)
        for widget in (
            self.analysis_previous_label, self.analysis_previous_start,
            self.analysis_previous_to, self.analysis_previous_end,
        ):
            widget.setVisible(custom and compare)
        if hasattr(self, "analysis_stat_previous"):
            self.analysis_stat_previous.setVisible(compare)
        if custom:
            return
        restoring = self._restoring_settings
        self._restoring_settings = True
        (current_start, current_end), (previous_start, previous_end) = compare_periods(
            mode, self.analysis_reference.date().toPython(), compare=compare,
        )
        self.analysis_current_start.setDate(QDate(current_start.year, current_start.month, current_start.day))
        self.analysis_current_end.setDate(QDate(current_end.year, current_end.month, current_end.day))
        self.analysis_previous_start.setDate(QDate(previous_start.year, previous_start.month, previous_start.day))
        self.analysis_previous_end.setDate(QDate(previous_end.year, previous_end.month, previous_end.day))
        self._restoring_settings = restoring

    def run_analysis(self) -> None:
        source_id = self.current_analysis_source_id()
        if not source_id:
            QMessageBox.warning(self, "请选择", "请选择你在「数据源」里已经配置好的表格。")
            return
        names = split_names(self.analysis_names.text())
        team = self.current_analysis_team()
        self.persist_workspace_settings()
        engine = DataEngine(self.store)
        if not engine.cache.has(source_id):
            QMessageBox.warning(self, "请先同步", "本地库还没有这份表。请先点「同步本地库」从表格下载，之后分析都走本地数据。")
            return
        self.analysis_summary.setText("正在分析本地库…")
        date_field = self.analysis_date_field.currentText().strip()
        name_field = self.analysis_name_field.currentText().strip()
        self.run_task(
            lambda: engine.analyze(
                source="direct",
                source_id=source_id,
                date_field=date_field,
                name_field=name_field,
                team=team,
                names=names,
                exclude_keywords=split_names(self.analysis_exclude.text()),
                compare_mode=self.current_analysis_range_mode(),
                compare=self.analysis_compare_on(),
                reference_date=self.analysis_reference.date().toPython(),
                current_start=self.analysis_current_start.date().toPython(),
                current_end=self.analysis_current_end.date().toPython(),
                previous_start=self.analysis_previous_start.date().toPython(),
                previous_end=self.analysis_previous_end.date().toPython(),
                refresh_cache=False,
            ),
            self.show_analysis_results,
            "正在分析本地库…",
            self.analysis_button,
        )

    def update_analysis_cache_status(self) -> None:
        if not hasattr(self, "analysis_cache_status"):
            return
        source_id = self.current_analysis_source_id()
        if not source_id:
            self.analysis_cache_status.setText("本地库：请选择数据源后同步。")
            return
        engine = DataEngine(self.store)
        if engine.cache.has(source_id):
            self.analysis_cache_status.setText(
                f"本地库：{engine.cache.row_count(source_id)} 行，更新于 {engine.cache.updated_at(source_id)}。分析走本地；有新数据再点同步。"
            )
        else:
            self.analysis_cache_status.setText("本地库：尚未同步。请先点「同步本地库」从表格下载。")

    def run_analysis_sync(self) -> None:
        source_id = self.current_analysis_source_id()
        if not source_id:
            QMessageBox.warning(self, "请选择", "请先选择数据源。")
            return
        self.persist_workspace_settings()
        engine = DataEngine(self.store)
        self.run_task(
            lambda: engine.refresh_cache([source_id]),
            self.analysis_sync_finished,
            "正在从表格同步到本地库…",
            self.analysis_sync_button,
        )

    def analysis_sync_finished(self, result: dict) -> None:
        self.update_analysis_cache_status()
        self.refresh_analysis_fields()
        self.refresh_logs()
        if result.get("unchanged"):
            extra = "内容与本地库一致，无需重写。"
        else:
            extra = "已用表格最新数据替换本地库。"
        QMessageBox.information(self, "本地库已同步", f"共 {result.get('rows', 0)} 行。{extra}")

    def show_analysis_results(self, result: AnalysisResult) -> None:
        self.refresh_logs()
        self._analysis_result = result
        if result.headers:
            self.rebuild_analysis_header_checks(result.headers)
        self.refresh_analysis_visuals(result)
        if result.names_found:
            restoring = self._restoring_settings
            self._restoring_settings = True
            typed = self.analysis_names.text().strip()
            self.analysis_names.setPlaceholderText("留空=整个队别")
            if typed:
                self.analysis_names.setText(typed)
            self._restoring_settings = restoring

    def refresh_analysis_visuals(self, result: AnalysisResult) -> None:
        current = result.current
        previous = result.previous
        compare = bool(result.compare_enabled)
        self.analysis_stat_previous.setVisible(compare)
        if compare:
            delta_text = format_delta(result.delta_count, result.delta_count_pct)
            extra = f"{result.scope}　较对比期 {delta_text}"
            self._set_stat_card(
                self.analysis_stat_current,
                f"本期  {current.start} 至 {current.end}",
                f"{format_number(current.count)} 条　　日均 {format_number(current.average)}",
                extra,
                result.delta_count,
            )
            range_label = RANGE_LABELS[COMPARE_MODES.index(result.compare_mode)] if result.compare_mode in COMPARE_MODES else result.compare_mode
            self._set_stat_card(
                self.analysis_stat_previous,
                f"对比期  {previous.start} 至 {previous.end}",
                f"{format_number(previous.count)} 条　　日均 {format_number(previous.average)}",
                range_label,
            )
            name_note = f" · 名字列「{result.name_field}」" if result.name_field else ""
            self.analysis_summary.setText(
                f"{result.scope} · 按「{result.date_field}」{name_note} · 本期 {current.count} 条，对比期 {previous.count} 条。"
            )
            self.analysis_daily_wrap.title_label.setText("每日增长（本期 vs 对比期）")
        else:
            self._set_stat_card(
                self.analysis_stat_current,
                f"{current.start} 至 {current.end}",
                f"{format_number(current.count)} 条　　日均 {format_number(current.average)}",
                result.scope,
            )
            name_note = f" · 名字列「{result.name_field}」" if result.name_field else ""
            self.analysis_summary.setText(
                f"{result.scope} · 按「{result.date_field}」{name_note} · {current.count} 条。"
            )
            self.analysis_daily_wrap.title_label.setText("每日记录")
        self.update_analysis_chart(result)
        self.fill_analysis_daily_table(result)
        self.fill_analysis_people_table(result)

    def update_analysis_chart(self, result: AnalysisResult) -> None:
        current = result.current
        previous = result.previous
        compare = bool(result.compare_enabled)
        labels = [point.day[5:] if len(point.day) >= 10 else point.day for point in current.daily]
        if not labels:
            labels = ["—"]
        series: list[tuple[str, list[tuple[str, float]], QColor]] = []
        show_count = bool(hasattr(self, "analysis_count_check") and self.analysis_count_check.isChecked())
        selected = self.selected_analysis_headers()
        if show_count or not selected:
            series.append(
                ("记录数", [(labels[index], float(point.count)) for index, point in enumerate(current.daily)], QColor("#087fbb")),
            )
            if compare and previous.daily:
                prev_labels = labels
                if len(previous.daily) == len(current.daily):
                    prev_labels = [
                        f"{cur[5:]}/{prev[5:]}" if len(cur) >= 10 and len(prev) >= 10 else labels[index]
                        for index, (cur, prev) in enumerate(
                            zip([p.day for p in current.daily], [p.day for p in previous.daily])
                        )
                    ]
                    series[0] = ("记录数", [(prev_labels[index], float(point.count)) for index, point in enumerate(current.daily)], QColor("#087fbb"))
                    labels = prev_labels
                series.append(
                    (
                        "对比期",
                        [
                            (labels[index] if index < len(labels) else f"第{index + 1}天", float(point.count))
                            for index, point in enumerate(previous.daily)
                        ],
                        QColor("#94a3b8"),
                    )
                )
        color_index = 0
        breakdowns = {item.header: item for item in result.breakdowns}
        session_selected = [header for header in selected if is_session_header(header)]
        for header in selected:
            item = breakdowns.get(header)
            if item is None:
                continue
            for part in item.series:
                daily = list(part.daily) + [0] * max(0, len(labels) - len(part.daily))
                if part.label == header or is_session_header(header):
                    series_label = header
                elif len(selected) > 1:
                    series_label = f"{header} · {part.label}"
                else:
                    series_label = part.label
                series.append(
                    (
                        series_label,
                        [(labels[index], float(daily[index])) for index in range(len(labels))],
                        PIE_COLORS[color_index % len(PIE_COLORS)],
                    )
                )
                color_index += 1
        title = "每日记录数"
        if selected:
            title = "每日 · " + "、".join(selected)
            if session_selected:
                title += "（场记=含D条数）"
        self.analysis_chart.set_series(series, title)
        if session_selected:
            slices = []
            for index, header in enumerate(session_selected):
                item = breakdowns.get(header)
                if item is None or not item.series:
                    continue
                count = float(item.series[0].count)
                if count > 0:
                    slices.append((header, count, PIE_COLORS[index % len(PIE_COLORS)]))
            pie_title = "本期各场含D数量"
        else:
            pie_header = selected[0] if selected else ""
            pie_item = breakdowns.get(pie_header) if pie_header else None
            if pie_item and pie_item.series:
                slices = [
                    (part.label, float(part.count), PIE_COLORS[index % len(PIE_COLORS)])
                    for index, part in enumerate(pie_item.series)
                    if part.count > 0
                ]
                pie_title = f"本期「{pie_item.header}」占比"
            else:
                pie_source = [item for item in result.people if item.count > 0][:8]
                rest = sum(item.count for item in result.people[8:] if item.count > 0)
                slices = [
                    (item.name, float(item.count), PIE_COLORS[index % len(PIE_COLORS)])
                    for index, item in enumerate(pie_source)
                ]
                if rest:
                    slices.append(("其他", float(rest), QColor("#cbd5e1")))
                pie_title = "本期人员占比"
        self.analysis_chart.set_slices(slices, pie_title)
        self.analysis_chart.set_mode("pie" if self.analysis_chart_pie.isChecked() else "line")

    def set_analysis_chart_mode(self, mode: str) -> None:
        pie = mode == "pie"
        self.analysis_chart_line.setChecked(not pie)
        self.analysis_chart_pie.setChecked(pie)
        self.analysis_chart.set_mode("pie" if pie else "line")
        self.persist_workspace_settings()

    def selected_analysis_headers(self) -> list[str]:
        return [name for name in getattr(self, "_analysis_header_selected", []) if name]

    def rebuild_analysis_header_checks(self, headers: list[str]) -> None:
        if not hasattr(self, "analysis_chip_layout"):
            return
        chartable = list_chart_headers(headers)
        self._analysis_chart_headers = chartable
        saved = [name for name in self.selected_analysis_headers() if name in chartable]
        if not saved:
            saved = [name for name in list(self.store.get("analysis_stat_headers", []) or []) if name in chartable]
        self._analysis_header_selected = saved
        visible: list[str] = []
        for name in [*saved, *chartable]:
            if name not in visible:
                visible.append(name)
            if len(visible) >= CHIP_VISIBLE:
                break
        while self.analysis_chip_layout.count():
            item = self.analysis_chip_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()
        restoring = self._restoring_settings
        self._restoring_settings = True
        for name in visible:
            box = QCheckBox(name)
            box.setObjectName("chipCheck")
            box.setChecked(name in saved)
            box.toggled.connect(lambda checked, header=name: self.toggle_analysis_header(header, checked))
            self.analysis_chip_layout.addWidget(box)
        self.analysis_more_headers.setVisible(len(chartable) > CHIP_VISIBLE)
        extra = len(saved) - sum(1 for name in visible if name in saved)
        if extra > 0:
            self.analysis_more_headers.setText(f"更多 +{extra}")
        else:
            self.analysis_more_headers.setText("更多")
        self._restoring_settings = restoring

    def toggle_analysis_header(self, header: str, checked: bool) -> None:
        if getattr(self, "_restoring_settings", False):
            return
        selected = self.selected_analysis_headers()
        if checked and header not in selected:
            selected.append(header)
        if not checked:
            selected = [name for name in selected if name != header]
        self._analysis_header_selected = selected
        self.on_analysis_header_checks()

    def on_analysis_header_checks(self) -> None:
        if getattr(self, "_restoring_settings", False):
            return
        self.persist_workspace_settings()
        if self._analysis_result is not None:
            self.update_analysis_chart(self._analysis_result)

    def open_analysis_header_more(self) -> None:
        headers = list(self._analysis_chart_headers)
        if not headers:
            QMessageBox.information(self, "没有可选项", "当前数据源没有可以拆到曲线上的表头。")
            return
        selected = set(self.selected_analysis_headers())
        dialog = QDialog(self)
        dialog.setWindowTitle("选择曲线显示的数据")
        dialog.resize(360, 420)
        box = QVBoxLayout(dialog)
        hint = QLabel("勾选后按该列的不同取值各画一条线。加友途径按渠道分开；场记（第一场、第二场等）只数含 D 的条数，不是时长。")
        hint.setObjectName("muted")
        hint.setWordWrap(True)
        box.addWidget(hint)
        listing = QListWidget()
        for name in headers:
            item = QListWidgetItem(name)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked if name in selected else Qt.Unchecked)
            listing.addItem(item)
        box.addWidget(listing, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        box.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return
        picked: list[str] = []
        for row in range(listing.count()):
            item = listing.item(row)
            if item and item.checkState() == Qt.Checked:
                picked.append(item.text())
        self._analysis_header_selected = picked
        self.rebuild_analysis_header_checks(self._analysis_chart_headers)
        self.on_analysis_header_checks()

    def fill_analysis_daily_table(self, result: AnalysisResult) -> None:
        current = result.current.daily
        previous = result.previous.daily
        if result.compare_enabled:
            width = max(len(current), len(previous))
            rows: list[list[object]] = []
            for index in range(width):
                cur = current[index] if index < len(current) else None
                prev = previous[index] if index < len(previous) else None
                cur_count = cur.count if cur else 0
                prev_count = prev.count if prev else 0
                rows.append([
                    str(index + 1),
                    cur.day if cur else "",
                    cur_count,
                    prev.day if prev else "",
                    prev_count,
                    cur_count - prev_count,
                ])
            self._fill_delta_table(
                self.analysis_daily_table,
                ["天", "本期日期", "本期", "对比日期", "对比", "增减"],
                rows,
                delta_column=5,
                number_columns={2, 4},
                current_column=2,
                previous_column=4,
            )
            return
        rows = []
        for index, point in enumerate(current):
            prev_count = current[index - 1].count if index else None
            delta = (point.count - prev_count) if prev_count is not None else 0
            rows.append([
                point.day,
                point.count,
                delta if index else 0,
            ])
        self._fill_delta_table(
            self.analysis_daily_table,
            ["日期", "记录数", "较前日"],
            rows,
            delta_column=2,
            number_columns={1},
            current_column=1,
            previous_column=None,
        )
        if rows:
            first = self.analysis_daily_table.item(0, 2)
            if first is not None:
                first.setText("—")
                first.setForeground(DELTA_FLAT)
                first.setBackground(QColor("#ffffff"))

    def fill_analysis_people_table(self, result: AnalysisResult) -> None:
        if result.compare_enabled:
            rows = [
                [item.name, item.count, item.previous_count, item.count - item.previous_count]
                for item in result.people
            ]
            self._fill_delta_table(
                self.analysis_people_table,
                ["人员", "本期", "对比", "增减"],
                rows,
                delta_column=3,
                number_columns={1, 2},
                current_column=1,
                previous_column=2,
            )
            return
        rows = [[item.name, item.count] for item in result.people]
        self._fill_delta_table(
            self.analysis_people_table,
            ["人员", "记录数"],
            rows,
            delta_column=-1,
            number_columns={1},
        )

    def _fill_delta_table(
        self,
        table: QTableWidget,
        headers: list[str],
        rows: list[list[object]],
        delta_column: int,
        number_columns: set[int],
        current_column: int | None = None,
        previous_column: int | None = None,
    ) -> None:
        table.clear()
        table.setColumnCount(len(headers))
        table.setHorizontalHeaderLabels(headers)
        table.setRowCount(len(rows))
        for row_index, row in enumerate(rows):
            for column, value in enumerate(row):
                if column == delta_column:
                    number = float(value or 0)
                    percent = None
                    if current_column is not None and previous_column is not None:
                        percent = delta_percent(
                            float(row[current_column] or 0),
                            float(row[previous_column] or 0),
                        )
                    item = QTableWidgetItem(format_delta(number, percent))
                    item.setForeground(self._delta_color(number))
                    item.setBackground(self._delta_background(number))
                    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                elif column in number_columns:
                    item = QTableWidgetItem(format_number(float(value or 0)))
                    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                else:
                    item = QTableWidgetItem(str(value))
                table.setItem(row_index, column, item)

    @staticmethod
    def _delta_color(value: float | None) -> QColor:
        number = float(value or 0)
        if number > 0:
            return DELTA_UP
        if number < 0:
            return DELTA_DOWN
        return DELTA_FLAT

    @staticmethod
    def _delta_background(value: float | None) -> QColor:
        number = float(value or 0)
        if number > 0:
            return QColor("#dcfce7")
        if number < 0:
            return QColor("#fee2e2")
        return QColor("#ffffff")

    def _set_stat_card(self, card: QFrame, title: str, value: str, extra: str = "", delta: float | None = None) -> None:
        card.title_label.setText(title)
        card.value_label.setText(value)
        card.extra_label.setText(extra)
        if delta is None:
            card.extra_label.setStyleSheet("")
        else:
            color = self._delta_color(delta).name()
            card.extra_label.setStyleSheet(f"color: {color}; font-weight: 600;")

    def _fill_table(self, table: QTableWidget, headers: list[str], rows: list[list[str]]) -> None:
        table.clear()
        table.setColumnCount(len(headers))
        table.setHorizontalHeaderLabels(headers)
        table.setRowCount(len(rows))
        for row_index, row in enumerate(rows):
            for column, value in enumerate(row):
                table.setItem(row_index, column, QTableWidgetItem(str(value)))

    def run_query(self) -> None:
        values = split_names(self.query_value.toPlainText())
        if not values:
            QMessageBox.warning(self, "请输入", "请输入一个或多个查询号码。")
            return
        engine = DataEngine(self.store)
        field = self.query_field.currentText().strip()
        if not field:
            QMessageBox.warning(self, "请选择", "请选择或输入要查询的字段。")
            return
        source = self.current_query_mode()
        exact = not self.query_fuzzy.isChecked()
        extract_target, extract_sheet = self.extract_query_target()
        source_id = self.current_query_source_id()
        date_field = self.query_date_field.currentText().strip() if self.query_date_enabled.isChecked() else ""
        start_date = self.query_start_date.date().toPython() if date_field else None
        end_date = self.query_end_date.date().toPython() if date_field else None
        if source == "extract" and not extract_target:
            QMessageBox.warning(self, "缺少提取表", "请先在「时间提取」页填写目标 Google 表格链接或本地输出文件。")
            return
        self.persist_workspace_settings()
        result_fields = query_result_headers(
            self.store,
            source,
            field,
            engine.list_query_fields(source, extract_target, extract_sheet, source_id),
            self.current_query_source_name(),
        )
        if hasattr(self, "query_result_summary"):
            self.query_result_summary.setText(f"正在查询：输入 {len(values)} 个值…")
        excludes = split_names(self.query_exclude.text()) if hasattr(self, "query_exclude") else []
        self.run_task(
            lambda: engine.query_many(
                field, values, source == "direct", exact, source, extract_target, extract_sheet, source_id,
                date_field, start_date, end_date, False, excludes,
            ),
            lambda results: self.show_query_results(results, result_fields, field),
            "正在查询本地库…",
        )

    def run_refresh_cache(self) -> None:
        self.persist_workspace_settings()
        engine = DataEngine(self.store)
        mode = self.current_query_mode()
        if mode == "extract":
            job = lambda: engine.refresh_cache([], include_extract=True)
        elif mode == "direct":
            source_id = self.current_query_source_id()
            ids = [source_id] if source_id else [source.id for source in self.store.load_sources() if source.enabled]
            job = lambda: engine.refresh_cache(ids)
        else:
            QMessageBox.information(self, "无需刷新", "汇总库本身就是本地数据，直接查询即可。")
            return
        self.run_task(job, self.refresh_cache_finished, "正在从表格同步到本地库…", self.refresh_cache_button)

    def refresh_cache_finished(self, result: dict[str, int]) -> None:
        self.update_query_cache_status()
        self.refresh_query_fields()
        self.refresh_logs()
        if result.get("unchanged"):
            extra = "内容与本地库一致，无需重写。"
        else:
            extra = "已用表格最新数据替换本地库。"
        QMessageBox.information(self, "本地库已同步", f"共 {result.get('rows', 0)} 行。{extra}")

    @staticmethod
    def corrected_source(source: str) -> str:
        return DataEngine.corrected_source(source)

    @staticmethod
    def record_value(record: Record, *names: str) -> str:
        for name in names:
            value = record.values.get(name, "")
            if value:
                return value
        return ""

    def show_query_results(
        self,
        results: list[tuple[str, Record | None]],
        fields: list[str] | None = None,
        query_field: str = "号码",
    ) -> None:
        if not fields:
            fields = self._current_result_fields() or [query_field]
        phone_table = self.current_table_is_phone(fields)
        display_fields = DEFAULT_QUERY_RESULT_FIELDS.copy() if phone_table else list(fields)
        self.query_table.setColumnCount(len(display_fields))
        self.query_table.setHorizontalHeaderLabels(display_fields)
        self.query_table.setRowCount(len(results))
        self.query_copy_rows: list[list[str]] = [display_fields]
        engine = DataEngine(self.store)
        query_canonical = engine._canonical_field(query_field).casefold()
        found = 0
        for row, (query_value, record) in enumerate(results):
            if record is None:
                values = ["" for _ in display_fields]
                for column, header_name in enumerate(display_fields):
                    if header_name.strip() in {"输入电话号码", "输入电话", "查询值"}:
                        values[column] = query_value
                    elif engine._canonical_field(header_name).casefold() == query_canonical:
                        values[column] = query_value
                        break
            else:
                found += 1
                values = [
                    query_value if header_name.strip() in {"输入电话号码", "输入电话", "查询值"}
                    else engine.query_display_value(record, header_name)
                    for header_name in display_fields
                ]
            self.query_copy_rows.append([str(item) for item in values])
            for column, value in enumerate(values):
                self.query_table.setItem(row, column, QTableWidgetItem(str(value)))
        header = self.query_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeToContents)
        for column, header_name in enumerate(display_fields):
            if header_name.strip().casefold() in {"评论贴文", "链接", "网址", "url", "link"}:
                header.setSectionResizeMode(column, QHeaderView.Stretch)
        self.copy_query_button.setEnabled(len(results) > 0)
        self.update_copy_button()
        self.update_query_cache_status()
        missing = len(results) - found
        input_count = len({str(query_value) for query_value, _ in results})
        summary = f"查询结果：共 {len(results)} 条；匹配 {found} 条；未找到 {missing} 条；输入 {input_count} 个值"
        if hasattr(self, "query_result_summary"):
            self.query_result_summary.setText(summary)
        message = f"查询完成：共 {len(results)} 条，匹配 {found} 条"
        if self.query_mode.currentIndex() == 0 and missing:
            message += "；未找到的可把表格来源改成「直接查询数据源」再查一次"
        self.statusBar().showMessage(message, 8000)
        self.refresh_logs()

    def copy_query_results(self) -> None:
        rows = getattr(self, "query_copy_rows", [])
        if not rows:
            QMessageBox.information(self, "没有结果", "请先查询。")
            return
        headers = rows[0] if rows and isinstance(rows[0], list) else []
        body = rows[1:] if headers else rows
        if headers and "修正格式" in headers:
            phone_index = next(
                (index for index, name in enumerate(headers) if str(name).strip().casefold() in {item.casefold() for item in PHONE_HEADERS}),
                0,
            )
            extra_index = headers.index("修正格式")
            lines = []
            for row in body:
                phone = row[phone_index] if phone_index < len(row) else ""
                extra = row[extra_index] if extra_index < len(row) else ""
                lines.append(f"{phone}\t{extra}".rstrip())
            QApplication.clipboard().setText("\n".join(lines))
            self.statusBar().showMessage(f"已复制 {len(body)} 行：号码和修正格式", 5000)
            return
        text_rows = rows if headers else [[str(item) for item in row] for row in rows]
        QApplication.clipboard().setText("\n".join("\t".join(row) for row in text_rows))
        self.statusBar().showMessage(f"已复制 {max(0, len(body))} 行当前表头结果", 5000)

    def choose_output(self) -> None:
        path, _ = QFileDialog.getSaveFileName(self, "保存提取结果", self.output_path.text(), "Excel 工作簿 (*.xlsx)")
        if path:
            if not path.lower().endswith(".xlsx"):
                path += ".xlsx"
            self.output_path.setText(path)
            self.persist_workspace_settings()

    def update_output_destination(self) -> None:
        google_mode = self.output_type.currentIndex() == 1
        self.google_output_url.setEnabled(google_mode)
        for widget in getattr(self, "local_output_widgets", []):
            widget.setEnabled(not google_mode)
        self.persist_workspace_settings()

    def run_extract(self) -> None:
        output = self.output_path.text().strip()
        google_mode = self.output_type.currentIndex() == 1
        google_url = self.google_output_url.text().strip()
        if not google_mode and not output:
            QMessageBox.warning(self, "缺少路径", "请选择输出文件。")
            return
        if google_mode and not google_url:
            QMessageBox.warning(self, "缺少链接", "请填写目标 Google 表格链接。")
            return
        self.persist_workspace_settings()
        engine = DataEngine(self.store)
        start = self.start_date.date().toPython()
        end = self.end_date.date().toPython()
        fields = split_names(self.dedup_fields.text())
        date_field = self.extract_date_field.currentText()
        direct = self.extract_mode.currentIndex() == 1
        source_id = self.current_extract_source_id()
        sheet_name = self.output_sheet_name.text().strip() or "提取结果"
        self.store.set("google_output_url", google_url)
        self.store.set("google_output_sheet", sheet_name)
        self.store.set("extract_destination_type", "google" if google_mode else "local")
        self.store.set("extract_output_path", output)
        self.extract_status.setText("正在提取…")
        self.run_task(
            lambda: engine.extract(
                date_field, start, end, output, fields, direct, sheet_name,
                "google" if google_mode else "local", google_url, source_id,
            ),
            self.extract_finished,
            "正在按时间提取…",
            self.extract_button,
            self.extract_failed,
        )

    def clear_extract_cache(self) -> None:
        total = self.store.count_extracted()
        if total <= 0:
            QMessageBox.information(self, "本地排重缓存", "当前没有本地提取排重缓存。")
            self.extract_status.setText("本地排重缓存为空。")
            return
        answer = QMessageBox.question(
            self,
            "清除本地排重缓存",
            f"确定清除 {total} 条本地提取排重缓存吗？\n\n"
            "这不会删除数据源、字段配置、汇总库或目标表格数据。清除后仍会根据目标提取表已有行继续排重。",
        )
        if answer != QMessageBox.Yes:
            return
        removed = self.store.clear_extracted()
        self.store.log("INFO", "时间提取", f"已清除本地提取排重缓存 {removed} 条")
        self.extract_status.setText(f"已清除本地排重缓存：{removed} 条。")
        self.refresh_logs()
        QMessageBox.information(self, "清除完成", f"已清除本地提取排重缓存 {removed} 条。")

    def extract_finished(self, result: dict[str, int]) -> None:
        message = f"提取已结束：已写入 {result['written']} 行；排除重复 {result['duplicates']} 行；无效日期 {result['invalid_dates']} 行。"
        self.extract_status.setText(message)
        self.refresh_logs()
        QMessageBox.information(self, "提取完成", message)

    def extract_failed(self, message: str) -> None:
        self.extract_status.setText(f"提取已结束：失败。{message}")

    def load_fields_table(self) -> None:
        self.schema_editor.set_schema(self.store.get("column_schema", []))
        self.update_schema_enabled(self.schema_enabled.isChecked())
        aliases = self.store.get("field_aliases", {})
        self.fields_table.setRowCount(len(aliases))
        for row, (field, names) in enumerate(aliases.items()):
            self.fields_table.setItem(row, 0, QTableWidgetItem(field))
            self.fields_table.setItem(row, 1, QTableWidgetItem(",".join(names)))

    def add_field_row(self) -> None:
        self.fields_table.insertRow(self.fields_table.rowCount())

    def update_schema_enabled(self, enabled: bool) -> None:
        self.schema_editor.setEnabled(enabled)

    def fill_example_schema(self) -> None:
        names = ["专页ID", "姓名", "标签", "订阅时间", "性别", "评论贴文", "手机号码", "日期"]
        self.schema_enabled.setChecked(True)
        self.schema_editor.set_schema([
            {"name": name, "column": excel_column(index), "enabled": True}
            for index, name in enumerate(names)
        ])

    def remove_field_row(self) -> None:
        if self.fields_table.currentRow() >= 0:
            self.fields_table.removeRow(self.fields_table.currentRow())

    def save_fields(self) -> None:
        if not self.save_fields_to_store():
            return
        self.refresh_field_controls()
        QMessageBox.information(self, "已保存", f"字段分配已经保存，共 {len(schema_field_names(self.schema_editor.schema()))} 个启用字段。")

    def refresh_field_controls(self) -> None:
        aliases = self.store.get("field_aliases", {})
        mode = "direct" if hasattr(self, "extract_mode") and self.extract_mode.currentIndex() == 1 else "aggregate"
        source_id = self.current_extract_source_id() if mode == "direct" else ""
        fields = DataEngine(self.store).list_query_fields(mode, source_id=source_id) or list(aliases.keys())
        restoring = self._restoring_settings
        self._restoring_settings = True
        current_date = self.extract_date_field.currentText() if hasattr(self, "extract_date_field") else ""
        saved_date = str(self.store.get("extract_date_field", "") or "")
        if hasattr(self, "extract_date_field"):
            self.extract_date_field.clear()
            self.extract_date_field.addItems(fields)
            date_field = current_date if current_date in fields else (saved_date if saved_date in fields else "")
            if not date_field:
                date_field = DataEngine(self.store).suggest_field(fields, "日期") or (fields[0] if fields else "")
            if date_field:
                self.extract_date_field.setCurrentText(date_field)
        self._restoring_settings = restoring
        self.refresh_query_source_picker()
        self.refresh_query_fields()
        self.refresh_analysis_source_picker()
        self.refresh_analysis_fields()

    def on_extract_mode_changed(self) -> None:
        if getattr(self, "_restoring_settings", False):
            return
        self.update_extract_source_visibility()
        self.refresh_field_controls()
        self.persist_workspace_settings()

    def on_extract_source_changed(self) -> None:
        if getattr(self, "_restoring_settings", False):
            return
        self.refresh_field_controls()
        self.persist_workspace_settings()

    def refresh_logs(self) -> None:
        rows = self.store.read_logs()
        self.log_table.setRowCount(len(rows))
        colors = {"ERROR": QColor("#dc2626"), "WARNING": QColor("#d97706"), "INFO": QColor("#2563eb")}
        for row_index, row in enumerate(rows):
            for column, key in enumerate(("created_at", "level", "operation", "message", "detail")):
                item = QTableWidgetItem(str(row[key]))
                if column == 1:
                    item.setForeground(colors.get(str(row[key]), QColor("#334155")))
                self.log_table.setItem(row_index, column, item)
        self.log_table.resizeColumnsToContents()
        self.log_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)

    def clear_logs(self) -> None:
        if QMessageBox.question(self, "清空日志", "确定清空全部运行日志吗？") == QMessageBox.Yes:
            self.store.clear_logs()
            self.refresh_logs()

    def check_updates_silent(self) -> None:
        task = TaskThread(fetch_latest_release, self)
        self.tasks.append(task)

        def done(info: object) -> None:
            self.tasks.remove(task)
            task.deleteLater()
            if not isinstance(info, dict):
                return
            if is_newer(str(info.get("version") or ""), APP_VERSION):
                self.statusBar().showMessage(
                    f"发现新版本 v{info['version']}，可点击「检查并安装更新」直接安装",
                    20000,
                )

        def failed(_message: str) -> None:
            if task in self.tasks:
                self.tasks.remove(task)
            task.deleteLater()

        task.succeeded.connect(done)
        task.failed.connect(failed)
        task.start()

    def _set_update_buttons_enabled(self, enabled: bool) -> None:
        self.sidebar_update_button.setEnabled(enabled)
        if hasattr(self, "update_button"):
            self.update_button.setEnabled(enabled)

    def check_updates(self) -> None:
        self._set_update_buttons_enabled(False)
        self.statusBar().showMessage("正在检查更新…")
        task = TaskThread(fetch_latest_release, self)
        self.tasks.append(task)

        def done(info: object) -> None:
            self._set_update_buttons_enabled(True)
            self.tasks.remove(task)
            task.deleteLater()
            if isinstance(info, dict):
                self.show_update_result(info)

        def failed(message: str) -> None:
            self._set_update_buttons_enabled(True)
            self.statusBar().showMessage("检查更新失败", 8000)
            if task in self.tasks:
                self.tasks.remove(task)
            task.deleteLater()
            QMessageBox.critical(self, "检查更新失败", message)

        task.succeeded.connect(done)
        task.failed.connect(failed)
        task.start()

    def show_update_result(self, info: dict[str, str]) -> None:
        self._set_update_buttons_enabled(True)
        latest = str(info.get("version") or "")
        if is_newer(latest, APP_VERSION):
            box = QMessageBox(self)
            box.setWindowTitle("发现新版本")
            box.setText(f"当前版本：v{APP_VERSION}\n最新版本：v{latest}")
            can_install = sys.platform == "win32" and bool(info.get("installer_url"))
            if can_install:
                box.setInformativeText("软件将自动下载安装包，随后关闭当前版本并启动安装。")
                install_button = box.addButton("立即下载安装", QMessageBox.AcceptRole)
                box.addButton("稍后", QMessageBox.RejectRole)
                box.exec()
                if box.clickedButton() is install_button:
                    self.download_and_install_update(info)
                return
            box.setInformativeText("请到 GitHub Releases 下载当前系统对应的安装包。")
            open_button = box.addButton("打开下载页", QMessageBox.AcceptRole)
            box.addButton("稍后", QMessageBox.RejectRole)
            box.exec()
            if box.clickedButton() is open_button:
                QDesktopServices.openUrl(QUrl(str(info.get("url") or RELEASES_URL)))
            return
        QMessageBox.information(self, "已是最新", f"当前已经是最新版本 v{APP_VERSION}。")

    def download_and_install_update(self, info: dict[str, str]) -> None:
        self._set_update_buttons_enabled(False)
        self.statusBar().showMessage("正在下载安装包，请稍候…")
        task = TaskThread(lambda: download_release_installer(info, self.store.data_dir), self)
        self.tasks.append(task)

        def done(path: object) -> None:
            self._set_update_buttons_enabled(True)
            self.tasks.remove(task)
            task.deleteLater()
            installer = Path(str(path)).resolve()
            updates = (self.store.data_dir / "updates").resolve()
            if installer.parent != updates:
                QMessageBox.critical(self, "更新失败", "安装包路径不在本机更新目录，已拒绝启动。")
                return
            QMessageBox.information(self, "下载完成", "安装包已下载，将关闭当前软件并启动安装程序。")
            subprocess.Popen([str(installer)], cwd=str(installer.parent))
            QApplication.quit()

        def failed(message: str) -> None:
            self._set_update_buttons_enabled(True)
            if task in self.tasks:
                self.tasks.remove(task)
            task.deleteLater()
            self.statusBar().showMessage("更新下载失败", 8000)
            QMessageBox.critical(self, "更新失败", message)

        task.succeeded.connect(done)
        task.failed.connect(failed)
        task.start()

    def _source_auth_label(self, source: SourceConfig) -> str:
        if source.credential_path.strip():
            return "指定账号"
        if self.store.list_credential_paths():
            return "账号池轮询"
        return "公开读取"

    def _reload_credential_list(self) -> None:
        if not hasattr(self, "credential_list"):
            return
        self.credential_list.clear()
        for path in self.store.list_credential_paths():
            item = QListWidgetItem(path)
            item.setToolTip(path)
            self.credential_list.addItem(item)

    def add_default_credentials(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(self, "选择服务账号 JSON", "", "JSON 文件 (*.json)")
        existing = {self.credential_list.item(index).text().casefold() for index in range(self.credential_list.count())}
        for path in paths:
            text = str(path).strip()
            if not text.lower().endswith(".json") or text.casefold() in existing:
                continue
            item = QListWidgetItem(text)
            item.setToolTip(text)
            self.credential_list.addItem(item)
            existing.add(text.casefold())

    def remove_default_credential(self) -> None:
        row = self.credential_list.currentRow()
        if row >= 0:
            self.credential_list.takeItem(row)

    def save_settings(self) -> None:
        self.store.set("global_excludes", split_names(self.global_excludes.toPlainText()))
        self.store.set_credential_paths([
            self.credential_list.item(index).text().strip()
            for index in range(self.credential_list.count())
        ])
        QMessageBox.information(self, "已保存", "设置已经保存。")

    def export_config_file(self) -> None:
        self.persist_workspace_settings()
        self.save_fields_to_store(silent=True)
        default_path = str(Path.home() / "Desktop" / "表数通配置.json")
        path, _ = QFileDialog.getSaveFileName(self, "导出配置", default_path, "JSON 配置文件 (*.json)")
        if not path:
            return
        if not path.lower().endswith(".json"):
            path += ".json"
        payload = self.store.export_config()
        Path(path).write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        self.store.log("INFO", "配置迁移", f"已导出配置：{path}")
        self.refresh_logs()
        QMessageBox.information(self, "导出完成", f"配置已导出：\n{path}")

    def import_config_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "导入配置", "", "JSON 配置文件 (*.json)")
        if not path:
            return
        answer = QMessageBox.question(
            self,
            "导入配置",
            "导入后会覆盖当前数据源、字段映射、列配置、查询和提取设置。\n\n"
            "不会删除本地排重缓存、运行日志或汇总数据库。确定继续吗？",
        )
        if answer != QMessageBox.Yes:
            return
        try:
            payload = json.loads(Path(path).read_text(encoding="utf-8"))
            settings_count, source_count = self.store.import_config(payload)
        except Exception as exc:
            QMessageBox.critical(self, "导入失败", str(exc))
            return
        self.store.log("INFO", "配置迁移", f"已导入配置：{path}，设置 {settings_count} 项，数据源 {source_count} 个")
        self.apply_store_to_ui()
        self.refresh_logs()
        QMessageBox.information(
            self,
            "导入完成",
            f"已导入配置：{source_count} 个数据源，{settings_count} 项设置。\n"
            "如果服务账号 JSON 路径在这台电脑不存在，请到设置里重新选择。",
        )

    def save_fields_to_store(self, silent: bool = False) -> bool:
        if not hasattr(self, "fields_table"):
            return True
        aliases: dict[str, list[str]] = {}
        for row in range(self.fields_table.rowCount()):
            field_item = self.fields_table.item(row, 0)
            names_item = self.fields_table.item(row, 1)
            field = field_item.text().strip() if field_item else ""
            if not field:
                continue
            aliases[field] = split_names(names_item.text() if names_item else "")
        if not aliases:
            if not silent:
                QMessageBox.warning(self, "不能保存", "至少保留一个标准字段。")
            return False
        schema = self.schema_editor.schema()
        self.store.set("field_aliases", aliases)
        self.store.set("column_schema_enabled", self.schema_enabled.isChecked())
        self.store.set("column_schema", schema)
        return True

    def apply_store_to_ui(self) -> None:
        restoring = self._restoring_settings
        self._restoring_settings = True
        try:
            if hasattr(self, "global_excludes"):
                self.global_excludes.setPlainText("\n".join(self.store.get("global_excludes", [])))
                self._reload_credential_list()
            if hasattr(self, "query_mode"):
                saved_source = str(self.store.get("query_source", "extract") or "extract")
                if saved_source in QUERY_SOURCES:
                    self.query_mode.setCurrentIndex(QUERY_SOURCES.index(saved_source))
                self.query_fuzzy.setChecked(bool(self.store.get("query_fuzzy", False)))
                self.query_date_enabled.setChecked(bool(self.store.get("query_date_enabled", False)))
                for widget, key in (
                    (self.query_start_date, "query_start_date"),
                    (self.query_end_date, "query_end_date"),
                ):
                    saved = QDate.fromString(str(self.store.get(key, "") or ""), "yyyy-MM-dd")
                    if saved.isValid():
                        widget.setDate(saved)
            if hasattr(self, "output_type"):
                saved_type = str(self.store.get("extract_destination_type", "local") or "local")
                self.output_type.setCurrentIndex(1 if saved_type == "google" else 0)
                self.google_output_url.setText(str(self.store.get("google_output_url", "") or ""))
                self.output_sheet_name.setText(str(self.store.get("google_output_sheet", "提取结果") or "提取结果"))
                self.output_path.setText(str(self.store.get("extract_output_path", "") or ""))
                self.extract_mode.setCurrentIndex(1 if str(self.store.get("extract_mode", "aggregate")) == "direct" else 0)
                self.dedup_fields.setText(str(self.store.get("extract_dedup_fields", "号码,日期") or "号码,日期"))
                self.extract_signature_enabled.setChecked(bool(self.store.get("extract_signature_enabled", False)))
                self.extract_signature_header.setText(str(self.store.get("extract_signature_header", "签字") or "签字"))
                self.extract_signature_value.setText(str(self.store.get("extract_signature_value", "") or ""))
                self.extract_signature_column.setCurrentText(str(self.store.get("extract_signature_column", "") or ""))
                self.extract_schema_enabled.setChecked(bool(self.store.get("extract_column_schema_enabled", False)))
                self.extract_schema_editor.set_schema(self.store.get("extract_column_schema", []) or [])
                self.extract_schema_editor.setEnabled(self.extract_schema_enabled.isChecked())
            if hasattr(self, "schema_enabled"):
                self.schema_enabled.setChecked(bool(self.store.get("column_schema_enabled", False)))
                self.schema_editor.set_schema(self.store.get("column_schema", []) or [])
                self.update_schema_enabled(self.schema_enabled.isChecked())
                self.load_fields_table()
            self.refresh_sources()
            self.refresh_extract_source_picker()
            self.refresh_query_source_picker()
            self.refresh_analysis_source_picker()
            self.update_output_destination()
            self.update_extract_source_visibility()
            self.update_query_date_controls()
            if hasattr(self, "analysis_source_pick"):
                self.analysis_names.setText(str(self.store.get("analysis_names", "") or ""))
                self.analysis_exclude.setText(str(self.store.get("analysis_exclude_keywords", "") or ""))
                saved_compare = UI_RANGE_ALIASES.get(str(self.store.get("analysis_compare_mode", "month") or "month"), "month")
                if saved_compare in COMPARE_MODES:
                    self.analysis_range.setCurrentIndex(COMPARE_MODES.index(saved_compare))
                self.analysis_compare_enabled.setChecked(bool(self.store.get("analysis_compare_enabled", False)))
                saved_ref = QDate.fromString(str(self.store.get("analysis_reference_date", "") or ""), "yyyy-MM-dd")
                if saved_ref.isValid():
                    self.analysis_reference.setDate(saved_ref)
                self.analysis_count_check.setChecked(bool(self.store.get("analysis_show_count", True)))
                chart_mode = str(self.store.get("analysis_chart_mode", "line") or "line")
                self.analysis_chart_line.setChecked(chart_mode != "pie")
                self.analysis_chart_pie.setChecked(chart_mode == "pie")
                self.analysis_chart.set_mode("pie" if chart_mode == "pie" else "line")
                self.sync_analysis_periods()
                self.refresh_analysis_fields()
            self.refresh_field_controls()
        finally:
            self._restoring_settings = restoring

    def run_task(
        self,
        job: Callable[[], object],
        success: Callable[[object], None],
        status: str,
        button: QPushButton | None = None,
        failure: Callable[[str], None] | None = None,
    ) -> None:
        self.statusBar().showMessage(status)
        if button:
            button.setEnabled(False)
        task = TaskThread(job, self)
        self.tasks.append(task)

        def done(result: object) -> None:
            self.statusBar().showMessage("操作完成", 5000)
            if button:
                button.setEnabled(True)
            success(result)
            self.tasks.remove(task)
            task.deleteLater()

        def failed(message: str) -> None:
            self.statusBar().showMessage("操作失败", 8000)
            if button:
                button.setEnabled(True)
            self.refresh_logs()
            if failure:
                failure(message)
            QMessageBox.critical(self, "操作失败", message)
            self.tasks.remove(task)
            task.deleteLater()

        task.succeeded.connect(done)
        task.failed.connect(failed)
        task.start()


def format_number(value: float | None, digits: int = 2) -> str:
    if value is None:
        return "—"
    number = float(value)
    if abs(number - round(number)) < 1e-9:
        return str(int(round(number)))
    return f"{number:.{digits}f}"


def format_delta(value: float | None, percent: float | None = None) -> str:
    if value is None:
        return "—"
    sign = "+" if value > 0 else ""
    text = f"{sign}{format_number(value)}"
    if percent is not None:
        percent_sign = "+" if percent > 0 else ""
        text += f" ({percent_sign}{percent:.1f}%)"
    return text


def delta_percent(current: float, previous: float) -> float | None:
    if previous == 0:
        return None if current == 0 else 100.0
    return (current - previous) / previous * 100.0


def _pie_contrast(color: QColor) -> QColor:
    luma = 0.2126 * color.red() + 0.7152 * color.green() + 0.0722 * color.blue()
    return QColor("#ffffff") if luma < 150 else QColor("#0f172a")


def _short_label(text: str, limit: int = 10) -> str:
    value = str(text or "").strip()
    if len(value) <= limit:
        return value
    return value[: limit - 1] + "…"


def format_percent(value: float | None) -> str:
    if value is None:
        return "—"
    sign = "+" if value > 0 else ""
    return f"{sign}{value:.1f}%"


class LineChartWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMinimumHeight(180)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.setMouseTracking(True)
        self.title = "每日曲线"
        self.line_title = "每日曲线"
        self.pie_title = "人员占比"
        self.mode = "line"
        self.series: list[tuple[str, list[tuple[str, float]], QColor]] = []
        self.slices: list[tuple[str, float, QColor]] = []
        self._line_hits: list[dict] = []
        self._pie_hits: list[dict] = []
        self._pie_center: tuple[float, float, float] | None = None
        self._hover_key = ""
        self._hover_x: float | None = None

    def set_series(self, series: list[tuple[str, list[tuple[str, float]], QColor]], title: str = "每日曲线") -> None:
        self.series = series
        self.line_title = title
        if self.mode == "line":
            self.title = title
        self._clear_hover()
        self.update()

    def set_slices(self, slices: list[tuple[str, float, QColor]], title: str = "人员占比") -> None:
        self.slices = slices
        self.pie_title = title
        if self.mode == "pie":
            self.title = title
        self._clear_hover()
        self.update()

    def set_mode(self, mode: str) -> None:
        self.mode = "pie" if mode == "pie" else "line"
        self.title = self.pie_title if self.mode == "pie" else self.line_title
        self._clear_hover()
        self.update()

    def _clear_hover(self) -> None:
        self._hover_key = ""
        self._hover_x = None
        QToolTip.hideText()

    def leaveEvent(self, event) -> None:
        if self._hover_key or self._hover_x is not None:
            self._clear_hover()
            self.update()
        super().leaveEvent(event)

    def mouseMoveEvent(self, event) -> None:
        pos = event.position() if hasattr(event, "position") else event.pos()
        if self.mode == "pie":
            self._hover_pie(float(pos.x()), float(pos.y()), event)
        else:
            self._hover_line(float(pos.x()), float(pos.y()), event)
        event.accept()

    def _show_tip(self, event, text: str) -> None:
        pos = event.globalPosition().toPoint() if hasattr(event, "globalPosition") else self.mapToGlobal(QPoint(int(event.pos().x()), int(event.pos().y())))
        QToolTip.showText(pos, text, self)

    def _hover_line(self, mx: float, my: float, event) -> None:
        if not self._line_hits:
            self._clear_hover()
            return
        nearest = min(self._line_hits, key=lambda item: (item["x"] - mx) ** 2 + (item["y"] - my) ** 2)
        dist = math.hypot(nearest["x"] - mx, nearest["y"] - my)
        x_dist = min(abs(item["x"] - mx) for item in self._line_hits)
        if dist > 16 and x_dist > 18:
            if self._hover_key:
                self._clear_hover()
                self.update()
            return
        slot = [item for item in self._line_hits if abs(item["x"] - nearest["x"]) < 0.5]
        lines = [str(nearest["x_label"])]
        for item in slot:
            lines.append(f"{item['label']}：{format_number(item['value'])}")
        text = "\n".join(lines)
        if text != self._hover_key or self._hover_x != nearest["x"]:
            self._hover_key = text
            self._hover_x = nearest["x"]
            self.update()
        self._show_tip(event, text)

    def _hover_pie(self, mx: float, my: float, event) -> None:
        if not self._pie_hits or not self._pie_center:
            self._clear_hover()
            return
        cx, cy, radius = self._pie_center
        if math.hypot(mx - cx, my - cy) > radius + 6:
            if self._hover_key:
                self._clear_hover()
                self.update()
            return
        deg = math.degrees(math.atan2(-(my - cy), mx - cx))
        if deg < 0:
            deg += 360
        angle16 = int(round(deg * 16)) % (360 * 16)
        hit = None
        for item in self._pie_hits:
            delta = (item["start"] - angle16) % (360 * 16)
            if delta <= abs(item["span"]):
                hit = item
                break
        if hit is None:
            if self._hover_key:
                self._clear_hover()
                self.update()
            return
        text = f"{hit['label']}\n{format_number(hit['value'])}（{hit['percent']:.0f}%）"
        if text != self._hover_key:
            self._hover_key = text
        self._show_tip(event, text)

    def paintEvent(self, event) -> None:  # noqa: ARG002
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        bounds = self.rect().adjusted(4, 4, -4, -4)
        painter.fillRect(bounds, QColor("#ffffff"))
        painter.setPen(QPen(QColor("#dce5ef")))
        painter.drawRoundedRect(bounds.adjusted(0, 0, -1, -1), 8, 8)
        self._line_hits = []
        self._pie_hits = []
        self._pie_center = None
        if self.mode == "pie":
            self._paint_pie(painter, bounds)
        else:
            self._paint_line(painter, bounds)

    def _paint_line(self, painter: QPainter, bounds) -> None:
        if not any(points for _label, points, _color in self.series):
            painter.setPen(QColor("#94a3b8"))
            painter.drawText(bounds, Qt.AlignCenter, "暂无曲线数据")
            return
        painter.setPen(QColor("#10213a"))
        painter.setFont(QFont("Microsoft YaHei UI", 10, QFont.DemiBold))
        painter.drawText(bounds.adjusted(12, 8, -12, 0), Qt.AlignTop | Qt.AlignLeft, self.title)
        painter.setFont(QFont("Microsoft YaHei UI", 8))
        metrics = QFontMetrics(painter.font())
        legend_x = bounds.left() + 12
        legend_y = bounds.top() + 28
        legend_right = bounds.right() - 12
        for label, _points, color in self.series:
            width = 18 + metrics.horizontalAdvance(label) + 12
            if legend_x + width > legend_right and legend_x > bounds.left() + 12:
                legend_x = bounds.left() + 12
                legend_y += 16
            painter.setPen(Qt.NoPen)
            painter.setBrush(color)
            painter.drawRoundedRect(legend_x, legend_y + 3, 10, 10, 2, 2)
            painter.setPen(QColor("#334155"))
            painter.drawText(legend_x + 14, legend_y, width - 14, 16, Qt.AlignLeft | Qt.AlignVCenter, label)
            legend_x += width
        plot = bounds.adjusted(48, legend_y - bounds.top() + 20, -16, -28)
        values = [value for _label, points, _color in self.series for _label_x, value in points]
        max_value = max(values) if values else 1.0
        if max_value <= 0:
            max_value = 1.0
        max_value *= 1.12
        painter.setFont(QFont("Microsoft YaHei UI", 8))
        painter.setPen(QColor("#94a3b8"))
        for step in range(5):
            ratio = step / 4
            y = plot.bottom() - (plot.height() * ratio)
            painter.drawLine(plot.left(), int(y), plot.right(), int(y))
            painter.drawText(6, int(y) - 8, plot.left() - 10, 16, Qt.AlignRight | Qt.AlignVCenter, format_number(max_value * ratio))
        labels = next((points for _label, points, _color in self.series if points), [])
        count = max((len(points) for _label, points, _color in self.series), default=1) or 1
        skip = max(1, (len(labels) + 7) // 8)
        for index, (label, _value) in enumerate(labels):
            if index % skip != 0 and index != len(labels) - 1:
                continue
            x = plot.left() if count == 1 else plot.left() + plot.width() * index / max(count - 1, 1)
            painter.drawText(int(x) - 36, plot.bottom() + 4, 72, 18, Qt.AlignHCenter | Qt.AlignTop, label)
        hits: list[dict] = []
        for label, points, color in self.series:
            painter.setPen(QPen(color, 2.4))
            last = None
            for index, (x_label, value) in enumerate(points):
                x = plot.left() if count == 1 else plot.left() + plot.width() * index / max(count - 1, 1)
                y = plot.bottom() - (value / max_value) * plot.height()
                point = QPointF(x, y)
                if last is not None:
                    painter.drawLine(last, point)
                painter.setBrush(color)
                painter.setPen(QPen(color, 2.4))
                painter.drawEllipse(point, 3.0, 3.0)
                hits.append({"x": x, "y": y, "label": label, "x_label": x_label, "value": value, "color": color})
                last = point
        self._line_hits = hits
        if self._hover_x is not None:
            for item in hits:
                if abs(item["x"] - self._hover_x) < 0.5:
                    painter.setPen(QPen(QColor("#ffffff"), 2))
                    painter.setBrush(item["color"])
                    painter.drawEllipse(QPointF(item["x"], item["y"]), 5.4, 5.4)
                    painter.setPen(QPen(item["color"], 1.6))
                    painter.drawEllipse(QPointF(item["x"], item["y"]), 3.2, 3.2)

    def _paint_pie(self, painter: QPainter, bounds) -> None:
        slices = [(label, value, color) for label, value, color in self.slices if value > 0]
        painter.setPen(QColor("#10213a"))
        painter.setFont(QFont("Microsoft YaHei UI", 10, QFont.DemiBold))
        painter.drawText(bounds.adjusted(12, 8, -12, 0), Qt.AlignTop | Qt.AlignLeft, self.title)
        if not slices:
            painter.setPen(QColor("#94a3b8"))
            painter.drawText(bounds, Qt.AlignCenter, "暂无饼图数据")
            return
        total = sum(value for _label, value, _color in slices)
        side = min(bounds.width() - 280, bounds.height() - 48)
        side = max(120, side)
        pie = QRect(bounds.center().x() - side // 2, bounds.top() + 32, side, side)
        if pie.bottom() > bounds.bottom() - 8:
            pie.moveTop(max(bounds.top() + 28, bounds.bottom() - 8 - pie.height()))
        cx = pie.center().x()
        cy = pie.center().y()
        radius = pie.width() / 2
        self._pie_center = (float(cx), float(cy), float(radius))
        items = []
        start = 90 * 16
        for label, value, color in slices:
            span = max(1, int(round(360 * 16 * value / total))) if total else 0
            painter.setBrush(color)
            painter.setPen(QPen(QColor("#ffffff"), 1.5))
            painter.drawPie(pie, start, -span)
            mid = (start - span / 2) / 16.0
            rad = math.radians(mid)
            percent = (value / total * 100) if total else 0
            items.append({
                "label": label,
                "value": value,
                "color": color,
                "percent": percent,
                "cos": math.cos(rad),
                "sin": math.sin(rad),
                "rad": rad,
            })
            self._pie_hits.append({
                "start": start,
                "span": -span,
                "label": label,
                "value": value,
                "percent": percent,
            })
            start -= span
        painter.setFont(QFont("Microsoft YaHei UI", 9, QFont.DemiBold))
        for item in items:
            if item["percent"] < 6:
                continue
            px = cx + radius * 0.62 * item["cos"]
            py = cy - radius * 0.62 * item["sin"]
            painter.setPen(_pie_contrast(item["color"]))
            painter.drawText(int(px - 22), int(py - 9), 44, 18, Qt.AlignCenter, f"{item['percent']:.0f}%")
        left_items = sorted((item for item in items if item["cos"] < 0), key=lambda item: -item["sin"])
        right_items = sorted((item for item in items if item["cos"] >= 0), key=lambda item: -item["sin"])
        painter.setFont(QFont("Microsoft YaHei UI", 8))
        self._paint_pie_labels(painter, left_items, pie, bounds, cx, cy, radius, False)
        self._paint_pie_labels(painter, right_items, pie, bounds, cx, cy, radius, True)

    def _paint_pie_labels(self, painter, items, pie, bounds, cx, cy, radius, right_side: bool) -> None:
        if not items:
            return
        top = pie.top()
        gap = max(18, min(26, (pie.height() - 8) / max(len(items), 1)))
        for index, item in enumerate(items):
            text = f"{_short_label(item['label'])}  {format_number(item['value'])}  {item['percent']:.0f}%"
            metrics = QFontMetrics(painter.font())
            text_w = metrics.horizontalAdvance(text) + 4
            y = top + index * gap + 4
            x_edge = cx + radius * item["cos"]
            y_edge = cy - radius * item["sin"]
            x_mid = cx + (radius + 14) * item["cos"]
            y_mid = cy - (radius + 14) * item["sin"]
            if right_side:
                x_text = min(bounds.right() - 10 - text_w, pie.right() + 18)
                x_elbow = x_text - 6
            else:
                x_text = max(bounds.left() + 10, pie.left() - 18 - text_w)
                x_elbow = x_text + text_w + 6
            painter.setPen(QPen(item["color"], 1.4))
            painter.drawLine(QPointF(x_edge, y_edge), QPointF(x_mid, y_mid))
            painter.drawLine(QPointF(x_mid, y_mid), QPointF(x_elbow, y + 8))
            painter.setPen(QColor("#1e293b"))
            painter.drawText(int(x_text), int(y), text_w, 16, Qt.AlignVCenter | (Qt.AlignLeft if right_side else Qt.AlignRight), text)


STYLE = """
QMainWindow, QWidget { background: #f7f9fc; color: #172033; font-family: "Microsoft YaHei UI"; font-size: 13px; }
#sidebar { background: #0b2748; }
#sidebar QLabel { background: transparent; }
#brand { color: white; font-size: 26px; font-weight: 700; }
#sidebarMuted { color: #9eb7d2; }
#sidebar QPushButton { background: transparent; border: 1px solid #3d6a94; color: #cfe1f5; padding: 6px 10px; }
#sidebar QPushButton:hover { border-color: #0ea5e9; color: white; }
#navigation { background: transparent; border: 0; color: #cfe1f5; outline: none; }
#navigation::item { padding: 12px 13px; border-radius: 7px; margin-bottom: 4px; }
#navigation::item:selected { background: #0ea5e9; color: white; }
#navigation::item:hover { background: #16446f; }
#pageTitle { font-size: 24px; font-weight: 700; color: #10213a; }
#muted { color: #64748b; }
#sectionTitle { font-size: 13px; font-weight: 700; color: #334155; }
#statCard { background: white; border: 1px solid #dce5ef; border-radius: 10px; }
#statValue { font-size: 15px; font-weight: 700; color: #10213a; }
#card { background: white; border: 1px solid #dce5ef; border-radius: 10px; padding: 16px; }
#columnMapRow { background: white; border: 1px solid #e2eaf2; border-radius: 8px; }
#columnLetter { font-weight: 600; }
#infoBox { background: #e8f5ff; color: #075985; border: 1px solid #b9e4ff; border-radius: 7px; padding: 11px; }
QPushButton { background: white; border: 1px solid #cbd5e1; border-radius: 7px; padding: 8px 14px; }
QPushButton:hover { border-color: #0ea5e9; color: #0369a1; }
QPushButton:disabled { color: #94a3b8; background: #e2e8f0; }
QPushButton#primary { background: #087fbb; border-color: #087fbb; color: white; font-weight: 600; }
QPushButton#primary:hover { background: #0369a1; }
QPushButton#chartToggle { padding: 4px 8px; }
QPushButton#chartToggle:checked { background: #087fbb; border-color: #087fbb; color: white; font-weight: 600; }
QCheckBox#chipCheck { spacing: 4px; padding: 2px 8px; background: white; border: 1px solid #cbd5e1; border-radius: 6px; }
QCheckBox#chipCheck:checked { background: #e8f5ff; border-color: #087fbb; color: #0369a1; }
QLineEdit, QTextEdit, QComboBox, QSpinBox, QDateEdit { background: white; border: 1px solid #cbd5e1; border-radius: 6px; padding: 7px; selection-background-color: #0ea5e9; }
QTableWidget { background: white; alternate-background-color: #f8fafc; border: 1px solid #dce5ef; border-radius: 8px; gridline-color: #e8edf3; }
QHeaderView::section { background: #edf3f8; color: #334155; border: 0; border-bottom: 1px solid #d7e0e9; padding: 9px; font-weight: 600; }
QCheckBox { spacing: 7px; }
QStatusBar { background: white; border-top: 1px solid #e2e8f0; color: #475569; }
"""


def create_app(icon_path: Path | None = None, data_dir: Path | None = None) -> tuple[QApplication, MainWindow]:
    app = QApplication.instance() or QApplication(sys.argv)
    app.setApplicationName(APP_TITLE)
    app.setApplicationVersion(APP_VERSION)
    app.setOrganizationName("SheetDataHub")
    app.setStyle("Fusion")
    # 某些精简 Windows/离屏环境不会自动枚举中文字体，显式注册系统字体。
    for font_path in (Path("C:/Windows/Fonts/msyh.ttc"), Path("C:/Windows/Fonts/simsun.ttc")):
        if font_path.exists() and QFontDatabase.addApplicationFont(str(font_path)) >= 0:
            break
    app.setStyleSheet(STYLE)
    if icon_path and icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))
    window = MainWindow(ConfigStore(data_dir), icon_path)
    return app, window
