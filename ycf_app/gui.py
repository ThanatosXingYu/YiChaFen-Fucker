from __future__ import annotations

import csv
import queue
import re
import threading
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse

import pandas as pd
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QApplication,
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListView,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QRadioButton,
    QScrollArea,
    QSizePolicy,
    QSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from .config import (
    APP_ROOT,
    DEFAULT_HOME_URL,
    DEFAULT_OCR_RETRIES,
    DEFAULT_REFRESH_CLASS,
    DEFAULT_REFRESH_NAME,
    DEFAULT_THREAD_COUNT,
    PATHS,
)
from .core import CacheCleaner, DataInputLoader, QueryRunner, YichafenClient, YichafenParser
from .models import QueryField, QueryInfo
from .utils import now_text


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.paths = PATHS
        self.events: queue.Queue = queue.Queue()
        self.queries: list[QueryInfo] = []
        self.fields: list[QueryField] = []
        self.post_url = ""
        self.dataframe: pd.DataFrame | None = None
        self.data_path: Path | None = None
        self.excel_sheets: list[str] = []
        self.mapping_boxes: dict[str, QComboBox] = {}
        self.runner: QueryRunner | None = None
        self.current_run_id = ""
        self.suppress_combo_signal = False
        self.query_details: dict[str, tuple[QueryInfo, list[QueryField], str]] = {}
        self.output_dir_manually_set = False
        self.log_file = self.paths.logs / f"app_{datetime.now().strftime('%Y%m%d')}.log"
        self.settings_stacked: bool | None = None

        self.setWindowTitle("YiChaFen-Fucker - 跨平台 GUI 查询工具")
        self._set_initial_size()
        self._build_ui()
        self._wire_signals()
        self._apply_style()

        self.timer = QTimer(self)
        self.timer.timeout.connect(self._drain_events)
        self.timer.start(100)
        self._update_start_state()

    def _set_initial_size(self) -> None:
        self.setMinimumSize(680, 520)
        screen = QApplication.primaryScreen()
        if not screen:
            self.resize(980, 760)
            return

        rect = screen.availableGeometry()
        width = min(1120, max(680, rect.width() - 120))
        height = min(860, max(520, rect.height() - 120))
        self.resize(width, height)

    def _build_ui(self) -> None:
        self.page = QWidget()
        root = QVBoxLayout(self.page)
        root.setContentsMargins(14, 14, 14, 14)
        root.setSpacing(10)

        link_group = QGroupBox("1. 链接解析")
        link_layout = QGridLayout(link_group)
        link_layout.setColumnStretch(1, 1)
        self.url_edit = QLineEdit(DEFAULT_HOME_URL)
        self.url_edit.setMinimumWidth(0)
        self.url_edit.setPlaceholderText("输入查询主页链接或单条查询链接")
        self.parse_button = QPushButton("解析链接")
        self.query_combo = QComboBox()
        self.query_combo.setMinimumWidth(0)
        self.query_combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.query_combo.addItem("请先解析链接", None)
        self._prepare_combo(self.query_combo, min_chars=28)
        self.query_status = QLabel("等待解析")
        self.query_status.setObjectName("subtleLabel")
        link_layout.addWidget(QLabel("链接"), 0, 0)
        link_layout.addWidget(self.url_edit, 0, 1)
        link_layout.addWidget(self.parse_button, 0, 2)
        link_layout.addWidget(QLabel("查询项"), 1, 0)
        link_layout.addWidget(self.query_combo, 1, 1)
        link_layout.addWidget(self.query_status, 1, 2)
        root.addWidget(link_group)

        mapping_group = QGroupBox("2. 查询条件与数据列匹配")
        mapping_layout = QVBoxLayout(mapping_group)
        csv_row = QHBoxLayout()
        self.csv_button = QPushButton("上传数据文件")
        self.csv_label = QLabel("未上传数据文件")
        self.csv_label.setObjectName("subtleLabel")
        self.csv_label.setWordWrap(True)
        csv_row.addWidget(self.csv_button)
        csv_row.addWidget(self.csv_label, 1)
        mapping_layout.addLayout(csv_row)

        import_row = QHBoxLayout()
        self.delimiter_label = QLabel("文本分隔符")
        self.delimiter_edit = QLineEdit(",")
        self.delimiter_edit.setFixedWidth(72)
        self.sheet_label = QLabel("Excel 工作表")
        self.sheet_combo = QComboBox()
        self._prepare_combo(self.sheet_combo, min_chars=16)
        self.sheet_combo.setEnabled(False)
        self.reload_data_button = QPushButton("重新解析并预览")
        import_row.addWidget(self.delimiter_label)
        import_row.addWidget(self.delimiter_edit)
        import_row.addSpacing(12)
        import_row.addWidget(self.sheet_label)
        import_row.addWidget(self.sheet_combo, 1)
        import_row.addWidget(self.reload_data_button)
        mapping_layout.addLayout(import_row)

        self.preview_table = QTableWidget()
        self.preview_table.setMinimumHeight(140)
        self.preview_table.setMaximumHeight(220)
        self.preview_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.preview_table.setAlternatingRowColors(True)
        mapping_layout.addWidget(self.preview_table)

        self.mapping_container = QWidget()
        self.mapping_rows = QVBoxLayout(self.mapping_container)
        self.mapping_rows.setContentsMargins(0, 0, 0, 0)
        self.mapping_rows.setSpacing(8)
        self._add_mapping_row("查询条件", QLabel("请先解析查询项并上传数据文件"))
        self.mapping_scroll = QScrollArea()
        self.mapping_scroll.setWidgetResizable(True)
        self.mapping_scroll.setMinimumHeight(150)
        self.mapping_scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.mapping_scroll.setWidget(self.mapping_container)
        mapping_layout.addWidget(self.mapping_scroll)
        root.addWidget(mapping_group, 1)

        self.settings_layout = QGridLayout()
        self.settings_layout.setContentsMargins(0, 0, 0, 0)
        self.captcha_group = QGroupBox("3. 验证码策略")
        captcha_layout = QGridLayout(self.captcha_group)
        self.strategy_group = QButtonGroup(self)
        self.ocr_radio = QRadioButton("纯 OCR 识别")
        self.refresh_radio = QRadioButton("自定义刷新请求")
        self.mixed_radio = QRadioButton("混合模式")
        self.mixed_radio.setChecked(True)
        self.strategy_group.addButton(self.ocr_radio)
        self.strategy_group.addButton(self.refresh_radio)
        self.strategy_group.addButton(self.mixed_radio)
        self.ocr_retry_spin = QSpinBox()
        self.ocr_retry_spin.setRange(1, 20)
        self.ocr_retry_spin.setValue(DEFAULT_OCR_RETRIES)
        self.refresh_name_edit = QLineEdit(DEFAULT_REFRESH_NAME)
        self.refresh_class_edit = QLineEdit(DEFAULT_REFRESH_CLASS)
        captcha_layout.addWidget(self.ocr_radio, 0, 0)
        captcha_layout.addWidget(self.refresh_radio, 0, 1)
        captcha_layout.addWidget(self.mixed_radio, 0, 2)
        captcha_layout.addWidget(QLabel("OCR 重试"), 1, 0)
        captcha_layout.addWidget(self.ocr_retry_spin, 1, 1)
        captcha_layout.addWidget(QLabel("次"), 1, 2)
        captcha_layout.addWidget(QLabel("刷新姓名"), 2, 0)
        captcha_layout.addWidget(self.refresh_name_edit, 2, 1)
        captcha_layout.addWidget(QLabel("刷新班级"), 2, 2)
        captcha_layout.addWidget(self.refresh_class_edit, 2, 3)

        self.thread_group = QGroupBox("4. 多线程控制")
        thread_layout = QGridLayout(self.thread_group)
        self.multithread_check = QCheckBox("启用多线程")
        self.thread_spin = QSpinBox()
        self.thread_spin.setRange(1, 32)
        self.thread_spin.setValue(DEFAULT_THREAD_COUNT)
        self.thread_spin.setEnabled(False)
        self.thread_tip = QLabel("默认单线程")
        self.thread_tip.setObjectName("subtleLabel")
        thread_layout.addWidget(self.multithread_check, 0, 0)
        thread_layout.addWidget(QLabel("线程数"), 1, 0)
        thread_layout.addWidget(self.thread_spin, 1, 1)
        thread_layout.addWidget(self.thread_tip, 2, 0, 1, 2)
        self._arrange_settings(stacked=self.width() < 900)
        root.addLayout(self.settings_layout)

        save_group = QGroupBox("5. 保存与进度")
        save_layout = QGridLayout(save_group)
        save_layout.setColumnStretch(1, 1)
        self.output_edit = QLineEdit(str(self.paths.output))
        self.output_edit.setMinimumWidth(0)
        self.output_edit.setReadOnly(True)
        self.output_format_combo = QComboBox()
        self.output_format_combo.addItem("CSV", ".csv")
        self.output_format_combo.addItem("XLSX", ".xlsx")
        self._prepare_combo(self.output_format_combo, min_chars=8)
        self.output_button = QPushButton("选择保存目录")
        self.xlsx_tip_label = QLabel("XLSX 会在每条结果写入后保存并套用居中与边框样式，速度可能较慢。")
        self.xlsx_tip_label.setObjectName("subtleLabel")
        self.xlsx_tip_label.setWordWrap(True)
        self.xlsx_tip_label.hide()
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.count_label = QLabel("已查询 0 / 0")
        self.time_label = QLabel("已用 00:00:00")
        self.eta_label = QLabel("预计剩余 --:--:--")
        self.speed_label = QLabel("速度 0.00 条/秒")
        for label in [self.count_label, self.time_label, self.eta_label, self.speed_label]:
            label.setObjectName("metricLabel")
        save_layout.addWidget(QLabel("导出格式"), 0, 0)
        save_layout.addWidget(self.output_format_combo, 0, 1)
        save_layout.addWidget(self.xlsx_tip_label, 0, 2)
        save_layout.addWidget(QLabel("保存目录"), 1, 0)
        save_layout.addWidget(self.output_edit, 1, 1)
        save_layout.addWidget(self.output_button, 1, 2)
        save_layout.addWidget(self.progress_bar, 2, 0, 1, 3)
        save_layout.addWidget(self.count_label, 3, 0)
        save_layout.addWidget(self.time_label, 3, 1)
        save_layout.addWidget(self.eta_label, 3, 2)
        save_layout.addWidget(self.speed_label, 4, 0)
        root.addWidget(save_group)

        log_group = QGroupBox("6. 实时日志")
        log_layout = QVBoxLayout(log_group)
        log_actions = QHBoxLayout()
        log_actions.addStretch(1)
        self.clear_log_button = QPushButton("清空日志")
        self.clear_log_button.setObjectName("secondaryButton")
        self.clear_cache_button = QPushButton("清理缓存")
        self.clear_cache_button.setObjectName("secondaryButton")
        log_actions.addWidget(self.clear_cache_button)
        log_actions.addWidget(self.clear_log_button)
        log_layout.addLayout(log_actions)
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMinimumHeight(220)
        self.log_text.setFont(QFont("Menlo", 11))
        log_layout.addWidget(self.log_text)
        root.addWidget(log_group, 3)

        controls = QHBoxLayout()
        controls.addStretch(1)
        self.start_button = QPushButton("开始查询")
        self.pause_button = QPushButton("暂停")
        self.pause_button.setEnabled(False)
        self.stop_button = QPushButton("停止")
        self.stop_button.setEnabled(False)
        controls.addWidget(self.start_button)
        controls.addWidget(self.pause_button)
        controls.addWidget(self.stop_button)
        root.addLayout(controls)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.scroll_area.setFrameShape(QFrame.Shape.NoFrame)
        self.scroll_area.setWidget(self.page)
        self.setCentralWidget(self.scroll_area)

    def _arrange_settings(self, *, stacked: bool) -> None:
        if self.settings_stacked == stacked:
            return
        for group in (self.captcha_group, self.thread_group):
            self.settings_layout.removeWidget(group)

        if stacked:
            self.settings_layout.addWidget(self.captcha_group, 0, 0)
            self.settings_layout.addWidget(self.thread_group, 1, 0)
            self.settings_layout.setColumnStretch(0, 1)
            self.settings_layout.setColumnStretch(1, 0)
        else:
            self.settings_layout.addWidget(self.captcha_group, 0, 0)
            self.settings_layout.addWidget(self.thread_group, 0, 1)
            self.settings_layout.setColumnStretch(0, 2)
            self.settings_layout.setColumnStretch(1, 1)
        self.settings_stacked = stacked

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        if hasattr(self, "settings_layout"):
            self._arrange_settings(stacked=self.width() < 900)

    def _wire_signals(self) -> None:
        self.parse_button.clicked.connect(self.parse_link)
        self.query_combo.currentIndexChanged.connect(self._on_query_changed)
        self.csv_button.clicked.connect(self.load_csv)
        self.reload_data_button.clicked.connect(self.reload_data_file)
        self.sheet_combo.currentTextChanged.connect(self.reload_data_file)
        self.delimiter_edit.returnPressed.connect(self.reload_data_file)
        self.output_button.clicked.connect(self.choose_output_path)
        self.output_format_combo.currentIndexChanged.connect(self._on_output_format_changed)
        self.output_edit.textChanged.connect(self._update_start_state)
        self.clear_log_button.clicked.connect(self.clear_log)
        self.clear_cache_button.clicked.connect(self.clear_cache)
        self.start_button.clicked.connect(self.start_query)
        self.pause_button.clicked.connect(self.toggle_pause_query)
        self.stop_button.clicked.connect(self.stop_query)
        self.multithread_check.toggled.connect(self._on_multithread_toggled)
        self.ocr_radio.toggled.connect(self._update_strategy_controls)
        self.refresh_radio.toggled.connect(self._update_strategy_controls)
        self.mixed_radio.toggled.connect(self._update_strategy_controls)
        self.ocr_retry_spin.valueChanged.connect(self._update_start_state)
        self.refresh_name_edit.textChanged.connect(self._update_start_state)
        self.refresh_class_edit.textChanged.connect(self._update_start_state)

    def _apply_style(self) -> None:
        self.setStyleSheet(
            """
            QWidget {
                font-size: 14px;
            }
            QMainWindow {
                background: #f5f7fb;
            }
            QGroupBox {
                background: #ffffff;
                border: 1px solid #d9e1ec;
                border-radius: 6px;
                margin-top: 10px;
                padding: 12px;
                font-weight: 600;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 4px;
            }
            QLineEdit, QComboBox, QSpinBox, QTextEdit {
                border: 1px solid #cbd5e1;
                border-radius: 4px;
                padding: 6px;
                background: #ffffff;
            }
            QPushButton {
                border: 1px solid #2f6fed;
                border-radius: 4px;
                padding: 8px 14px;
                background: #2f6fed;
                color: white;
                font-weight: 600;
            }
            QPushButton:disabled {
                background: #d8dee9;
                border-color: #c6cfda;
                color: #7b8794;
            }
            QPushButton#secondaryButton {
                background: #ffffff;
                color: #2f4057;
                border-color: #b8c4d2;
                font-weight: 500;
            }
            QComboBox {
                color: #172033;
                selection-color: #ffffff;
                selection-background-color: #2f6fed;
                combobox-popup: 0;
            }
            QComboBox::drop-down {
                width: 28px;
                border-left: 1px solid #cbd5e1;
                background: #f7f9fc;
            }
            QComboBox QAbstractItemView {
                color: #172033;
                background: #ffffff;
                border: 1px solid #b8c4d2;
                outline: 0;
                selection-color: #ffffff;
                selection-background-color: #2f6fed;
                padding: 4px;
            }
            QComboBox QAbstractItemView::item {
                min-height: 28px;
                padding: 6px 8px;
            }
            QComboBox QAbstractItemView::item:hover {
                color: #ffffff;
                background: #2f6fed;
            }
            QComboBox QAbstractItemView::item:selected {
                color: #ffffff;
                background: #2f6fed;
            }
            QProgressBar {
                border: 1px solid #cbd5e1;
                border-radius: 4px;
                text-align: center;
                height: 20px;
                background: #eef2f7;
            }
            QProgressBar::chunk {
                background: #19a974;
                border-radius: 3px;
            }
            QLabel#subtleLabel {
                color: #667085;
            }
            QLabel#metricLabel {
                color: #1f2937;
                font-weight: 600;
            }
            QLabel#pendingLabel {
                color: #667085;
                padding: 7px 4px;
            }
            """
        )

    def _prepare_combo(self, combo: QComboBox, *, min_chars: int = 18, max_visible: int = 12) -> None:
        if not combo.property("customListView"):
            combo.setView(QListView(combo))
            combo.setProperty("customListView", True)
        combo.setMinimumContentsLength(min_chars)
        combo.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon)
        combo.setMaxVisibleItems(max_visible)
        combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        combo.view().setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        combo.view().setTextElideMode(Qt.TextElideMode.ElideNone)
        combo.view().setMaximumHeight(max_visible * 34 + 12)
        if combo.count():
            metrics = combo.fontMetrics()
            widest = max(metrics.horizontalAdvance(combo.itemText(index)) for index in range(combo.count()))
            combo.view().setMinimumWidth(max(260, min(widest + 80, 720)))
        if not combo.property("tooltipSignalConnected"):
            combo.currentTextChanged.connect(combo.setToolTip)
            combo.setProperty("tooltipSignalConnected", True)
        combo.setToolTip(combo.currentText())

    def _safe_filename(self, value: str) -> str:
        name = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", value.strip())
        name = re.sub(r"\s+", "_", name).strip(" ._")
        if not name:
            name = "未命名查询"
        reserved = {
            "CON",
            "PRN",
            "AUX",
            "NUL",
            *(f"COM{index}" for index in range(1, 10)),
            *(f"LPT{index}" for index in range(1, 10)),
        }
        if name.upper() in reserved:
            name = f"{name}_query"
        return name[:80]

    def _default_output_path(self, query_name: str | None = None) -> Path:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        title = self._safe_filename(query_name or "未选择查询")
        return self._output_directory() / f"{timestamp}_{title}{self._selected_output_ext()}"

    def _selected_output_ext(self) -> str:
        return self.output_format_combo.currentData() if hasattr(self, "output_format_combo") else ".csv"

    def _output_directory(self) -> Path:
        path_text = self.output_edit.text().strip()
        path = Path(path_text) if path_text else self.paths.output
        if path.suffix.lower() in {".csv", ".xlsx"}:
            return path.parent
        return path

    def _refresh_default_output_path(self) -> None:
        if self.output_dir_manually_set:
            return
        self.output_edit.setText(str(self.paths.output))

    def _on_output_format_changed(self) -> None:
        ext = self._selected_output_ext()
        self.xlsx_tip_label.setVisible(ext == ".xlsx")
        self._refresh_default_output_path()

    def _next_available_output_path(self, path: Path) -> Path:
        if not path.exists():
            return path
        for index in range(2, 1000):
            candidate = path.with_name(f"{path.stem}_{index}{path.suffix}")
            if not candidate.exists():
                return candidate
        return path.with_name(f"{path.stem}_{datetime.now().strftime('%f')}{path.suffix}")

    def parse_link(self) -> None:
        url = self.url_edit.text().strip()
        if not url:
            QMessageBox.warning(self, "提示", "请先输入查询主页链接或单条查询链接")
            return
        parsed = urlparse(url)
        if not parsed.scheme or not parsed.netloc:
            QMessageBox.warning(self, "提示", "请输入完整 URL，例如 https://r0l2pzad.yichafen.com/")
            return

        self.parse_button.setEnabled(False)
        self.query_status.setText("解析中...")
        self._append_log(f"开始解析链接：{url}")
        threading.Thread(target=self._parse_link_worker, args=(url,), daemon=True).start()

    def _parse_link_worker(self, url: str) -> None:
        try:
            client = YichafenClient(self.paths, "parser")
            path = urlparse(url).path
            if "/qz/" in path:
                html = client.open_query_page(url)
                title, fields, post_url = YichafenParser.parse_query_page(url, html)
                info = QueryInfo(name=title, url=url)
                self.events.put(
                    {
                        "kind": "direct_query_done",
                        "query": info,
                        "fields": fields,
                        "post_url": post_url,
                    }
                )
            else:
                html = client.get_html(url)
                queries = YichafenParser.parse_home(url, html)
                if not queries:
                    raise ValueError("主页没有解析到任何 /qz/ 查询链接")
                self.events.put({"kind": "home_done", "queries": queries})
        except Exception as exc:
            self.events.put({"kind": "parse_error", "message": str(exc)})

    def _on_query_changed(self, index: int) -> None:
        if self.suppress_combo_signal or index < 0 or index >= len(self.queries):
            return
        info = self.queries[index]
        cached = self.query_details.get(info.url)
        if cached:
            self._append_log(f"切换查询项：{cached[0].name}")
            self._apply_query_fields(*cached)
            return

        self.fields = []
        self.post_url = ""
        self._refresh_mapping_area()
        self._update_start_state()
        self._append_log(f"切换查询项：{info.name}")
        self.query_status.setText("解析查询条件中...")
        threading.Thread(target=self._parse_query_worker, args=(info,), daemon=True).start()

    def _parse_query_worker(self, info: QueryInfo) -> None:
        try:
            client = YichafenClient(self.paths, "parser_query")
            html = client.open_query_page(info.url)
            title, fields, post_url = YichafenParser.parse_query_page(info.url, html)
            parsed = QueryInfo(name=title or info.name, url=info.url)
            self.events.put(
                {
                    "kind": "query_done",
                    "query": parsed,
                    "fields": fields,
                    "post_url": post_url,
                }
            )
        except Exception as exc:
            self.events.put({"kind": "parse_error", "message": str(exc)})

    def load_csv(self) -> None:
        path_text, _ = QFileDialog.getOpenFileName(
            self,
            "选择数据文件",
            str(APP_ROOT),
            "数据文件 (*.txt *.csv *.xls *.xlsx);;文本/CSV (*.txt *.csv);;Excel (*.xls *.xlsx)",
        )
        if not path_text:
            return
        self.data_path = Path(path_text)
        self._prepare_data_options()
        self.reload_data_file()

    def _prepare_data_options(self) -> None:
        if self.data_path is None:
            return
        suffix = self.data_path.suffix.lower()
        is_excel = suffix in {".xls", ".xlsx"}
        self.delimiter_edit.setEnabled(not is_excel)
        self.sheet_combo.setEnabled(is_excel)
        self.excel_sheets = []
        self.sheet_combo.blockSignals(True)
        self.sheet_combo.clear()
        if is_excel:
            try:
                self.excel_sheets = DataInputLoader.excel_sheets(self.data_path)
                self.sheet_combo.addItems(self.excel_sheets)
            except Exception as exc:
                QMessageBox.critical(self, "Excel 读取失败", str(exc))
        self.sheet_combo.blockSignals(False)

    def reload_data_file(self) -> None:
        if self.data_path is None:
            return
        try:
            df = self._load_current_data_file()
        except Exception as exc:
            QMessageBox.critical(self, "数据读取失败", str(exc))
            return
        self.dataframe = df
        column_text = ", ".join(map(str, df.columns)) or "无有效列头"
        self.csv_label.setText(f"{self.data_path.name}，{len(df)} 条，列：{column_text}")
        self._append_log(f"已加载数据文件：{self.data_path}，共 {len(df)} 条")
        self._refresh_preview_table()
        self._refresh_mapping_area()
        self._update_start_state()

    def _load_current_data_file(self) -> pd.DataFrame:
        if self.data_path is None:
            raise ValueError("请先上传数据文件")
        suffix = self.data_path.suffix.lower()
        if suffix in {".txt", ".csv"}:
            return DataInputLoader.load_delimited(self.data_path, self._delimiter_value())
        if suffix in {".xls", ".xlsx"}:
            sheet_name = self.sheet_combo.currentText() or 0
            return DataInputLoader.load_excel(self.data_path, sheet_name)
        raise ValueError("仅支持 txt、csv、xls、xlsx 文件")

    def _delimiter_value(self) -> str:
        value = self.delimiter_edit.text()
        escapes = {"\\t": "\t", "\\n": "\n", "\\s": r"\s+"}
        return escapes.get(value, value)

    def _refresh_preview_table(self) -> None:
        dataframe = self.dataframe
        if dataframe is None:
            self.preview_table.clear()
            self.preview_table.setRowCount(0)
            self.preview_table.setColumnCount(0)
            return

        preview = dataframe.head(20)
        self.preview_table.setColumnCount(len(preview.columns))
        self.preview_table.setRowCount(len(preview))
        self.preview_table.setHorizontalHeaderLabels([str(column) for column in preview.columns])
        for row_index, (_, row) in enumerate(preview.iterrows()):
            for column_index, value in enumerate(row):
                self.preview_table.setItem(row_index, column_index, QTableWidgetItem(str(value)))
        self.preview_table.resizeColumnsToContents()

    def choose_output_path(self) -> None:
        directory = QFileDialog.getExistingDirectory(
            self,
            "选择保存目录",
            str(self._output_directory()),
            QFileDialog.Option.ShowDirsOnly,
        )
        if not directory:
            return
        self.output_dir_manually_set = True
        self.output_edit.setText(str(Path(directory)))

    def clear_log(self) -> None:
        self.log_text.clear()
        self.log_file.write_text("", encoding="utf-8")

    def clear_cache(self) -> None:
        removed = CacheCleaner.clean(self.paths)
        self._append_log(f"缓存清理完成：删除 {removed} 个文件或目录")
        QMessageBox.information(self, "清理完成", f"已清理 cookies、验证码图片和临时缓存，共 {removed} 项。")

    def _on_multithread_toggled(self, checked: bool) -> None:
        self.thread_spin.setEnabled(checked)
        if checked:
            QMessageBox.information(
                self,
                "多线程提示",
                "开启多线程后，若错误数据过多，会导致刷新请求大量失效，验证码将强制走 OCR 识别。",
            )
        self._update_start_state()

    def _update_strategy_controls(self) -> None:
        self.ocr_retry_spin.setEnabled(self.ocr_radio.isChecked() or self.mixed_radio.isChecked())
        refresh_enabled = self.refresh_radio.isChecked() or self.mixed_radio.isChecked()
        self.refresh_name_edit.setEnabled(refresh_enabled)
        self.refresh_class_edit.setEnabled(refresh_enabled)
        self._update_start_state()

    def _current_query(self) -> QueryInfo | None:
        index = self.query_combo.currentIndex()
        if 0 <= index < len(self.queries):
            return self.queries[index]
        return None

    def _current_strategy(self) -> str:
        if self.ocr_radio.isChecked():
            return "ocr"
        if self.refresh_radio.isChecked():
            return "refresh"
        return "mixed"

    def _refresh_mapping_area(self) -> None:
        while self.mapping_rows.count():
            item = self.mapping_rows.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.setParent(None)
                widget.deleteLater()
        self.mapping_boxes.clear()

        if not self.fields:
            self._add_mapping_row("查询条件", QLabel("请先解析查询项"))
            return
        if self.dataframe is None:
            for field in self.fields:
                pending = QLabel("已解析，上传数据文件后选择对应列头")
                pending.setObjectName("pendingLabel")
                pending.setWordWrap(True)
                self._add_mapping_row(f"{field.label} ({field.name})", pending)
            return

        columns = [str(column) for column in self.dataframe.columns]
        for field in self.fields:
            combo = QComboBox()
            combo.addItem("")
            combo.addItems(columns)
            self._prepare_combo(combo, min_chars=24, max_visible=8)
            match = self._best_column_match(field, columns)
            if match:
                combo.setCurrentText(match)
            combo.currentIndexChanged.connect(self._update_start_state)
            self.mapping_boxes[field.name] = combo
            self._add_mapping_row(f"{field.label} ({field.name})", combo)

    def _add_mapping_row(self, label_text: str, field_widget: QWidget) -> None:
        row = QWidget()
        row_layout = QGridLayout(row)
        row_layout.setContentsMargins(0, 2, 0, 2)
        row_layout.setHorizontalSpacing(12)
        row_layout.setColumnStretch(0, 0)
        row_layout.setColumnStretch(1, 1)

        label = QLabel(label_text)
        label.setWordWrap(True)
        label.setMinimumWidth(210)
        label.setMinimumHeight(48)
        label.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.MinimumExpanding)
        label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        field_widget.setMinimumHeight(36)
        if isinstance(field_widget, QLabel):
            field_widget.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        else:
            field_widget.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

        row_layout.addWidget(label, 0, 0)
        row_layout.addWidget(field_widget, 0, 1)
        self.mapping_rows.addWidget(row)

    def _best_column_match(self, field: QueryField, columns: list[str]) -> str:
        probes = [field.label, field.name]
        for probe in probes:
            if probe in columns:
                return probe
        normalized = {re.sub(r"\s+", "", col).lower(): col for col in columns}
        for probe in probes:
            key = re.sub(r"\s+", "", probe).lower()
            if key in normalized:
                return normalized[key]
        for column in columns:
            if field.label and (field.label in column or column in field.label):
                return column
        return ""

    def _get_mappings(self) -> dict[str, str] | None:
        if not self.fields:
            return None
        mappings: dict[str, str] = {}
        for field in self.fields:
            combo = self.mapping_boxes.get(field.name)
            value = combo.currentText().strip() if combo else ""
            if not value:
                return None
            mappings[field.name] = value
        return mappings

    def _update_start_state(self) -> None:
        mappings_ready = self._get_mappings() is not None
        strategy = self._current_strategy()
        refresh_ready = True
        if strategy in {"refresh", "mixed"}:
            refresh_ready = bool(self.refresh_name_edit.text().strip()) and bool(self.refresh_class_edit.text().strip())
        ready = (
            self.runner is None
            and self._current_query() is not None
            and bool(self.post_url)
            and bool(self.fields)
            and self.dataframe is not None
            and len(self.dataframe) > 0
            and mappings_ready
            and refresh_ready
            and bool(self.output_edit.text().strip())
        )
        self.start_button.setEnabled(ready)

    def start_query(self) -> None:
        query_info = self._current_query()
        mappings = self._get_mappings()
        if query_info is None or mappings is None or self.dataframe is None:
            QMessageBox.warning(self, "配置不完整", "请完成查询项、数据文件和所有查询条件列匹配")
            return

        output_dir = self._output_directory()
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = self._next_available_output_path(self._default_output_path(query_info.name))
        ext = self._selected_output_ext()
        self.output_edit.setText(str(output_dir))

        if ext == ".xlsx":
            QMessageBox.information(
                self,
                "XLSX 导出提示",
                "XLSX 会在每条结果写入后保存并套用居中与黑色边框样式，数据较多时速度会比 CSV 慢。",
            )

        rows, skipped_duplicates = self._deduplicate_rows(self.dataframe.to_dict(orient="records"), mappings)
        if skipped_duplicates:
            self._append_log(f"已跳过 {skipped_duplicates} 条重复查询条件")
        self._append_log(f"本次输出文件：{output_path}")
        self.progress_bar.setValue(0)
        self.count_label.setText(f"已查询 0 / {len(rows)}")
        self.time_label.setText("已用 00:00:00")
        self.eta_label.setText("预计剩余 --:--:--")
        self.speed_label.setText("速度 0.00 条/秒")

        self.runner = QueryRunner(
            paths=self.paths,
            events=self.events,
            query_info=query_info,
            fields=self.fields,
            post_url=self.post_url,
            rows=rows,
            mappings=mappings,
            output_path=output_path,
            strategy=self._current_strategy(),
            ocr_retries=self.ocr_retry_spin.value(),
            refresh_name=self.refresh_name_edit.text(),
            refresh_class=self.refresh_class_edit.text(),
            multithread=self.multithread_check.isChecked(),
            thread_count=self.thread_spin.value(),
        )
        self.current_run_id = self.runner.run_id
        self.start_button.setEnabled(False)
        self.pause_button.setText("暂停")
        self.pause_button.setEnabled(True)
        self.stop_button.setEnabled(True)
        self.parse_button.setEnabled(False)
        self.csv_button.setEnabled(False)
        self.runner.start()

    def _deduplicate_rows(
        self,
        rows: list[dict[str, str]],
        mappings: dict[str, str],
    ) -> tuple[list[dict[str, str]], int]:
        seen: set[tuple[str, ...]] = set()
        unique_rows: list[dict[str, str]] = []
        field_names = [field.name for field in self.fields]
        for row in rows:
            key = tuple(str(row.get(mappings[field_name], "")).strip() for field_name in field_names)
            if key in seen:
                continue
            seen.add(key)
            unique_rows.append(row)
        return unique_rows, len(rows) - len(unique_rows)

    def stop_query(self) -> None:
        if self.runner:
            self.runner.stop()
            self._append_log("已请求停止，正在等待后台任务收尾...")
            self.pause_button.setEnabled(False)
            self.pause_button.setText("暂停")
            self.stop_button.setEnabled(False)

    def toggle_pause_query(self) -> None:
        if not self.runner:
            return
        if self.runner.is_paused():
            self.runner.resume()
            self.pause_button.setText("暂停")
            self._append_log("已继续查询")
        else:
            self.runner.pause()
            self.pause_button.setText("继续")
            self._append_log("已请求暂停，当前请求结束后暂停新任务")

    def _drain_events(self) -> None:
        while True:
            try:
                event = self.events.get_nowait()
            except queue.Empty:
                break
            self._handle_event(event)

    def _handle_event(self, event: dict[str, Any]) -> None:
        kind = event.get("kind")
        if kind == "log":
            self._append_log(event.get("message", ""))
        elif kind == "parse_error":
            self.parse_button.setEnabled(True)
            self.query_status.setText("解析失败")
            message = event.get("message", "未知错误")
            self._append_log(f"解析失败：{message}")
            QMessageBox.critical(self, "解析失败", message)
        elif kind == "home_done":
            self._apply_home_queries(event["queries"])
        elif kind == "direct_query_done":
            self._apply_direct_query(event["query"], event["fields"], event["post_url"])
        elif kind == "query_done":
            self._apply_query_fields(event["query"], event["fields"], event["post_url"])
        elif kind == "progress":
            if event.get("run_id") == self.current_run_id:
                self._apply_progress(event)
        elif kind == "error":
            self._append_log(event.get("message", "后台错误"))
        elif kind == "finished":
            if event.get("run_id") == self.current_run_id:
                self._finish_run(event)

    def _apply_home_queries(self, queries: list[QueryInfo]) -> None:
        self.queries = queries
        self.fields = []
        self.post_url = ""
        self.suppress_combo_signal = True
        self.query_combo.clear()
        for item in queries:
            self.query_combo.addItem(item.name, item.url)
            self.query_combo.setItemData(self.query_combo.count() - 1, item.url, Qt.ItemDataRole.ToolTipRole)
        self._prepare_combo(self.query_combo, min_chars=28)
        self.suppress_combo_signal = False
        self.parse_button.setEnabled(True)
        self.query_status.setText(f"解析到 {len(queries)} 个查询项")
        self._append_log(f"主页解析完成：{len(queries)} 个查询项")
        if queries:
            self.query_combo.setCurrentIndex(0)
            self._on_query_changed(0)
        self._refresh_mapping_area()
        self._update_start_state()

    def _apply_direct_query(self, query: QueryInfo, fields: list[QueryField], post_url: str) -> None:
        self.query_details[query.url] = (query, fields, post_url)
        existing = next((index for index, item in enumerate(self.queries) if item.url == query.url), -1)
        if existing >= 0:
            self.queries[existing] = query
            selected = existing
        else:
            if not self.queries:
                self.query_combo.clear()
            self.queries.append(query)
            selected = len(self.queries) - 1

        self.suppress_combo_signal = True
        self._reload_query_combo(selected)
        self.suppress_combo_signal = False
        self.parse_button.setEnabled(True)
        self._apply_query_fields(query, fields, post_url)

    def _apply_query_fields(self, query: QueryInfo, fields: list[QueryField], post_url: str) -> None:
        current = self._current_query()
        if current and current.url != query.url:
            self._append_log(f"忽略过期查询解析结果：{query.name}")
            return
        if current and current.url == query.url:
            self.queries[self.query_combo.currentIndex()] = query
            self.query_combo.setItemText(self.query_combo.currentIndex(), query.name)
            self.query_combo.setItemData(self.query_combo.currentIndex(), query.url, Qt.ItemDataRole.ToolTipRole)
        self.query_details[query.url] = (query, fields, post_url)
        self.fields = fields
        self.post_url = post_url
        names = "，".join(f"{field.label}({field.name})" for field in fields)
        self.query_status.setText(f"已解析 {len(fields)} 个条件")
        self._append_log(f"查询条件解析完成：{names}")
        self._append_log(f"提交接口：{post_url}")
        self._refresh_default_output_path()
        self._refresh_mapping_area()
        self._update_start_state()

    def _reload_query_combo(self, selected: int = 0) -> None:
        self.query_combo.clear()
        for item in self.queries:
            self.query_combo.addItem(item.name, item.url)
            self.query_combo.setItemData(self.query_combo.count() - 1, item.url, Qt.ItemDataRole.ToolTipRole)
        self._prepare_combo(self.query_combo, min_chars=28)
        if self.queries:
            self.query_combo.setCurrentIndex(max(0, min(selected, len(self.queries) - 1)))

    def _apply_progress(self, event: dict[str, Any]) -> None:
        done = int(event.get("done", 0))
        total = int(event.get("total", 0))
        speed = float(event.get("speed", 0.0))
        elapsed = float(event.get("elapsed", 0.0))
        eta = float(event.get("eta", 0.0))
        percent = int(done * 100 / total) if total else 0
        self.progress_bar.setValue(percent)
        self.count_label.setText(f"已查询 {done} / {total}，成功 {event.get('success', 0)}，失败 {event.get('failed', 0)}")
        self.time_label.setText(f"已用 {self._format_seconds(elapsed)}")
        eta_text = self._format_seconds(eta) if speed > 0 and done < total else "00:00:00"
        self.eta_label.setText(f"预计剩余 {eta_text}")
        self.speed_label.setText(f"速度 {speed:.2f} 条/秒")

    def _finish_run(self, event: dict[str, Any]) -> None:
        stopped = bool(event.get("stopped"))
        output = event.get("output", "")
        failures = event.get("failures") or []
        self.pause_button.setEnabled(False)
        self.pause_button.setText("暂停")
        self.stop_button.setEnabled(False)
        self.parse_button.setEnabled(True)
        self.csv_button.setEnabled(True)
        self.runner = None
        self.current_run_id = ""
        self._update_start_state()
        title = "查询已停止" if stopped else "查询完成"
        self._append_log(f"{title}，输出文件：{output}")
        failure_log = self._save_failure_report(output, failures)
        QMessageBox.information(
            self,
            title,
            "成功 "
            f"{event.get('success', 0)} 条，失败 {event.get('failed', 0)} 条\n"
            f"结果文件：{output}\n"
            f"失败日志：{failure_log}",
        )

    def _save_failure_report(self, output: str, failures: list[dict[str, str]]) -> Path:
        output_stem = self._safe_filename(Path(output).stem or "未命名查询")
        report_path = self.paths.logs / f"failures_{output_stem}.csv"
        with report_path.open("w", encoding="utf-8-sig", newline="") as file:
            writer = csv.DictWriter(file, fieldnames=["失败时间", "序号", "数据", "原因"])
            writer.writeheader()
            for item in failures:
                writer.writerow(
                    {
                        "失败时间": self._spreadsheet_safe_text(item.get("time", "")),
                        "序号": self._spreadsheet_safe_text(item.get("index", "")),
                        "数据": self._spreadsheet_safe_text(item.get("data", "")),
                        "原因": self._spreadsheet_safe_text(item.get("reason", "")),
                    }
                )
        return report_path

    @staticmethod
    def _spreadsheet_safe_text(value: object) -> str:
        text = str(value or "")
        if text.lstrip().startswith(("=", "+", "-", "@")):
            return "'" + text
        return text

    def _append_log(self, message: str) -> None:
        line = f"[{now_text()}] {message}"
        self.log_text.append(line)
        with self.log_file.open("a", encoding="utf-8") as file:
            file.write(line + "\n")

    @staticmethod
    def _format_seconds(seconds: float) -> str:
        seconds = max(0, int(seconds))
        hours = seconds // 3600
        minutes = (seconds % 3600) // 60
        secs = seconds % 60
        return f"{hours:02d}:{minutes:02d}:{secs:02d}"
