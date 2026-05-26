from __future__ import annotations

from pathlib import Path

from PyQt5.QtCore import QObject, QThread, pyqtSignal
from PyQt5.QtGui import QDesktopServices
from PyQt5.QtWidgets import (
    QCheckBox,
    QDoubleSpinBox,
    QFileDialog,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QProgressBar,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)
from PyQt5.QtCore import QUrl

from config import DEFAULT_CONFIG, RuntimeConfig, clone_config
from excel_processor import ExcelProcessor


class ProcessorWorker(QObject):
    progress = pyqtSignal(int, str)
    log = pyqtSignal(str)
    finished = pyqtSignal(str, list)
    failed = pyqtSignal(str)

    def __init__(self, input_path: str, output_dir: str, config: RuntimeConfig) -> None:
        super().__init__()
        self.input_path = input_path
        self.output_dir = output_dir
        self.config = config

    def run(self) -> None:
        try:
            processor = ExcelProcessor(self.config)
            output_path, results = processor.process_workbook(
                input_path=self.input_path,
                output_dir=self.output_dir,
                drill_length=self.config.default_drill_length_m,
                progress_callback=self.progress.emit,
                log_callback=self.log.emit,
            )
        except Exception as exc:  # noqa: BLE001
            self.failed.emit(str(exc))
            return
        self.finished.emit(str(output_path), results)


class MainWindow(QMainWindow):
    RESULT_HEADERS = ["Sheet", "Số cọc", "CĐ đỉnh", "LLV", "VX", "VL", "Trạng thái", "Ghi chú"]

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Phần mềm tính chỉ số khoan/phun Excel")
        self.resize(1100, 760)

        self.last_output_path: str | None = None
        self.thread: QThread | None = None
        self.worker: ProcessorWorker | None = None

        self.file_edit = QLineEdit()
        self.output_edit = QLineEdit(str(Path.cwd()))
        self.drill_length_spin = QDoubleSpinBox()
        self.drill_length_spin.setRange(0.01, 999.0)
        self.drill_length_spin.setDecimals(2)
        self.drill_length_spin.setValue(DEFAULT_CONFIG.default_drill_length_m)
        self.drill_length_spin.setSuffix(" m")

        self.create_summary_checkbox = QCheckBox("Tạo sheet tổng hợp")
        self.create_summary_checkbox.setChecked(DEFAULT_CONFIG.create_summary_sheet)
        self.highlight_checkbox = QCheckBox("Tô màu vùng tính")
        self.highlight_checkbox.setChecked(DEFAULT_CONFIG.highlight_ranges)
        self.no_overwrite_checkbox = QCheckBox("Không ghi đè file gốc")
        self.no_overwrite_checkbox.setChecked(DEFAULT_CONFIG.do_not_overwrite_original)
        self.include_equal_checkbox = QCheckBox("Tính cả giá trị lưu lượng đoạn = 1")
        self.include_equal_checkbox.setChecked(DEFAULT_CONFIG.include_flow_equal_threshold)

        self.process_button = QPushButton("Phân tích & Xuất file")
        self.open_output_button = QPushButton("Mở file kết quả")
        self.open_output_button.setEnabled(False)
        self.progress_bar = QProgressBar()
        self.result_table = QTableWidget(0, len(self.RESULT_HEADERS))
        self.result_table.setHorizontalHeaderLabels(self.RESULT_HEADERS)
        self.result_table.horizontalHeader().setStretchLastSection(True)
        self.log_edit = QTextEdit()
        self.log_edit.setReadOnly(True)

        self._build_ui()
        self._connect_signals()

    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)

        root_layout = QVBoxLayout(central)
        form_layout = QGridLayout()

        file_button = QPushButton("Chọn file")
        file_button.clicked.connect(self.choose_file)
        self.file_button = file_button

        output_button = QPushButton("Chọn thư mục")
        output_button.clicked.connect(self.choose_output_dir)
        self.output_button = output_button

        form_layout.addWidget(QLabel("File Excel:"), 0, 0)
        form_layout.addWidget(self.file_edit, 0, 1)
        form_layout.addWidget(file_button, 0, 2)

        form_layout.addWidget(QLabel("Độ dài mũi khoan:"), 1, 0)
        form_layout.addWidget(self.drill_length_spin, 1, 1)

        form_layout.addWidget(QLabel("Thư mục lưu:"), 2, 0)
        form_layout.addWidget(self.output_edit, 2, 1)
        form_layout.addWidget(output_button, 2, 2)

        options_layout = QHBoxLayout()
        options_layout.addWidget(self.create_summary_checkbox)
        options_layout.addWidget(self.highlight_checkbox)
        options_layout.addWidget(self.no_overwrite_checkbox)
        options_layout.addWidget(self.include_equal_checkbox)
        options_layout.addStretch(1)

        button_layout = QHBoxLayout()
        button_layout.addWidget(self.process_button)
        button_layout.addWidget(self.open_output_button)
        button_layout.addStretch(1)

        root_layout.addLayout(form_layout)
        root_layout.addLayout(options_layout)
        root_layout.addLayout(button_layout)
        root_layout.addWidget(self.progress_bar)
        root_layout.addWidget(QLabel("Bảng kết quả:"))
        root_layout.addWidget(self.result_table, stretch=2)
        root_layout.addWidget(QLabel("Log:"))
        root_layout.addWidget(self.log_edit, stretch=1)

    def _connect_signals(self) -> None:
        self.process_button.clicked.connect(self.start_processing)
        self.open_output_button.clicked.connect(self.open_output_file)

    def choose_file(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(self, "Chọn file Excel", str(Path.cwd()), "Excel Files (*.xlsx)")
        if file_path:
            self.file_edit.setText(file_path)
            if not self.output_edit.text().strip():
                self.output_edit.setText(str(Path(file_path).parent))

    def choose_output_dir(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "Chọn thư mục lưu", self.output_edit.text().strip() or str(Path.cwd()))
        if folder:
            self.output_edit.setText(folder)

    def start_processing(self) -> None:
        input_path = self.file_edit.text().strip()
        output_dir = self.output_edit.text().strip()
        validation_error = self._validate_inputs(input_path, output_dir)
        if validation_error:
            QMessageBox.warning(self, "Thiếu thông tin", validation_error)
            return

        self.result_table.setRowCount(0)
        self.log_edit.clear()
        self.progress_bar.setValue(0)
        self.open_output_button.setEnabled(False)
        self.last_output_path = None
        self._set_busy(True)

        runtime_config = clone_config()
        runtime_config.default_drill_length_m = self.drill_length_spin.value()
        runtime_config.create_summary_sheet = self.create_summary_checkbox.isChecked()
        runtime_config.highlight_ranges = self.highlight_checkbox.isChecked()
        runtime_config.do_not_overwrite_original = self.no_overwrite_checkbox.isChecked()
        runtime_config.include_flow_equal_threshold = self.include_equal_checkbox.isChecked()

        self.thread = QThread(self)
        self.worker = ProcessorWorker(input_path, output_dir, runtime_config)
        self.worker.moveToThread(self.thread)

        self.thread.started.connect(self.worker.run)
        self.worker.progress.connect(self.on_progress)
        self.worker.log.connect(self.append_log)
        self.worker.finished.connect(self.on_finished)
        self.worker.failed.connect(self.on_failed)
        self.worker.finished.connect(self.thread.quit)
        self.worker.failed.connect(self.thread.quit)
        self.thread.finished.connect(self.thread.deleteLater)
        self.thread.start()

    def _validate_inputs(self, input_path: str, output_dir: str) -> str | None:
        if not input_path:
            return "Hãy chọn file Excel."
        if not input_path.lower().endswith(".xlsx"):
            return "File đầu vào phải là .xlsx."
        if not Path(input_path).exists():
            return "File Excel không tồn tại."
        if self.drill_length_spin.value() <= 0:
            return "Độ dài mũi khoan phải lớn hơn 0."
        if not output_dir:
            return "Hãy chọn thư mục lưu."
        if not Path(output_dir).exists():
            return "Thư mục lưu không tồn tại."
        return None

    def _set_busy(self, busy: bool) -> None:
        self.process_button.setEnabled(not busy)
        self.file_button.setEnabled(not busy)
        self.output_button.setEnabled(not busy)

    def on_progress(self, percent: int, message: str) -> None:
        self.progress_bar.setValue(percent)
        if message:
            self.append_log(message)

    def append_log(self, message: str) -> None:
        self.log_edit.append(message)

    def on_finished(self, output_path: str, results: list) -> None:
        self.last_output_path = output_path
        self.progress_bar.setValue(100)
        self._fill_result_table(results)
        self.open_output_button.setEnabled(True)
        self._set_busy(False)
        QMessageBox.information(self, "Hoàn tất", f"Đã xử lý xong.\nFile kết quả:\n{output_path}")

    def on_failed(self, error_message: str) -> None:
        self._set_busy(False)
        self.append_log(f"Lỗi: {error_message}")
        QMessageBox.critical(self, "Lỗi xử lý", error_message)

    def _fill_result_table(self, results: list) -> None:
        self.result_table.setRowCount(len(results))
        for row_idx, result in enumerate(results):
            for col_idx, value in enumerate(result.as_table_row()):
                self.result_table.setItem(row_idx, col_idx, QTableWidgetItem(value))
        self.result_table.resizeColumnsToContents()

    def open_output_file(self) -> None:
        if not self.last_output_path:
            return
        output_path = Path(self.last_output_path)
        if output_path.exists():
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(output_path)))
        else:
            QMessageBox.warning(self, "Không tìm thấy file", "File kết quả không còn tồn tại.")
