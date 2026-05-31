from __future__ import annotations

from pathlib import Path

from PyQt5.QtCore import QObject, QThread, Qt, pyqtSignal
from PyQt5.QtGui import QDesktopServices
from PyQt5.QtWidgets import (
    QCheckBox,
    QDoubleSpinBox,
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
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
from pdf_exporter import export_sheets_to_pdf, list_sheet_names, preview_sheet, print_sheets


class ProcessorWorker(QObject):
    progress = pyqtSignal(int, str)
    log = pyqtSignal(str)
    finished = pyqtSignal(str, list)
    failed = pyqtSignal(str)

    def __init__(self, input_paths: list[str], output_dir: str, config: RuntimeConfig) -> None:
        super().__init__()
        self.input_paths = input_paths
        self.output_dir = output_dir
        self.config = config

    def run(self) -> None:
        try:
            processor = ExcelProcessor(self.config)
            if len(self.input_paths) == 1 and self.input_paths[0].lower().endswith(".xlsx"):
                output_path, results = processor.process_workbook(
                    input_path=self.input_paths[0],
                    output_dir=self.output_dir,
                    drill_length=self.config.default_drill_length_m,
                    progress_callback=self.progress.emit,
                    log_callback=self.log.emit,
                )
            else:
                output_path, results = processor.process_csv_files(
                    csv_paths=self.input_paths,
                    output_dir=self.output_dir,
                    drill_length=self.config.default_drill_length_m,
                    progress_callback=self.progress.emit,
                    log_callback=self.log.emit,
                )
        except Exception as exc:  # noqa: BLE001
            self.failed.emit(str(exc))
            return
        self.finished.emit(str(output_path), results)


class PdfExportWorker(QObject):
    progress = pyqtSignal(int, str)
    log = pyqtSignal(str)
    finished = pyqtSignal(list)
    failed = pyqtSignal(str)

    def __init__(self, input_path: str, sheet_names: list[str], output_dir: str) -> None:
        super().__init__()
        self.input_path = input_path
        self.sheet_names = sheet_names
        self.output_dir = output_dir

    def run(self) -> None:
        try:
            output_paths = export_sheets_to_pdf(
                input_path=self.input_path,
                sheet_names=self.sheet_names,
                output_dir=self.output_dir,
                progress_callback=self.progress.emit,
                log_callback=self.log.emit,
            )
        except Exception as exc:  # noqa: BLE001
            self.failed.emit(str(exc))
            return
        self.finished.emit([str(path) for path in output_paths])


class PrintWorker(QObject):
    progress = pyqtSignal(int, str)
    log = pyqtSignal(str)
    finished = pyqtSignal(int)
    failed = pyqtSignal(str)

    def __init__(self, input_path: str, sheet_names: list[str]) -> None:
        super().__init__()
        self.input_path = input_path
        self.sheet_names = sheet_names

    def run(self) -> None:
        try:
            print_sheets(
                input_path=self.input_path,
                sheet_names=self.sheet_names,
                progress_callback=self.progress.emit,
                log_callback=self.log.emit,
            )
        except Exception as exc:  # noqa: BLE001
            self.failed.emit(str(exc))
            return
        self.finished.emit(len(self.sheet_names))


class MainWindow(QMainWindow):
    RESULT_HEADERS = [
        "Sheet",
        "Số cọc",
        "CĐ đỉnh",
        "LLV",
        "VX",
        "VL",
        "Cao độ đỉnh TT",
        "Cao độ đáy TT",
        "Chiều dài cọc TT",
        "Trạng thái",
        "Ghi chú",
    ]

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Phần mềm tính chỉ số khoan/phun Excel")
        self.resize(1100, 760)

        self.last_output_path: str | None = None
        self.thread: QThread | None = None
        self.worker: ProcessorWorker | None = None
        self.pdf_thread: QThread | None = None
        self.pdf_worker: PdfExportWorker | None = None
        self.print_thread: QThread | None = None
        self.print_worker: PrintWorker | None = None
        self.input_paths: list[str] = []

        self.file_edit = QLineEdit()
        self.file_edit.setReadOnly(True)
        self.output_edit = QLineEdit(str(Path.cwd()))
        self.drill_length_spin = QDoubleSpinBox()
        self.drill_length_spin.setRange(0.01, 999.0)
        self.drill_length_spin.setDecimals(2)
        self.drill_length_spin.setValue(DEFAULT_CONFIG.default_drill_length_m)
        self.drill_length_spin.setSuffix(" m")
        self.design_top_spin = QDoubleSpinBox()
        self.design_top_spin.setRange(-9999.0, 9999.0)
        self.design_top_spin.setDecimals(2)
        self.design_top_spin.setValue(DEFAULT_CONFIG.default_design_top_elevation)
        self.design_bottom_spin = QDoubleSpinBox()
        self.design_bottom_spin.setRange(-9999.0, 9999.0)
        self.design_bottom_spin.setDecimals(2)
        self.design_bottom_spin.setValue(DEFAULT_CONFIG.default_design_bottom_elevation)
        self.llv_threshold_spin = QDoubleSpinBox()
        self.llv_threshold_spin.setRange(0.0, 999999.0)
        self.llv_threshold_spin.setDecimals(2)
        self.llv_threshold_spin.setValue(DEFAULT_CONFIG.flow_min_for_llv)

        self.create_summary_checkbox = QCheckBox("Tạo sheet tổng hợp")
        self.create_summary_checkbox.setChecked(DEFAULT_CONFIG.create_summary_sheet)
        self.highlight_checkbox = QCheckBox("Tô màu vùng tính")
        self.highlight_checkbox.setChecked(DEFAULT_CONFIG.highlight_ranges)
        self.no_overwrite_checkbox = QCheckBox("Không ghi đè file gốc")
        self.no_overwrite_checkbox.setChecked(DEFAULT_CONFIG.do_not_overwrite_original)
        self.include_equal_checkbox = QCheckBox("Tính cả lưu lượng bằng ngưỡng")
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

        self.pdf_file_edit = QLineEdit()
        self.pdf_file_edit.setReadOnly(True)
        self.pdf_output_edit = QLineEdit(str(Path.cwd()))
        self.pdf_sheet_list = QListWidget()
        self.pdf_sheet_list.setMinimumHeight(120)
        self.pdf_load_button = QPushButton("Tải sheet")
        self.pdf_select_all_button = QPushButton("Chọn tất cả")
        self.pdf_clear_button = QPushButton("Bỏ chọn")
        self.pdf_preview_button = QPushButton("Xem trước sheet")
        self.pdf_export_button = QPushButton("Xuất PDF")
        self.pdf_print_button = QPushButton("In sheet đã chọn")

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

        csv_folder_button = QPushButton("Chọn thư mục CSV")
        csv_folder_button.clicked.connect(self.choose_csv_folder)
        self.csv_folder_button = csv_folder_button

        output_button = QPushButton("Chọn thư mục")
        output_button.clicked.connect(self.choose_output_dir)
        self.output_button = output_button

        form_layout.addWidget(QLabel("File Excel / CSV:"), 0, 0)
        form_layout.addWidget(self.file_edit, 0, 1)
        form_layout.addWidget(file_button, 0, 2)
        form_layout.addWidget(csv_folder_button, 0, 3)

        form_layout.addWidget(QLabel("Độ dài mũi khoan/cọc TK:"), 1, 0)
        form_layout.addWidget(self.drill_length_spin, 1, 1)
        form_layout.addWidget(QLabel("Cao độ đỉnh thiết kế:"), 1, 2)
        form_layout.addWidget(self.design_top_spin, 1, 3)

        form_layout.addWidget(QLabel("Cao độ đáy thiết kế:"), 2, 0)
        form_layout.addWidget(self.design_bottom_spin, 2, 1)
        form_layout.addWidget(QLabel("Ngưỡng LLV:"), 2, 2)
        form_layout.addWidget(self.llv_threshold_spin, 2, 3)

        form_layout.addWidget(QLabel("Thư mục lưu:"), 3, 0)
        form_layout.addWidget(self.output_edit, 3, 1)
        form_layout.addWidget(output_button, 3, 2)

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
        root_layout.addWidget(self._build_pdf_export_group(), stretch=1)
        root_layout.addWidget(QLabel("Log:"))
        root_layout.addWidget(self.log_edit, stretch=1)

    def _build_pdf_export_group(self) -> QGroupBox:
        group = QGroupBox("Tool xuất PDF theo sheet")
        layout = QVBoxLayout(group)
        form_layout = QGridLayout()

        pdf_file_button = QPushButton("Chọn Excel")
        pdf_file_button.clicked.connect(self.choose_pdf_file)
        self.pdf_file_button = pdf_file_button

        pdf_output_button = QPushButton("Chọn thư mục")
        pdf_output_button.clicked.connect(self.choose_pdf_output_dir)
        self.pdf_output_button = pdf_output_button

        form_layout.addWidget(QLabel("File Excel:"), 0, 0)
        form_layout.addWidget(self.pdf_file_edit, 0, 1)
        form_layout.addWidget(pdf_file_button, 0, 2)
        form_layout.addWidget(self.pdf_load_button, 0, 3)
        form_layout.addWidget(QLabel("Thư mục PDF:"), 1, 0)
        form_layout.addWidget(self.pdf_output_edit, 1, 1)
        form_layout.addWidget(pdf_output_button, 1, 2)

        button_layout = QHBoxLayout()
        button_layout.addWidget(self.pdf_select_all_button)
        button_layout.addWidget(self.pdf_clear_button)
        button_layout.addWidget(self.pdf_preview_button)
        button_layout.addWidget(self.pdf_export_button)
        button_layout.addWidget(self.pdf_print_button)
        button_layout.addStretch(1)

        layout.addLayout(form_layout)
        layout.addWidget(self.pdf_sheet_list)
        layout.addLayout(button_layout)
        return group

    def _connect_signals(self) -> None:
        self.process_button.clicked.connect(self.start_processing)
        self.open_output_button.clicked.connect(self.open_output_file)
        self.pdf_load_button.clicked.connect(self.load_pdf_sheets)
        self.pdf_select_all_button.clicked.connect(lambda: self.set_pdf_sheets_checked(True))
        self.pdf_clear_button.clicked.connect(lambda: self.set_pdf_sheets_checked(False))
        self.pdf_preview_button.clicked.connect(self.preview_selected_pdf_sheet)
        self.pdf_export_button.clicked.connect(self.start_pdf_export)
        self.pdf_print_button.clicked.connect(self.start_direct_print)

    def choose_file(self) -> None:
        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Chọn file Excel hoặc CSV",
            str(Path.cwd()),
            "Excel/CSV Files (*.xlsx *.csv);;Excel Files (*.xlsx);;CSV Files (*.csv)",
        )
        if file_paths:
            self.input_paths = file_paths
            self.file_edit.setText(self._format_input_display(file_paths))
            if not self.output_edit.text().strip():
                self.output_edit.setText(str(Path(file_paths[0]).parent))

    def choose_csv_folder(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "Chọn thư mục chứa CSV", self.output_edit.text().strip() or str(Path.cwd()))
        if not folder:
            return
        csv_paths = sorted(str(path) for path in Path(folder).glob("*.csv"))
        if not csv_paths:
            QMessageBox.warning(self, "Không có CSV", "Thư mục đã chọn không có file .csv.")
            return
        self.input_paths = csv_paths
        self.file_edit.setText(self._format_input_display(csv_paths))
        self.output_edit.setText(folder)

    def choose_output_dir(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "Chọn thư mục lưu", self.output_edit.text().strip() or str(Path.cwd()))
        if folder:
            self.output_edit.setText(folder)

    def start_processing(self) -> None:
        output_dir = self.output_edit.text().strip()
        validation_error = self._validate_inputs(self.input_paths, output_dir)
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
        runtime_config.default_design_top_elevation = self.design_top_spin.value()
        runtime_config.default_design_bottom_elevation = self.design_bottom_spin.value()
        runtime_config.flow_min_for_llv = self.llv_threshold_spin.value()
        runtime_config.create_summary_sheet = self.create_summary_checkbox.isChecked()
        runtime_config.highlight_ranges = self.highlight_checkbox.isChecked()
        runtime_config.do_not_overwrite_original = self.no_overwrite_checkbox.isChecked()
        runtime_config.include_flow_equal_threshold = self.include_equal_checkbox.isChecked()

        self.thread = QThread(self)
        self.worker = ProcessorWorker(self.input_paths, output_dir, runtime_config)
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

    def _validate_inputs(self, input_paths: list[str], output_dir: str) -> str | None:
        if not input_paths:
            return "Hãy chọn file Excel hoặc CSV."
        suffixes = {Path(path).suffix.lower() for path in input_paths}
        if ".xlsx" in suffixes and (len(input_paths) > 1 or suffixes != {".xlsx"}):
            return "Chỉ chọn 1 file Excel, hoặc chọn nhiều file CSV. Không trộn Excel và CSV."
        if suffixes not in ({".xlsx"}, {".csv"}):
            return "File đầu vào phải là .xlsx hoặc .csv."
        missing_paths = [path for path in input_paths if not Path(path).exists()]
        if missing_paths:
            return f"File không tồn tại: {missing_paths[0]}"
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
        self.csv_folder_button.setEnabled(not busy)
        self.output_button.setEnabled(not busy)

    def _set_pdf_busy(self, busy: bool) -> None:
        self.pdf_file_button.setEnabled(not busy)
        self.pdf_output_button.setEnabled(not busy)
        self.pdf_load_button.setEnabled(not busy)
        self.pdf_select_all_button.setEnabled(not busy)
        self.pdf_clear_button.setEnabled(not busy)
        self.pdf_preview_button.setEnabled(not busy)
        self.pdf_export_button.setEnabled(not busy)
        self.pdf_print_button.setEnabled(not busy)

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

    def _format_input_display(self, input_paths: list[str]) -> str:
        if len(input_paths) == 1:
            return input_paths[0]
        first_path = Path(input_paths[0])
        return f"{len(input_paths)} file CSV - {first_path.parent}"

    def choose_pdf_file(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn file Excel để xuất PDF",
            str(Path.cwd()),
            "Excel Files (*.xlsx)",
        )
        if not file_path:
            return
        self.pdf_file_edit.setText(file_path)
        if not self.pdf_output_edit.text().strip():
            self.pdf_output_edit.setText(str(Path(file_path).parent))
        self.load_pdf_sheets()

    def choose_pdf_output_dir(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "Chọn thư mục lưu PDF", self.pdf_output_edit.text().strip() or str(Path.cwd()))
        if folder:
            self.pdf_output_edit.setText(folder)

    def load_pdf_sheets(self) -> None:
        file_path = self.pdf_file_edit.text().strip()
        if not file_path:
            QMessageBox.warning(self, "Thiếu file", "Hãy chọn file Excel để tải danh sách sheet.")
            return
        if Path(file_path).suffix.lower() != ".xlsx" or not Path(file_path).exists():
            QMessageBox.warning(self, "File không hợp lệ", "File xuất PDF phải là file .xlsx đang tồn tại.")
            return
        try:
            sheet_names = list_sheet_names(file_path)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Lỗi đọc sheet", str(exc))
            return
        self.pdf_sheet_list.clear()
        for sheet_name in sheet_names:
            item = QListWidgetItem(sheet_name)
            item.setCheckState(Qt.Checked)
            self.pdf_sheet_list.addItem(item)
        self.append_log(f"Đã tải {len(sheet_names)} sheet từ file PDF input.")

    def set_pdf_sheets_checked(self, checked: bool) -> None:
        state = Qt.Checked if checked else Qt.Unchecked
        for row_idx in range(self.pdf_sheet_list.count()):
            self.pdf_sheet_list.item(row_idx).setCheckState(state)

    def selected_pdf_sheets(self) -> list[str]:
        selected: list[str] = []
        for row_idx in range(self.pdf_sheet_list.count()):
            item = self.pdf_sheet_list.item(row_idx)
            if item.checkState() == Qt.Checked:
                selected.append(item.text())
        return selected

    def preview_selected_pdf_sheet(self) -> None:
        file_path = self.pdf_file_edit.text().strip()
        sheet_names = self.selected_pdf_sheets()
        if not file_path or not sheet_names:
            QMessageBox.warning(self, "Thiếu thông tin", "Hãy chọn file Excel và ít nhất một sheet để xem trước.")
            return
        try:
            preview_sheet(file_path, sheet_names[0])
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Không xem trước được", str(exc))

    def start_pdf_export(self) -> None:
        file_path = self.pdf_file_edit.text().strip()
        output_dir = self.pdf_output_edit.text().strip()
        sheet_names = self.selected_pdf_sheets()
        if not file_path or Path(file_path).suffix.lower() != ".xlsx" or not Path(file_path).exists():
            QMessageBox.warning(self, "File không hợp lệ", "Hãy chọn file Excel .xlsx để xuất PDF.")
            return
        if not sheet_names:
            QMessageBox.warning(self, "Thiếu sheet", "Hãy chọn ít nhất một sheet để xuất PDF.")
            return
        if not output_dir or not Path(output_dir).exists():
            QMessageBox.warning(self, "Thư mục không hợp lệ", "Hãy chọn thư mục lưu PDF đang tồn tại.")
            return

        self.progress_bar.setValue(0)
        self._set_pdf_busy(True)
        self.pdf_thread = QThread(self)
        self.pdf_worker = PdfExportWorker(file_path, sheet_names, output_dir)
        self.pdf_worker.moveToThread(self.pdf_thread)
        self.pdf_thread.started.connect(self.pdf_worker.run)
        self.pdf_worker.progress.connect(self.on_progress)
        self.pdf_worker.log.connect(self.append_log)
        self.pdf_worker.finished.connect(self.on_pdf_finished)
        self.pdf_worker.failed.connect(self.on_pdf_failed)
        self.pdf_worker.finished.connect(self.pdf_thread.quit)
        self.pdf_worker.failed.connect(self.pdf_thread.quit)
        self.pdf_thread.finished.connect(self.pdf_thread.deleteLater)
        self.pdf_thread.start()

    def start_direct_print(self) -> None:
        file_path = self.pdf_file_edit.text().strip()
        sheet_names = self.selected_pdf_sheets()
        if not file_path or Path(file_path).suffix.lower() != ".xlsx" or not Path(file_path).exists():
            QMessageBox.warning(self, "File không hợp lệ", "Hãy chọn file Excel .xlsx để in.")
            return
        if not sheet_names:
            QMessageBox.warning(self, "Thiếu sheet", "Hãy chọn ít nhất một sheet để in.")
            return
        confirm = QMessageBox.question(
            self,
            "Xác nhận in",
            f"Gửi {len(sheet_names)} sheet tới máy in mặc định?",
        )
        if confirm != QMessageBox.Yes:
            return

        self.progress_bar.setValue(0)
        self._set_pdf_busy(True)
        self.print_thread = QThread(self)
        self.print_worker = PrintWorker(file_path, sheet_names)
        self.print_worker.moveToThread(self.print_thread)
        self.print_thread.started.connect(self.print_worker.run)
        self.print_worker.progress.connect(self.on_progress)
        self.print_worker.log.connect(self.append_log)
        self.print_worker.finished.connect(self.on_print_finished)
        self.print_worker.failed.connect(self.on_print_failed)
        self.print_worker.finished.connect(self.print_thread.quit)
        self.print_worker.failed.connect(self.print_thread.quit)
        self.print_thread.finished.connect(self.print_thread.deleteLater)
        self.print_thread.start()

    def on_pdf_finished(self, output_paths: list) -> None:
        self.progress_bar.setValue(100)
        self._set_pdf_busy(False)
        QMessageBox.information(self, "Hoàn tất", f"Đã xuất {len(output_paths)} file PDF.")

    def on_pdf_failed(self, error_message: str) -> None:
        self._set_pdf_busy(False)
        self.append_log(f"Lỗi xuất PDF: {error_message}")
        QMessageBox.critical(self, "Lỗi xuất PDF", error_message)

    def on_print_finished(self, printed_count: int) -> None:
        self.progress_bar.setValue(100)
        self._set_pdf_busy(False)
        QMessageBox.information(self, "Hoàn tất", f"Đã gửi {printed_count} sheet tới máy in mặc định.")

    def on_print_failed(self, error_message: str) -> None:
        self._set_pdf_busy(False)
        self.append_log(f"Lỗi in trực tiếp: {error_message}")
        QMessageBox.critical(self, "Lỗi in trực tiếp", error_message)
