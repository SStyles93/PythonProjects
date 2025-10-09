# app.py (PyQt5) - Professional UI/UX Version

import sys
import os
import logging
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QListWidgetItem
from PyQt5.QtCore import QThread, QObject, pyqtSignal, pyqtSlot

# Import our generated UI and converter engine
from main_window_ui import Ui_MainWindow
from converter_engine import ConverterEngine

# --- Worker for Background Processing (No changes needed here) ---
class Worker(QObject):
    """
    A worker object that runs tasks in a separate thread.
    """
    finished = pyqtSignal()
    progress = pyqtSignal(int)
    status_update = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, files, merge, output_filename):
        super().__init__()
        self.files = files
        self.merge = merge
        self.output_filename = output_filename

    @pyqtSlot()
    def run(self):
        """The main task execution method."""
        try:
            engine = ConverterEngine(self.files)
            
            class QtLogHandler(logging.Handler):
                def __init__(self, status_signal):
                    super().__init__()
                    self.status_signal = status_signal
                
                def emit(self, record):
                    msg = self.format(record)
                    self.status_signal.emit(msg)

            logger = logging.getLogger() 
            log_handler = QtLogHandler(self.status_update)
            logger.addHandler(log_handler)
            
            self.status_update.emit("Starting conversion process...")
            converted_files = engine.convert_to_pdf()
            self.progress.emit(50)

            if not converted_files:
                raise ValueError("No files were successfully converted.")

            if self.merge:
                self.status_update.emit("Starting merge process...")
                success = engine.merge_pdfs(self.output_filename)
                if not success:
                    raise RuntimeError("Failed to merge the PDF files.")
            
            self.progress.emit(100)
            self.status_update.emit("Process completed successfully!")

        except Exception as e:
            error_message = f"An error occurred: {e}"
            logging.error(error_message)
            self.error.emit(error_message)
        finally:
            logger.removeHandler(log_handler)
            self.finished.emit()


# --- Main Application Window ---
class AppWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.setWindowTitle("FileFusion Pro")
        
        # --- NEW: Enable Drag and Drop ---
        self.setAcceptDrops(True)

        # --- NEW: Apply a professional dark theme ---
        self.apply_stylesheet()

        self.thread = None
        self.worker = None

        # --- Connect signals to slots ---
        self.ui.addButton.clicked.connect(self.add_files)
        self.ui.removeButton.clicked.connect(self.remove_selected_file)
        self.ui.clearButton.clicked.connect(self.clear_all_files)
        self.ui.upButton.clicked.connect(self.move_file_up)
        self.ui.downButton.clicked.connect(self.move_file_down)
        self.ui.browseButton.clicked.connect(self.browse_output_file)
        self.ui.startButton.clicked.connect(self.start_processing)
        self.ui.mergeCheckbox.stateChanged.connect(self.toggle_merge_options)

        # --- Initial UI state ---
        self.toggle_merge_options()
        self.ui.progressBar.setValue(0)

    # --- NEW: Method to apply the dark theme stylesheet ---
    def apply_stylesheet(self):
        """Applies a dark theme stylesheet to the application."""
        self.setStyleSheet("""
            QWidget {
                background-color: #2b2b2b;
                color: #ffffff;
                font-family: Segoe UI;
                font-size: 10pt;
            }
            QMainWindow {
                border: 1px solid #1e1e1e;
            }
            QListWidget {
                background-color: #3c3c3c;
                border: 1px solid #1e1e1e;
                padding: 5px;
            }
            QListWidget::item {
                padding: 5px;
            }
            QListWidget::item:hover {
                background-color: #4a4a4a;
            }
            QListWidget::item:selected {
                background-color: #0078d7;
                color: #ffffff;
            }
            QPushButton {
                background-color: #0078d7;
                color: #ffffff;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #005a9e;
            }
            QPushButton:pressed {
                background-color: #003a64;
            }
            QPushButton:disabled {
                background-color: #555555;
                color: #aaaaaa;
            }
            QLineEdit, QCheckBox {
                padding: 5px;
            }
            QLineEdit {
                border: 1px solid #1e1e1e;
                background-color: #3c3c3c;
                border-radius: 4px;
            }
            QProgressBar {
                border: 1px solid #1e1e1e;
                border-radius: 4px;
                text-align: center;
                background-color: #3c3c3c;
            }
            QProgressBar::chunk {
                background-color: #0078d7;
                border-radius: 4px;
            }
            QLabel#statusLabel {
                color: #cccccc;
                font-style: italic;
            }
        """)

    # --- NEW: Drag and Drop Event Handlers ---
    def dragEnterEvent(self, event):
        """Handles files being dragged over the window."""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        """Handles files being dropped onto the window."""
        urls = event.mimeData().urls()
        files_to_add = []
        valid_extensions = ['.doc', '.docx', '.pdf']
        
        for url in urls:
            file_path = url.toLocalFile()
            if os.path.isfile(file_path):
                if any(file_path.lower().endswith(ext) for ext in valid_extensions):
                    files_to_add.append(file_path)
        
        if files_to_add:
            self.add_files_to_list(files_to_add)

    # --- MODIFIED: `add_files` now uses the helper method ---
    def add_files(self):
        """Opens a file dialog to add files to the list."""
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Select Files to Convert",
            "",
            "Documents (*.doc *.docx *.pdf)"
        )
        if files:
            self.add_files_to_list(files)

    # --- NEW: Helper method to add files and prevent duplicates ---
    def add_files_to_list(self, files: list):
        """Adds a list of file paths to the list widget, avoiding duplicates."""
        current_files = {self.ui.fileListWidget.item(i).text() for i in range(self.ui.fileListWidget.count())}
        
        for file in files:
            if file not in current_files:
                self.ui.fileListWidget.addItem(QListWidgetItem(file))
                current_files.add(file)

    def remove_selected_file(self):
        selected_items = self.ui.fileListWidget.selectedItems()
        if not selected_items: return
        for item in selected_items:
            self.ui.fileListWidget.takeItem(self.ui.fileListWidget.row(item))

    def clear_all_files(self):
        self.ui.fileListWidget.clear()

    def move_file_up(self):
        current_row = self.ui.fileListWidget.currentRow()
        if current_row > 0:
            item = self.ui.fileListWidget.takeItem(current_row)
            self.ui.fileListWidget.insertItem(current_row - 1, item)
            self.ui.fileListWidget.setCurrentRow(current_row - 1)

    def move_file_down(self):
        current_row = self.ui.fileListWidget.currentRow()
        if current_row < self.ui.fileListWidget.count() - 1:
            item = self.ui.fileListWidget.takeItem(current_row)
            self.ui.fileListWidget.insertItem(current_row + 1, item)
            self.ui.fileListWidget.setCurrentRow(current_row + 1)

    def toggle_merge_options(self):
        is_merge_enabled = self.ui.mergeCheckbox.isChecked()
        self.ui.outputFilenameEdit.setEnabled(is_merge_enabled)
        self.ui.browseButton.setEnabled(is_merge_enabled)

    def browse_output_file(self):
        output_file, _ = QFileDialog.getSaveFileName(
            self,
            "Save Merged PDF As...",
            "",
            "PDF Files (*.pdf)"
        )
        if output_file:
            self.ui.outputFilenameEdit.setText(output_file)

    def start_processing(self):
        files = [self.ui.fileListWidget.item(i).text() for i in range(self.ui.fileListWidget.count())]
        if not files:
            self.ui.statusLabel.setText("Error: No files selected.")
            return

        merge = self.ui.mergeCheckbox.isChecked()
        output_filename = self.ui.outputFilenameEdit.text()

        if merge and not output_filename:
            self.ui.statusLabel.setText("Error: Please specify an output filename for merging.")
            return

        self.ui.startButton.setEnabled(False)
        self.ui.progressBar.setValue(0)
        self.ui.statusLabel.setText("Preparing to process...")

        self.thread = QThread()
        self.worker = Worker(files, merge, output_filename)
        self.worker.moveToThread(self.thread)

        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.worker.progress.connect(self.ui.progressBar.setValue)
        self.worker.status_update.connect(self.ui.statusLabel.setText)
        self.worker.error.connect(self.ui.statusLabel.setText)
        
        self.thread.finished.connect(lambda: self.ui.startButton.setEnabled(True))

        self.thread.start()

# --- Application Entry Point ---
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = AppWindow()
    window.show()
    sys.exit(app.exec_())
