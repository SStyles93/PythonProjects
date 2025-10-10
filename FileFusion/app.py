import sys
import os
import logging
import tempfile
import shutil
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QListWidgetItem, QDialog, QWidget
from PyQt5.QtCore import QThread, QObject, pyqtSignal, pyqtSlot

# Import your custom and generated UI files
from converter_engine import ConverterEngine
from main_window_ui import Ui_MainWindow
from preview_window import PreviewWindow


# --- Worker for Background Processing (No changes needed) ---
class Worker(QObject):
    finished = pyqtSignal(bool, str, str)
    progress = pyqtSignal(int)
    status_update = pyqtSignal(str)

    def __init__(self, files, merge, output_path):
        super().__init__()
        self.files = files
        self.merge = merge
        self.output_path = output_path

    @pyqtSlot()
    def run(self):
        log_handler = None
        temp_file_handle = None
        try:
            engine = ConverterEngine(self.files)
            
            class QtLogHandler(logging.Handler):
                def __init__(self, status_signal):
                    super().__init__()
                    self.status_signal = status_signal
                def emit(self, record):
                    self.status_signal.emit(self.format(record))

            log_handler = QtLogHandler(self.status_update)
            logger = logging.getLogger() 
            logger.addHandler(log_handler)
            
            self.status_update.emit("Starting conversion process...")

            # Report the progress of the engine
            def report_progress(percentage):
                self.progress.emit(percentage)
            
            # Pass the output folder for individual conversions
            output_folder = None if self.merge else self.output_path
            converted_files = engine.convert_to_pdf(
                output_folder=output_folder,
                progress_callback=report_progress # Pass our function here
            )

            if not converted_files:
                raise ValueError("No files were successfully converted.")
            
            if self.merge:
                self.status_update.emit("Starting merge process...")
                temp_file_handle = open(self.output_path, 'wb')
                engine.merge_pdfs(temp_file_handle)
                temp_file_handle.close()
                temp_file_handle = None
            
            self.progress.emit(100)
            self.finished.emit(True, self.output_path, "")

        except Exception as e:
            error_message = f"An error occurred: {e}"
            logging.error(error_message)
            self.finished.emit(False, "", error_message)
        finally:
            if temp_file_handle:
                temp_file_handle.close()
            if log_handler:
                logger.removeHandler(log_handler)


# --- Main Application Window ---
class AppWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.setWindowTitle("FileFusion")
        self.setAcceptDrops(True)
        self.apply_stylesheet()

        self.thread = None
        self.worker = None
        self.temp_pdf_file = None

        # --- Connect signals to slots ---
        self.ui.addButton.clicked.connect(self.add_files)
        self.ui.removeButton.clicked.connect(self.remove_selected_file)
        self.ui.clearButton.clicked.connect(self.clear_all_files)
        self.ui.upButton.clicked.connect(self.move_file_up)
        self.ui.downButton.clicked.connect(self.move_file_down)
        self.ui.browseButton.clicked.connect(self.browse_output_folder)
        self.ui.startButton.clicked.connect(self.start_processing)
        self.ui.mergeCheckbox.stateChanged.connect(self.update_ui_state)

        # --- Connect file list changes to the UI update method ---
        self.ui.fileListWidget.model().rowsInserted.connect(self.update_ui_state)
        self.ui.fileListWidget.model().rowsRemoved.connect(self.update_ui_state)

        # --- Initial UI state ---
        self.update_ui_state()
        self.ui.progressBar.setValue(0)

    def update_ui_state(self):
        """
        Updates the visibility of widgets based on the number of files and merge state.
        """
        num_files = self.ui.fileListWidget.count()
        
        # Rule 1: 'mergeCheckbox' is only visible if there are 2 or more files.
        self.ui.mergeCheckbox.setVisible(num_files > 1)

        # If there's only one file, ensure the merge box is unchecked.
        if num_files <= 1 and self.ui.mergeCheckbox.isChecked():
            self.ui.mergeCheckbox.setChecked(False)
            # The setChecked(False) will re-trigger this method, so we can return early.
            return

        merge_is_checked = self.ui.mergeCheckbox.isChecked()

        # Rule 2: The output folder widgets are visible only if 'merge' is NOT checked.
        # We use the container widget from your UI file: 'horizontalLayoutWidget'
        show_output_folder_widgets = not merge_is_checked
        self.ui.horizontalLayoutWidget.setVisible(show_output_folder_widgets)

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
                           
            QCheckBox {
                /* Use less horizontal padding to prevent the label from being cut off */
                padding: 0px;
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

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls(): event.acceptProposedAction()
        else: event.ignore()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        files_to_add = []
        valid_extensions = ['.doc', '.docx', '.pdf', '.png', '.jpg', '.jpeg', '.bmp']
        for url in urls:
            file_path = url.toLocalFile()
            if os.path.isfile(file_path) and any(file_path.lower().endswith(ext) for ext in valid_extensions):
                files_to_add.append(file_path)
        if files_to_add: self.add_files_to_list(files_to_add)

    def add_files(self):
        file_filter = "All Supported Files (*.doc *.docx *.pdf *.png *.jpg *.jpeg);;Documents (*.doc *.docx *.pdf);;Images (*.png *.jpg *.jpeg *.bmp)"
        files, _ = QFileDialog.getOpenFileNames(self, "Select Files to Convert", "", file_filter)
        if files: self.add_files_to_list(files)

    def add_files_to_list(self, files: list):
        current_files = {self.ui.fileListWidget.item(i).text() for i in range(self.ui.fileListWidget.count())}
        for file in files:
            if file not in current_files:
                self.ui.fileListWidget.addItem(QListWidgetItem(file))
                current_files.add(file)

    def remove_selected_file(self):
        selected_items = self.ui.fileListWidget.selectedItems()
        if not selected_items: return
        for item in selected_items: self.ui.fileListWidget.takeItem(self.ui.fileListWidget.row(item))

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

    def browse_output_folder(self):
        """Opens a dialog to select an output folder."""
        folder = QFileDialog.getExistingDirectory(self, "Select Output Folder", "")
        if folder:
            self.ui.outputFolderEdit.setText(folder)

    def closeEvent(self, event):
        self.cleanup_temp_file()
        super().closeEvent(event)

    def start_processing(self):
        files = [self.ui.fileListWidget.item(i).text() for i in range(self.ui.fileListWidget.count())]
        if not files:
            self.ui.statusLabel.setText("Error: No files selected.")
            return

        merge = self.ui.mergeCheckbox.isChecked() and self.ui.fileListWidget.count() > 1
        output_path = ""

        if merge:
            temp_file_handle = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
            self.temp_pdf_file = temp_file_handle.name
            temp_file_handle.close()
            output_path = self.temp_pdf_file
        else:
            output_path = self.ui.outputFolderEdit.text()
            if not output_path or not os.path.isdir(output_path):
                self.ui.statusLabel.setText("Error: Please select a valid output folder.")
                return

        self.ui.startButton.setEnabled(False)
        self.ui.progressBar.setValue(0)
        self.ui.statusLabel.setText("Preparing to process...")

        self.thread = QThread()
        self.worker = Worker(files, merge, output_path)
        self.worker.moveToThread(self.thread)

        self.worker.finished.connect(self.on_processing_finished)
        self.thread.started.connect(self.worker.run)
        self.worker.progress.connect(self.ui.progressBar.setValue)
        self.worker.status_update.connect(self.ui.statusLabel.setText)
        
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        
        self.thread.start()

    def on_processing_finished(self, success, result_path, error_msg):
        self.ui.startButton.setEnabled(True)
        if not success:
            self.ui.statusLabel.setText(f"Error: {error_msg}")
            self.cleanup_temp_file()
            return
            
        # Check if we were in merge mode
        if self.ui.mergeCheckbox.isChecked() and self.ui.fileListWidget.count() > 1:
            self.ui.statusLabel.setText("Merge complete. Opening preview...")
            try:
                preview = PreviewWindow(result_path, self)
                preview.closing.connect(self.cleanup_temp_file)
                if preview.exec_() == QDialog.Accepted:
                    final_path, _ = QFileDialog.getSaveFileName(self, "Save Merged PDF", "", "PDF Files (*.pdf)")
                    if final_path:
                        shutil.copy(self.temp_pdf_file, final_path)
                        self.ui.statusLabel.setText(f"File saved successfully to {final_path}")
                else:
                    self.ui.statusLabel.setText("Save cancelled by user.")
            except Exception as e:
                self.ui.statusLabel.setText(f"Error opening preview: {e}")
            finally:
                self.cleanup_temp_file()
        else:
            # This was an individual conversion
            self.ui.statusLabel.setText(f"Individual conversions complete! Files saved in {self.ui.outputFolderEdit.text()}")

    def cleanup_temp_file(self):
        if self.temp_pdf_file and os.path.exists(self.temp_pdf_file):
            try:
                os.remove(self.temp_pdf_file)
                logging.info(f"Successfully cleaned up temporary file: {self.temp_pdf_file}")
                self.temp_pdf_file = None
            except Exception as e:
                logging.error(f"Failed to delete temporary file on cleanup: {e}")

# --- Application Entry Point ---
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = AppWindow()
    window.show()
    sys.exit(app.exec_())
