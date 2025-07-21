#!/usr/bin/env python3
"""
CSV to Excel Converter
A desktop application with PyQt GUI for converting CSV files to Excel format.
"""

import sys
import os
from pathlib import Path
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, 
                             QWidget, QPushButton, QLabel, QFileDialog, QTextEdit, 
                             QProgressBar, QMessageBox, QGroupBox, QCheckBox,
                             QSpinBox, QComboBox)
from PyQt5.QtCore import QThread, pyqtSignal, Qt
from PyQt5.QtGui import QFont, QIcon


class ConversionWorker(QThread):
    """Worker thread for CSV to Excel conversion to prevent GUI freezing"""
    
    progress = pyqtSignal(int)
    status = pyqtSignal(str)
    finished = pyqtSignal(bool, str)
    
    def __init__(self, csv_files, output_path, combine_sheets, sheet_names):
        super().__init__()
        self.csv_files = csv_files
        self.output_path = output_path
        self.combine_sheets = combine_sheets
        self.sheet_names = sheet_names
    
    def run(self):
        try:
            if self.combine_sheets:
                # Combine all CSV files into one Excel file with multiple sheets
                self.status.emit("Creating combined Excel file...")
                with pd.ExcelWriter(self.output_path, engine='openpyxl') as writer:
                    for i, csv_file in enumerate(self.csv_files):
                        self.status.emit(f"Processing {os.path.basename(csv_file)}...")
                        
                        # Read CSV file
                        df = pd.read_csv(csv_file)
                        
                        # Use custom sheet name or file name
                        if i < len(self.sheet_names) and self.sheet_names[i].strip():
                            sheet_name = self.sheet_names[i].strip()
                        else:
                            sheet_name = Path(csv_file).stem
                        
                        # Ensure sheet name is valid (Excel has restrictions)
                        sheet_name = self.sanitize_sheet_name(sheet_name)
                        
                        # Write to Excel sheet
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                        
                        # Update progress
                        progress_value = int((i + 1) / len(self.csv_files) * 100)
                        self.progress.emit(progress_value)
                
                self.finished.emit(True, f"Successfully created combined Excel file: {self.output_path}")
            
            else:
                # Convert each CSV to separate Excel file
                for i, csv_file in enumerate(self.csv_files):
                    self.status.emit(f"Converting {os.path.basename(csv_file)}...")
                    
                    # Read CSV file
                    df = pd.read_csv(csv_file)
                    
                    # Create output path for individual file
                    csv_path = Path(csv_file)
                    output_file = Path(self.output_path) / f"{csv_path.stem}.xlsx"
                    
                    # Write to Excel
                    df.to_excel(output_file, index=False)
                    
                    # Update progress
                    progress_value = int((i + 1) / len(self.csv_files) * 100)
                    self.progress.emit(progress_value)
                
                self.finished.emit(True, f"Successfully converted {len(self.csv_files)} files to Excel format")
        
        except Exception as e:
            self.finished.emit(False, f"Error during conversion: {str(e)}")
    
    def sanitize_sheet_name(self, name):
        """Sanitize sheet name to comply with Excel restrictions"""
        # Excel sheet names cannot contain: \ / ? * [ ] :
        invalid_chars = ['\\', '/', '?', '*', '[', ']', ':']
        for char in invalid_chars:
            name = name.replace(char, '_')
        
        # Excel sheet names cannot be longer than 31 characters
        if len(name) > 31:
            name = name[:31]
        
        return name


class CSVToExcelConverter(QMainWindow):
    def __init__(self):
        super().__init__()
        self.csv_files = []
        self.output_path = ""
        self.worker = None
        
        self.init_ui()
    
    def init_ui(self):
        """Initialize the user interface"""
        self.setWindowTitle("CSV to Excel Converter")
        self.setGeometry(100, 100, 800, 600)
        
        # Create central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # Title
        title_label = QLabel("CSV to Excel Converter")
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)
        
        # File selection group
        file_group = QGroupBox("File Selection")
        file_layout = QVBoxLayout(file_group)
        
        # CSV files selection
        csv_layout = QHBoxLayout()
        self.csv_label = QLabel("No CSV files selected")
        self.csv_button = QPushButton("Select CSV Files")
        self.csv_button.clicked.connect(self.select_csv_files)
        csv_layout.addWidget(self.csv_label)
        csv_layout.addWidget(self.csv_button)
        file_layout.addLayout(csv_layout)
        
        # Output selection
        output_layout = QHBoxLayout()
        self.output_label = QLabel("No output location selected")
        self.output_button = QPushButton("Select Output Location")
        self.output_button.clicked.connect(self.select_output_location)
        output_layout.addWidget(self.output_label)
        output_layout.addWidget(self.output_button)
        file_layout.addLayout(output_layout)
        
        main_layout.addWidget(file_group)
        
        # Conversion options group
        options_group = QGroupBox("Conversion Options")
        options_layout = QVBoxLayout(options_group)
        
        # Combine sheets option
        self.combine_checkbox = QCheckBox("Combine all CSV files into one Excel file with multiple sheets")
        self.combine_checkbox.setChecked(True)
        self.combine_checkbox.stateChanged.connect(self.on_combine_changed)
        options_layout.addWidget(self.combine_checkbox)
        
        # Sheet names (only visible when combining)
        self.sheet_names_label = QLabel("Sheet names (optional, leave blank to use file names):")
        self.sheet_names_text = QTextEdit()
        self.sheet_names_text.setMaximumHeight(100)
        self.sheet_names_text.setPlaceholderText("Enter sheet names, one per line...")
        options_layout.addWidget(self.sheet_names_label)
        options_layout.addWidget(self.sheet_names_text)
        
        main_layout.addWidget(options_group)
        
        # Progress section
        progress_group = QGroupBox("Progress")
        progress_layout = QVBoxLayout(progress_group)
        
        self.progress_bar = QProgressBar()
        self.status_label = QLabel("Ready to convert")
        
        progress_layout.addWidget(self.progress_bar)
        progress_layout.addWidget(self.status_label)
        
        main_layout.addWidget(progress_group)
        
        # Convert button
        self.convert_button = QPushButton("Convert to Excel")
        self.convert_button.clicked.connect(self.start_conversion)
        self.convert_button.setMinimumHeight(40)
        convert_font = QFont()
        convert_font.setPointSize(12)
        convert_font.setBold(True)
        self.convert_button.setFont(convert_font)
        main_layout.addWidget(self.convert_button)
        
        # Log area
        log_group = QGroupBox("Log")
        log_layout = QVBoxLayout(log_group)
        
        self.log_text = QTextEdit()
        self.log_text.setMaximumHeight(150)
        self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text)
        
        main_layout.addWidget(log_group)
        
        # Initial state
        self.update_ui_state()
        self.log("Application started. Select CSV files to begin.")
    
    def select_csv_files(self):
        """Open file dialog to select CSV files"""
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Select CSV Files",
            "",
            "CSV Files (*.csv);;All Files (*)"
        )
        
        if files:
            self.csv_files = files
            self.csv_label.setText(f"{len(files)} CSV file(s) selected")
            self.log(f"Selected {len(files)} CSV files:")
            for file in files:
                self.log(f"  - {os.path.basename(file)}")
            self.update_ui_state()
    
    def select_output_location(self):
        """Select output location based on conversion mode"""
        if self.combine_checkbox.isChecked():
            # Single file output
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Save Combined Excel File As",
                "combined_data.xlsx",
                "Excel Files (*.xlsx);;All Files (*)"
            )
            if file_path:
                self.output_path = file_path
                self.output_label.setText(f"Output: {os.path.basename(file_path)}")
                self.log(f"Output file: {file_path}")
        else:
            # Directory output
            directory = QFileDialog.getExistingDirectory(
                self,
                "Select Output Directory"
            )
            if directory:
                self.output_path = directory
                self.output_label.setText(f"Output directory: {os.path.basename(directory)}")
                self.log(f"Output directory: {directory}")
        
        self.update_ui_state()
    
    def on_combine_changed(self):
        """Handle combine checkbox state change"""
        is_combine = self.combine_checkbox.isChecked()
        self.sheet_names_label.setVisible(is_combine)
        self.sheet_names_text.setVisible(is_combine)
        
        # Reset output selection when mode changes
        self.output_path = ""
        self.output_label.setText("No output location selected")
        
        self.update_ui_state()
    
    def update_ui_state(self):
        """Update UI elements based on current state"""
        has_files = len(self.csv_files) > 0
        has_output = self.output_path != ""
        
        self.convert_button.setEnabled(has_files and has_output)
        
        if not has_files:
            self.status_label.setText("Please select CSV files")
        elif not has_output:
            self.status_label.setText("Please select output location")
        else:
            self.status_label.setText("Ready to convert")
    
    def start_conversion(self):
        """Start the conversion process"""
        if not self.csv_files or not self.output_path:
            QMessageBox.warning(self, "Warning", "Please select CSV files and output location.")
            return
        
        # Disable UI during conversion
        self.convert_button.setEnabled(False)
        self.csv_button.setEnabled(False)
        self.output_button.setEnabled(False)
        
        # Reset progress
        self.progress_bar.setValue(0)
        
        # Get sheet names if combining
        sheet_names = []
        if self.combine_checkbox.isChecked():
            sheet_names_text = self.sheet_names_text.toPlainText().strip()
            if sheet_names_text:
                sheet_names = [name.strip() for name in sheet_names_text.split('\n')]
        
        # Start worker thread
        self.worker = ConversionWorker(
            self.csv_files,
            self.output_path,
            self.combine_checkbox.isChecked(),
            sheet_names
        )
        
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.status.connect(self.status_label.setText)
        self.worker.finished.connect(self.on_conversion_finished)
        
        self.worker.start()
        self.log("Starting conversion...")
    
    def on_conversion_finished(self, success, message):
        """Handle conversion completion"""
        # Re-enable UI
        self.convert_button.setEnabled(True)
        self.csv_button.setEnabled(True)
        self.output_button.setEnabled(True)
        
        if success:
            self.progress_bar.setValue(100)
            self.status_label.setText("Conversion completed successfully!")
            self.log("✓ " + message)
            QMessageBox.information(self, "Success", message)
        else:
            self.status_label.setText("Conversion failed!")
            self.log("✗ " + message)
            QMessageBox.critical(self, "Error", message)
        
        self.worker = None
    
    def log(self, message):
        """Add message to log area"""
        self.log_text.append(message)
        # Auto-scroll to bottom
        scrollbar = self.log_text.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())


def main():
    """Main application entry point"""
    app = QApplication(sys.argv)
    
    # Set application properties
    app.setApplicationName("CSV to Excel Converter")
    app.setApplicationVersion("1.0")
    
    # Create and show main window
    window = CSVToExcelConverter()
    window.show()
    
    # Start event loop
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()

