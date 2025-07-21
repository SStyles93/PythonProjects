#!/usr/bin/env python3
"""
Enhanced CSV to Excel Converter
A desktop application with PyQt GUI for converting CSV files to Excel format.
Includes feature to detect and merge CSV files with similar names.
"""

import sys
import os
import re
from pathlib import Path
from collections import defaultdict
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
    
    def __init__(self, csv_files, output_path, combine_sheets, sheet_names, detect_similar):
        super().__init__()
        self.csv_files = csv_files
        self.output_path = output_path
        self.combine_sheets = combine_sheets
        self.sheet_names = sheet_names
        self.detect_similar = detect_similar
    
    def run(self):
        try:
            if self.detect_similar:
                # Group similar files and process them
                grouped_files = self.group_similar_files(self.csv_files)
                self.process_grouped_files(grouped_files)
            elif self.combine_sheets:
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
    
    def group_similar_files(self, csv_files):
        """Group CSV files with similar names (ignoring dates and numbers)"""
        groups = defaultdict(list)
        
        for file_path in csv_files:
            file_name = Path(file_path).stem
            # Remove common date patterns and numbers to find base name
            base_name = self.extract_base_name(file_name)
            groups[base_name].append(file_path)
        
        return groups
    
    def extract_base_name(self, file_name):
        """Extract base name by removing date patterns and trailing numbers"""
        # Remove common date patterns: YYYYMMDD, YYYY-MM-DD, YYYY_MM_DD, etc.
        patterns = [
            r'_\d{8}$',           # _20250401
            r'_\d{4}-\d{2}-\d{2}$', # _2025-04-01
            r'_\d{4}_\d{2}_\d{2}$', # _2025_04_01
            r'_\d{6}$',           # _202504
            r'_\d{4}$',           # _2025
            r'-\d{8}$',           # -20250401
            r'-\d{4}-\d{2}-\d{2}$', # -2025-04-01
            r'-\d{4}_\d{2}_\d{2}$', # -2025_04_01
            r'-\d{6}$',           # -202504
            r'-\d{4}$',           # -2025
            r'\d{8}$',            # 20250401 (at end)
            r'\d{4}-\d{2}-\d{2}$', # 2025-04-01 (at end)
            r'\d{4}_\d{2}_\d{2}$', # 2025_04_01 (at end)
            r'\d{6}$',            # 202504 (at end)
            r'_\d+$',             # Any trailing underscore + numbers
            r'-\d+$',             # Any trailing dash + numbers
        ]
        
        base_name = file_name
        for pattern in patterns:
            base_name = re.sub(pattern, '', base_name)
        
        return base_name.strip('_-')
    
    def process_grouped_files(self, grouped_files):
        """Process grouped files and merge similar ones"""
        if self.combine_sheets:
            # Create one Excel file with sheets for each group
            self.status.emit("Creating Excel file with merged similar files...")
            with pd.ExcelWriter(self.output_path, engine='openpyxl') as writer:
                sheet_index = 0
                total_groups = len(grouped_files)
                
                for base_name, file_list in grouped_files.items():
                    if len(file_list) > 1:
                        # Merge similar files
                        self.status.emit(f"Merging {len(file_list)} similar files for '{base_name}'...")
                        merged_df = self.merge_csv_files(file_list)
                        sheet_name = self.sanitize_sheet_name(base_name)
                        merged_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    else:
                        # Single file, process normally
                        csv_file = file_list[0]
                        self.status.emit(f"Processing {os.path.basename(csv_file)}...")
                        df = pd.read_csv(csv_file)
                        
                        # Use custom sheet name if available
                        if sheet_index < len(self.sheet_names) and self.sheet_names[sheet_index].strip():
                            sheet_name = self.sheet_names[sheet_index].strip()
                        else:
                            sheet_name = Path(csv_file).stem
                        
                        sheet_name = self.sanitize_sheet_name(sheet_name)
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    sheet_index += 1
                    progress_value = int(sheet_index / total_groups * 100)
                    self.progress.emit(progress_value)
            
            merged_count = sum(1 for files in grouped_files.values() if len(files) > 1)
            self.finished.emit(True, f"Successfully created Excel file with {merged_count} merged file groups: {self.output_path}")
        
        else:
            # Create separate Excel files, but merge similar ones
            output_dir = Path(self.output_path)
            total_groups = len(grouped_files)
            processed_groups = 0
            
            for base_name, file_list in grouped_files.items():
                if len(file_list) > 1:
                    # Merge similar files into one Excel file
                    self.status.emit(f"Merging {len(file_list)} similar files for '{base_name}'...")
                    merged_df = self.merge_csv_files(file_list)
                    output_file = output_dir / f"{base_name}_merged.xlsx"
                    merged_df.to_excel(output_file, index=False)
                else:
                    # Single file, convert normally
                    csv_file = file_list[0]
                    self.status.emit(f"Converting {os.path.basename(csv_file)}...")
                    df = pd.read_csv(csv_file)
                    csv_path = Path(csv_file)
                    output_file = output_dir / f"{csv_path.stem}.xlsx"
                    df.to_excel(output_file, index=False)
                
                processed_groups += 1
                progress_value = int(processed_groups / total_groups * 100)
                self.progress.emit(progress_value)
            
            merged_count = sum(1 for files in grouped_files.values() if len(files) > 1)
            self.finished.emit(True, f"Successfully processed {total_groups} file groups with {merged_count} merged groups")
    
    def merge_csv_files(self, file_list):
        """Merge multiple CSV files with similar names into one DataFrame"""
        dataframes = []
        
        for csv_file in file_list:
            df = pd.read_csv(csv_file)
            # Add a column to identify the source file
            df['Source_File'] = os.path.basename(csv_file)
            dataframes.append(df)
        
        # Concatenate all dataframes
        merged_df = pd.concat(dataframes, ignore_index=True, sort=False)
        
        # Move Source_File column to the beginning
        cols = merged_df.columns.tolist()
        cols = ['Source_File'] + [col for col in cols if col != 'Source_File']
        merged_df = merged_df[cols]
        
        return merged_df
    
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
        self.setWindowTitle("Enhanced CSV to Excel Converter")
        self.setGeometry(100, 100, 850, 700)
        
        # Create central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # Title
        title_label = QLabel("Enhanced CSV to Excel Converter")
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
        
        # Detect similar files option
        self.detect_similar_checkbox = QCheckBox("Detect and merge files with similar names (e.g., Downloads_20250401 & Downloads_20250501)")
        self.detect_similar_checkbox.setChecked(False)
        self.detect_similar_checkbox.stateChanged.connect(self.on_detect_similar_changed)
        options_layout.addWidget(self.detect_similar_checkbox)
        
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
        self.log("Enhanced application started. Select CSV files to begin.")
    
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
            
            # Show similar file detection preview if enabled
            if self.detect_similar_checkbox.isChecked():
                self.preview_similar_files()
            
            self.update_ui_state()
    
    def preview_similar_files(self):
        """Preview which files will be grouped together"""
        if not self.csv_files:
            return
        
        groups = defaultdict(list)
        for file_path in self.csv_files:
            file_name = Path(file_path).stem
            base_name = self.extract_base_name_preview(file_name)
            groups[base_name].append(os.path.basename(file_path))
        
        similar_groups = {k: v for k, v in groups.items() if len(v) > 1}
        
        if similar_groups:
            self.log("Similar files detected:")
            for base_name, files in similar_groups.items():
                self.log(f"  Group '{base_name}': {', '.join(files)}")
        else:
            self.log("No similar files detected for merging.")
    
    def extract_base_name_preview(self, file_name):
        """Extract base name for preview (same logic as worker)"""
        patterns = [
            r'_\d{8}$', r'_\d{4}-\d{2}-\d{2}$', r'_\d{4}_\d{2}_\d{2}$',
            r'_\d{6}$', r'_\d{4}$', r'-\d{8}$', r'-\d{4}-\d{2}-\d{2}$',
            r'-\d{4}_\d{2}_\d{2}$', r'-\d{6}$', r'-\d{4}$',
            r'\d{8}$', r'\d{4}-\d{2}-\d{2}$', r'\d{4}_\d{2}_\d{2}$',
            r'\d{6}$', r'_\d+$', r'-\d+$'
        ]
        
        base_name = file_name
        for pattern in patterns:
            base_name = re.sub(pattern, '', base_name)
        
        return base_name.strip('_-')
    
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
    
    def on_detect_similar_changed(self):
        """Handle detect similar files checkbox state change"""
        if self.detect_similar_checkbox.isChecked() and self.csv_files:
            self.preview_similar_files()
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
            sheet_names,
            self.detect_similar_checkbox.isChecked()
        )
        
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.status.connect(self.status_label.setText)
        self.worker.finished.connect(self.on_conversion_finished)
        
        self.worker.start()
        self.log("Starting conversion with similar file detection..." if self.detect_similar_checkbox.isChecked() else "Starting conversion...")
    
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
    app.setApplicationName("Enhanced CSV to Excel Converter")
    app.setApplicationVersion("2.0")
    
    # Create and show main window
    window = CSVToExcelConverter()
    window.show()
    
    # Start event loop
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
