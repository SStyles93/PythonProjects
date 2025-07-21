#!/usr/bin/env python3
"""
Advanced CSV to Excel Converter
A desktop application with PyQt GUI for converting CSV files to Excel format.
Features:
- Recursive folder selection for CSV files
- Similar file detection and merging
- Append to existing Excel files with duplicate detection
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
                             QSpinBox, QComboBox, QListWidget, QRadioButton,
                             QButtonGroup, QScrollArea)
from PyQt5.QtCore import QThread, pyqtSignal, Qt
from PyQt5.QtGui import QFont, QIcon


class ConversionWorker(QThread):
    """Worker thread for CSV to Excel conversion to prevent GUI freezing"""
    
    progress = pyqtSignal(int)
    status = pyqtSignal(str)
    finished = pyqtSignal(bool, str)
    
    def __init__(self, csv_files, output_path, combine_sheets, sheet_names, 
                 detect_similar, append_mode, existing_file_path=None):
        super().__init__()
        self.csv_files = csv_files
        self.output_path = output_path
        self.combine_sheets = combine_sheets
        self.sheet_names = sheet_names
        self.detect_similar = detect_similar
        self.append_mode = append_mode
        self.existing_file_path = existing_file_path
    
    def run(self):
        try:
            if self.append_mode and self.existing_file_path:
                self.append_to_existing_file()
            elif self.detect_similar:
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
    
    def append_to_existing_file(self):
        """Append CSV data to existing Excel file with duplicate detection"""
        self.status.emit("Loading existing Excel file...")
        
        # Read existing Excel file
        existing_sheets = pd.read_excel(self.existing_file_path, sheet_name=None)
        
        # Create a copy to work with
        updated_sheets = {}
        total_files = len(self.csv_files)
        duplicates_found = 0
        new_rows_added = 0
        
        for i, csv_file in enumerate(self.csv_files):
            self.status.emit(f"Processing {os.path.basename(csv_file)} for append...")
            
            # Read CSV file
            new_df = pd.read_csv(csv_file)
            
            # Determine target sheet name
            if self.detect_similar:
                # Use base name for similar file detection
                base_name = self.extract_base_name(Path(csv_file).stem)
                target_sheet = self.sanitize_sheet_name(base_name)
            else:
                # Use file name or custom sheet name
                if i < len(self.sheet_names) and self.sheet_names[i].strip():
                    target_sheet = self.sanitize_sheet_name(self.sheet_names[i].strip())
                else:
                    target_sheet = self.sanitize_sheet_name(Path(csv_file).stem)
            
            # Check if target sheet exists in existing file
            if target_sheet in existing_sheets:
                existing_df = existing_sheets[target_sheet].copy()
                
                # Detect and remove duplicates
                merged_df, file_duplicates, file_new_rows = self.merge_with_duplicate_detection(
                    existing_df, new_df, csv_file
                )
                
                duplicates_found += file_duplicates
                new_rows_added += file_new_rows
                updated_sheets[target_sheet] = merged_df
            else:
                # New sheet, add source file column and use as-is
                new_df["Source_File"] = os.path.basename(csv_file)
                # Move Source_File column to the beginning
                cols = new_df.columns.tolist()
                cols = ["Source_File"] + [col for col in cols if col != "Source_File"]
                new_df = new_df[cols]
                
                updated_sheets[target_sheet] = new_df
                new_rows_added += len(new_df)
            
            # Update progress
            progress_value = int((i + 1) / total_files * 100)
            self.progress.emit(progress_value)
        
        # Add any existing sheets that weren\'t updated
        for sheet_name, sheet_df in existing_sheets.items():
            if sheet_name not in updated_sheets:
                updated_sheets[sheet_name] = sheet_df
        
        # Write updated file
        self.status.emit("Saving updated Excel file...")
        with pd.ExcelWriter(self.output_path, engine='openpyxl') as writer:
            for sheet_name, df in updated_sheets.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        message = f"Successfully appended data to Excel file. Added {new_rows_added} new rows, skipped {duplicates_found} duplicates."
        self.finished.emit(True, message)
    
    def merge_with_duplicate_detection(self, existing_df, new_df, source_file):
        """Merge new data with existing data, detecting and avoiding duplicates"""
        # Add source file column to new data
        new_df = new_df.copy()
        new_df["Source_File"] = os.path.basename(source_file)
        
        # Ensure existing_df has Source_File column
        if "Source_File" not in existing_df.columns:
            existing_df["Source_File"] = "Unknown"
        
        # Get the original column order from existing_df
        original_columns = existing_df.columns.tolist()
        
        # Identify all unique columns from both dataframes
        all_columns = list(set(original_columns) | set(new_df.columns))
        
        # Reindex new_df to match all_columns, filling missing with empty string
        new_df = new_df.reindex(columns=all_columns, fill_value="")
        
        # Reindex existing_df to match all_columns, filling missing with empty string
        existing_df = existing_df.reindex(columns=all_columns, fill_value="")
        
        # Reorder columns of both dataframes to match the original_columns order
        # Any new columns will be appended at the end in their original order from new_df
        final_column_order = original_columns + [col for col in new_df.columns if col not in original_columns]
        
        existing_df = existing_df[final_column_order]
        new_df = new_df[final_column_order]
        
        # Detect duplicates based on all columns except Source_File
        data_columns = [col for col in final_column_order if col != "Source_File"]
        
        # Find rows in new_df that already exist in existing_df
        # Convert to tuples for efficient comparison of rows
        existing_rows = set(tuple(row) for row in existing_df[data_columns].values)
        
        duplicates_mask = new_df[data_columns].apply(
            lambda row: tuple(row) in existing_rows, axis=1
        )
        
        duplicates_count = duplicates_mask.sum()
        new_rows = new_df[~duplicates_mask]
        new_rows_count = len(new_rows)
        
        # Combine existing data with new non-duplicate data
        merged_df = pd.concat([existing_df, new_rows], ignore_index=True, sort=False)
        
        return merged_df, duplicates_count, new_rows_count
    
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
        #Uses raw strings (r'') with for regex patterns.
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
            r'\d+$'               # Any trailing numbers without separator
        ]
    
        base_name = file_name
        for pattern in patterns:
            base_name = re.sub(pattern, '', base_name)
    
        # Use rstrip for cleaner removal of trailing characters
        return base_name.rstrip('_-\t ')
    
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
                        self.status.emit(f"Merging {len(file_list)} similar files for \'{base_name}\'...")
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
                    self.status.emit(f"Merging {len(file_list)} similar files for \'{base_name}\'...")
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
        # Excel sheet names cannot contain: \\ / ? * [ ] :
        invalid_chars = ['\\\\\', \'/\', \'?\', \'*\', \'[\', \']\', \':']
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
        self.existing_file_path = ""
        self.worker = None
        
        self.init_ui()
    
    def init_ui(self):
        """Initialize the user interface"""
        self.setWindowTitle("Advanced CSV to Excel Converter")
        self.setGeometry(100, 100, 900, 800)
        
        # Create central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # Title
        title_label = QLabel("Advanced CSV to Excel Converter")
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)
        
        # File selection group
        file_group = QGroupBox("File Selection")
        file_layout = QVBoxLayout(file_group)
        
        # Selection mode radio buttons
        selection_mode_layout = QHBoxLayout()
        self.selection_mode_group = QButtonGroup()
        
        self.files_radio = QRadioButton("Select individual CSV files")
        self.folder_radio = QRadioButton("Select folder (recursive search)")
        self.files_radio.setChecked(True)
        
        self.selection_mode_group.addButton(self.files_radio, 0)
        self.selection_mode_group.addButton(self.folder_radio, 1)
        
        self.files_radio.toggled.connect(self.on_selection_mode_changed)
        
        selection_mode_layout.addWidget(self.files_radio)
        selection_mode_layout.addWidget(self.folder_radio)
        file_layout.addLayout(selection_mode_layout)
        
        # CSV files/folder selection
        csv_layout = QHBoxLayout()
        self.csv_label = QLabel("No CSV files selected")
        self.csv_button = QPushButton("Select CSV Files")
        self.csv_button.clicked.connect(self.select_csv_files)
        csv_layout.addWidget(self.csv_label)
        csv_layout.addWidget(self.csv_button)
        file_layout.addLayout(csv_layout)
        
        # File list display
        self.file_list = QListWidget()
        self.file_list.setMaximumHeight(120)
        file_layout.addWidget(QLabel("Selected files:"))
        file_layout.addWidget(self.file_list)
        
        # Output mode selection
        output_mode_layout = QHBoxLayout()
        self.output_mode_group = QButtonGroup()
        
        self.new_file_radio = QRadioButton("Create new Excel file")
        self.append_file_radio = QRadioButton("Append to existing Excel file")
        self.new_file_radio.setChecked(True)
        
        self.output_mode_group.addButton(self.new_file_radio, 0)
        self.output_mode_group.addButton(self.append_file_radio, 1)
        
        self.new_file_radio.toggled.connect(self.on_output_mode_changed)
        
        output_mode_layout.addWidget(self.new_file_radio)
        output_mode_layout.addWidget(self.append_file_radio)
        file_layout.addLayout(output_mode_layout)
        
        # Existing file selection (for append mode)
        existing_file_layout = QHBoxLayout()
        self.existing_file_label = QLabel("No existing file selected")
        self.existing_file_button = QPushButton("Select Existing Excel File")
        self.existing_file_button.clicked.connect(self.select_existing_file)
        self.existing_file_button.setVisible(False)
        self.existing_file_label.setVisible(False)
        existing_file_layout.addWidget(self.existing_file_label)
        existing_file_layout.addWidget(self.existing_file_button)
        file_layout.addLayout(existing_file_layout)
        
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
        self.log("Advanced application started. Select CSV files or folder to begin.")
    
    def on_selection_mode_changed(self):
        """Handle selection mode change between files and folder"""
        if self.files_radio.isChecked():
            self.csv_button.setText("Select CSV Files")
        else:
            self.csv_button.setText("Select Folder")
        
        # Clear current selection
        self.csv_files = []
        self.csv_label.setText("No CSV files selected")
        self.file_list.clear()
        self.update_ui_state()
    
    def on_output_mode_changed(self):
        """Handle output mode change between new file and append"""
        is_append = self.append_file_radio.isChecked()
        
        self.existing_file_button.setVisible(is_append)
        self.existing_file_label.setVisible(is_append)
        
        if is_append:
            self.output_button.setText("Select Output Location (Copy)")
            self.combine_checkbox.setChecked(True)  # Force combine mode for append
            self.combine_checkbox.setEnabled(False)
        else:
            self.output_button.setText("Select Output Location")
            self.combine_checkbox.setEnabled(True)
        
        # Reset selections
        self.output_path = ""
        self.existing_file_path = ""
        self.output_label.setText("No output location selected")
        self.existing_file_label.setText("No existing file selected")
        
        self.update_ui_state()
    
    def select_csv_files(self):
        """Open file dialog to select CSV files or folder"""
        if self.files_radio.isChecked():
            # Select individual files
            files, _ = QFileDialog.getOpenFileNames(
                self,
                "Select CSV Files",
                "",
                "CSV Files (*.csv);;All Files (*)"
            )
            
            if files:
                self.csv_files = files
                self.csv_label.setText(f"{len(files)} CSV file(s) selected")
                self.update_file_list()
                self.log(f"Selected {len(files)} CSV files")
        else:
            # Select folder and search recursively
            folder = QFileDialog.getExistingDirectory(
                self,
                "Select Folder to Search for CSV Files"
            )
            
            if folder:
                self.csv_files = self.find_csv_files_recursive(folder)
                if self.csv_files:
                    self.csv_label.setText(f"{len(self.csv_files)} CSV file(s) found in folder")
                    self.update_file_list()
                    self.log(f"Found {len(self.csv_files)} CSV files in folder: {folder}")
                else:
                    self.csv_label.setText("No CSV files found in selected folder")
                    self.file_list.clear()
                    self.log(f"No CSV files found in folder: {folder}")
        
        # Show similar file detection preview if enabled
        if self.detect_similar_checkbox.isChecked() and self.csv_files:
            self.preview_similar_files()
        
        self.update_ui_state()
    
    def find_csv_files_recursive(self, folder_path):
        """Recursively find all CSV files in the given folder"""
        csv_files = []
        folder = Path(folder_path)
        
        for csv_file in folder.rglob("*.csv"):
            csv_files.append(str(csv_file))
        
        return sorted(csv_files)
    
    def update_file_list(self):
        """Update the file list widget"""
        self.file_list.clear()
        for file_path in self.csv_files:
            self.file_list.addItem(os.path.basename(file_path))
    
    def select_existing_file(self):
        """Select existing Excel file for append mode"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Existing Excel File",
            "",
            "Excel Files (*.xlsx *.xls);;All Files (*)"
        )
        
        if file_path:
            self.existing_file_path = file_path
            self.existing_file_label.setText(f"Existing file: {os.path.basename(file_path)}")
            self.log(f"Selected existing Excel file: {file_path}")
            self.update_ui_state()
    
    def select_output_location(self):
        """Select output location based on conversion mode"""
        if self.append_file_radio.isChecked():
            # For append mode, select where to save the updated file
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Save Updated Excel File As",
                "updated_data.xlsx",
                "Excel Files (*.xlsx);;All Files (*)"
            )
            if file_path:
                self.output_path = file_path
                self.output_label.setText(f"Output: {os.path.basename(file_path)}")
                self.log(f"Output file: {file_path}")
        elif self.combine_checkbox.isChecked():
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
                self.log(f"  Group \'{base_name}\': {', '.join(files)}")
        else:
            self.log("No similar files detected for merging.")
    
    def extract_base_name_preview(self, file_name):
        """Extract base name for preview (same logic as worker)"""
        # Uses raw strings (r'') for regex patterns.
        patterns = [
            r'_\d{8}$',
            r'_\d{4}-\d{2}-\d{2}$',
            r'_\d{4}_\d{2}_\d{2}$',
            r'_\d{6}$',
            r'_\d{4}$',
            r'-\d{8}$', 
            r'-\d{4}-\d{2}-\d{2}$',
            r'-\d{4}_\d{2}_\d{2}$', 
            r'-\d{6}$', 
            r'-\d{4}$',
            r'\d{8}$', 
            r'\d{4}-\d{2}-\d{2}$', 
            r'\d{4}_\d{2}_\d{2}$',
            r'\d{6}$', 
            r'_\d+$', 
            r'-\d+$',
            r'\d+$'
        ]
        
        base_name = file_name
        for pattern in patterns:
            base_name = re.sub(pattern, '', base_name)
    
        # Use rstrip for cleaner removal of trailing characters
        return base_name.rstrip('_-\t ')
    
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
        
        # Reset output selection when mode changes (except in append mode)
        if not self.append_file_radio.isChecked():
            self.output_path = ""
            self.output_label.setText("No output location selected")
        
        self.update_ui_state()
    
    def update_ui_state(self):
        """Update UI elements based on current state"""
        has_files = len(self.csv_files) > 0
        has_output = self.output_path != ""
        has_existing = self.existing_file_path != "" if self.append_file_radio.isChecked() else True
        
        self.convert_button.setEnabled(has_files and has_output and has_existing)
        
        if not has_files:
            self.status_label.setText("Please select CSV files or folder")
        elif self.append_file_radio.isChecked() and not has_existing:
            self.status_label.setText("Please select existing Excel file")
        elif not has_output:
            self.status_label.setText("Please select output location")
        else:
            self.status_label.setText("Ready to convert")
    
    def start_conversion(self):
        """Start the conversion process"""
        if not self.csv_files or not self.output_path:
            QMessageBox.warning(self, "Warning", "Please select CSV files and output location.")
            return
        
        if self.append_file_radio.isChecked() and not self.existing_file_path:
            QMessageBox.warning(self, "Warning", "Please select an existing Excel file for append mode.")
            return
        
        # Disable UI during conversion
        self.convert_button.setEnabled(False)
        self.csv_button.setEnabled(False)
        self.output_button.setEnabled(False)
        self.existing_file_button.setEnabled(False)
        
        # Reset progress
        self.progress_bar.setValue(0)
        
        # Get sheet names if combining
        sheet_names = []
        if self.combine_checkbox.isChecked():
            sheet_names_text = self.sheet_names_text.toPlainText().strip()
            if sheet_names_text:
                sheet_names = [name.strip() for name in sheet_names_text.split('\\n')]
        
        # Start worker thread
        self.worker = ConversionWorker(
            self.csv_files,
            self.output_path,
            self.combine_checkbox.isChecked(),
            sheet_names,
            self.detect_similar_checkbox.isChecked(),
            self.append_file_radio.isChecked(),
            self.existing_file_path if self.append_file_radio.isChecked() else None
        )
        
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.status.connect(self.status_label.setText)
        self.worker.finished.connect(self.on_conversion_finished)
        
        self.worker.start()
        
        mode_desc = "append mode" if self.append_file_radio.isChecked() else "conversion"
        similar_desc = " with similar file detection" if self.detect_similar_checkbox.isChecked() else ""
        self.log(f"Starting {mode_desc}{similar_desc}...")
    
    def on_conversion_finished(self, success, message):
        """Handle conversion completion"""
        # Re-enable UI
        self.convert_button.setEnabled(True)
        self.csv_button.setEnabled(True)
        self.output_button.setEnabled(True)
        self.existing_file_button.setEnabled(True)
        
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
    app.setApplicationName("Advanced CSV to Excel Converter")
    app.setApplicationVersion("3.0")
    
    # Create and show main window
    window = CSVToExcelConverter()
    window.show()
    
    # Start event loop
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
