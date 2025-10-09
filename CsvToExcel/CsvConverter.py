#!/usr/bin/env python3
"""
Advanced CSV to Excel Converter
A desktop application with PyQt GUI for converting CSV files to Excel format.
Features:
- Recursive folder selection for CSV files
- Similar file detection and merging
- Append to existing Excel files with duplicate detection
- Option to override and merge into existing files
- User-defined key columns for duplicate detection
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
                             QListWidget, QRadioButton, QButtonGroup, QLineEdit) # Added QLineEdit
from PyQt5.QtCore import QThread, pyqtSignal, Qt
from PyQt5.QtGui import QFont

class ConversionWorker(QThread):
    """Worker thread for CSV to Excel conversion to prevent GUI freezing"""

    progress = pyqtSignal(int)
    status = pyqtSignal(str)
    finished = pyqtSignal(bool, str)

    def __init__(self, csv_files, output_path, combine_sheets, sheet_names,
                 detect_similar, append_mode, override_mode, existing_file_path,
                 duplicate_keys):
        super().__init__()
        self.csv_files = csv_files
        self.output_path = output_path
        self.combine_sheets = combine_sheets
        self.sheet_names = sheet_names
        self.detect_similar = detect_similar
        self.append_mode = append_mode
        self.override_mode = override_mode
        self.existing_file_path = existing_file_path
        self.duplicate_keys = duplicate_keys

    def run(self):
        try:
            if self.append_mode and self.existing_file_path:
                self.append_to_existing_file()
            elif self.detect_similar:
                grouped_files = self.group_similar_files(self.csv_files)
                self.process_grouped_files(grouped_files)
            elif self.combine_sheets:
                self.status.emit("Creating combined Excel file...")
                output_file = self.get_unique_filename(self.output_path)
                with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                    for i, csv_file in enumerate(self.csv_files):
                        self.status.emit(f"Processing {os.path.basename(csv_file)}...")
                        df = pd.read_csv(csv_file)
                        if i < len(self.sheet_names) and self.sheet_names[i].strip():
                            sheet_name = self.sheet_names[i].strip()
                        else:
                            sheet_name = Path(csv_file).stem
                        sheet_name = self.sanitize_sheet_name(sheet_name)
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                        progress_value = int((i + 1) / len(self.csv_files) * 100)
                        self.progress.emit(progress_value)
                self.finished.emit(True, f"Successfully created combined Excel file: {output_file}")
            else:
                for i, csv_file in enumerate(self.csv_files):
                    self.status.emit(f"Converting {os.path.basename(csv_file)}...")
                    df = pd.read_csv(csv_file)
                    csv_path = Path(csv_file)
                    output_file = Path(self.output_path) / f"{csv_path.stem}.xlsx"
                    final_output_path = self.get_unique_filename(output_file)
                    df.to_excel(final_output_path, index=False)
                    progress_value = int((i + 1) / len(self.csv_files) * 100)
                    self.progress.emit(progress_value)
                self.finished.emit(True, f"Successfully converted {len(self.csv_files)} files to Excel format")
        except Exception as e:
            self.finished.emit(False, f"Error during conversion: {str(e)}")

    def get_unique_filename(self, file_path):
        if self.override_mode:
            return file_path

        p = Path(file_path)
        if not p.exists():
            return file_path

        parent = p.parent
        stem = p.stem
        suffix = p.suffix
        counter = 1
        while True:
            new_stem = f"{stem}_updated_{counter}"
            new_path = parent / f"{new_stem}{suffix}"
            if not new_path.exists():
                return new_path
            counter += 1

    def append_to_existing_file(self):
        self.status.emit("Loading existing Excel file...")
        
        source_file = self.existing_file_path
        if self.override_mode:
            self.output_path = self.existing_file_path

        if not os.path.exists(source_file):
            self.finished.emit(False, f"Existing file not found: {source_file}")
            return
            
        existing_sheets = pd.read_excel(source_file, sheet_name=None)
        # --- FIX: Create a copy of the sheets to modify. This is the key change. ---
        updated_sheets = existing_sheets.copy()
        
        total_files = len(self.csv_files)
        duplicates_found = 0
        new_rows_added = 0

        for i, csv_file in enumerate(self.csv_files):
            self.status.emit(f"Processing {os.path.basename(csv_file)} for append...")
            new_df = pd.read_csv(csv_file)
            
            csv_stem = Path(csv_file).stem
            
            # Determine the target sheet name based on whether similar file detection is on
            if self.detect_similar:
                target_sheet_name = self.sanitize_sheet_name(self.extract_base_name(csv_stem))
            else:
                target_sheet_name = self.sanitize_sheet_name(csv_stem)

            # Check if this sheet exists in our dictionary of sheets to be written
            if target_sheet_name in updated_sheets:
                # If it exists, merge the new data into it
                existing_df = updated_sheets[target_sheet_name]
                merged_df, file_duplicates, file_new_rows = self.merge_with_duplicate_detection(
                    existing_df, new_df, csv_file
                )
                duplicates_found += file_duplicates
                new_rows_added += file_new_rows
                updated_sheets[target_sheet_name] = merged_df
            else:
                # If it's a new sheet, just add it
                new_df["Source_File"] = os.path.basename(csv_file)
                cols = new_df.columns.tolist()
                cols = ["Source_File"] + [col for col in cols if col != "Source_File"]
                new_df = new_df[cols]
                updated_sheets[target_sheet_name] = new_df
                new_rows_added += len(new_df)

            progress_value = int((i + 1) / total_files * 100)
            self.progress.emit(progress_value)

        self.status.emit("Saving updated Excel file...")
        final_output_path = self.get_unique_filename(self.output_path)
        with pd.ExcelWriter(final_output_path, engine='openpyxl') as writer:
            # Write all sheets from the updated_sheets dictionary, preserving unmodified ones
            for sheet_name, df in updated_sheets.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)

        message = f"Successfully processed data. Added {new_rows_added} new rows, skipped {duplicates_found} duplicates. Saved to {os.path.basename(final_output_path)}"
        self.finished.emit(True, message)

    def merge_with_duplicate_detection(self, existing_df, new_df, source_file):
        new_df = new_df.copy()
        new_df["Source_File"] = os.path.basename(source_file)
        if "Source_File" not in existing_df.columns:
            existing_df["Source_File"] = "Unknown"
        
        original_columns = existing_df.columns.tolist()
        all_columns = list(set(original_columns) | set(new_df.columns))
        
        new_df = new_df.reindex(columns=all_columns, fill_value="")
        existing_df = existing_df.reindex(columns=all_columns, fill_value="")
        
        final_column_order = original_columns + [col for col in new_df.columns if col not in original_columns]
        existing_df = existing_df[final_column_order]
        new_df = new_df[final_column_order]
        
        if self.duplicate_keys:
            check_cols = [col for col in self.duplicate_keys if col in new_df.columns]
            if not check_cols:
                check_cols = [col for col in final_column_order if col != "Source_File"]
        else:
            check_cols = [col for col in final_column_order if col != "Source_File"]

        existing_rows = set(tuple(row) for row in existing_df[check_cols].to_numpy())
        
        duplicates_mask = new_df[check_cols].apply(
            lambda row: tuple(row) in existing_rows, axis=1
        )
        
        duplicates_count = duplicates_mask.sum()
        new_rows = new_df[~duplicates_mask]
        new_rows_count = len(new_rows)
        
        merged_df = pd.concat([existing_df, new_rows], ignore_index=True, sort=False)
        return merged_df, duplicates_count, new_rows_count

    def group_similar_files(self, csv_files):
        groups = defaultdict(list)
        for file_path in csv_files:
            file_name = Path(file_path).stem
            base_name = self.extract_base_name(file_name)
            groups[base_name].append(file_path)
        return groups

    def extract_base_name(self, file_name):
        patterns = [
            r'_\d{8}$', r'_\d{4}-\d{2}-\d{2}$', r'_\d{4}_\d{2}_\d{2}$', r'_\d{6}$',
            r'_\d{4}$', r'-\d{8}$', r'-\d{4}-\d{2}-\d{2}$', r'-\d{4}_\d{2}_\d{2}$',
            r'-\d{6}$', r'-\d{4}$', r'\d{8}$', r'\d{4}-\d{2}-\d{2}$',
            r'\d{4}_\d{2}_\d{2}$', r'\d{6}$', r'_\d+$', r'-\d+$', r'\d+$'
        ]
        base_name = file_name
        for pattern in patterns:
            base_name = re.sub(pattern, '', base_name)
        return base_name.rstrip('_-\t ')

    def process_grouped_files(self, grouped_files):
        if self.combine_sheets:
            self.status.emit("Creating Excel file with merged similar files...")
            output_file = self.get_unique_filename(self.output_path)
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                sheet_index = 0
                total_groups = len(grouped_files)
                for base_name, file_list in grouped_files.items():
                    if len(file_list) > 1:
                        self.status.emit(f"Merging {len(file_list)} similar files for '{base_name}'...")
                        merged_df = self.merge_csv_files(file_list)
                        sheet_name = self.sanitize_sheet_name(base_name)
                        merged_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    else:
                        csv_file = file_list[0]
                        self.status.emit(f"Processing {os.path.basename(csv_file)}...")
                        df = pd.read_csv(csv_file)
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
            self.finished.emit(True, f"Successfully created Excel file with {merged_count} merged groups: {os.path.basename(output_file)}")
        else:
            output_dir = Path(self.output_path)
            total_groups = len(grouped_files)
            processed_groups = 0
            for base_name, file_list in grouped_files.items():
                if len(file_list) > 1:
                    self.status.emit(f"Merging {len(file_list)} similar files for '{base_name}'...")
                    merged_df = self.merge_csv_files(file_list)
                    output_file = output_dir / f"{base_name}_merged.xlsx"
                    final_output_path = self.get_unique_filename(output_file)
                    merged_df.to_excel(final_output_path, index=False)
                else:
                    csv_file = file_list[0]
                    self.status.emit(f"Converting {os.path.basename(csv_file)}...")
                    df = pd.read_csv(csv_file)
                    csv_path = Path(csv_file)
                    output_file = output_dir / f"{csv_path.stem}.xlsx"
                    final_output_path = self.get_unique_filename(output_file)
                    df.to_excel(final_output_path, index=False)
                processed_groups += 1
                progress_value = int(processed_groups / total_groups * 100)
                self.progress.emit(progress_value)
            merged_count = sum(1 for files in grouped_files.values() if len(files) > 1)
            self.finished.emit(True, f"Successfully processed {total_groups} file groups with {merged_count} merged groups")

    def merge_csv_files(self, file_list):
        dataframes = []
        for csv_file in file_list:
            df = pd.read_csv(csv_file)
            df['Source_File'] = os.path.basename(csv_file)
            dataframes.append(df)
        merged_df = pd.concat(dataframes, ignore_index=True, sort=False)
        cols = merged_df.columns.tolist()
        cols = ['Source_File'] + [col for col in cols if col != 'Source_File']
        merged_df = merged_df[cols]
        return merged_df

    def sanitize_sheet_name(self, name):
        invalid_chars = ['\\', '/', '?', '*', '[', ']', ':']
        for char in invalid_chars:
            name = name.replace(char, '_')
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
        self.setWindowTitle("Advanced CSV to Excel Converter")
        self.setGeometry(100, 100, 900, 850) # Increased height for new field
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        title_label = QLabel("Advanced CSV to Excel Converter")
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)
        file_group = QGroupBox("File Selection")
        file_layout = QVBoxLayout(file_group)
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
        csv_layout = QHBoxLayout()
        self.csv_label = QLabel("No CSV files selected")
        self.csv_button = QPushButton("Select CSV Files")
        self.csv_button.clicked.connect(self.select_csv_files)
        csv_layout.addWidget(self.csv_label)
        csv_layout.addWidget(self.csv_button)
        file_layout.addLayout(csv_layout)
        self.file_list = QListWidget()
        self.file_list.setMaximumHeight(120)
        file_layout.addWidget(QLabel("Selected files:"))
        file_layout.addWidget(self.file_list)
        main_layout.addWidget(file_group)
        
        options_group = QGroupBox("Conversion Options")
        options_layout = QVBoxLayout(options_group)
        
        self.new_file_radio = QRadioButton("Create new Excel file(s)")
        self.append_file_radio = QRadioButton("Append to existing Excel file")
        self.new_file_radio.setChecked(True)
        self.output_mode_group = QButtonGroup()
        self.output_mode_group.addButton(self.new_file_radio, 0)
        self.output_mode_group.addButton(self.append_file_radio, 1)
        self.new_file_radio.toggled.connect(self.on_output_mode_changed)
        options_layout.addWidget(self.new_file_radio)
        
        self.new_file_options_widget = QWidget()
        new_file_layout = QVBoxLayout(self.new_file_options_widget)
        new_file_layout.setContentsMargins(20, 0, 0, 0)
        self.combine_checkbox = QCheckBox("Combine all CSVs into one Excel file with multiple sheets")
        self.combine_checkbox.setChecked(True)
        self.combine_checkbox.stateChanged.connect(self.on_combine_changed)
        new_file_layout.addWidget(self.combine_checkbox)
        self.sheet_names_label = QLabel("Sheet names (optional, one per line):")
        self.sheet_names_text = QTextEdit()
        self.sheet_names_text.setMaximumHeight(100)
        self.sheet_names_text.setPlaceholderText("Enter sheet names, one per line...")
        new_file_layout.addWidget(self.sheet_names_label)
        new_file_layout.addWidget(self.sheet_names_text)
        self.output_new_layout = QHBoxLayout()
        self.output_new_label = QLabel("No output location selected")
        self.output_new_button = QPushButton("Select Output Location")
        self.output_new_button.clicked.connect(self.select_output_location)
        self.output_new_layout.addWidget(self.output_new_label)
        self.output_new_layout.addWidget(self.output_new_button)
        new_file_layout.addLayout(self.output_new_layout)
        options_layout.addWidget(self.new_file_options_widget)
        
        options_layout.addWidget(self.append_file_radio)
        
        self.append_options_widget = QWidget()
        append_layout = QVBoxLayout(self.append_options_widget)
        append_layout.setContentsMargins(20, 0, 0, 0)
        self.existing_file_layout = QHBoxLayout()
        self.existing_file_label = QLabel("No existing file selected")
        self.existing_file_button = QPushButton("Select Existing Excel File")
        self.existing_file_button.clicked.connect(self.select_existing_file)
        self.existing_file_layout.addWidget(self.existing_file_label)
        self.existing_file_layout.addWidget(self.existing_file_button)
        append_layout.addLayout(self.existing_file_layout)
        
        self.override_checkbox = QCheckBox("Override the selected file (merges new data into existing sheets)")
        self.override_checkbox.setChecked(True)
        self.override_checkbox.stateChanged.connect(self.on_override_changed)
        append_layout.addWidget(self.override_checkbox)
        
        self.output_copy_widget = QWidget()
        output_copy_layout = QHBoxLayout(self.output_copy_widget)
        output_copy_layout.setContentsMargins(0, 0, 0, 0)
        self.output_copy_label = QLabel("No output location selected")
        self.output_copy_button = QPushButton("Select Output Location (Copy)")
        self.output_copy_button.clicked.connect(self.select_output_location)
        output_copy_layout.addWidget(self.output_copy_label)
        output_copy_layout.addWidget(self.output_copy_button)
        append_layout.addWidget(self.output_copy_widget)
        
        options_layout.addWidget(self.append_options_widget)
        
        self.detect_similar_checkbox = QCheckBox("Detect and merge files with similar names (e.g., file_2024 & file_2025)")
        self.detect_similar_checkbox.setChecked(False)
        self.detect_similar_checkbox.stateChanged.connect(self.on_detect_similar_changed)
        options_layout.addWidget(self.detect_similar_checkbox)

        # --- FIX 2: Add UI for duplicate detection keys ---
        self.duplicate_keys_label = QLabel("Duplicate check columns (comma-separated, optional):")
        self.duplicate_keys_input = QLineEdit()
        self.duplicate_keys_input.setPlaceholderText("e.g., ID,Name,Date")
        options_layout.addWidget(self.duplicate_keys_label)
        options_layout.addWidget(self.duplicate_keys_input)
        
        main_layout.addWidget(options_group)
        
        progress_group = QGroupBox("Progress")
        progress_layout = QVBoxLayout(progress_group)
        self.progress_bar = QProgressBar()
        self.status_label = QLabel("Ready to convert")
        progress_layout.addWidget(self.progress_bar)
        progress_layout.addWidget(self.status_label)
        main_layout.addWidget(progress_group)
        self.convert_button = QPushButton("Convert to Excel")
        self.convert_button.clicked.connect(self.start_conversion)
        self.convert_button.setMinimumHeight(40)
        convert_font = QFont()
        convert_font.setPointSize(12)
        convert_font.setBold(True)
        self.convert_button.setFont(convert_font)
        main_layout.addWidget(self.convert_button)
        log_group = QGroupBox("Log")
        log_layout = QVBoxLayout(log_group)
        self.log_text = QTextEdit()
        self.log_text.setMaximumHeight(150)
        self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text)
        main_layout.addWidget(log_group)
        
        self.update_ui_state()
        self.log("Application started. Select CSV files or folder to begin.")

    def on_selection_mode_changed(self):
        if self.files_radio.isChecked():
            self.csv_button.setText("Select CSV Files")
        else:
            self.csv_button.setText("Select Folder")
        self.csv_files = []
        self.csv_label.setText("No CSV files selected")
        self.file_list.clear()
        self.update_ui_state()

    def on_output_mode_changed(self):
        self.output_path = ""
        self.existing_file_path = ""
        self.output_new_label.setText("No output location selected")
        self.existing_file_label.setText("No existing file selected")
        self.output_copy_label.setText("No output location selected")
        self.update_ui_state()

    def on_override_changed(self):
        self.output_path = ""
        self.output_copy_label.setText("No output location selected")
        self.update_ui_state()

    def select_csv_files(self):
        if self.files_radio.isChecked():
            files, _ = QFileDialog.getOpenFileNames(self, "Select CSV Files", "", "CSV Files (*.csv);;All Files (*)")
            if files:
                self.csv_files = files
                self.csv_label.setText(f"{len(files)} CSV file(s) selected")
                self.update_file_list()
                self.log(f"Selected {len(files)} CSV files")
        else:
            folder = QFileDialog.getExistingDirectory(self, "Select Folder to Search for CSV Files")
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
        if self.detect_similar_checkbox.isChecked() and self.csv_files:
            self.preview_similar_files()
        self.update_ui_state()

    def find_csv_files_recursive(self, folder_path):
        csv_files = []
        folder = Path(folder_path)
        for csv_file in folder.rglob("*.csv"):
            csv_files.append(str(csv_file))
        return sorted(csv_files)

    def update_file_list(self):
        self.file_list.clear()
        for file_path in self.csv_files:
            self.file_list.addItem(os.path.basename(file_path))

    def select_existing_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Existing Excel File", "", "Excel Files (*.xlsx *.xls);;All Files (*)")
        if file_path:
            self.existing_file_path = file_path
            self.existing_file_label.setText(f"Existing file: {os.path.basename(file_path)}")
            self.log(f"Selected existing Excel file: {file_path}")
            self.update_ui_state()

    def select_output_location(self):
        if self.append_file_radio.isChecked():
            file_path, _ = QFileDialog.getSaveFileName(self, "Save Updated Excel File As", "updated_data.xlsx", "Excel Files (*.xlsx);;All Files (*)")
            if file_path:
                self.output_path = file_path
                self.output_copy_label.setText(f"Output: {os.path.basename(file_path)}")
                self.log(f"Output file for copy: {file_path}")
        elif self.combine_checkbox.isChecked():
            file_path, _ = QFileDialog.getSaveFileName(self, "Save Combined Excel File As", "combined_data.xlsx", "Excel Files (*.xlsx);;All Files (*)")
            if file_path:
                self.output_path = file_path
                self.output_new_label.setText(f"Output: {os.path.basename(file_path)}")
                self.log(f"Output file: {file_path}")
        else:
            directory = QFileDialog.getExistingDirectory(self, "Select Output Directory")
            if directory:
                self.output_path = directory
                self.output_new_label.setText(f"Output directory: {os.path.basename(directory)}")
                self.log(f"Output directory: {directory}")
        self.update_ui_state()

    def preview_similar_files(self):
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
        patterns = [
            r'_\d{8}$', r'_\d{4}-\d{2}-\d{2}$', r'_\d{4}_\d{2}_\d{2}$', r'_\d{6}$',
            r'_\d{4}$', r'-\d{8}$', r'-\d{4}-\d{2}-\d{2}$', r'-\d{4}_\d{2}_\d{2}$',
            r'-\d{6}$', r'-\d{4}$', r'\d{8}$', r'\d{4}-\d{2}-\d{2}$',
            r'\d{4}_\d{2}_\d{2}$', r'\d{6}$', r'_\d+$', r'-\d+$', r'\d+$'
        ]
        base_name = file_name
        for pattern in patterns:
            base_name = re.sub(pattern, '', base_name)
        return base_name.rstrip('_-\t ')

    def on_detect_similar_changed(self):
        if self.detect_similar_checkbox.isChecked() and self.csv_files:
            self.preview_similar_files()
        self.update_ui_state()

    def on_combine_changed(self):
        is_combine = self.combine_checkbox.isChecked()
        self.sheet_names_label.setVisible(is_combine)
        self.sheet_names_text.setVisible(is_combine)
        self.output_path = ""
        self.output_new_label.setText("No output location selected")
        self.update_ui_state()

    def update_ui_state(self):
        is_new_mode = self.new_file_radio.isChecked()
        is_append_mode = self.append_file_radio.isChecked()
        is_override = self.override_checkbox.isChecked()

        self.new_file_options_widget.setVisible(is_new_mode)
        self.append_options_widget.setVisible(is_append_mode)
        
        self.output_copy_widget.setVisible(is_append_mode and not is_override)

        # Show duplicate key input only when appending/merging
        self.duplicate_keys_label.setVisible(is_append_mode)
        self.duplicate_keys_input.setVisible(is_append_mode)

        has_files = len(self.csv_files) > 0
        ready_to_convert = False
        
        if is_new_mode:
            if has_files and self.output_path:
                ready_to_convert = True
        elif is_append_mode:
            if has_files and self.existing_file_path:
                if is_override:
                    ready_to_convert = True
                elif not is_override and self.output_path:
                    ready_to_convert = True

        self.convert_button.setEnabled(ready_to_convert)
        
        if not has_files:
            self.status_label.setText("Please select CSV files or folder")
        elif not ready_to_convert:
            self.status_label.setText("Please complete all required selections for the chosen mode")
        else:
            self.status_label.setText("Ready to convert")

    def start_conversion(self):
        self.update_ui_state() 
        if not self.convert_button.isEnabled():
            QMessageBox.warning(self, "Warning", "Please ensure all required files and output locations are selected.")
            return

        self.convert_button.setEnabled(False)
        self.csv_button.setEnabled(False)
        self.output_new_button.setEnabled(False)
        self.existing_file_button.setEnabled(False)
        self.output_copy_button.setEnabled(False)
        
        self.progress_bar.setValue(0)
        
        sheet_names = []
        if self.new_file_radio.isChecked() and self.combine_checkbox.isChecked():
            sheet_names_text = self.sheet_names_text.toPlainText().strip()
            if sheet_names_text:
                sheet_names = [name.strip() for name in sheet_names_text.split('\n')]

        # Get duplicate keys from input field
        duplicate_keys = [key.strip() for key in self.duplicate_keys_input.text().split(',') if key.strip()]

        # Start worker thread
        self.worker = ConversionWorker(
            self.csv_files,
            self.output_path,
            self.combine_checkbox.isChecked() if self.new_file_radio.isChecked() else True,
            sheet_names,
            self.detect_similar_checkbox.isChecked(),
            self.append_file_radio.isChecked(),
            self.override_checkbox.isChecked(),
            self.existing_file_path,
            duplicate_keys
        )

        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.status.connect(self.status_label.setText)
        self.worker.finished.connect(self.on_conversion_finished)

        self.worker.start()

        mode_desc = "append mode" if self.append_file_radio.isChecked() else "new file mode"
        similar_desc = " with similar file detection" if self.detect_similar_checkbox.isChecked() else ""
        override_desc = ""
        if self.append_file_radio.isChecked():
            override_desc = " (override enabled)" if self.override_checkbox.isChecked() else " (creating a copy)"
        
        self.log(f"Starting {mode_desc}{similar_desc}{override_desc}...")

    def on_conversion_finished(self, success, message):
        """Handle conversion completion"""
        # Re-enable UI
        self.convert_button.setEnabled(True)
        self.csv_button.setEnabled(True)
        self.output_new_button.setEnabled(True)
        self.existing_file_button.setEnabled(True)
        self.output_copy_button.setEnabled(True)

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
        self.update_ui_state()

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
    app.setApplicationVersion("3.3") # Version bump for new features

    # Create and show main window
    window = CSVToExcelConverter()
    window.show()

    # Start event loop
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
