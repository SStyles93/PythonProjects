# Advanced CSV to Excel Converter - User Manual & Documentation

## Table of Contents

1.  **Introduction**
    *   What is the Advanced CSV to Excel Converter?
2.  **User Manual**
    *   System Requirements
    *   Installation
    *   Main Window Overview
    *   Step-by-Step Guides
        *   Guide 1: Combining Multiple CSVs into a Single Excel File
        *   Guide 2: Converting Each CSV to a Separate Excel File
        *   Guide 3: Merging New CSV Data into an Existing Excel File (Override Mode)
        *   Guide 4: Appending New CSV Data as a Copy
    *   Understanding the Options
        *   File Selection
        *   Conversion Options
        *   Duplicate Detection
    *   Troubleshooting
3.  **Technical Documentation**
    *   Application Architecture
    *   Core Classes and Their Roles
        *   `CSVToExcelConverter` (Main UI Class)
        *   `ConversionWorker` (Processing Thread)
    *   Key Functions Explained
    *   Dependencies
4.  **Future Improvements & Development Roadmap**
    *   Profile Management
    *   Advanced Duplicate Handling
    *   Interactive Column Mapping
    *   Data Preview Pane
    *   Command-Line Interface (CLI) Mode
5.  **Conclusion**

---

### 1. Introduction

#### What is the Advanced CSV to Excel Converter?

The Advanced CSV to Excel Converter is a desktop application designed to simplify and automate the process of converting and managing data from CSV (Comma-Separated Values) files into Microsoft Excel (`.xlsx`) format.

It goes beyond simple one-to-one conversion, offering powerful features for data management, including:
*   **Bulk Conversion:** Convert hundreds of CSV files from a folder and its subfolders at once.
*   **File Merging:** Intelligently group and merge similarly named files (e.g., `report_2024-01.csv` and `report_2024-02.csv`) into a single, clean dataset.
*   **Data Appending:** Add new data from CSVs to an existing Excel file without losing your original data.
*   **Smart Override:** Merge new rows into the correct sheets of an existing Excel file, automatically skipping duplicates to keep your master file clean and up-to-date.
*   **Custom Duplicate Checks:** Define specific columns (like `ID` or `Email`) to identify duplicate rows, giving you full control over your data integrity.

This tool is ideal for anyone who regularly works with data exports, reports, or logs that are generated as multiple CSV files.

### 2. User Manual

#### System Requirements
*   **Operating System:** Windows, macOS, or Linux.
*   **Python:** Python 3.6 or newer.
*   **Libraries:** `PyQt5` and `pandas`.

#### Installation
1.  Ensure you have Python 3 installed on your system.
2.  Install the required libraries by opening a terminal or command prompt and running:
    ```sh
    pip install PyQt5 pandas openpyxl
    ```
3.  Save the application code as a Python file (e.g., `converter_app.py`).
4.  Run the application from your terminal:
    ```sh
    python converter_app.py
    ```

#### Main Window Overview

The application window is divided into four main sections:

1.  **File Selection:** Where you select your source CSV files or folders.
2.  **Conversion Options:** The core section where you define *how* the conversion or merge should happen.
3.  **Progress:** A real-time view of the conversion process, including a progress bar and status messages.
4.  **Log:** A detailed log of all actions, selections, and outcomes (success or errors).

#### Step-by-Step Guides

##### Guide 1: Combining Multiple CSVs into a Single Excel File

Use this when you have several CSV files that you want to become individual sheets in one new Excel file.

1.  **Select Files:** In the "File Selection" area, click **"Select CSV Files"** and choose all the files you want to combine.
2.  **Set Mode:** In "Conversion Options," ensure **"Create new Excel file(s)"** is selected.
3.  **Configure Options:**
    *   Check the box for **"Combine all CSVs into one Excel file..."**.
    *   (Optional) If you want custom sheet names, enter them in the text box, one per line, in the same order as your selected files.
4.  **Set Output:** Click **"Select Output Location"** and choose a name and location for your new combined Excel file (e.g., `Combined_Reports.xlsx`).
5.  **Convert:** Click the **"Convert to Excel"** button.

##### Guide 2: Converting Each CSV to a Separate Excel File

Use this for simple bulk conversion where each CSV becomes its own Excel file.

1.  **Select Files:** Select your CSV files or use the **"Select folder"** option to find all CSVs in a directory.
2.  **Set Mode:** Ensure **"Create new Excel file(s)"** is selected.
3.  **Configure Options:**
    *   **Uncheck** the box for **"Combine all CSVs into one Excel file..."**.
4.  **Set Output:** Click **"Select Output Location"** and choose a **folder** where the new Excel files will be saved.
5.  **Convert:** Click **"Convert to Excel"**.

##### Guide 3: Merging New CSV Data into an Existing Excel File (Override Mode)

This is the most powerful feature for maintaining a "master" spreadsheet. It adds new rows from CSVs to the correct sheets in your existing file and skips duplicates.

1.  **Select Files:** Select the new CSV files containing the data you want to add.
2.  **Set Mode:** In "Conversion Options," select **"Append to existing Excel file"**.
3.  **Configure Options:**
    *   Click **"Select Existing Excel File"** and choose your master spreadsheet.
    *   Ensure the **"Override the selected file..."** checkbox is **checked**. This tells the app to modify the master file directly.
    *   (Optional but Recommended) In the **"Duplicate check columns"** field, enter the column names that uniquely identify a row (e.g., `Transaction ID,Date`). This prevents adding data that's already there.
4.  **Convert:** Click **"Convert to Excel"**. The application will read your master file, merge the new data, and save the changes back to the same file.

##### Guide 4: Appending New CSV Data as a Copy

Use this if you want to merge data but keep your original master file untouched.

1.  **Select Files:** Select the new CSV files.
2.  **Set Mode:** Select **"Append to existing Excel file"**.
3.  **Configure Options:**
    *   Click **"Select Existing Excel File"** to choose your master spreadsheet.
    *   **Uncheck** the **"Override the selected file..."** checkbox.
    *   A new option, **"Select Output Location (Copy)"**, will appear. Click it and choose a name and location for the new, merged file.
4.  **Convert:** Click **"Convert to Excel"**. A new file will be created containing all the data from your master file plus the new data from the CSVs.

#### Understanding the Options

*   **Detect and merge files with similar names:** If checked, the app will automatically group files like `Sales-Jan.csv` and `Sales-Feb.csv` into a single sheet named `Sales`. This is useful for combining monthly or daily reports.
*   **Duplicate check columns:** When appending data, this tells the app how to identify a duplicate. If you provide column names (e.g., `ID,Name`), a row from a new CSV will be skipped if another row with the same `ID` and `Name` already exists in the target sheet. If left blank, a row is only considered a duplicate if *all* its values are identical to an existing row.

#### Troubleshooting

*   **`AttributeError: 'QHBoxLayout' object has no attribute 'setVisible'`:** This is a known bug from a previous version. Please ensure you are running the latest version of the script.
*   **File Not Found Error:** Make sure the selected files have not been moved or deleted after being selected in the app. Re-select the files if necessary.
*   **Permission Denied Error:** This usually happens if the Excel file you are trying to override is currently open in Excel or another program. Close the file and try again.
*   **Incorrect Merging:** If data is not merging as expected, double-check your "Duplicate check columns." Ensure the column names are spelled exactly as they appear in your CSV/Excel files.

---

### 3. Technical Documentation

#### Application Architecture

The application follows a standard GUI architecture that separates the user interface from the business logic to ensure the UI remains responsive during long operations.

*   **Main Thread:** Runs the `CSVToExcelConverter` class, which manages the PyQt5 window, handles user input (button clicks, selections), and updates the UI.
*   **Worker Thread:** When the "Convert" button is clicked, a `ConversionWorker` object is created and moved to a separate `QThread`. This thread performs all the heavy lifting: reading CSVs, processing data with `pandas`, and writing Excel files. This prevents the GUI from freezing.
*   **Signal and Slot Mechanism:** The worker thread communicates back to the main thread using PyQt's signals (`progress`, `status`, `finished`). The main thread has "slots" (functions) connected to these signals to update the progress bar, status label, and display final messages.

#### Core Classes and Their Roles

##### `CSVToExcelConverter(QMainWindow)`
*   **Role:** The main application window and controller.
*   **Responsibilities:**
    *   `init_ui()`: Builds the entire user interface, including layouts, widgets, and groups.
    *   `on_*` methods (e.g., `on_output_mode_changed`, `on_override_changed`): Event handlers that respond to user interactions, like clicking radio buttons or checkboxes.
    *   `update_ui_state()`: A critical method that dynamically shows, hides, enables, or disables UI elements based on the user's current selections. This provides an intuitive user experience.
    *   `select_*` methods (e.g., `select_csv_files`): Open file dialogs to get input/output paths from the user.
    *   `start_conversion()`: Gathers all settings from the UI, instantiates the `ConversionWorker`, connects signals to slots, and starts the worker thread.
    *   `on_conversion_finished()`: A slot that is called when the worker thread emits the `finished` signal. It re-enables the UI and shows a success or error message.

##### `ConversionWorker(QThread)`
*   **Role:** The data processing engine. It runs independently of the UI.
*   **Responsibilities:**
    *   `run()`: The main entry point for the thread. It contains the primary logic that decides which conversion method to call based on the user's settings.
    *   `append_to_existing_file()`: Contains the logic for the most complex use case. It reads an existing Excel file into a dictionary of pandas DataFrames, iterates through the new CSVs, merges data into the appropriate DataFrame, and writes the entire dictionary back to an Excel file, preserving all sheets.
    *   `merge_with_duplicate_detection()`: Compares a new DataFrame against an existing one. It uses a user-provided list of key columns to identify duplicates. If no keys are provided, it performs a full-row comparison.
    *   `group_similar_files()` & `extract_base_name()`: Work together to implement the "Detect similar files" feature using regular expressions to strip dates and numbers from filenames.
    *   `process_grouped_files()`: Handles the logic for processing files that have been grouped by the `group_similar_files` method.

#### Key Functions Explained

*   `append_to_existing_file()`: The core of the override logic. The critical step is `updated_sheets = existing_sheets.copy()`. This creates a shallow copy of the dictionary of DataFrames. The loop then modifies this `updated_sheets` dictionary. Finally, the entire `updated_sheets` dictionary is written to the new Excel file, which ensures that sheets that were never touched during the loop are still included in the final output.
*   `merge_with_duplicate_detection()`: The logic `existing_rows = set(tuple(row) for row in existing_df[check_cols].to_numpy())` is a highly efficient way to check for duplicates. It creates a hash set of existing rows (based on key columns), allowing for near-instantaneous lookups (`O(1)` average time complexity) when checking if a new row exists.

#### Dependencies

*   **PyQt5:** The GUI toolkit used to build the application's front-end.
*   **pandas:** The primary data manipulation library. It is used for reading CSVs, creating and managing DataFrames, and writing to Excel files.
*   **openpyxl:** The engine used by pandas to write to the modern `.xlsx` Excel format. It is required for the `ExcelWriter`.
