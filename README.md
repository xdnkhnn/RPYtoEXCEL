# RPYtoEXCEL - Ren'Py Translation Tool

RPYtoEXCEL is a desktop utility designed to streamline the translation process for Ren'Py games. It extracts dialogue, character tags, and string blocks from script files (.rpy) into organized Excel spreadsheets (.xlsx), and merges translated text back into the game script.

---

## Features

* Smart Extraction (Export RPY -> Excel):
  * Scans and extracts dialogues, character tags, strings, and old/new translation blocks from .rpy files.
  * Generates formatted Excel spreadsheets with auto-fitted columns, centered headers, and separate sheets for each script.
  * Creates an automatic Notes/Summary sheet with real-time translation progress tracking.
* Flexible Merging (Merge Excel -> RPY):
  * Merges translations back into individual .rpy files or full game directories based on line index.
  * Recognizes translation column names dynamically across multiple languages.
* User Interface:
  * Drag and Drop support for files (.rpy, .xlsx) and folders.
  * Multilingual UI supporting 16 languages.

---

## Output Excel Structure

The generated Excel file contains the following columns for each script sheet:

| Column | Name | Description |
| --- | --- | --- |
| A | File | Name of the source script file |
| B | Line | Line number of the string in the original script |
| C | Original Location | Script filepath and location reference |
| D | Label | Ren'Py translation label name |
| E | Tag Char | Speaker character ID / tag |
| F | Original | Original untranslated text |
| G | Translated | Column where translations should be entered |
| H | Notes | Optional notes for translators |

---

## Installation & Setup

### 1. Prerequisites
* Python 3.8 or higher.

### 2. Install Dependencies
Open your Terminal or Command Prompt and run:

pip install pandas openpyxl tkinterdnd2

### 3. Run from Source
python RPYtoEXCEL.py

---

## Usage Guide

### 1. Extracting Text (Export Tab)
1. Add .rpy files or drag and drop your game directory into the application.
2. Click START EXPORT TO EXCEL and choose a save location.
3. Open the generated .xlsx file and fill in your translations in the Translated column.

### 2. Applying Translations (Merge Tab)
1. Drag and drop the translated Excel file into Step 1.
2. Drag and drop the target .rpy files or game folder into Step 2.
3. Click START MERGING TRANSLATIONS.

---

## Important Notes & Troubleshooting

* Line Numbers: Do not modify the Line column (Column B) in Excel, as the tool relies on line indexing to place translations back into .rpy files correctly.
* Double Quotes: Preserve internal formatting tags (such as {w}, {p}, [player_name]) and escape quotes (\") inside your translated text.
* Admin Privilege Issue: If drag-and-drop does not work on Windows, ensure the application is NOT running as Administrator (Windows blocks drag-and-drop from standard File Explorer to elevated processes).

---

## Building Executable (.exe)

To package the tool into a standalone Windows executable:

pip install pyinstaller
python -m PyInstaller --noconsole --onefile --collect-data tkinterdnd2 -n RPYtoEXCEL RPYtoEXCEL.py

The output file will be saved in the dist/ directory.

---

## License
Distributed under the MIT License.