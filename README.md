RPYtoEXCEL

A simple desktop tool for Ren'Py game translation.

RPYtoEXCEL helps translation teams move text between Ren'Py .rpy scripts and Excel:

.rpy → .xlsx → translate → .xlsx → .rpy

The application uses a graphical interface built with Tkinter and supports drag-and-drop when tkinterdnd2 is installed.

Features

Export translatable text from .rpy files to Excel.

Import/merge translated Excel files back into .rpy files.

Select individual .rpy files or entire folders.

Recursively scan folders for .rpy files.

Drag and drop .rpy files/folders.

Drag and drop translated Excel files.

Keep original file name and line number information.

Detect Ren'Py translate blocks.

Handle old / new translation pairs.

Include character/tag information when available.

Add an Addition sheet for extra translation entries.

Generate a Notes sheet with translation progress.

Automatically format Excel sheets as tables.

Multilingual interface with 16 languages.

Supported Interface Languages

English

Tiếng Việt

Español

中文

Deutsch

Français

Bahasa Indonesia

العربية

Türkçe

Português

Polski

Русский

Українська

日本語

한국어

Italiano

Requirements

Python 3

pandas

openpyxl

tkinter / tkinter.ttk

tkinterdnd2 (optional, for drag-and-drop support)

Install dependencies

pip install pandas openpyxl tkinterdnd2

On some Linux distributions, Tkinter may need to be installed separately.

Usage

Run the application:

python RPYtoEXCEL.py

1. Export RPY → Excel

Open the Export RPY -> Excel tab.

Add .rpy files, or add a folder containing .rpy files.

You can also drag and drop files/folders into the list.

Click Start Export to Excel.

Choose where to save the .xlsx file.

The generated workbook contains translation data such as:

Column

Description

File

Original .rpy file

Line

Source line number

Original Location

Detected # game/... location

Label

Detected Ren'Py translate label

Tag Char

Character/tag information when available

Original

Original text

Translated

Translation

Notes

Notes for translators

The workbook also contains a notes/progress sheet and an Addition sheet.

2. Translate the Excel file

Open the generated .xlsx file in Excel, LibreOffice Calc, Google Sheets, or another compatible spreadsheet application.

Put your translation in the Translated column (or the equivalent translated column for the selected interface language).

Do not change the File and Line information unless you know exactly what you are doing, because these values are used to locate the original text when merging.

3. Merge Excel → RPY

Open the Merge Excel -> RPY tab.

Select the translated .xlsx file.

Add the target .rpy files or a folder containing them.

Click Start Merging Translations.

The program matches translation entries using the .rpy file name and line number, then writes the translated text back into the target .rpy files.

Excel Compatibility

The merge function recognizes translated/original/file/line column names used by the supported interface languages, so an exported workbook can be processed even after changing the application's interface language.

The Notes-style sheets are ignored during the merge process.

Drag & Drop

Drag-and-drop support is optional.

If tkinterdnd2 is unavailable, the application still works normally through the file/folder selection buttons.

Install it with:

pip install tkinterdnd2

Important Notes

Back up your .rpy files before merging translations.

The program writes changes directly to the selected target .rpy files.

The tool reads .rpy files using UTF-8 encoding.

Only use the merge function on the intended game files.

Always test the game after merging translations.

Ren'Py syntax and special characters can affect whether a translated line works correctly.

If a translation contains characters that are meaningful to Ren'Py or Python strings, make sure they are escaped/handled correctly.

Project Structure

A minimal setup can look like:

RPYtoEXCEL/
├── RPYtoEXCEL.py
└── README.md

How It Works

Export

Ren'Py .rpy files
       │
       ▼
  RPYtoEXCEL
       │
       ▼
   Excel .xlsx
       │
       ▼
  Translator edits
  "Translated" column

Merge

Translated Excel
       │
       ▼
  RPYtoEXCEL
       │
       ▼
Match File + Line
       │
       ▼
Updated .rpy files
       │
       ▼
   Test in game

Tech Stack

Python

Tkinter — graphical user interface

pandas — spreadsheet/data processing

openpyxl — Excel workbook creation and formatting

tkinterdnd2 — optional drag-and-drop support

Status

This is a practical tool for Ren'Py translation workflows. It is intended to make extracting, translating, and reinserting game dialogue easier for translation teams.

License

Add your preferred license here before publishing the project publicly.