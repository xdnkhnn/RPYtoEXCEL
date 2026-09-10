# RPYtoEXCEL

A simple desktop tool for **Ren'Py game translation**.

RPYtoEXCEL helps translation teams move text between Ren'Py `.rpy` scripts and Excel:

**`.rpy` → `.xlsx` → translate → `.xlsx` → `.rpy`**

The application uses a graphical interface built with Tkinter and supports drag-and-drop when `tkinterdnd2` is installed.

## Download

You can download the latest version from the GitHub Releases page:

[![Latest Release](https://img.shields.io/github/v/release/xdnkhnn/RenPyTranslatorApp?label=Latest%20Release)](https://github.com/xdnkhnn/RenPyTranslatorApp/releases/latest)

👉 **[Download the latest release](https://github.com/xdnkhnn/RenPyTranslatorApp/releases/latest)**

## Features

- Export translatable text from `.rpy` files to Excel.
- Import/merge translated Excel files back into `.rpy` files.
- Select individual `.rpy` files or entire folders.
- Recursively scan folders for `.rpy` files.
- Drag and drop `.rpy` files/folders.
- Drag and drop translated Excel files.
- Keep original file name and line number information.
- Detect Ren'Py `translate` blocks.
- Handle `old` / `new` translation pairs.
- Include character/tag information when available.
- Add an `Addition` sheet for extra translation entries.
- Generate a `Notes` sheet with translation progress.
- Automatically format Excel sheets as tables.
- Multilingual interface with 16 languages.

## Supported Interface Languages

1. English
2. Tiếng Việt
3. Español
4. 中文
5. Deutsch
6. Français
7. Bahasa Indonesia
8. العربية
9. Türkçe
10. Português
11. Polski
12. Русский
13. Українська
14. 日本語
15. 한국어
16. Italiano

## Requirements

- Python 3
- `pandas`
- `openpyxl`
- `tkinter` / `tkinter.ttk`
- `tkinterdnd2` *(optional, for drag-and-drop support)*

### Install dependencies

```bash
pip install pandas openpyxl tkinterdnd2
```

On some Linux distributions, Tkinter may need to be installed separately.

## Usage

Run the application:

```bash
python RPYtoEXCEL.py
```

### 1. Export RPY → Excel

1. Open the **Export RPY -> Excel** tab.
2. Add `.rpy` files, or add a folder containing `.rpy` files.
3. You can also drag and drop files/folders into the list.
4. Click **Start Export to Excel**.
5. Choose where to save the `.xlsx` file.

The generated workbook contains translation data such as:

| Column | Description |
|---|---|
| File | Original `.rpy` file |
| Line | Source line number |
| Original Location | Detected `# game/...` location |
| Label | Detected Ren'Py translate label |
| Tag Char | Character/tag information when available |
| Original | Original text |
| Translated | Translation |
| Notes | Notes for translators |

The workbook also contains a notes/progress sheet and an `Addition` sheet.

### 2. Translate the Excel file

Open the generated `.xlsx` file in Excel, LibreOffice Calc, Google Sheets, or another compatible spreadsheet application.

Put your translation in the **Translated** column (or the equivalent translated column for the selected interface language).

Do not change the **File** and **Line** information unless you know exactly what you are doing, because these values are used to locate the original text when merging.

### 3. Merge Excel → RPY

1. Open the **Merge Excel -> RPY** tab.
2. Select the translated `.xlsx` file.
3. Add the target `.rpy` files or a folder containing them.
4. Click **Start Merging Translations**.

The program matches translation entries using the `.rpy` file name and line number, then writes the translated text back into the target `.rpy` files.

## Excel Compatibility

The merge function recognizes translated/original/file/line column names used by the supported interface languages, so an exported workbook can be processed even after changing the application's interface language.

The `Notes`-style sheets are ignored during the merge process.

## Drag & Drop

Drag-and-drop support is optional.

If `tkinterdnd2` is unavailable, the application still works normally through the file/folder selection buttons.

Install it with:

```bash
pip install tkinterdnd2
```

## Important Notes

- **Back up your `.rpy` files before merging translations.**
- The program writes changes directly to the selected target `.rpy` files.
- The tool reads `.rpy` files using UTF-8 encoding.
- Only use the merge function on the intended game files.
- Always test the game after merging translations.
- Ren'Py syntax and special characters can affect whether a translated line works correctly.
- If a translation contains characters that are meaningful to Ren'Py or Python strings, make sure they are escaped/handled correctly.

## Project Structure

A minimal setup can look like:

```text
RPYtoEXCEL/
├── RPYtoEXCEL.py
└── README.md
```

## How It Works

### Export

```text
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
```

### Merge

```text
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
```

## Tech Stack

- **Python**
- **Tkinter** — graphical user interface
- **pandas** — spreadsheet/data processing
- **openpyxl** — Excel workbook creation and formatting
- **tkinterdnd2** — optional drag-and-drop support

## Status

This is a practical tool for Ren'Py translation workflows. It is intended to make extracting, translating, and reinserting game dialogue easier for translation teams.

## License

Add your preferred license here before publishing the project publicly.
