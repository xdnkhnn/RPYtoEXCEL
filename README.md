# Ren'Py Translation Tool

GUI tool for extracting Ren'Py `.rpy` translation files into formatted Excel spreadsheets and merging them back into game scripts seamlessly.

## Key Features

* **Multi-Language UI:** Built-in support for 16 languages (English, Vietnamese, Spanish, Chinese, French, German, Japanese, Russian, etc.).
* **Folder & Subdirectory Processing:** Supports recursive folder scanning (`.rpy` files located deep inside nested subdirectories are automatically detected and updated).
* **Excel Formatting & Progress Tracking:** Exports structured Excel Tables with clean gridlines, auto-fitted columns, and built-in formulas in the `Notes` sheet to track translation completion (% progress).

---

## How to Use

1. Download `app.exe` or `app.zip` from [Releases](https://github.com/xdnkhnn/RenPyTranslatorApp/releases).
2. **Export (RPY -> Excel):**
   * Open the app and select your preferred language.
   * Add individual `.rpy` files or use **Add Folder...** to scan subdirectories automatically.
   * Click **START EXPORT TO EXCEL** to generate the `.xlsx` file.
3. Fill in your translations under the `Translated` column in Excel.
4. **Merge (Excel -> RPY):**
   * Switch to the **Merge Excel -> RPY** tab.
   * Select your translated Excel file and the target `.rpy` game directory.
   * Click **START MERGING TRANSLATIONS**.

---

## Build from Source

```bash
# Clone repository
git clone [https://github.com/xdnkhnn/RenPyTranslatorApp.git](https://github.com/xdnkhnn/RenPyTranslatorApp.git)
cd RenPyTranslatorApp

# Install dependencies
pip install pandas openpyxl pyinstaller

# Run application
python app.py

# Build single executable (.exe)
python -m PyInstaller --noconsole --onefile app.py