# RPYtoEXCEL

A lightweight GUI utility designed to extract Ren'Py `.rpy` script files into formatted Excel spreadsheets and seamlessly merge translations back into your game code without breaking syntax.

---

## Key Features

* **RPY to Excel Extraction:** Automatically parses dialogue lines, `label` blocks, and original file paths from `.rpy` files into organized Excel tables.
* **Automated Translation Merging:** Safely merges translated text from the `Translated` column back into your `.rpy` script files.
* **Recursive Folder Scanning:** Automatically searches and detects `.rpy` files located deep inside nested subdirectories.
* **Progress Tracking Sheet:** Generates an automated `Notes` tab with Excel formulas to track completion progress (%) per file and for the entire project.
* **Multi-Language UI:** Built-in localization support for 16 languages (English, Vietnamese, Spanish, Chinese, French, German, Japanese, Russian, etc.).

---

## How to Use

1. Download `RPYtoEXCEL.exe` or `RPYtoEXCEL.zip` from [Releases](https://github.com/xdnkhnn/RenPyTranslatorApp/releases).
2. **Export (RPY -> Excel):**
   * Open the app and select your preferred language.
   * Select individual `.rpy` files or click **Add Folder...** to scan subdirectories automatically.
   * Click **START EXPORT TO EXCEL** to generate your `.xlsx` file.
3. Fill in your translations under the **`Translated`** column in Excel.
4. **Merge (Excel -> RPY):**
   * Switch to the **Merge Excel -> RPY** tab.
   * Select your translated Excel file and your target `.rpy` game folder.
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
python RPYtoEXCEL.py

# Build single executable (.exe)
python -m PyInstaller --noconsole --onefile RPYtoEXCEL.py