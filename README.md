# RenPy Translation App

GUI tool for extracting Ren'Py `.rpy` files to Excel and merging them back.

## How to use

1. Download `app.exe` from [Releases](https://github.com/xdnkhnn/RenPyTranslatorApp/releases).
2. **Export:** Drop `.rpy` files -> click Export -> get `.xlsx`.
3. Fill translations in the `Translated` column.
4. **Merge:** Pick the Excel file + target `.rpy` folder -> click Merge.

## Build from source

```bash
pip install pandas openpyxl pyinstaller
python app.py
