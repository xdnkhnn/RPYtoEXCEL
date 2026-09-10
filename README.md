# RPYtoEXCEL - Ren'Py Translation Tool

RPYtoEXCEL is a utility designed to extract translation dialogue and strings from Ren'Py script files (.rpy) into Excel spreadsheets (.xlsx) for easy translation, and merge the translated text back into the game scripts.

---

## Features

* Smart Extraction (Export RPY -> Excel):
  * Scans and extracts dialogues, character tags, strings, and old/new translation blocks from .rpy files.
  * Generates structured Excel spreadsheets with centered headers, auto-adjusted column widths, and separate sheets for each script file.
  * Automatically creates a Notes sheet with a progress tracking table.
* Flexible Merging (Merge Excel -> RPY):
  * Supports merging translations back into individual .rpy files or entire game directories.
  * Dynamically recognizes translation column names across various languages.
* User Interface:
  * Drag and Drop support for files (.rpy, .xlsx) and directories.
  * Multilingual UI supporting 16 languages.

---

## Installation & Setup

### 1. Prerequisites
* Python 3.8 or higher.

### 2. Install Required Libraries
Open your Terminal or Command Prompt and run:

```bash
pip install pandas openpyxl tkinterdnd2