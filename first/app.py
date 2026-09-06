import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment

# Từ điển đa ngôn ngữ (Multilingual Dictionary)
TRANSLATIONS = {
    "English": {
        "title": "Ren'Py Translation Tool",
        "language_label": "Language / Ngôn ngữ:",
        "tab_export": " 1. Export RPY -> Excel ",
        "tab_merge": " 2. Merge Excel -> RPY ",
        "export_frame": " Extract Translation Files ",
        "export_desc": "Select or drop .rpy files to export into a single Excel file:",
        "btn_add_files": "Add RPY Files...",
        "btn_clear_list": "Clear List",
        "btn_start_export": " START EXPORT TO EXCEL ",
        "merge_frame": " Merge Translations into RPY Files ",
        "lbl_excel_step": "1. Select translated Excel file:",
        "btn_browse_excel": "Browse Excel...",
        "lbl_rpy_step": "2. Select directory containing original .rpy files:",
        "btn_browse_folder": "Browse Folder...",
        "btn_start_merge": " START MERGING TRANSLATIONS ",
        "excel_notes_sheet": "Please provide translations in the 'Translated' column. Once finished, save the file and use the 'Merge Excel -> RPY' tab to apply translations to the game.",
        "msg_warn_select_rpy": "Please select at least one .rpy file!",
        "msg_success_export": "File successfully exported to:\n{}",
        "msg_err_export": "An error occurred while exporting the Excel file:\n{}",
        "msg_warn_excel": "Please select a valid Excel file!",
        "msg_warn_rpy_folder": "Please select a valid directory containing .rpy files!",
        "msg_success_merge": "Successfully merged translations into {} .rpy file(s)!",
        "msg_err_merge": "An error occurred while merging:\n{}",
        "dlg_select_rpy": "Select .rpy Files",
        "dlg_save_excel": "Save Output Excel File",
        "dlg_select_excel": "Select Translated Excel File",
        "dlg_select_folder": "Select Folder Containing .rpy Files",
        "warning": "Warning",
        "success": "Success",
        "error": "Error"
    },
    "Tiếng Việt": {
        "title": "Công Cụ Dịch Ren'Py",
        "language_label": "Language / Ngôn ngữ:",
        "tab_export": " 1. Trích Xuất RPY -> Excel ",
        "tab_merge": " 2. Ghép Excel -> RPY ",
        "export_frame": " Trích xuất file dịch ",
        "export_desc": "Chọn các file .rpy để xuất ra 1 file Excel duy nhất:",
        "btn_add_files": "Thêm File RPY...",
        "btn_clear_list": "Xóa Danh Sách",
        "btn_start_export": " BẮT ĐẦU XUẤT EXCEL ",
        "merge_frame": " Ghép bản dịch vào File RPY ",
        "lbl_excel_step": "1. Chọn file Excel đã dịch:",
        "btn_browse_excel": "Chọn Excel...",
        "lbl_rpy_step": "2. Chọn thư mục chứa các file RPY gốc cần cập nhật:",
        "btn_browse_folder": "Chọn Thư Mục...",
        "btn_start_merge": " BẮT ĐẦU GỘP BẢN DỊCH ",
        "excel_notes_sheet": "Vui lòng dịch ở cột 'Translated'. Sau khi hoàn tất, lưu file và sử dụng tab 'Ghép Excel -> RPY' để hợp nhất bản dịch vào game.",
        "msg_warn_select_rpy": "Vui lòng chọn ít nhất 1 file .rpy!",
        "msg_success_export": "Đã xuất file thành công tại:\n{}",
        "msg_err_export": "Có lỗi xảy ra khi xuất file Excel:\n{}",
        "msg_warn_excel": "Vui lòng chọn file Excel bản dịch hợp lệ!",
        "msg_warn_rpy_folder": "Vui lòng chọn thư mục chứa file .rpy hợp lệ!",
        "msg_success_merge": "Đã hợp nhất thành công bản dịch vào {} file .rpy!",
        "msg_err_merge": "Có lỗi xảy ra khi merge:\n{}",
        "dlg_select_rpy": "Chọn các file .rpy",
        "dlg_save_excel": "Lưu file Excel kết quả",
        "dlg_select_excel": "Chọn file Excel bản dịch",
        "dlg_select_folder": "Chọn thư mục chứa các file .rpy cần cập nhật",
        "warning": "Cảnh báo",
        "success": "Thành công",
        "error": "Lỗi"
    }
}

class RenPyTranslatorApp:
    def __init__(self, root):
        self.root = root
        self.current_lang = "English"  # Ngôn ngữ mặc định
        
        self.root.title("Ren'Py Translation Tool")
        self.root.geometry("680x560")
        self.root.resizable(True, True)

        style = ttk.Style()
        style.theme_use('clam')

        # Language Selector Bar Header
        lang_frame = ttk.Frame(self.root)
        lang_frame.pack(fill='x', padx=15, pady=(10, 0))

        self.lbl_lang_select = ttk.Label(lang_frame, text="Language / Ngôn ngữ:", font=('Segoe UI', 9, 'bold'))
        self.lbl_lang_select.pack(side='left', padx=(0, 5))

        self.lang_var = tk.StringVar(value=self.current_lang)
        self.lang_menu = ttk.OptionMenu(
            lang_frame, 
            self.lang_var, 
            self.current_lang, 
            "English", 
            "Tiếng Việt", 
            command=self.change_language
        )
        self.lang_menu.pack(side='left')

        # Notebook (Tabs)
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)

        # Tab 1: Export
        self.tab_export = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_export, text="")
        self.setup_export_tab()

        # Tab 2: Merge
        self.tab_merge = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_merge, text="")
        self.setup_merge_tab()

        # Cập nhật văn bản theo ngôn ngữ khởi tạo
        self.update_ui_text()

    def change_language(self, selected_lang):
        self.current_lang = selected_lang
        self.update_ui_text()

    def t(self, key):
        return TRANSLATIONS[self.current_lang].get(key, key)

    def update_ui_text(self):
        self.root.title(self.t("title"))
        self.lbl_lang_select.config(text=self.t("language_label"))
        
        # Tabs
        self.notebook.tab(self.tab_export, text=self.t("tab_export"))
        self.notebook.tab(self.tab_merge, text=self.t("tab_merge"))

        # Export Tab
        self.export_lf.config(text=self.t("export_frame"))
        self.lbl_export_desc.config(text=self.t("export_desc"))
        self.btn_add.config(text=self.t("btn_add_files"))
        self.btn_clear.config(text=self.t("btn_clear_list"))
        self.btn_export.config(text=self.t("btn_start_export"))

        # Merge Tab
        self.merge_lf.config(text=self.t("merge_frame"))
        self.lbl_excel_step.config(text=self.t("lbl_excel_step"))
        self.btn_browse_excel.config(text=self.t("btn_browse_excel"))
        self.lbl_rpy_step.config(text=self.t("lbl_rpy_step"))
        self.btn_browse_folder.config(text=self.t("btn_browse_folder"))
        self.btn_merge.config(text=self.t("btn_start_merge"))

    # ----------------------------------------------------
    # TAB 1: EXPORT (RPY -> EXCEL)
    # ----------------------------------------------------
    def setup_export_tab(self):
        self.export_lf = ttk.LabelFrame(self.tab_export, text="")
        self.export_lf.pack(fill='both', expand=True, padx=15, pady=15)

        self.lbl_export_desc = ttk.Label(self.export_lf, text="", font=('Segoe UI', 10))
        self.lbl_export_desc.pack(anchor='w', padx=10, pady=5)

        list_frame = ttk.Frame(self.export_lf)
        list_frame.pack(fill='both', expand=True, padx=10, pady=5)

        self.export_listbox = tk.Listbox(list_frame, selectmode=tk.EXTENDED, height=8)
        self.export_listbox.pack(side='left', fill='both', expand=True)

        scrollbar = ttk.Scrollbar(list_frame, orient='vertical', command=self.export_listbox.yview)
        scrollbar.pack(side='right', fill='y')
        self.export_listbox.config(yscrollcommand=scrollbar.set)

        btn_box = ttk.Frame(self.export_lf)
        btn_box.pack(fill='x', padx=10, pady=5)

        self.btn_add = ttk.Button(btn_box, text="", command=self.add_rpy_files)
        self.btn_add.pack(side='left', padx=5)

        self.btn_clear = ttk.Button(btn_box, text="", command=lambda: self.export_listbox.delete(0, tk.END))
        self.btn_clear.pack(side='left', padx=5)

        out_frame = ttk.Frame(self.export_lf)
        out_frame.pack(fill='x', padx=10, pady=10)

        self.btn_export = ttk.Button(out_frame, text="", command=self.process_export)
        self.btn_export.pack(fill='x', ipady=5)

    def add_rpy_files(self):
        files = filedialog.askopenfilenames(
            title=self.t("dlg_select_rpy"),
            filetypes=[("Ren'Py Script", "*.rpy"), ("All Files", "*.*")]
        )
        for f in files:
            if f not in self.export_listbox.get(0, tk.END):
                self.export_listbox.insert(tk.END, f)

    def process_export(self):
        rpy_files = list(self.export_listbox.get(0, tk.END))
        if not rpy_files:
            messagebox.showwarning(self.t("warning"), self.t("msg_warn_select_rpy"))
            return

        save_path = filedialog.asksaveasfilename(
            title=self.t("dlg_save_excel"),
            defaultextension=".xlsx",
            initialfile="Translation.xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if not save_path:
            return

        try:
            with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                # Sheet Notes
                df_notes = pd.DataFrame([self.t("excel_notes_sheet")])
                df_notes.to_excel(writer, sheet_name="Notes", index=False)

                for file_path in rpy_files:
                    file_name = os.path.basename(file_path)
                    if not os.path.exists(file_path):
                        continue
                        
                    with open(file_path, "r", encoding="utf-8") as f:
                        lines = f.readlines()
                        
                    current_game_location = ""
                    current_label = ""
                    extracted_data = []
                        
                    for idx, line in enumerate(lines, start=1):
                        line_strip = line.strip()
                        
                        if line_strip.startswith("# game/") and ":" in line_strip:
                            current_game_location = line_strip.replace("#", "").strip()
                            
                        if line_strip.startswith("translate ") and line_strip.endswith(":"):
                            parts = line_strip.split()
                            if len(parts) >= 3:
                                current_label = parts[2].replace(":", "").strip()
                        
                        # dialogue lines
                        if line_strip.startswith("#") and '"' in line_strip:
                            start_pos = line.find('"')
                            end_pos = line.rfind('"')
                            
                            if start_pos != end_pos:
                                english_text = line[start_pos + 1 : end_pos]
                                
                                if english_text.startswith("game/") or english_text.endswith(".rpy"):
                                    continue
                                
                                character_name = ""
                                translated_text = ""
                                
                                for offset in range(1, 4):
                                    next_idx = (idx - 1) + offset
                                    if next_idx >= len(lines):
                                        break
                                        
                                    next_line_strip = lines[next_idx].strip()
                                    
                                    if next_line_strip.startswith("#") or next_line_strip.startswith("translate"):
                                        break
                                    
                                    if next_line_strip.endswith('""'):
                                        possible_name = next_line_strip[:-2].strip()
                                        if possible_name:
                                            character_name = possible_name
                                    
                                    elif '"' in next_line_strip and not next_line_strip.startswith("#"):
                                        t_start = next_line_strip.find('"')
                                        t_end = next_line_strip.rfind('"')
                                        if t_start != t_end:
                                            translated_text = next_line_strip[t_start + 1 : t_end]
                                            if not character_name:
                                                tag = next_line_strip[:t_start].strip()
                                                if '"' not in tag:
                                                    character_name = tag
                                            break
                                
                                if not character_name:
                                    tag = line_strip[1:start_pos].strip()
                                    if '"' in tag or tag.startswith("game/"):
                                        character_name = ""
                                    else:
                                        character_name = tag

                                extracted_data.append([file_name, idx, current_game_location, current_label, character_name, english_text, translated_text, ""])

                        # old new
                        elif line_strip.startswith("old") and '"' in line_strip:
                            start_pos = line.find('"')
                            end_pos = line.rfind('"')
                            
                            if start_pos != end_pos:
                                english_text = line[start_pos + 1 : end_pos]
                                translated_text = ""
                                
                                for offset in range(1, 3):
                                    next_idx = (idx - 1) + offset
                                    if next_idx >= len(lines):
                                        break
                                    
                                    next_line_strip = lines[next_idx].strip()
                                    if next_line_strip.startswith("new") and '"' in next_line_strip:
                                        t_start = next_line_strip.find('"')
                                        t_end = next_line_strip.rfind('"')
                                        if t_start != t_end:
                                            translated_text = next_line_strip[t_start + 1 : t_end]
                                            break

                                extracted_data.append([file_name, idx, current_game_location, current_label, "", english_text, translated_text, ""])

                    if extracted_data:
                        columns = ["File", "Line", "Original Location", "Label", "Tag Char", "Original", "Translated", "Notes"]
                        df = pd.DataFrame(extracted_data, columns=columns)
                        sheet_name = file_name.replace(".rpy", "")
                        sheet_name = sheet_name[:30]
                        df.to_excel(writer, sheet_name=sheet_name, index=False)

                columns = ["File", "Line", "Original Location", "Label", "Tag Char", "Original", "Translated", "Notes"]
                df_bo_sung = pd.DataFrame(columns=columns)
                df_bo_sung.to_excel(writer, sheet_name="Addition", index=False)

                workbook = writer.book
                col_widths = {
                    "A": 12, "B": 10, "C": 22, "D": 22,
                    "E": 10, "F": 45, "G": 45, "H": 25
                }

                for sheetname in workbook.sheetnames:
                    ws = workbook[sheetname]
                    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
                        for cell in row:
                            cell.alignment = Alignment(wrap_text=True, vertical="top")
                    for col_letter, width in col_widths.items():
                        ws.column_dimensions[col_letter].width = width
                    if sheetname == "Notes":
                        continue
                    max_row = max(ws.max_row, 2)
                    ws.auto_filter.ref = f"A1:H{max_row}"

            messagebox.showinfo(self.t("success"), self.t("msg_success_export").format(save_path))

        except Exception as e:
            messagebox.showerror(self.t("error"), self.t("msg_err_export").format(str(e)))

    # ----------------------------------------------------
    # TAB 2: MERGE (EXCEL -> RPY)
    # ----------------------------------------------------
    def setup_merge_tab(self):
        self.merge_lf = ttk.LabelFrame(self.tab_merge, text="")
        self.merge_lf.pack(fill='both', expand=True, padx=15, pady=15)

        self.lbl_excel_step = ttk.Label(self.merge_lf, text="", font=('Segoe UI', 10, 'bold'))
        self.lbl_excel_step.pack(anchor='w', padx=10, pady=(10, 2))

        excel_box = ttk.Frame(self.merge_lf)
        excel_box.pack(fill='x', padx=10, pady=2)

        self.entry_excel = ttk.Entry(excel_box)
        self.entry_excel.pack(side='left', fill='x', expand=True, padx=(0, 5))

        self.btn_browse_excel = ttk.Button(excel_box, text="", command=self.browse_excel_file)
        self.btn_browse_excel.pack(side='right')

        self.lbl_rpy_step = ttk.Label(self.merge_lf, text="", font=('Segoe UI', 10, 'bold'))
        self.lbl_rpy_step.pack(anchor='w', padx=10, pady=(15, 2))

        rpy_box = ttk.Frame(self.merge_lf)
        rpy_box.pack(fill='x', padx=10, pady=2)

        self.entry_rpy_folder = ttk.Entry(rpy_box)
        self.entry_rpy_folder.pack(side='left', fill='x', expand=True, padx=(0, 5))

        self.btn_browse_folder = ttk.Button(rpy_box, text="", command=self.browse_rpy_folder)
        self.btn_browse_folder.pack(side='right')

        self.btn_merge = ttk.Button(self.merge_lf, text="", command=self.process_merge)
        self.btn_merge.pack(fill='x', padx=10, pady=25, ipady=5)

    def browse_excel_file(self):
        f = filedialog.askopenfilename(
            title=self.t("dlg_select_excel"),
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if f:
            self.entry_excel.delete(0, tk.END)
            self.entry_excel.insert(0, f)

    def browse_rpy_folder(self):
        d = filedialog.askdirectory(title=self.t("dlg_select_folder"))
        if d:
            self.entry_rpy_folder.delete(0, tk.END)
            self.entry_rpy_folder.insert(0, d)

    def process_merge(self):
        excel_file = self.entry_excel.get().strip()
        rpy_dir = self.entry_rpy_folder.get().strip()

        if not excel_file or not os.path.exists(excel_file):
            messagebox.showwarning(self.t("warning"), self.t("msg_warn_excel"))
            return

        if not rpy_dir or not os.path.isdir(rpy_dir):
            messagebox.showwarning(self.t("warning"), self.t("msg_warn_rpy_folder"))
            return

        try:
            excel_sheets = pd.read_excel(excel_file, sheet_name=None, engine="openpyxl")
            translations = {}
            ignored_sheets = {"Notes"}

            for sheet_name, df in excel_sheets.items():
                if sheet_name in ignored_sheets:
                    continue

                df.columns = [str(col).strip() for col in df.columns]

                for _, row in df.iterrows():
                    file_name = str(row.get("File", "")).strip()
                    if not file_name or file_name == "nan":
                        if sheet_name not in {"Addition", "Bổ sung"}:
                            file_name = f"{sheet_name}.rpy" if not sheet_name.endswith(".rpy") else sheet_name
                        else:
                            continue

                    line_val = row.get("Line", row.get("Dòng số", ""))
                    if pd.isna(line_val):
                        continue
                    try:
                        line_no = int(float(line_val))
                    except ValueError:
                        continue

                    orig_val = row.get("Original", row.get("Bản gốc", ""))
                    original_text = "" if pd.isna(orig_val) else str(orig_val).strip()

                    trans_val = row.get("Translated", row.get("Bản dịch", ""))
                    translated_text = "" if pd.isna(trans_val) else str(trans_val).strip()

                    if original_text.startswith('"') and original_text.endswith('"') and len(original_text) >= 2:
                        original_text = original_text[1:-1]
                    if translated_text.startswith('"') and translated_text.endswith('"') and len(translated_text) >= 2:
                        translated_text = translated_text[1:-1]

                    translations[(file_name, line_no)] = {
                        "original": original_text,
                        "translated": translated_text
                    }

            all_files = {k[0] for k in translations.keys()}
            updated_count = 0

            for file_name in all_files:
                file_path = os.path.join(rpy_dir, file_name)
                
                if not os.path.exists(file_path):
                    continue

                with open(file_path, "r", encoding="utf-8") as f:
                    lines = f.readlines()

                file_trans = {k[1]: v for k, v in translations.items() if k[0] == file_name}
                if not file_trans:
                    continue

                max_line_excel = max(file_trans.keys())
                while len(lines) < max_line_excel + 5:
                    lines.append("\n")

                new_lines = list(lines)

                for line_num, data in sorted(file_trans.items()):
                    orig = data["original"]
                    trans = data["translated"]
                    
                    if not orig and not trans:
                        continue

                    idx = line_num - 1
                    line_content = new_lines[idx].strip()

                    # Case 1: old new
                    if line_content.startswith("old "):
                        indent = new_lines[idx][:new_lines[idx].find("old")]
                        if orig:
                            new_lines[idx] = f'{indent}old "{orig}"\n'
                        
                        if idx + 1 < len(new_lines) and new_lines[idx + 1].strip().startswith("new "):
                            next_indent = new_lines[idx + 1][:new_lines[idx + 1].find("new")]
                            if trans:
                                new_lines[idx + 1] = f'{next_indent}new "{trans}"\n'
                        else:
                            if trans:
                                new_lines.insert(idx + 1, f'{indent}new "{trans}"\n')

                    # Case 2: #
                    elif line_content.startswith("#"):
                        if orig and '"' in new_lines[idx]:
                            first_q = new_lines[idx].find('"')
                            last_q = new_lines[idx].rfind('"')
                            if first_q != -1 and last_q > first_q:
                                new_lines[idx] = f'{new_lines[idx][:first_q + 1]}{orig}{new_lines[idx][last_q:]}'

                        if trans and idx + 1 < len(new_lines):
                            next_l = new_lines[idx + 1]
                            if '"' in next_l:
                                first_q = next_l.find('"')
                                last_q = next_l.rfind('"')
                                if first_q != -1 and last_q >= first_q:
                                    prefix = next_l[:first_q + 1]
                                    suffix = next_l[last_q:] if last_q > first_q else '"\n'
                                    new_lines[idx + 1] = f'{prefix}{trans}{suffix}'

                    # Case 3: Empty line
                    elif line_content == "" or not ('"' in line_content):
                        indent = "    "
                        if orig and trans:
                            new_lines[idx] = f'{indent}old "{orig}"\n'
                            if idx + 1 < len(new_lines) and new_lines[idx + 1].strip().startswith("new "):
                                new_lines[idx + 1] = f'{indent}new "{trans}"\n'
                            else:
                                new_lines.insert(idx + 1, f'{indent}new "{trans}"\n')
                        elif trans:
                            new_lines[idx] = f'{indent}"{trans}"\n'

                    # Case 4: Normal dialogue line
                    elif '"' in line_content:
                        first_q = new_lines[idx].find('"')
                        last_q = new_lines[idx].rfind('"')
                        if first_q != -1 and last_q >= first_q:
                            prefix = new_lines[idx][:first_q + 1]
                            suffix = new_lines[idx][last_q:] if last_q > first_q else '"\n'
                            if trans:
                                new_lines[idx] = f'{prefix}{trans}{suffix}'

                with open(file_path, "w", encoding="utf-8") as f:
                    f.writelines(new_lines)
                updated_count += 1

            messagebox.showinfo(self.t("success"), self.t("msg_success_merge").format(updated_count))

        except Exception as e:
            messagebox.showerror(self.t("error"), self.t("msg_err_merge").format(str(e)))

if __name__ == "__main__":
    root = tk.Tk()
    app = RenPyTranslatorApp(root)
    root.mainloop()