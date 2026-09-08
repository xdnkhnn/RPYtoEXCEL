import os
import sys
import glob
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Font
from openpyxl.worksheet.table import Table, TableStyleInfo

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
        "btn_add_folder": "Add Folder...",
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
        "btn_add_folder": "Thêm Thư Mục...",
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
    },
    "Español": {
        "title": "Herramienta de Traducción Ren'Py",
        "language_label": "Idioma / Language:",
        "tab_export": " 1. Exportar RPY -> Excel ",
        "tab_merge": " 2. Combinar Excel -> RPY ",
        "export_frame": " Extraer archivos de traducción ",
        "export_desc": "Seleccione archivos .rpy para exportar a un único archivo Excel:",
        "btn_add_files": "Añadir archivos RPY...",
        "btn_add_folder": "Añadir carpeta...",
        "btn_clear_list": "Limpiar lista",
        "btn_start_export": " INICIAR EXPORTACIÓN ",
        "merge_frame": " Combinar traducciones en archivos RPY ",
        "lbl_excel_step": "1. Seleccione archivo Excel traducido:",
        "btn_browse_excel": "Buscar Excel...",
        "lbl_rpy_step": "2. Seleccione carpeta con archivos .rpy originales:",
        "btn_browse_folder": "Buscar carpeta...",
        "btn_start_merge": " INICIAR COMBINACIÓN ",
        "excel_notes_sheet": "Traduzca en la columna 'Translated'. Guarde el archivo y use la pestaña 'Combinar Excel -> RPY'.",
        "msg_warn_select_rpy": "¡Seleccione al menos un archivo .rpy!",
        "msg_success_export": "Archivo exportado con éxito en:\n{}",
        "msg_err_export": "Error al exportar el archivo Excel:\n{}",
        "msg_warn_excel": "¡Seleccione un archivo Excel válido!",
        "msg_warn_rpy_folder": "¡Seleccione una carpeta .rpy válida!",
        "msg_success_merge": "¡Traducciones combinadas con éxito en {} archivo(s)!",
        "msg_err_merge": "Error al combinar:\n{}",
        "dlg_select_rpy": "Seleccionar archivos .rpy",
        "dlg_save_excel": "Guardar archivo Excel",
        "dlg_select_excel": "Seleccionar archivo Excel",
        "dlg_select_folder": "Seleccionar carpeta .rpy",
        "warning": "Advertencia",
        "success": "Éxito",
        "error": "Error"
    },
    "中文": {
        "title": "Ren'Py 翻译工具",
        "language_label": "语言 / Language:",
        "tab_export": " 1. 导出 RPY -> Excel ",
        "tab_merge": " 2. 合并 Excel -> RPY ",
        "export_frame": " 提取翻译文件 ",
        "export_desc": "选择 .rpy 文件导出到单个 Excel 文件：",
        "btn_add_files": "添加 RPY 文件...",
        "btn_add_folder": "添加文件夹...",
        "btn_clear_list": "清空列表",
        "btn_start_export": " 开始导出 Excel ",
        "merge_frame": " 将翻译合并到 RPY 文件 ",
        "lbl_excel_step": "1. 选择已翻译的 Excel 文件：",
        "btn_browse_excel": "浏览 Excel...",
        "lbl_rpy_step": "2. 选择包含原始 .rpy 文件的目录：",
        "btn_browse_folder": "浏览文件夹...",
        "btn_start_merge": " 开始合并翻译 ",
        "excel_notes_sheet": "请在 'Translated' 列中提供翻译。完成后保存并使用 '合并 Excel -> RPY' 选项卡。",
        "msg_warn_select_rpy": "请至少选择一个 .rpy 文件！",
        "msg_success_export": "成功导出至：\n{}",
        "msg_err_export": "导出 Excel 文件时出错：\n{}",
        "msg_warn_excel": "请选择有效的 Excel 文件！",
        "msg_warn_rpy_folder": "请选择包含 .rpy 文件的有效目录！",
        "msg_success_merge": "成功合并翻译至 {} 个 .rpy 文件！",
        "msg_err_merge": "合并时出错：\n{}",
        "dlg_select_rpy": "选择 .rpy 文件",
        "dlg_save_excel": "保存输出 Excel 文件",
        "dlg_select_excel": "选择翻译好的 Excel 文件",
        "dlg_select_folder": "选择包含 .rpy 文件的文件夹",
        "warning": "警告",
        "success": "成功",
        "error": "错误"
    },
    "Deutsch": {
        "title": "Ren'Py Übersetzungswerkzeug",
        "language_label": "Sprache / Language:",
        "tab_export": " 1. Exportieren RPY -> Excel ",
        "tab_merge": " 2. Zusammenführen Excel -> RPY ",
        "export_frame": " Übersetzungsdateien extrahieren ",
        "export_desc": "Wählen Sie .rpy-Dateien zum Exportieren aus:",
        "btn_add_files": "RPY-Dateien hinzufügen...",
        "btn_add_folder": "Ordner hinzufügen...",
        "btn_clear_list": "Liste leeren",
        "btn_start_export": " EXPORT STARTEN ",
        "merge_frame": " Übersetzungen in RPY-Dateien zusammenführen ",
        "lbl_excel_step": "1. Übersetzte Excel-Datei auswählen:",
        "btn_browse_excel": "Excel durchsuchen...",
        "lbl_rpy_step": "2. Ordner mit originalen .rpy-Dateien auswählen:",
        "btn_browse_folder": "Ordner durchsuchen...",
        "btn_start_merge": " ZUSAMMENFÜHRUNG STARTEN ",
        "excel_notes_sheet": "Tragen Sie Übersetzungen in die Spalte 'Translated' ein. Speichern und nutzen Sie den Tab 'Zusammenführen'.",
        "msg_warn_select_rpy": "Bitte wählen Sie mindestens eine .rpy-Datei aus!",
        "msg_success_export": "Erfolgreich exportiert nach:\n{}",
        "msg_err_export": "Fehler beim Exportieren der Excel-Datei:\n{}",
        "msg_warn_excel": "Bitte wählen Sie eine gültige Excel-Datei aus!",
        "msg_warn_rpy_folder": "Bitte wählen Sie einen gültigen Ordner aus!",
        "msg_success_merge": "Erfolgreich in {} .rpy-Datei(en) zusammengeführt!",
        "msg_err_merge": "Fehler beim Zusammenführen:\n{}",
        "dlg_select_rpy": ".rpy-Dateien auswählen",
        "dlg_save_excel": "Excel-Datei speichern",
        "dlg_select_excel": "Excel-Datei auswählen",
        "dlg_select_folder": "Ordner auswählen",
        "warning": "Warnung",
        "success": "Erfolg",
        "error": "Fehler"
    },
    "Français": {
        "title": "Outil de Traduction Ren'Py",
        "language_label": "Langue / Language:",
        "tab_export": " 1. Exporter RPY -> Excel ",
        "tab_merge": " 2. Fusionner Excel -> RPY ",
        "export_frame": " Extraire les fichiers de traduction ",
        "export_desc": "Sélectionnez les fichiers .rpy à exporter :",
        "btn_add_files": "Ajouter fichiers RPY...",
        "btn_add_folder": "Ajouter dossier...",
        "btn_clear_list": "Effacer la liste",
        "btn_start_export": " LANCER L'EXPORTATION ",
        "merge_frame": " Fusionner dans les fichiers RPY ",
        "lbl_excel_step": "1. Sélectionner le fichier Excel traduit :",
        "btn_browse_excel": "Parcourir Excel...",
        "lbl_rpy_step": "2. Sélectionner le dossier des fichiers .rpy d'origine :",
        "btn_browse_folder": "Parcourir dossier...",
        "btn_start_merge": " LANCER LA FUSION ",
        "excel_notes_sheet": "Traduisez dans la colonne 'Translated'. Enregistrez et utilisez l'onglet 'Fusionner'.",
        "msg_warn_select_rpy": "Veuillez sélectionner au moins un fichier .rpy !",
        "msg_success_export": "Exporté avec succès vers :\n{}",
        "msg_err_export": "Erreur lors de l'exportation :\n{}",
        "msg_warn_excel": "Veuillez sélectionner un fichier Excel valide !",
        "msg_warn_rpy_folder": "Veuillez sélectionner un dossier valide !",
        "msg_success_merge": "Fusion réussie dans {} fichier(s) .rpy !",
        "msg_err_merge": "Erreur lors de la fusion :\n{}",
        "dlg_select_rpy": "Sélectionner fichiers .rpy",
        "dlg_save_excel": "Enregistrer le fichier Excel",
        "dlg_select_excel": "Sélectionner le fichier Excel",
        "dlg_select_folder": "Sélectionner le dossier",
        "warning": "Avertissement",
        "success": "Succès",
        "error": "Erreur"
    },
    "Bahasa Indonesia": {
        "title": "Alat Penerjemah Ren'Py",
        "language_label": "Bahasa / Language:",
        "tab_export": " 1. Ekspor RPY -> Excel ",
        "tab_merge": " 2. Gabung Excel -> RPY ",
        "export_frame": " Ekstrak Berkas Terjemahan ",
        "export_desc": "Pilih berkas .rpy untuk diekspor ke satu berkas Excel:",
        "btn_add_files": "Tambah Berkas RPY...",
        "btn_add_folder": "Tambah Folder...",
        "btn_clear_list": "Bersihkan Daftar",
        "btn_start_export": " MULAI EKSPOR ",
        "merge_frame": " Gabung Terjemahan ke Berkas RPY ",
        "lbl_excel_step": "1. Pilih berkas Excel yang sudah diterjemahkan:",
        "btn_browse_excel": "Cari Excel...",
        "lbl_rpy_step": "2. Pilih direktori berisi berkas .rpy asli:",
        "btn_browse_folder": "Cari Folder...",
        "btn_start_merge": " MULAI PENGGABUNGAN ",
        "excel_notes_sheet": "Isi terjemahan di kolom 'Translated'. Simpan dan gunakan tab 'Gabung'.",
        "msg_warn_select_rpy": "Pilih setidaknya satu berkas .rpy!",
        "msg_success_export": "Berhasil diekspor ke:\n{}",
        "msg_err_export": "Galat saat mengekspor berkas Excel:\n{}",
        "msg_warn_excel": "Pilih berkas Excel yang valid!",
        "msg_warn_rpy_folder": "Pilih folder yang valid!",
        "msg_success_merge": "Berhasil menggabungkan ke {} berkas .rpy!",
        "msg_err_merge": "Galat saat menggabungkan:\n{}",
        "dlg_select_rpy": "Pilih Berkas .rpy",
        "dlg_save_excel": "Simpan Berkas Excel",
        "dlg_select_excel": "Pilih Berkas Excel",
        "dlg_select_folder": "Pilih Folder",
        "warning": "Peringatan",
        "success": "Berhasil",
        "error": "Galat"
    },
    "العربية": {
        "title": "أداة ترجمة Ren'Py",
        "language_label": "اللغة / Language:",
        "tab_export": " 1. تصدير RPY -> Excel ",
        "tab_merge": " 2. دمج Excel -> RPY ",
        "export_frame": " استخراج ملفات الترجمة ",
        "export_desc": "حدد ملفات .rpy للتصدير إلى ملف Excel واحد:",
        "btn_add_files": "إضافة ملفات RPY...",
        "btn_add_folder": "إضافة مجلد...",
        "btn_clear_list": "مسح القائمة",
        "btn_start_export": " بدء التصدير ",
        "merge_frame": " دمج الترجمات في ملفات RPY ",
        "lbl_excel_step": "1. حدد ملف Excel المترجم:",
        "btn_browse_excel": "تصفح Excel...",
        "lbl_rpy_step": "2. حدد المجلد الذي يحتوي على ملفات .rpy الأصلية:",
        "btn_browse_folder": "تصفح المجلد...",
        "btn_start_merge": " بدء الدمج ",
        "excel_notes_sheet": "يرجى توفير الترجمات في عمود 'Translated'. واحفظ الملف ثم استخدم علامة التبويب الدمج.",
        "msg_warn_select_rpy": "يرجى تحديد ملف .rpy واحد على الأقل!",
        "msg_success_export": "تم التصدير بنجاح إلى:\n{}",
        "msg_err_export": "حدث خطأ أثناء تصدير Excel:\n{}",
        "msg_warn_excel": "يرجى تحديد ملف Excel صالحة!",
        "msg_warn_rpy_folder": "يرجى تحديد مجلد صالح!",
        "msg_success_merge": "تم الدمج بنجاح في {} ملف(ملفات) .rpy!",
        "msg_err_merge": "حدث خطأ أثناء الدمج:\n{}",
        "dlg_select_rpy": "حدد ملفات .rpy",
        "dlg_save_excel": "حفظ ملف Excel",
        "dlg_select_excel": "حدد ملف Excel",
        "dlg_select_folder": "حدد المجلد",
        "warning": "تحذير",
        "success": "نجاح",
        "error": "خطأ"
    },
    "Türkçe": {
        "title": "Ren'Py Çeviri Aracı",
        "language_label": "Dil / Language:",
        "tab_export": " 1. Dışa Aktar RPY -> Excel ",
        "tab_merge": " 2. Birleştir Excel -> RPY ",
        "export_frame": " Çeviri Dosyalarını Ayıkla ",
        "export_desc": "Tek bir Excel dosyasına aktarmak için .rpy dosyalarını seçin:",
        "btn_add_files": "RPY Dosyası Ekle...",
        "btn_add_folder": "Klasör Ekle...",
        "btn_clear_list": "Listeyi Temizle",
        "btn_start_export": " DIŞA AKTARMAYI BAŞLAT ",
        "merge_frame": " Çevirileri RPY Dosyalarına Birleştir ",
        "lbl_excel_step": "1. Çevrilmiş Excel dosyasını seçin:",
        "btn_browse_excel": "Excel Gözat...",
        "lbl_rpy_step": "2. Orijinal .rpy dosyalarını içeren klasörü seçin:",
        "btn_browse_folder": "Klasör Gözat...",
        "btn_start_merge": " BİRLEŞTİRMEYİ BAŞLAT ",
        "excel_notes_sheet": "Lütfen çevirileri 'Translated' sütununa yazın. Kaydedin ve birleştirme sekmesini kullanın.",
        "msg_warn_select_rpy": "Lütfen en az bir .rpy dosyası seçin!",
        "msg_success_export": "Başarıyla şuraya aktarıldı:\n{}",
        "msg_err_export": "Excel aktarılırken bir hata oluştu:\n{}",
        "msg_warn_excel": "Lütfen geçerli bir Excel dosyası seçin!",
        "msg_warn_rpy_folder": "Lütfen geçerli bir klasör seçin!",
        "msg_success_merge": "Çeviriler {} .rpy dosyasına başarıyla birleştirildi!",
        "msg_err_merge": "Birleştirilirken hata oluştu:\n{}",
        "dlg_select_rpy": ".rpy Dosyalarını Seç",
        "dlg_save_excel": "Excel Dosyasını Kaydet",
        "dlg_select_excel": "Excel Dosyasını Seç",
        "dlg_select_folder": "Klasör Seç",
        "warning": "Uyarı",
        "success": "Başarılı",
        "error": "Hata"
    },
    "Português": {
        "title": "Ferramenta de Tradução Ren'Py",
        "language_label": "Idioma / Language:",
        "tab_export": " 1. Exportar RPY -> Excel ",
        "tab_merge": " 2. Mesclar Excel -> RPY ",
        "export_frame": " Extrair arquivos de tradução ",
        "export_desc": "Selecione arquivos .rpy para exportar para Excel:",
        "btn_add_files": "Adicionar arquivos RPY...",
        "btn_add_folder": "Adicionar pasta...",
        "btn_clear_list": "Limpar lista",
        "btn_start_export": " INICIAR EXPORTAÇÃO ",
        "merge_frame": " Mesclar traduções nos arquivos RPY ",
        "lbl_excel_step": "1. Selecione o arquivo Excel traduzido:",
        "btn_browse_excel": "Procurar Excel...",
        "lbl_rpy_step": "2. Selecione a pasta com os arquivos .rpy originais:",
        "btn_browse_folder": "Procurar pasta...",
        "btn_start_merge": " INICIAR MESCLAGEM ",
        "excel_notes_sheet": "Traduza na coluna 'Translated'. Salve e use a aba de mesclagem.",
        "msg_warn_select_rpy": "Selecione pelo menos um arquivo .rpy!",
        "msg_success_export": "Exportado com sucesso para:\n{}",
        "msg_err_export": "Erro ao exportar arquivo Excel:\n{}",
        "msg_warn_excel": "Selecione um arquivo Excel válido!",
        "msg_warn_rpy_folder": "Selecione uma pasta válida!",
        "msg_success_merge": "Traduções mescladas com sucesso em {} arquivo(s) .rpy!",
        "msg_err_merge": "Erro ao mesclar:\n{}",
        "dlg_select_rpy": "Selecionar arquivos .rpy",
        "dlg_save_excel": "Salvar arquivo Excel",
        "dlg_select_excel": "Selecionar arquivo Excel",
        "dlg_select_folder": "Selecionar pasta",
        "warning": "Aviso",
        "success": "Sucesso",
        "error": "Erro"
    },
    "Polski": {
        "title": "Narzędzie do Tłumaczenia Ren'Py",
        "language_label": "Język / Language:",
        "tab_export": " 1. Eksportuj RPY -> Excel ",
        "tab_merge": " 2. Scal Excel -> RPY ",
        "export_frame": " Wyodrębnij pliki tłumaczeń ",
        "export_desc": "Wybierz pliki .rpy, aby wyeksportować do Excela:",
        "btn_add_files": "Dodaj pliki RPY...",
        "btn_add_folder": "Dodaj folder...",
        "btn_clear_list": "Wyczyść listę",
        "btn_start_export": " ROZPOCZNIJ EKSPORT ",
        "merge_frame": " Scal tłumaczenia z plikami RPY ",
        "lbl_excel_step": "1. Wybierz przetłumaczony plik Excel:",
        "btn_browse_excel": "Przeglądaj Excel...",
        "lbl_rpy_step": "2. Wybierz folder z oryginalnymi plikami .rpy:",
        "btn_browse_folder": "Przeglądaj folder...",
        "btn_start_merge": " ROZPOCZNIJ SCALANIE ",
        "excel_notes_sheet": "Wpisz tłumaczenia w kolumnie 'Translated'. Zapisz i użyj zakładki scalania.",
        "msg_warn_select_rpy": "Wybierz co najmniej jeden plik .rpy!",
        "msg_success_export": "Pomyślnie wyeksportowano do:\n{}",
        "msg_err_export": "Błąd podczas eksportowania pliku Excel:\n{}",
        "msg_warn_excel": "Wybierz prawidłowy plik Excel!",
        "msg_warn_rpy_folder": "Wybierz prawidłowy folder!",
        "msg_success_merge": "Pomyślnie scalono tłumaczenia w {} plikach .rpy!",
        "msg_err_merge": "Błąd podczas scalania:\n{}",
        "dlg_select_rpy": "Wybierz pliki .rpy",
        "dlg_save_excel": "Zapisz plik Excel",
        "dlg_select_excel": "Wybierz plik Excel",
        "dlg_select_folder": "Wybierz folder",
        "warning": "Ostrzeżenie",
        "success": "Sukces",
        "error": "Błąd"
    },
    "Русский": {
        "title": "Инструмент Перевода Ren'Py",
        "language_label": "Язык / Language:",
        "tab_export": " 1. Экспорт RPY -> Excel ",
        "tab_merge": " 2. Слияние Excel -> RPY ",
        "export_frame": " Извлечь файлы перевода ",
        "export_desc": "Выберите файлы .rpy для экспорта в один файл Excel:",
        "btn_add_files": "Добавить файлы RPY...",
        "btn_add_folder": "Добавить папку...",
        "btn_clear_list": "Очистить список",
        "btn_start_export": " НАЧАТЬ ЭКСПОРТ В EXCEL ",
        "merge_frame": " Объединить переводы с файлами RPY ",
        "lbl_excel_step": "1. Выберите переведенный файл Excel:",
        "btn_browse_excel": "Обзор Excel...",
        "lbl_rpy_step": "2. Выберите папку с исходными файлами .rpy:",
        "btn_browse_folder": "Обзор папки...",
        "btn_start_merge": " НАЧАТЬ СЛИЯНИЕ ПЕРЕВОДОВ ",
        "excel_notes_sheet": "Укажите перевод в столбце 'Translated'. Сохраните файл и используйте вкладку слияния.",
        "msg_warn_select_rpy": "Выберите хотя бы один файл .rpy!",
        "msg_success_export": "Успешно экспортировано в:\n{}",
        "msg_err_export": "Ошибка при экспорте файла Excel:\n{}",
        "msg_warn_excel": "Выберите корректный файл Excel!",
        "msg_warn_rpy_folder": "Выберите корректную папку с файлами .rpy!",
        "msg_success_merge": "Успешно объединены переводы в {} файл(ах) .rpy!",
        "msg_err_merge": "Ошибка при объединении:\n{}",
        "dlg_select_rpy": "Выберите файлы .rpy",
        "dlg_save_excel": "Сохранить файл Excel",
        "dlg_select_excel": "Выберите файл Excel",
        "dlg_select_folder": "Выберите папку с файлами .rpy",
        "warning": "Предупреждение",
        "success": "Успех",
        "error": "Ошибка"
    },
    "Українська": {
        "title": "Інструмент Перекладу Ren'Py",
        "language_label": "Мова / Language:",
        "tab_export": " 1. Експорт RPY -> Excel ",
        "tab_merge": " 2. Злиття Excel -> RPY ",
        "export_frame": " Витягти файли перекладу ",
        "export_desc": "Виберіть файли .rpy для експорту в один файл Excel:",
        "btn_add_files": "Додати файли RPY...",
        "btn_add_folder": "Додати папку...",
        "btn_clear_list": "Очистити список",
        "btn_start_export": " ПОЧАТИ ЕКСПОРТ ",
        "merge_frame": " Об'єднати переклади з файлами RPY ",
        "lbl_excel_step": "1. Виберіть перекладений файл Excel:",
        "btn_browse_excel": "Огляд Excel...",
        "lbl_rpy_step": "2. Виберіть папку з оригінальними файлами .rpy:",
        "btn_browse_folder": "Огляд папки...",
        "btn_start_merge": " ПОЧАТИ ЗЛИТТЯ ",
        "excel_notes_sheet": "Вкажіть переклад у стовпчику 'Translated'. Збережіть та використайте вкладку злиття.",
        "msg_warn_select_rpy": "Виберіть принаймні один файл .rpy!",
        "msg_success_export": "Успішно експортовано в:\n{}",
        "msg_err_export": "Помилка під час експорту Excel:\n{}",
        "msg_warn_excel": "Виберіть коректний файл Excel!",
        "msg_warn_rpy_folder": "Виберіть коректну папку з файлами .rpy!",
        "msg_success_merge": "Успішно об'єднано переклади в {} файл(ах) .rpy!",
        "msg_err_merge": "Помилка під час злиття:\n{}",
        "dlg_select_rpy": "Виберіть файли .rpy",
        "dlg_save_excel": "Зберегти файл Excel",
        "dlg_select_excel": "Виберіть файл Excel",
        "dlg_select_folder": "Виберіть папку з файлами .rpy",
        "warning": "Попередження",
        "success": "Успіх",
        "error": "Помилка"
    },
    "日本語": {
        "title": "Ren'Py 翻訳ツール",
        "language_label": "言語 / Language:",
        "tab_export": " 1. 抽出 RPY -> Excel ",
        "tab_merge": " 2. 統合 Excel -> RPY ",
        "export_frame": " 翻訳ファイルの抽出 ",
        "export_desc": ".rpy ファイルを選択して1つの Excel ファイルに出力します:",
        "btn_add_files": "RPY ファイルを追加...",
        "btn_add_folder": "フォルダを追加...",
        "btn_clear_list": "リストをクリア",
        "btn_start_export": " EXCEL 抽出を開始 ",
        "merge_frame": " 翻訳を RPY ファイルに統合 ",
        "lbl_excel_step": "1. 翻訳済み Excel ファイルを選択:",
        "btn_browse_excel": "Excel を参照...",
        "lbl_rpy_step": "2. 元の .rpy ファイルがあるフォルダを選択:",
        "btn_browse_folder": "フォルダを参照...",
        "btn_start_merge": " 翻訳の統合を開始 ",
        "excel_notes_sheet": "'Translated' 列に翻訳を入力してください。保存後、統合タブを使用してください。",
        "msg_warn_select_rpy": "少なくとも1つの .rpy ファイルを選択してください！",
        "msg_success_export": "正常に抽出されました:\n{}",
        "msg_err_export": "Excel の抽出中にエラーが発生しました:\n{}",
        "msg_warn_excel": "有効な Excel ファイルを選択してください！",
        "msg_warn_rpy_folder": "有効なフォルダを選択してください！",
        "msg_success_merge": "{} 個の .rpy ファイルに翻訳を正常に統合しました！",
        "msg_err_merge": "統合中にエラーが発生しました:\n{}",
        "dlg_select_rpy": ".rpy ファイルを選択",
        "dlg_save_excel": "Excel ファイルを保存",
        "dlg_select_excel": "Excel ファイルを選択",
        "dlg_select_folder": "フォルダを選択",
        "warning": "警告",
        "success": "成功",
        "error": "エラー"
    },
    "Italiano": {
        "title": "Strumento di Traduzione Ren'Py",
        "language_label": "Lingua / Language:",
        "tab_export": " 1. Esporta RPY -> Excel ",
        "tab_merge": " 2. Unisci Excel -> RPY ",
        "export_frame": " Estrai file di traduzione ",
        "export_desc": "Seleziona i file .rpy da esportare in un singolo file Excel:",
        "btn_add_files": "Aggiungi file RPY...",
        "btn_add_folder": "Aggiungi cartella...",
        "btn_clear_list": "Svuota lista",
        "btn_start_export": " AVVIA ESPORTAZIONE ",
        "merge_frame": " Unisci traduzioni nei file RPY ",
        "lbl_excel_step": "1. Seleziona il file Excel tradotto:",
        "btn_browse_excel": "Sfoglia Excel...",
        "lbl_rpy_step": "2. Seleziona la cartella contenente i file .rpy originali:",
        "btn_browse_folder": "Sfoglia cartella...",
        "btn_start_merge": " AVVIA UNIONE TRADUZIONI ",
        "excel_notes_sheet": "Inserisci le traduzioni nella colonna 'Translated'. Salva e usa la scheda di unione.",
        "msg_warn_select_rpy": "Seleziona almeno un file .rpy!",
        "msg_success_export": "Esportato con successo in:\n{}",
        "msg_err_export": "Errore durante l'esportazione Excel:\n{}",
        "msg_warn_excel": "Seleziona un file Excel valido!",
        "msg_warn_rpy_folder": "Seleziona una cartella valida!",
        "msg_success_merge": "Traduzioni unite con successo in {} file .rpy!",
        "msg_err_merge": "Errore durante l'unione:\n{}",
        "dlg_select_rpy": "Seleziona file .rpy",
        "dlg_save_excel": "Salva file Excel",
        "dlg_select_excel": "Seleziona file Excel",
        "dlg_select_folder": "Seleziona cartella",
        "warning": "Avviso",
        "success": "Successo",
        "error": "Erro"
    }
}

class RenPyTranslatorApp:
    def __init__(self, root):
        self.root = root
        self.current_lang = "English"
        
        self.root.title("Ren'Py Translation Tool")
        self.root.geometry("680x560")
        self.root.resizable(True, True)

        style = ttk.Style()
        style.theme_use('clam')

        lang_frame = ttk.Frame(self.root)
        lang_frame.pack(fill='x', padx=15, pady=(10, 0))

        self.lbl_lang_select = ttk.Label(lang_frame, text="Language / Ngôn ngữ:", font=('Segoe UI', 9, 'bold'))
        self.lbl_lang_select.pack(side='left', padx=(0, 5))

        self.lang_var = tk.StringVar(value=self.current_lang)
        self.lang_menu = ttk.OptionMenu(
            lang_frame, 
            self.lang_var, 
            self.current_lang, 
            *list(TRANSLATIONS.keys()), 
            command=self.change_language
        )
        self.lang_menu.pack(side='left')

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)

        self.tab_export = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_export, text="")
        self.setup_export_tab()

        self.tab_merge = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_merge, text="")
        self.setup_merge_tab()

        self.update_ui_text()

    def change_language(self, selected_lang):
        self.current_lang = selected_lang
        self.update_ui_text()

    def t(self, key):
        return TRANSLATIONS[self.current_lang].get(key, key)

    def update_ui_text(self):
        self.root.title(self.t("title"))
        self.lbl_lang_select.config(text=self.t("language_label"))
        
        self.notebook.tab(self.tab_export, text=self.t("tab_export"))
        self.notebook.tab(self.tab_merge, text=self.t("tab_merge"))

        self.export_lf.config(text=self.t("export_frame"))
        self.lbl_export_desc.config(text=self.t("export_desc"))
        self.btn_add.config(text=self.t("btn_add_files"))
        self.btn_add_folder.config(text=self.t("btn_add_folder"))
        self.btn_clear.config(text=self.t("btn_clear_list"))
        self.btn_export.config(text=self.t("btn_start_export"))

        self.merge_lf.config(text=self.t("merge_frame"))
        self.lbl_excel_step.config(text=self.t("lbl_excel_step"))
        self.btn_browse_excel.config(text=self.t("btn_browse_excel"))
        self.lbl_rpy_step.config(text=self.t("lbl_rpy_step"))
        self.btn_browse_folder.config(text=self.t("btn_browse_folder"))
        self.btn_merge.config(text=self.t("btn_start_merge"))

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

        self.btn_add_folder = ttk.Button(btn_box, text="", command=self.add_rpy_folder)
        self.btn_add_folder.pack(side='left', padx=5)

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

    def add_rpy_folder(self):
        d = filedialog.askdirectory(title=self.t("dlg_select_folder"))
        if d:
            rpy_files = glob.glob(os.path.join(d, "**", "*.rpy"), recursive=True)
            for f in rpy_files:
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
                df_notes = pd.DataFrame([self.t("excel_notes_sheet")])
                df_notes.to_excel(writer, sheet_name="Notes", index=False)

                table_mapping = []

                for table_idx, file_path in enumerate(rpy_files, start=1):
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
                        sheet_name = file_name.replace(".rpy", "")[:30]
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                        
                        table_name = f"Bảng_{table_idx}"
                        table_mapping.append((sheet_name, table_name))

                columns = ["File", "Line", "Original Location", "Label", "Tag Char", "Original", "Translated", "Notes"]
                df_bo_sung = pd.DataFrame(columns=columns)
                df_bo_sung.to_excel(writer, sheet_name="Addition", index=False)

                workbook = writer.book

                # Style Bảng KHÔNG MÀU (TableStyleLight1)
                for sheet_name, table_name in table_mapping:
                    ws = workbook[sheet_name]
                    max_row = max(ws.max_row, 2)
                    tab = Table(displayName=table_name, ref=f"A1:H{max_row}")
                    style = TableStyleInfo(name="TableStyleLight1", showFirstColumn=False, showLastColumn=False, showRowStripes=False, showColumnStripes=False)
                    tab.tableStyleInfo = style
                    ws.add_table(tab)

                # Sheet Notes & Bảng Progress
                ws_notes = workbook["Notes"]
                
                ws_notes.cell(row=3, column=3, value="File").font = Font(bold=True)
                ws_notes.cell(row=3, column=4, value="Progress").font = Font(bold=True)

                start_row = 4
                for idx, (sheet_name, table_name) in enumerate(table_mapping, start=start_row):
                    ws_notes.cell(row=idx, column=3, value=sheet_name)
                    
                    cell_prog = ws_notes.cell(row=idx, column=4)
                    cell_prog.value = f"=COUNTA({table_name}[Translated])/COUNTA({table_name}[Original])"
                    cell_prog.number_format = '0.00%'

                sum_row = start_row + len(table_mapping)
                ws_notes.cell(row=sum_row, column=3, value="Sum").font = Font(bold=True)
                
                if table_mapping:
                    trans_parts = [f"COUNTA({tname}[Translated])" for _, tname in table_mapping]
                    orig_parts = [f"COUNTA({tname}[Original])" for _, tname in table_mapping]
                    sum_formula = f"=({' + '.join(trans_parts)})/({' + '.join(orig_parts)})"
                    
                    cell_sum = ws_notes.cell(row=sum_row, column=4)
                    cell_sum.value = sum_formula
                    cell_sum.font = Font(bold=True)
                    cell_sum.number_format = '0.00%'

                # Định dạng độ rộng & Alignment đơn giản
                col_widths = {
                    "A": 28,
                    "B": 10,
                    "C": 35,
                    "D": 15,
                    "E": 12,
                    "F": 50,
                    "G": 50,
                    "H": 25
                }

                for sheetname in workbook.sheetnames:
                    if sheetname == "Notes":
                        continue
                        
                    ws = workbook[sheetname]
                    for col_letter, width in col_widths.items():
                        ws.column_dimensions[col_letter].width = width

                    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
                        for cell in row:
                            if cell.column_letter in ['F', 'G']:
                                cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="left")
                            elif cell.column_letter == 'B':
                                cell.alignment = Alignment(vertical="center", horizontal="center")
                            else:
                                cell.alignment = Alignment(vertical="center", horizontal="left")

            messagebox.showinfo(self.t("success"), self.t("msg_success_export").format(save_path))

        except Exception as e:
            messagebox.showerror(self.t("error"), self.t("msg_err_export").format(str(e)))

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
                
                # Tìm đệ quy trong toàn bộ thư mục cháu/chắt
                if not os.path.exists(file_path):
                    matching = glob.glob(os.path.join(rpy_dir, "**", file_name), recursive=True)
                    if matching:
                        file_path = matching[0]
                    else:
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