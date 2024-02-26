# .module/file_select_handler.py
from PyQt5.QtWidgets import QFileDialog

def on_file_select_button_clicked(main_window):
    filenames, _ = QFileDialog.getOpenFileNames(main_window, "XLSX 파일 선택", "", "XLSX (*.xlsx)")

    for filename in filenames:
        main_window.file_list.append(filename)
        main_window.file_combo.addItem(filename)
