# .module/convert_handler.py
import os
import pandas as pd
from PyQt5.QtWidgets import QFileDialog, QMessageBox

def on_convert_button_clicked(main_window):
    if not main_window.selected_columns:
        QMessageBox.warning(main_window, "에러", "변환할 열을 선택해주세요.")
        return

    selected_file = main_window.file_list[main_window.file_combo.currentIndex()]
    data = pd.read_excel(selected_file)

    selected_column_names = [data.columns[index] for index in main_window.selected_columns]
    selected_data = data[selected_column_names]

    base_filename = os.path.splitext(os.path.basename(selected_file))[0]
    csv_filename, _ = QFileDialog.getSaveFileName(main_window, "저장", base_filename, "CSV (*.csv)")
    if csv_filename:
        selected_data.to_csv(csv_filename, index=False, encoding='utf-8-sig')
        QMessageBox.information(main_window, "성공", "CSV 파일로 저장되었습니다.")
