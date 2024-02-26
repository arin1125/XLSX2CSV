# .module/all_convert_handler.py
import os
import zipfile
import pandas as pd
from PyQt5.QtWidgets import QFileDialog, QMessageBox

def on_all_convert_button_clicked(main_window):
    if not main_window.selected_columns:
        QMessageBox.warning(main_window, "에러", "변환할 열을 선택해주세요.")
        return

    zip_filename, _ = QFileDialog.getSaveFileName(main_window, "저장", "", "ZIP (*.zip)")
    if not zip_filename:
        return

    with zipfile.ZipFile(zip_filename, 'w') as zipf:
        for file in main_window.file_list:
            data = pd.read_excel(file)

            selected_column_names = [data.columns[index] for index in main_window.selected_columns]
            selected_data = data[selected_column_names]

            csv_filename = os.path.splitext(os.path.basename(file))[0] + '.csv'
            selected_data.to_csv(csv_filename, index=False, encoding='utf-8-sig')

            zipf.write(csv_filename)
            os.remove(csv_filename)

    QMessageBox.information(main_window, "성공", "ZIP 파일로 저장되었습니다.")
