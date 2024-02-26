# .handler/drag_drop_handler.py
from PyQt5.QtWidgets import QMessageBox

def drag_enter_event(main_window, event):
    if event.mimeData().hasUrls():
        event.accept()
    else:
        event.ignore()

def drop_event(main_window, event):
    files = [u.toLocalFile() for u in event.mimeData().urls()]
    for filename in files:
        if not filename.endswith('.xlsx'):  # 파일 확장자 확인
            QMessageBox.warning(main_window, "에러", "사용할 수 없는 확장자입니다. xlsx파일이 맞는지 확인 해주세요.")
            continue  # 현재 파일은 스킵하고 다음 파일로 넘어감

        main_window.file_list.append(filename)
        main_window.file_combo.addItem(filename)
