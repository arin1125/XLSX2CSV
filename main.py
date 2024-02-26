# main.py
import sys
import os
import zipfile
import pandas as pd
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog, QComboBox


from module.TableWidget import MyTableWidget
from module.convert_handler import on_convert_button_clicked
from module.all_convert_handler import on_all_convert_button_clicked
from module.file_select_handler import on_file_select_button_clicked
from module.delete_handler import on_delete_button_clicked
from module.all_delete_handler import on_all_delete_button_clicked
from module.drag_drop_handler import drag_enter_event, drop_event


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()

        # 윈도우 설정
        self.setWindowTitle("XLSX 변환")
        self.setGeometry(100, 100, 800, 400)

        # Drag & Drop 활성화
        self.setAcceptDrops(True)

        # 위젯 생성
        self.central_widget = QtWidgets.QWidget()
        self.setCentralWidget(self.central_widget)

        # 코드 모듈화 진행
        self.file_select_button = QtWidgets.QPushButton("파일 선택")
        self.file_select_button.clicked.connect(lambda: on_file_select_button_clicked)

        self.convert_button = QtWidgets.QPushButton("변환")
        self.convert_button.clicked.connect(lambda: on_convert_button_clicked)

        self.all_convert_button = QtWidgets.QPushButton("모든파일 변환")
        self.all_convert_button.clicked.connect(lambda: on_all_convert_button_clicked)

        self.delete_button = QtWidgets.QPushButton("삭제")
        self.delete_button.clicked.connect(lambda: on_delete_button_clicked)

        self.all_delete_button = QtWidgets.QPushButton("모든파일 삭제")
        self.all_delete_button.clicked.connect(lambda: on_all_delete_button_clicked)

        self.label = QtWidgets.QLabel()

        self.file_list = []
        self.file_combo = QComboBox()
        self.file_combo.currentIndexChanged.connect(self.on_file_combo_currentIndexChanged)

        # 레이아웃 설정
        self.layout = QtWidgets.QVBoxLayout()
        self.layout.addWidget(self.file_select_button)
        self.layout.addSpacing(5)
        self.layout.addWidget(self.convert_button)
        self.layout.addWidget(self.all_convert_button)
        self.layout.addSpacing(14)
        self.layout.addWidget(self.delete_button)
        self.layout.addWidget(self.all_delete_button)
        self.layout.addSpacing(14)
        self.layout.addWidget(self.file_combo)
        self.layout.addWidget(self.label)
        self.central_widget.setLayout(self.layout)

        # 선택된 열 추적
        self.selected_columns = []
        self.selected_row = None

    def dragEnterEvent(self, event):
        drag_enter_event(self, event)

    def dropEvent(self, event):
        drop_event(self, event)

    def on_file_combo_currentIndexChanged(self):
        self.selected_columns = []  # 선택된 열 기록 초기화

        selected_file = self.file_list[self.file_combo.currentIndex()]
        data = pd.read_excel(selected_file)

        if hasattr(self, 'header_table'):
            self.header_table.deleteLater()

        self.header_table = MyTableWidget()
        self.header_table.setRowCount(1)
        self.header_table.setColumnCount(len(data.columns))
        self.header_table.setSelectionMode(QtWidgets.QAbstractItemView.MultiSelection)
        self.header_table.itemClicked.connect(self.on_item_clicked)  # 아이템 클릭 이벤트 연결

        for i, col_name in enumerate(data.columns):
            item = QtWidgets.QTableWidgetItem(str(col_name))
            item.setFlags(QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)  # 헤더 수정 불가능
            self.header_table.setItem(0, i, item)

        self.header_table.horizontalHeader().setStretchLastSection(True)
        self.header_table.verticalHeader().setVisible(False)
        self.connect_header_click_event()  # 헤더 클릭 이벤트 연결

        self.layout.addWidget(self.header_table)

    def connect_header_click_event(self):
        self.header_table.horizontalHeader().sectionClicked.connect(self.on_header_clicked)

    def disconnect_header_click_event(self):
        self.header_table.horizontalHeader().sectionClicked.disconnect(self.on_header_clicked)

    def on_header_clicked(self, logical_index):
        if logical_index < self.header_table.columnCount():  # 테이블 데이터에 속한 열인지 확인
            self.disconnect_header_click_event()  # 헤더 클릭 이벤트 해제

            # 같은 열을 다시 클릭한 경우
            if logical_index in self.selected_columns:
                self.selected_columns.remove(logical_index)  # 선택된 열 목록에서 제거
            else:  # 새로운 열을 클릭한 경우
                self.selected_columns.append(logical_index)  # 선택된 열 목록에 추가

            self.connect_header_click_event()  # 헤더 클릭 이벤트 재연결


# 헤더만 선택 열에 반응하고 1행은 반응을 하지않아서 1행도 똑같이 반응할 수 있도록 추가함
    def on_item_clicked(self, item):  
        if item.column() in self.selected_columns:  # 같은 열을 다시 클릭한 경우
            self.selected_columns.remove(item.column())  # 선택된 열 목록에서 제거
        else:  # 새로운 열을 클릭 경우
            self.selected_columns.append(item.column())  # 선택된 열 목록에 추가

        self.selected_row = item.row()  # 선택된 행 기록

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
