import sys
import os
import zipfile
import pandas as pd
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog, QComboBox


from module.TableWidget import MyTableWidget

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

        self.file_select_button = QtWidgets.QPushButton("파일 선택")
        self.file_select_button.clicked.connect(self.on_file_select_button_clicked)

        #추출 및 삭제 버튼 추가
        self.convert_button = QtWidgets.QPushButton("변환")
        self.convert_button.clicked.connect(self.on_convert_button_clicked)

        self.all_convert_button = QtWidgets.QPushButton("모든파일 변환")
        self.all_convert_button.clicked.connect(self.on_all_convert_button_clicked)

        self.delete_button = QtWidgets.QPushButton("삭제")
        self.delete_button.clicked.connect(self.on_delete_button_clicked)

        self.all_delete_button = QtWidgets.QPushButton("모든파일 삭제")
        self.all_delete_button.clicked.connect(self.on_all_delete_button_clicked)

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
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        for filename in files:
            if not filename.endswith('.xlsx'):  # 파일 확장자 확인
                QtWidgets.QMessageBox.warning(self, "에러", "사용할 수 없는 확장자입니다. xlsx파일이 맞는지 확인 해주세요.")
                continue  # 현재 파일은 스킵하고 다음 파일로 넘어감

            self.file_list.append(filename)
            self.file_combo.addItem(filename)

    def on_file_select_button_clicked(self):
        filenames, _ = QFileDialog.getOpenFileNames(self, "XLSX 파일 선택", "", "XLSX (*.xlsx)")

        for filename in filenames:
            self.file_list.append(filename)
            self.file_combo.addItem(filename)

    def on_convert_button_clicked(self):
        if not self.selected_columns:
            QtWidgets.QMessageBox.warning(self, "에러", "변환할 열을 선택해주세요.")
            return

        selected_file = self.file_list[self.file_combo.currentIndex()]
        data = pd.read_excel(selected_file)

        selected_column_names = [data.columns[index] for index in self.selected_columns]
        selected_data = data[selected_column_names]

        # xlsx 파일 이름에서 확장자를 제거하고 기본 파일 이름으로 사용
        base_filename = os.path.splitext(os.path.basename(selected_file))[0]
        csv_filename, _ = QFileDialog.getSaveFileName(self, "저장", base_filename, "CSV (*.csv)")
        if csv_filename:  # 사용자가 "저장" 버튼을 눌렀을 때
            selected_data.to_csv(csv_filename, index=False, encoding='utf-8-sig')  # encoding 파라미터를 'utf-8-sig'로 설정
            QtWidgets.QMessageBox.information(self, "성공", "CSV 파일로 저장되었습니다.")

    def on_all_convert_button_clicked(self):
        if not self.selected_columns:
            QtWidgets.QMessageBox.warning(self, "에러", "변환할 열을 선택해주세요.")
            return

        zip_filename, _ = QFileDialog.getSaveFileName(self, "저장", "", "ZIP (*.zip)")
        if not zip_filename:  # 사용자가 "취소" 버튼을 눌렀을 때
            return

        with zipfile.ZipFile(zip_filename, 'w') as zipf:
            for file in self.file_list:
                data = pd.read_excel(file)

                selected_column_names = [data.columns[index] for index in self.selected_columns]
                selected_data = data[selected_column_names]

                csv_filename = os.path.splitext(os.path.basename(file))[0] + '.csv'
                selected_data.to_csv(csv_filename, index=False, encoding='utf-8-sig')  # encoding 파라미터를 'utf-8-sig'로 설정

                zipf.write(csv_filename)
                os.remove(csv_filename)

        QtWidgets.QMessageBox.information(self, "성공", "ZIP 파일로 저장되었습니다.")

    def on_delete_button_clicked(self):
        currentIndex = self.file_combo.currentIndex()
        if currentIndex >= 0:  # 선택된 파일이 있는지 확인
            # 현재 인덱스 변경 시그널 일시적으로 끊기
            self.file_combo.currentIndexChanged.disconnect()

            # 선택된 파일을 file_list와 file_combo에서 삭제
            del self.file_list[currentIndex]
            self.file_combo.removeItem(currentIndex)

            # 테이블이 선택된 파일을 표시하고 있었다면, 테이블도 삭제
            if hasattr(self, 'header_table'):
                self.header_table.deleteLater()
                del self.header_table

            # 삭제 후에도 파일이 남아있다면 다음 파일을 선택
            if len(self.file_list) > 0:
                newIndex = 0 if currentIndex == 0 else currentIndex - 1
                self.file_combo.setCurrentIndex(newIndex)
                self.on_file_combo_currentIndexChanged()  # 다음 파일을 테이블에 표시

            # 현재 인덱스 변경 시그널 다시 연결
            self.file_combo.currentIndexChanged.connect(self.on_file_combo_currentIndexChanged)
            
        self.selected_columns = []  # 선택된 열 기록 초기화

    def on_all_delete_button_clicked(self):
        # 현재 인덱스 변경 시그널 일시적으로 끊기
        self.file_combo.currentIndexChanged.disconnect()

        # 모든 파일을 file_list와 file_combo에서 삭제
        self.file_list.clear()
        self.file_combo.clear()

        # 테이블이 파일을 표시하고 있었다면, 테이블도 삭제
        if hasattr(self, 'header_table'):
            self.header_table.deleteLater()
            del self.header_table

        # 현재 인덱스 변경 시그널 다시 연결
        self.file_combo.currentIndexChanged.connect(self.on_file_combo_currentIndexChanged)
        
        self.selected_columns = []  # 선택된 열 기록 초기화

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
