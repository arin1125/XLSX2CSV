from PyQt5 import QtCore, QtGui, QtWidgets

class MyTableWidget(QtWidgets.QTableWidget):
    def __init__(self, parent=None):
        super().__init__(parent)

    def wheelEvent(self, event):
        if QtWidgets.QApplication.keyboardModifiers() == QtCore.Qt.ControlModifier:
            self.horizontalScrollBar().event(event)
        else:
            super().wheelEvent(event)
