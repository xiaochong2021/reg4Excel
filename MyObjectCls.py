from PyQt5.QtCore import *


class MyObjectCls(QObject):
    sigGetRegFilePath = pyqtSignal(str)
    sigSetVueRegFilePath = pyqtSignal(str, str)
    sigInitPage2 = pyqtSignal()
    sigSetVueRegColumns = pyqtSignal(list, list)
    sigStart = pyqtSignal(int, int, str)
    sigInfo = pyqtSignal(str, str)
    sigLoadingTip = pyqtSignal(str)

    def __init__(self, parent=None):
        QObject.__init__(self, parent)

    @pyqtSlot(str)
    def chooseExcel(self, msg):
        self.sigGetRegFilePath.emit(msg)

    @pyqtSlot()
    def initPage2(self):
        self.sigInitPage2.emit()

    @pyqtSlot(int, int, str)
    def start(self, text_column_index, reg_column_index, logic_code):
        self.sigStart.emit(text_column_index, reg_column_index, logic_code)