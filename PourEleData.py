import sys

from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QApplication, QDialog
from QPourEleData import Ui_Dialog

class QPourEleInputForm(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.ui = Ui_Dialog()
        self.ui.setupUi(self)
        self.DamSecEle = {}

    @pyqtSlot()
    def on_DamSecOkButton_clicked(self):
        DamSecEleTem = {}
        RowNum = self.ui.tableWidget.rowCount()
        for num in range(RowNum):
            item = self.ui.tableWidget.item(num, 1)
            key = self.ui.tableWidget.item(num, 0)
            if item:
                DamSecEleTem[key.text()] = item.text()
        for num in range(RowNum):
            item = self.ui.tableWidget.item(num, 3)
            key = self.ui.tableWidget.item(num, 2)
            if item:
                DamSecEleTem[key.text()] = item.text()
        self.DamSecEle = DamSecEleTem
        QPourEleInputForm.close(self)

    def initialize(self):
        DamSecEleTem = {}
        RowNum = self.ui.tableWidget.rowCount()
        for num in range(RowNum):
            item = self.ui.tableWidget.item(num, 1)
            key = self.ui.tableWidget.item(num, 0)
            if item:
                DamSecEleTem[key.text()] = item.text()
        for num in range(RowNum):
            item = self.ui.tableWidget.item(num, 3)
            key = self.ui.tableWidget.item(num, 2)
            if item:
                DamSecEleTem[key.text()] = item.text()
        self.DamSecEle = DamSecEleTem
        return DamSecEleTem


if __name__== "__main__":
    app = QApplication(sys.argv)
    form = QPourEleInputForm()
    form.show()
    sys.exit(app.exec_())