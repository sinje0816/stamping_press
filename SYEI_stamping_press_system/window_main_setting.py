from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget, QPushButton, QLineEdit, QLabel, \
    QTableWidget, QTableWidgetItem, QHeaderView, QSpinBox, QComboBox, QAbstractItemView
from PyQt5.QtCore import QObject, QTimer, Qt
from PyQt5.QtGui import QColor, QBrush
from window_main import Ui_Form
import sys
import parameter as par

class main(QtWidgets.QWidget, Ui_Form):
    def __init__(self):
        super(main, self).__init__()
        self.ui = Ui_Form()
        self.ui.setupUi(self)
        self.setting()

    def setting(self):
        for x in range(0, 2):
            self.ui.window_main_table.setSpan(x, 0, 1, 5) #以某格為基準向下向左合併儲存格
            newItem = QTableWidgetItem(par.series1[x]) #儲存格內文字
            newItem.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter) #設定置中
            self.ui.window_main_table.setItem(x, 0, newItem) #指定文字放置位置
        for x in range(2, 8):
            self.ui.window_main_table.setSpan(x, 0, 1, 3)
            newItem = QTableWidgetItem(par.series2[x-2])
            newItem.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
            self.ui.window_main_table.setItem(x, 0, newItem)
        for x in range(2, 6):
            combo_box = QComboBox()
            combo_box.addItem()

if __name__ == "__main__":
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)  # 自適應屏幕分辨率
    app = QtWidgets.QApplication([])
    window = main()
    window.show()
    sys.exit(app.exec_())