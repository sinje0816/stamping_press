from seyi_stamping_die_GUI import Ui_Form
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget, QPushButton, QLineEdit, QLabel, \
    QTableWidget, QTableWidgetItem
from PyQt5.QtCore import QObject
import seyi_stamping_die as ssd
import datetime
import os
import sys
import pathlib


class windows(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.ui = Ui_Form()
        self.ui.setupUi(self)
        self.setWindowTitle('模型自動化設計系統')
        dir = os.path.dirname(os.path.realpath(__file__))
        self.ui.picture.setPixmap(QtGui.QPixmap(dir + "\\bead.png"))
        self.ui.reset_value.clicked.connect(self.reset)
        self.ui.escape.clicked.connect(self.out)
        self.ui.setup_value.clicked.connect(self.start)

    def start(self):
        H1 = self.ui.H1.text()
        R1 = self.ui.R1.text()
        R2 = self.ui.R2.text()
        A = self.ui.A.text()
        B = self.ui.B.text()
        C = self.ui.C.text()
        if C == '':
            C = 4
        if B == '':
            B = 4
        print(H1, R1, R2)
        self.create_dir('crank_bead')
        env = ssd.set_CATIA_workbench_env()
        ssd.create_new(H1, R1, R2, A, B, C, self.path)
        ssd.saveas(self.path, 'Product1', '.CATProduct', '.CATProduct')
        ssd.saveas(self.path, 'Product1', '.CATProduct', '.igs')
        # ssd.close_window()
        self.finish()

    def finish(self):
        Form = QtWidgets.QWidget()
        # Form.setWindowTitle('警告')
        Form.resize(400, 300)
        mbox = QtWidgets.QMessageBox(Form)
        mbox.warning(Form, '完成', '已完成，請至桌面查看檔案')

    def reset(self):
        self.ui.H1.setText('')
        self.ui.R1.setText('')
        self.ui.R2.setText('')
        self.ui.A.setText('')
        self.ui.B.setText('')
        self.ui.C.setText('')

    def out(self):
        sys.exit(app.exec_())

    def create_dir(self, type):  # 創建資料夾
        time_now = datetime.datetime.now()
        dir_name = '{}_{}_{}_{}_{}'.format(type, time_now.day, time_now.hour, time_now.minute, time_now.second)
        desktop = str(pathlib.Path.home() / 'Desktop')
        path = desktop + '\\' + dir_name
        print(path)
        os.mkdir(path)
        self.path = path


if __name__ == "__main__":
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)  # 自適應屏幕分辨率
    app = QtWidgets.QApplication([])
    window = windows()
    window.show()
    sys.exit(app.exec_())
