from PyQt5 import QtCore, QtGui, QtWidgets
from DEMOGUI import Ui_Dialog
import main_program as mprog
import drafting_part as dp
import drafting_part_calculate as dpc
import file_path as fp
import parameter as par
import engineering_drawing as eng
import Assembly_diagram as Assdig
import sys
import datetime
import os
import time




class main(QtWidgets.QWidget, Ui_Dialog):
    def __init__(self):
        super(main, self).__init__()
        self.ui = Ui_Dialog()
        self.ui.setupUi(self)
        self.ui.pushButton.clicked.connect(self.start)


    def start(self):
        type = str(self.ui.comboBox_4.currentText())
        alpha = str(self.ui.lineEdit.text())
        beta = str(self.ui.lineEdit_2.text())
        zeta = str(self.ui.lineEdit_3.text())
        delta = str(self.ui.lineEdit_4.text())
        print(type, alpha, beta, zeta, delta)
        self.create_dir(type)
        if alpha == "":
            self.aplha = 0
        else:
            self.aplha = int(alpha)
        if beta == "":
            self.beta = 0
        else:
            self.beta = int(beta)
        if zeta == "":
            self.zeta = 0
        else:
            self.zeta = int(zeta)
        if delta == "":
            self.delta = 0
        else:
            self.delta = int(delta)
        self.i= self.choos(type)
        self.change_dir(self.i,self.aplha, self.beta, self.zeta, self.delta, self.part_path)

    def choos(self, type):
        # 確認型號"輸入型號"
        if type == "SN1-25" or type == "sn1-25" or type == "25":
            i = 0
        elif type == "SN1-35" or type == "sn1-35" or type == "35":
            i = 1
        elif type == "SN1-45" or type == "sn1-45" or type == "45":
            i = 2
        elif type == "SN1-60" or type == "sn1-60" or type == "60":
            i = 3
        elif type == "SN1-80" or type == "sn1-80" or type == "80":
            i = 4
        elif type == "SN1-110" or type == "sn1-110" or type == "110":
            i = 5
        elif type == "SN1-160" or type == "sn1-160" or type == "160":
            i = 6
        elif type == "SN1-200" or type == "sn1-200" or type == "200":
            i = 7
        elif type == "SN1-250" or type == "sn1-250" or type == "250":
            i = 8
        elif type == "SN1-300" or type == "sn1-300" or type == "300":
            i = 9

        return i

    def create_dir(self, type):  # 創建資料夾
        time_now = datetime.datetime.now()
        dir_name = '{}_{}_{}_{}_{}'.format(type, time_now.day, time_now.hour, time_now.minute, time_now.second)
        desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
        path = desktop + '\\' + dir_name
        os.mkdir(path)
        part_path = path + "\\" + "part"
        os.mkdir(part_path)
        self.path = path
        self.part_path = part_path

    def change_dir(self, i, alpha, beta, gamma, delta, path):
        # 開啟CATIA
        env = mprog.set_CATIA_workbench_env()





if __name__ == "__main__":
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)  # 自適應屏幕分辨率
    app = QtWidgets.QApplication([])
    window = main()
    window.show()
    sys.exit(app.exec_())