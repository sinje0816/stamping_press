from PyQt5 import QtCore, QtGui, QtWidgets
from GUI import Ui_Dialog
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

change = ()
height = ()
type = ()
hole = ()
close = ()

class main(QtWidgets.QWidget, Ui_Dialog):
    def __init__(self):
        super(main, self).__init__()
        self.ui = Ui_Dialog()
        self.ui.setupUi(self)
        self.ui.pushButton.clicked.connect(self.start)
        self.add_item_for_comboBox()
        self.path = str()
        self.part_path = str()
        BASE_DIR = os.path.dirname(os.path.realpath(__file__))
        self.setWindowIcon(QtGui.QIcon(BASE_DIR + '\\ico.ico'))
        self.a = int()
        self.ui.comboBox_4.currentIndexChanged.connect(self.change_combobox4)
        self.ui.lineEdit.textChanged.connect(self.boundary_value)

    def insert_data_combobox_change(self):
        data = {250: {230: [720, 1058], 200: [1080, 1587]},
                350: {250: [830, 1125], 220: [1245, 1688]},
                450: {270: [890, 1210], 240: [1335, 1815]},
                600: {300: [940, 1315], 270: [1410, 1973]},
                800: {330: [1050, 1480], 300: [1575, 2220]},
                1100: {350: [1160, 1680], 320: [1740, 2520]},
                1600: {400: [1300, 1985], 360: [1950, 2978]},
                2000: {450: [1480, 2113], 400: [2220, 3170]},
                2500: {450: [1560, 2400], 400: [2340, 3600]},
                3000: {500: [1760, 2700], 450: [2640, 4050]}
                }
        return data

    def boundary_value(self):
        Form = QtWidgets.QWidget()
        Form.setWindowTitle('oxxo.studio')
        Form.resize(500, 200)
        mbox = QtWidgets.QMessageBox(Form)
        try:
            if int(self.ui.lineEdit.text()) > 10 or int(self.ui.lineEdit.text()) < -10:
                mbox.warning(Form,'warning' , '超出界限')
                self.ui.lineEdit.clear()
        except:
            alpha = str(self.ui.lineEdit.text())

    def change_combobox4(self):
        print(self.ui.comboBox_4.currentText())
        all_data = self.insert_data_combobox_change()
        type = str(self.ui.comboBox_4.currentText())
        ton = int(type.split('-')[-1] + '0')
        close_h = list(all_data[ton].keys())
        table_size = [str(all_data[ton][close_h[0]][0]) + 'x' + str(all_data[ton][close_h[0]][1]),
                      str(all_data[ton][close_h[1]][0]) + 'x' + str(all_data[ton][close_h[1]][1])]
        close_h = list(map(str, close_h))

        self.ui.comboBox_2.clear()
        self.ui.comboBox_2.addItems(table_size)
        self.ui.comboBox_3.clear()
        self.ui.comboBox_3.addItems(close_h)

    def start(self):
        l = str(self.ui.comboBox_2.currentText())
        height = str(self.ui.comboBox_3.currentText())
        type = str(self.ui.comboBox_4.currentText())
        hole = str(self.ui.comboBox_5.currentText())
        alpha = str(self.ui.lineEdit.text())
        print(type, l, height, hole , alpha)
        self.create_dir(type)
        if alpha == "":
            self.aplha = 0
        else:
            self.aplha = int(alpha)
        self.l, self.h, self.i, self.j = self.choos(l, height, type, hole)
        self.change_dir(self.l, self.h, self.i, self.j, self.aplha, self.part_path)
        self.ass_(self.l, self.h, self.i, self.j, self.path, self.part_path, self.aplha)

    def choos(self, l, height, type, hole):
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

        # 合模高
        if type == "SN1-25" and height == '230':
            h = 0
        elif type == "SN1-25" and height == '200':
            h = 1
        elif type == "SN1-35" and height == '250':
            h = 0
        elif type == "SN1-35" and height == '220':
            h = 1
        elif type == "SN1-45" and height == '270':
            h = 0
        elif type == "SN1-45" and height == '240':
            h = 1
        elif type == "SN1-60" and height == '300':
            h = 0
        elif type == "SN1-60" and height == '270':
            h = 1
        elif type == "SN1-80" and height == '330':
            h = 0
        elif type == "SN1-80" and height == '300':
            h = 1
        elif type == "SN1-110" and height == '350':
            h = 0
        elif type == "SN1-110" and height == '320':
            h = 1
        elif type == "SN1-160" and height == '400':
            h = 0
        elif type == "SN1-160" and height == '360':
            h = 1
        elif type == "SN1-200" and height == '450':
            h = 0
        elif type == "SN1-200" and height == '400':
            h = 1
        elif type == "SN1-250" and height == '450':
            h = 0
        elif type == "SN1-250" and height == '400':
            h = 1
        elif type == "SN1-300" and height == '500':
            h = 0
        elif type == "SN1-300" and height == '450':
            h = 1

        # 判斷長寬
        if l == '720x1058' or l == '830x1125' or l == '890x1210' or l == '940x1315' or l == '1050x1480' or l == '1160x1680' or l == '1300x1985' or l == '1480x2113' or l == '1560x2400' or l == '1760x2700':
            l = 1
        elif l == '1080x1587' or l == '1245x1688' or l == '1335x1815' or l == '1410x1973' or l == '1575x2220' or l == '1740x2520' or l == '1950x2978' or l == '2220x3170' or l == '2340x3600' or l == '2640x4050':
            l = 0

        # 輸入平板型號
        if hole == "圓型平板":
            j = 0
        elif hole == "方型平板":
            j = 1
        elif hole == "模墊型平板":
            j = 2

        return h, i, l, j


    def add_item_for_comboBox(self):
        print('insert')
        # data = ['sd1', 'sd2']
        # for item in data:
        #     self.ui.comboBox.addItem(item)

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

    # 針對母檔零件進行變數變換
    def change_dir(self, h, i, l, j, alpha, path):
        print(alpha)
        # 開啟CATIA
        env = mprog.set_CATIA_workbench_env()
        # 開啟零件檔更改變數後儲存並關閉
        for name in par.file_name_list:
            mprog.import_part(fp.system_root + fp.part, name)

    def ass_(self, h, i, l, j, path, part_path, alpha):
        print(l, h, i)
        type = str(self.ui.comboBox_4.currentText())  # 沖床噸數類型

        # ----------生成沖床組合圖---------
        Assdig.assembly_create(l, type, i, part_path, alpha, h)
        # ---------- 生成爆炸圖--------
        eng.explosion_diagram(l, type, i, h)
        eng.balloons(i, l, h)
        # 存檔
        mprog.PDF_save(path, "Exploded_Views")
        mprog.save_detail_drawing(path , "Exploded_Views")

        # -----------零件圖生成--------
        mprog.OPEN_detail_drawing()#開啟零件圖
        dp.Parts_drawing_generation(i, l, part_path, alpha)
        dpc.bom_text_create()
        mprog.PDF_save(path, "detail_drawing")
        mprog.save_detail_drawing(path , "detail_drawing")

        # -------------焊接圖----------
        eng.welding_drawing(l, type, i, alpha)
        # 存檔關閉
        mprog.PDF_save(path, "Welding_diagram")
        mprog.save_detail_drawing(path, "Welding_diagram")

        #恢復隱藏零配件
        for part_name in par.hide_part_name:
            mprog.hide_show_part(part_name, 0)

if __name__ == "__main__":
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)  # 自適應屏幕分辨率
    app = QtWidgets.QApplication([])
    window = main()
    window.show()
    sys.exit(app.exec_())