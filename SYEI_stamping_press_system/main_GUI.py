from PyQt5 import QtCore, QtGui, QtWidgets
import sys, datetime, os, math
from GUI import Ui_Dialog
import main_program as mprog
import drafting as draft
import drafting_part as dp
import drafting_part_calculate as dpc
import file_path as fp
import parameter as par

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
        print(type, l, height, hole)
        self.create_dir(type)
        self.l, self.h, self.i, self.j = self.choos(l, height, type, hole)
        self.change_dir(self.l, self.h, self.i, self.j, self.part_path)
        self.ass_(self.l, self.h, self.i, self.j, self.path , self.part_path)

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

    def change_dir(self, h, i, l, j, path):
        # 開啟CATIA
        env = mprog.set_CATIA_workbench_env()

        # 匯入零件檔
        file_name_list = ['BOLSTER1', 'BOLSTER2', 'BOLSTER3', 'Fixture', 'FRAME1', 'FRAME2', 'FRAME3', 'FRAME4',
                          'FRAME5', 'FRAME6', 'FRAME7', 'FRAME8', 'FRAME9', 'FRAME10', 'FRAME11', 'FRAME12', 'FRAME13',
                          'FRAME14', 'FRAME15', 'FRAME16', 'FRAME17', 'FRAME18', 'FRAME19', 'FRAME20', 'FRAME21',
                          'FRAME22', 'FRAME23', 'FRAME24', 'FRAME25', 'FRAME26', 'FRAME27', 'FRAME28', 'FRAME29',
                          'FRAME30', 'FRAME31', 'FRAME32', 'FRAME33', 'FRAME34', 'FRAME35', 'FRAME36', 'FRAME36',
                          'FRAME37', 'FRAME38', 'FRAME39', 'FRAME40', 'FRAME41', 'FRAME42', 'FRAME43', 'FRAME44',
                          'FRAME45', 'FRAME46', 'GIB1', 'GIB2', 'BALANCER_LEFT_All', 'BALANCER_RIGHT_ALL', 'CRANK_SHAFT_CLOCK',
                          'CLUCTH_ASSEMBLY_All', 'SLIDE_UNIT_All', 'CRANK_SHAFT', 'JOINT_All', 'MAIN_GEAR1',
                          'MAIN_GEAR2', 'MAIN_GEAR3', 'MAIN_GEAR4', 'JOINT1', 'FRAME47', 'FRAME48', 'FRAME49']

        # 開啟零件檔更改變數後儲存並關閉
        for name in file_name_list:
            mprog.import_part(fp.system_root + fp.part, name)

            if i == 4:
                if name == 'SLIDE_UNIT_All':
                    mprog.axis_system()
                    mprog.scaling(0.8)
                    if l == 0:
                        mprog.param_change(name, 'P', par.P_15[i])
                    else:
                        mprog.param_change(name, 'P', par.P[i])
                    mprog.save_file_part(path, name)
                elif name == 'BALANCER_LEFT_All' or name == 'BALANCER_RIGHT_ALL':
                    mprog.axis_system()
                    mprog.scaling(0.7)
                    mprog.save_file_part(path, name)
                # elif name == 'GIB1' or name == 'GIB2':
                #     mprog.axis_system()
                #     mprog.scaling(0.9)
                #     mprog.save_file_part(path, name)
                else:
                    if l == 0:
                        if name == 'FRAME1' or name == 'FRAME2' or name == 'FRAME44' or name == 'GIB1' or name == 'GIB2':  # 更改零件變數B
                            mprog.param_change(name, 'B', par.B_15[i])
                            mprog.save_file_part(path, name)
                        elif name == 'FRAME3' or name == 'FRAME4' or name == 'FRAME9' or name == 'FRAME32' or name == 'FRAME41' or name == 'FRAME43' or name == 'FRAME20' or name == 'FRAME30' or name == 'FRAME29' or name == 'FRAME42' or name == 'FRAME43' or name == 'FRAME45':  # 更改零件變數R
                            mprog.param_change(name, 'R', par.R_15[i])
                            mprog.save_file_part(path, name)
                        elif name == 'FRAME10' or name == 'FRAME11' or name == 'FRAME12' or name == 'FRAME13' or name == 'BOLSTER1':  # 更改零件變數E
                            mprog.param_change(name, 'E', par.E_15[i])
                            if name == 'BOLSTER1':
                                mprog.param_change('BOLSTER1', "hole_type", par.hole_type[j])
                            mprog.save_file_part(path, name)
                        elif name == 'FRAME29' or name == 'FRAME8' or name == 'FRAME5' or name == 'FRAME6' or name == 'FRAME7':  # 更改零件變數A
                            mprog.param_change(name, 'A', par.A_15[i])
                            mprog.save_file_part(path, name)
                        elif name == 'SLIDE_UNIT_All':  # 更改零件變數P
                            mprog.param_change(name, 'P', par.P_15[i])
                            mprog.save_file_part(path, name)
                        elif name == 'BOLSTER3' or name == 'BOLSTER2':  # 更改零件變數Q
                            mprog.param_change(name, 'Q', par.Q_15[i])
                            mprog.save_file_part(path, name)
                        else:
                            mprog.save_file_part(path, name)
                    else:
                        if name == 'FRAME1' or name == 'FRAME2' or name == 'FRAME44' or name == 'GIB1' or name == 'GIB2':  # 更改零件變數B
                            mprog.param_change(name, 'B', par.B[i])
                            mprog.save_file_part(path, name)
                        elif name == 'FRAME3' or name == 'FRAME4' or name == 'FRAME9' or name == 'FRAME32' or name == 'FRAME41' or name == 'FRAME43' or name == 'FRAME20' or name == 'FRAME30' or name == 'FRAME29' or name == 'FRAME42' or name == 'FRAME42' or name == 'FRAME43' or name == 'FRAME45':  # 更改零件變數R
                            mprog.param_change(name, 'R', par.R[i])
                            mprog.save_file_part(path, name)
                        elif name == 'FRAME10' or name == 'FRAME11' or name == 'FRAME12' or name == 'FRAME13' or name == 'BOLSTER1':  # 更改零件變數E
                            mprog.param_change(name, 'E', par.E[i])
                            if name == 'BOLSTER1':
                                mprog.param_change('BOLSTER1', "hole_type", par.hole_type[j])
                            mprog.save_file_part(path, name)
                        elif name == 'FRAME29' or name == 'FRAME8' or name == 'FRAME5' or name == 'FRAME6' or name == 'FRAME7':  # 更改零件變數A
                            mprog.param_change(name, 'A', par.A[i])
                            mprog.save_file_part(path, name)
                        elif name == 'SLIDE_UNIT_All':  # 更改零件變數P
                            mprog.param_change(name, 'P', par.P[i])
                            mprog.save_file_part(path, name)
                        elif name == 'BOLSTER3' or name == 'BOLSTER2':  # 更改零件變數Q
                            mprog.param_change(name, 'Q', par.Q[i])
                            mprog.save_file_part(path, name)
                        else:
                            mprog.save_file_part(path, name)
            else:
                if l == 0:
                    if name == 'FRAME1' or name == 'FRAME2' or name == 'FRAME44' or name == 'GIB1' or name == 'GIB2':  # 更改零件變數B
                        mprog.param_change(name, 'B', par.B_15[i])
                        mprog.save_file_part(path, name)
                    elif name == 'FRAME3' or name == 'FRAME4' or name == 'FRAME9' or name == 'FRAME32' or name == 'FRAME41' or name == 'FRAME43' or name == 'FRAME20' or name == 'FRAME30' or name == 'FRAME29' or name == 'FRAME42' or name == 'FRAME42' or name == 'FRAME43' or name == 'FRAME45':  # 更改零件變數R
                        mprog.param_change(name, 'R', par.R_15[i])
                        mprog.save_file_part(path, name)
                    elif name == 'FRAME10' or name == 'FRAME11' or name == 'FRAME12' or name == 'FRAME13' or name == 'BOLSTER1':  # 更改零件變數E
                        mprog.param_change(name, 'E', par.E_15[i])
                        if name == 'BOLSTER1':
                            mprog.param_change('BOLSTER1', "hole_type", par.hole_type[j])
                        mprog.save_file_part(path, name)
                    elif name == 'FRAME29' or name == 'FRAME8' or name == 'FRAME5' or name == 'FRAME6' or name == 'FRAME7':  # 更改零件變數A
                        mprog.param_change(name, 'A', par.A_15[i])
                        mprog.save_file_part(path, name)
                    elif name == 'SLIDE_UNIT_All':  # 更改零件變數P
                        mprog.param_change(name, 'P', par.P_15[i])
                        mprog.save_file_part(path, name)
                    elif name == 'BOLSTER3' or name == 'BOLSTER2':  # 更改零件變數Q
                        mprog.param_change(name, 'Q', par.Q_15[i])
                        mprog.save_file_part(path, name)
                    else:
                        mprog.save_file_part(path, name)
                else:
                    if name == 'FRAME1' or name == 'FRAME2' or name == 'FRAME44' or name == 'GIB1' or name == 'GIB2':  # 更改零件變數B
                        mprog.param_change(name, 'B', par.B[i])
                        mprog.save_file_part(path, name)
                    elif name == 'FRAME3' or name == 'FRAME4' or name == 'FRAME9' or name == 'FRAME32' or name == 'FRAME41' or name == 'FRAME43' or name == 'FRAME20' or name == 'FRAME30' or name == 'FRAME29' or name == 'FRAME42' or name == 'FRAME42' or name == 'FRAME43' or name == 'FRAME45':  # 更改零件變數R
                        mprog.param_change(name, 'R', par.R[i])
                        mprog.save_file_part(path, name)
                    elif name == 'FRAME10' or name == 'FRAME11' or name == 'FRAME12' or name == 'FRAME13' or name == 'BOLSTER1':  # 更改零件變數E
                        mprog.param_change(name, 'E', par.E[i])
                        if name == 'BOLSTER1':
                            mprog.param_change('BOLSTER1', "hole_type", par.hole_type[j])
                        mprog.save_file_part(path, name)
                    elif name == 'FRAME29' or name == 'FRAME8' or name == 'FRAME5' or name == 'FRAME6' or name == 'FRAME7':  # 更改零件變數A
                        mprog.param_change(name, 'A', par.A[i])
                        mprog.save_file_part(path, name)
                    elif name == 'SLIDE_UNIT_All':  # 更改零件變數P
                        mprog.param_change(name, 'P', par.P[i])
                        mprog.save_file_part(path, name)
                    elif name == 'BOLSTER3' or name == 'BOLSTER2':  # 更改零件變數Q
                        mprog.param_change(name, 'Q', par.Q[i])
                        mprog.save_file_part(path, name)
                    else:
                        mprog.save_file_part(path, name)

    def ass_(self, h, i, l, j, path , part_path):
        type = str(self.ui.comboBox_4.currentText())  # 沖床噸數類型
        # 開啟新組合檔
        mprog.assembly_create()

        # 匯入待組合零件檔
        file_Assembly_name_list = ['BOLSTER1', 'BOLSTER2', 'BOLSTER3', 'Fixture', 'FRAME1', 'FRAME2', 'FRAME3',
                                   'FRAME4', 'FRAME5', 'FRAME6', 'FRAME7', 'FRAME8', 'FRAME9', 'FRAME10', 'FRAME11',
                                   'FRAME12', 'FRAME13', 'FRAME14', 'FRAME15', 'FRAME16', 'FRAME17', 'FRAME18',
                                   'FRAME19', 'FRAME20', 'FRAME21', 'FRAME22', 'FRAME23', 'FRAME24', 'FRAME25',
                                   'FRAME26', 'FRAME27', 'FRAME28', 'FRAME29', 'FRAME30', 'FRAME31', 'FRAME31',
                                   'FRAME31', 'FRAME31', 'FRAME32', 'FRAME33', 'FRAME33', 'FRAME34', 'FRAME34', 'FRAME35',
                                   'FRAME36', 'FRAME36', 'FRAME37', 'FRAME37', 'FRAME37', 'FRAME37', 'FRAME38',
                                   'FRAME39', 'FRAME40', 'FRAME41', 'FRAME42', 'FRAME43', 'FRAME44', 'FRAME45',
                                   'FRAME46', 'FRAME47', 'FRAME47', 'FRAME48', 'FRAME49', 'GIB1', 'GIB2',
                                   'BALANCER_LEFT_All', 'BALANCER_RIGHT_ALL', 'CRANK_SHAFT_CLOCK', 'CLUCTH_ASSEMBLY_All',
                                   'SLIDE_UNIT_ALL', 'CRANK_SHAFT', 'JOINT_All', 'MAIN_GEAR1', 'MAIN_GEAR2', 'MAIN_GEAR3',
                                   'MAIN_GEAR4', 'JOINT1']

        for x in file_Assembly_name_list:  # 讀取串列名稱並匯入檔案
            mprog.import_file_Part(part_path, x)

        mprog.base_lock('FRAME20.1', 'FRAME20.1', 0)  # 基準零件(定海神針)

        # (0表示SAME, 1表示OPPOSITE)
        # 平板-四底座
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME6.1', par.A_15[i] / 2, 'XZ.PLANE', 1, 1)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME6.1', par.A[i] / 2, 'XZ.PLANE', 1, 1)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME6.1', -par.Z[i], 'XY.PLANE', 1, 2)
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME6.1', -80 - par.B_15[i] + par.F[i] / 2, 'YZ.PLANE', 1, 3)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME6.1', -80 - par.B[i] + par.F[i] / 2, 'YZ.PLANE', 1, 3)
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME7.1', -par.A_15[i] / 2, 'XZ.PLANE', 0, 4)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME7.1', -par.A[i] / 2, 'XZ.PLANE', 0, 4)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME7.1', -par.Z[i], 'XY.PLANE', 0, 5)
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME7.1', -80 - par.B_15[i] + par.F[i] / 2, 'YZ.PLANE', 1, 6)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME7.1', -80 - par.B[i] + par.F[i] / 2, 'YZ.PLANE', 1, 6)
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME5.1', -par.A_15[i] / 2, 'XZ.PLANE', 0, 7)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME5.1', -par.A[i] / 2, 'XZ.PLANE', 0, 7)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME5.1', -par.Z[i], 'XY.PLANE', 1, 8)
        if l == 0:
            mprog.add_offset_assembly('FRAME7.1', 'FRAME5.1', -par.B_15[i], 'YZ.PLANE', 1, 9)
        else:
            mprog.add_offset_assembly('FRAME7.1', 'FRAME5.1', -par.B[i], 'YZ.PLANE', 1, 9)
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME8.1', par.A_15[i] / 2, 'XZ.PLANE', 0, 10)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME8.1', par.A[i] / 2, 'XZ.PLANE', 0, 10)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME8.1', -par.Z[i], 'XY.PLANE', 0, 11)
        if l == 0:
            mprog.add_offset_assembly('FRAME6.1', 'FRAME8.1', -par.B_15[i], 'YZ.PLANE', 0, 12)
        else:
            mprog.add_offset_assembly('FRAME6.1', 'FRAME8.1', -par.B[i], 'YZ.PLANE', 0, 12)
        # 左右側板
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME2.1', par.R_15[i] / 2 + 140, 'XZ.PLANE', 1, 13)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME2.1', par.R[i] / 2 + 140, 'XZ.PLANE', 1, 13)
        mprog.add_offset_assembly('FRAME8.1', 'FRAME2.1', par.H[i] / 2, 'XY.PLANE', 0, 14)
        if l == 0:
            mprog.add_offset_assembly('FRAME8.1', 'FRAME2.1', par.B_15[i] / 2, 'YZ.PLANE', 0, 15)
        else:
            mprog.add_offset_assembly('FRAME8.1', 'FRAME2.1', par.B[i] / 2, 'YZ.PLANE', 0, 15)
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME1.1', -par.R_15[i] / 2 - 140, 'XZ.PLANE', 1, 16)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME1.1', -par.R[i] / 2 - 140, 'XZ.PLANE', 1, 16)
        mprog.add_offset_assembly('FRAME5.1', 'FRAME1.1', -par.H[i] / 2, 'XY.PLANE', 1, 17)
        if l == 0:
            mprog.add_offset_assembly('FRAME5.1', 'FRAME1.1', -par.B_15[i] / 2, 'YZ.PLANE', 1, 18)
        else:
            mprog.add_offset_assembly('FRAME5.1', 'FRAME1.1', -par.B[i] / 2, 'YZ.PLANE', 1, 18)
        # 底部前、中板
        if l == 0:
            mprog.add_offset_assembly('FRAME5.1', 'FRAME3.1', par.A_15[i] / 2, 'XZ.PLANE', 1, 19)
        else:
            mprog.add_offset_assembly('FRAME5.1', 'FRAME3.1', par.A[i] / 2, 'XZ.PLANE', 1, 19)
        mprog.add_offset_assembly('FRAME5.1', 'FRAME3.1', 0, 'XY.PLANE', 0, 20)
        mprog.add_offset_assembly('FRAME5.1', 'FRAME3.1', 0, 'YZ.PLANE', 1, 21)
        mprog.add_offset_assembly('FRAME9.1', 'FRAME3.1', 0, 'XZ.PLANE', 0, 22)
        mprog.add_offset_assembly('FRAME9.1', 'FRAME3.1', 0, 'XY.PLANE', 0, 23)
        if l == 0:
            mprog.add_offset_assembly('FRAME9.1', 'FRAME3.1', -par.FRAME2_lower_depth_15[i] + 5, 'YZ.PLANE', 0, 24)
        else:
            mprog.add_offset_assembly('FRAME9.1', 'FRAME3.1', -par.FRAME2_lower_depth[i] + 5, 'YZ.PLANE', 0, 24)
        # 中間左右側板
        if l == 0:
            mprog.add_offset_assembly('FRAME11.1', 'FRAME3.1', par.R_15[i] / 2, 'XZ.PLANE', 0, 25)
        else:
            mprog.add_offset_assembly('FRAME11.1', 'FRAME3.1', par.R[i] / 2, 'XZ.PLANE', 0, 25)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME11.1', -par.T[i], 'XY.PLANE', 1, 26)
        if l == 0:
            if i == 4:
                mprog.add_offset_assembly('FRAME3.1', 'FRAME11.1', par.FRAME2_lower_depth_15[i] + 80, 'YZ.PLANE', 1, 27)
            else :
                mprog.add_offset_assembly('FRAME3.1', 'FRAME11.1', par.FRAME2_lower_depth_15[i], 'YZ.PLANE', 1, 27)
        else:
            if i == 4:
                mprog.add_offset_assembly('FRAME3.1', 'FRAME11.1', par.FRAME2_lower_depth[i] + 80, 'YZ.PLANE', 1, 27)
            else :
                mprog.add_offset_assembly('FRAME3.1', 'FRAME11.1', par.FRAME2_lower_depth[i], 'YZ.PLANE', 1, 27)
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME10.1', -par.R_15[i] / 2 - 90, 'XZ.PLANE', 0, 28)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME10.1', -par.R[i] / 2 - 90, 'XZ.PLANE', 0, 28)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME10.1', -par.T[i], 'XY.PLANE', 0, 29)
        if l == 0:
            if i == 4:
                mprog.add_offset_assembly('FRAME3.1', 'FRAME10.1', par.FRAME2_lower_depth_15[i] + 80, 'YZ.PLANE', 1, 30)
            else:
                mprog.add_offset_assembly('FRAME3.1', 'FRAME10.1', par.FRAME2_lower_depth_15[i], 'YZ.PLANE', 1, 30)
        else:
            if i == 4:
                mprog.add_offset_assembly('FRAME3.1', 'FRAME10.1', par.FRAME2_lower_depth[i] + 80, 'YZ.PLANE', 1, 30)
            else:
                mprog.add_offset_assembly('FRAME3.1', 'FRAME10.1', par.FRAME2_lower_depth[i], 'YZ.PLANE', 1, 30)
        # 底部後面ㄇ形角鐵
        mprog.add_offset_assembly('FRAME3.1', 'FRAME4.1', 0, 'XZ.PLANE', 1, 31)
        mprog.add_offset_assembly('FRAME3.1', 'FRAME4.1', 0, 'XY.PLANE', 0, 32)
        if l == 0:
            mprog.add_offset_assembly('FRAME3.1', 'FRAME4.1', par.B_15[i], 'YZ.PLANE', 0, 33)
        else:
            mprog.add_offset_assembly('FRAME3.1', 'FRAME4.1', par.B[i], 'YZ.PLANE', 0, 33)
        # 平板底部零件
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME12.1', -par.R_15[i] / 2, 'XZ.PLANE', 1, 34)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME12.1', -par.R[i] / 2, 'XZ.PLANE', 1, 34)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME12.1', -par.T[i], 'XY.PLANE', 0, 35)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME12.1', 0, 'YZ.PLANE', 1, 36)
        # 平板底部零件
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME13.1', par.R_15[i] / 2, 'XZ.PLANE', 1, 37)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME13.1', par.R[i] / 2, 'XZ.PLANE', 1, 37)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME13.1', -par.T[i], 'XY.PLANE', 0, 38)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME13.1', 0, 'YZ.PLANE', 1, 39)
        # 前中上軸承板&角鐵
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME20.1', 0, 'XZ.PLANE', 1, 40)
        mprog.add_offset_assembly('FRAME3.1', 'FRAME20.1', -par.H[i] + 40, 'XY.PLANE', 0, 41)
        mprog.add_offset_assembly('FRAME3.1', 'FRAME20.1', 0, 'YZ.PLANE', 0, 42)
        mprog.add_offset_assembly('FRAME30.1', 'FRAME20.1', 0, 'XZ.PLANE', 0, 43)
        mprog.add_offset_assembly('FRAME30.1', 'FRAME20.1', -5, 'XY.PLANE', 0, 44)
        mprog.add_offset_assembly('FRAME30.1', 'FRAME20.1', -550, 'YZ.PLANE', 0, 45)
        mprog.add_offset_assembly('FRAME29.1', 'FRAME20.1', 0, 'XZ.PLANE', 1, 46)
        mprog.add_offset_assembly('FRAME29.1', 'FRAME20.1', -par.FRAME20_H[i], 'XY.PLANE', 0, 47)
        mprog.add_offset_assembly('FRAME29.1', 'FRAME20.1', 0, 'YZ.PLANE', 1, 48)
        # 氣壓缸鎖固板左右
        if i == 4:
            mprog.add_offset_assembly('FRAME2.1', 'FRAME21.1', 183, 'XZ.PLANE', 1, 49)
        else:
            mprog.add_offset_assembly('FRAME2.1', 'FRAME21.1', 242, 'XZ.PLANE', 1, 49)
        mprog.add_offset_assembly('FRAME3.1', 'FRAME21.1', -par.H[i] + 40, 'XY.PLANE', 1, 50)
        mprog.add_offset_assembly('FRAME20.1', 'FRAME21.1', 280, 'YZ.PLANE', 1, 51)
        if i == 4:
            mprog.add_offset_assembly('FRAME1.1', 'FRAME22.1', -183, 'XZ.PLANE', 1, 52)
        else:
            mprog.add_offset_assembly('FRAME1.1', 'FRAME22.1', -242, 'XZ.PLANE', 1, 52)
        mprog.add_offset_assembly('FRAME3.1', 'FRAME22.1', -par.H[i] + 40, 'XY.PLANE', 0, 53)
        mprog.add_offset_assembly('FRAME20.1', 'FRAME22.1', 280, 'YZ.PLANE', 1, 54)
        # 鎖固平板六兄弟
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME14.1', 0, 'XZ.PLANE', 0, 55)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME14.1', -par.T[i], 'XY.PLANE', 0, 56)
        mprog.add_offset_assembly('FRAME3.1', 'FRAME14.1', -75, 'YZ.PLANE', 1, 57)
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME15.1', par.R_15[i] / 2 + 75 + 140, 'XZ.PLANE', 0, 58)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME15.1', par.R[i] / 2 + 75 + 140, 'XZ.PLANE', 0, 58)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME15.1', -par.T[i], 'XY.PLANE', 0, 59)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME15.1', -par.F[i] / 2 + 80 + 37.5, 'YZ.PLANE', 1, 60)
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME16.1', par.R_15[i] / 2 + 75 + 140, 'XZ.PLANE', 0, 61)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME16.1', par.R[i] / 2 + 75 + 140, 'XZ.PLANE', 0, 61)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME16.1', -par.T[i], 'XY.PLANE', 0, 62)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME16.1', par.F[i] / 2 - 80 - 37.5, 'YZ.PLANE', 1, 63)
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME18.1', -par.R_15[i] / 2 - 75 - 140, 'XZ.PLANE', 1, 64)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME18.1', -par.R[i] / 2 - 75 - 140, 'XZ.PLANE', 1, 64)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME18.1', -par.T[i], 'XY.PLANE', 0, 65)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME18.1', par.F[i] / 2 - 80 - 37.5, 'YZ.PLANE', 0, 66)
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME19.1', -par.R_15[i] / 2 - 75 - 140, 'XZ.PLANE', 1, 67)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME19.1', -par.R[i] / 2 - 75 - 140, 'XZ.PLANE', 1, 67)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME19.1', -par.T[i], 'XY.PLANE', 0, 68)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME19.1', -par.F[i] / 2 + 80 + 37.5, 'YZ.PLANE', 0, 69)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME17.1', 0, 'XZ.PLANE', 1, 70)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME17.1', -par.T[i], 'XY.PLANE', 0, 71)
        mprog.add_offset_assembly('FRAME9.1', 'FRAME17.1', 0, 'YZ.PLANE', 0, 72)
        # 左右側板前GIB
        if i == 4:
            mprog.add_offset_assembly('GIB1.1', 'FRAME3.1', par.FRAME1_lower_high[i] + 40 - 34.5, 'XY.PLANE', 0, 73)
        else:
            mprog.add_offset_assembly('GIB1.1', 'FRAME3.1', par.FRAME1_lower_high[i] + 40, 'XY.PLANE', 0, 73)
        mprog.add_offset_assembly('GIB1.1', 'FRAME1.1', 72.5, 'XZ.PLANE', 0, 74)
        mprog.add_offset_assembly('GIB1.1', 'FRAME5.1', 334.65, 'YZ.PLANE', 0, 75)
        if i == 4:
            mprog.add_offset_assembly('GIB2.1', 'FRAME3.1', par.FRAME1_lower_high[i] + 40 - 34.5, 'XY.PLANE', 0, 76)
        else:
            mprog.add_offset_assembly('GIB2.1', 'FRAME3.1', par.FRAME1_lower_high[i] + 40, 'XY.PLANE', 0, 76)
        mprog.add_offset_assembly('GIB2.1', 'FRAME2.1', -72.5, 'XZ.PLANE', 0, 77)
        mprog.add_offset_assembly('GIB2.1', 'FRAME5.1', 334.65, 'YZ.PLANE', 0, 78)
        # 左GIB後鎖固用方塊
        if i == 4:
            mprog.add_offset_assembly('GIB2.1', 'FRAME23.1', -par.pocket_1_upper_hole[i] + 80 + 3, 'XY.PLANE', 1, 79)
        else:
            mprog.add_offset_assembly('GIB2.1', 'FRAME23.1', -par.pocket_1_upper_hole[i] + 80 + 3, 'XY.PLANE', 1, 79)
        mprog.add_offset_assembly('FRAME2.1', 'FRAME23.1', 50 + 35, 'XZ.PLANE', 1, 80)
        mprog.add_offset_assembly('GIB2.1', 'FRAME23.1', 0, 'YZ.PLANE', 0, 81)
        if i == 4:
            mprog.add_offset_assembly('GIB2.1', 'FRAME24.1', -34.5, 'XY.PLANE', 1, 82)
        else:
            mprog.add_offset_assembly('GIB2.1', 'FRAME24.1', 0, 'XY.PLANE', 1, 82)
        mprog.add_offset_assembly('FRAME2.1', 'FRAME24.1', 50 + 35, 'XZ.PLANE', 0, 83)
        mprog.add_offset_assembly('GIB2.1', 'FRAME24.1', 0, 'YZ.PLANE', 1, 84)
        if i == 4:
            mprog.add_offset_assembly('GIB2.1', 'FRAME27.1', -par.pocket_1_upper_hole[i] + 340 + 80, 'XY.PLANE', 1, 85)
        else:
            mprog.add_offset_assembly('GIB2.1', 'FRAME27.1', -par.pocket_1_upper_hole[i] + 340 + 80, 'XY.PLANE', 1, 85)
        mprog.add_offset_assembly('FRAME2.1', 'FRAME27.1', 50, 'XZ.PLANE', 1, 86)
        mprog.add_offset_assembly('GIB2.1', 'FRAME27.1', 0, 'YZ.PLANE', 1, 87)
        if i == 4:
            mprog.add_offset_assembly('GIB2.1', 'FRAME31.1', -par.pocket_1_upper_hole[i], 'XY.PLANE', 1, 88)
        else:
            mprog.add_offset_assembly('GIB2.1', 'FRAME31.1', -par.pocket_1_upper_hole[i], 'XY.PLANE', 1, 88)
        mprog.add_offset_assembly('FRAME23.1', 'FRAME31.1', -35, 'XZ.PLANE', 0, 89)
        mprog.add_offset_assembly('FRAME23.1', 'FRAME31.1', -130.35 - 22.5, 'YZ.PLANE', 0, 90)
        if i == 4:
            mprog.add_offset_assembly('GIB2.1', 'FRAME31.2', -150 - 34.5, 'XY.PLANE', 1, 91)
        else:
            mprog.add_offset_assembly('GIB2.1', 'FRAME31.2', -150, 'XY.PLANE', 1, 91)
        mprog.add_offset_assembly('FRAME23.1', 'FRAME31.2', -130.35 - 22.5, 'YZ.PLANE', 0, 92)
        mprog.add_offset_assembly('FRAME23.1', 'FRAME31.2', -35, 'XZ.PLANE', 0, 93)
        # 右GIB後鎖固用方塊
        if i == 4:
            mprog.add_offset_assembly('GIB1.1', 'FRAME25.1', -par.pocket_1_upper_hole[i] + 80 + 3, 'XY.PLANE', 1, 94)
        else:
            mprog.add_offset_assembly('GIB1.1', 'FRAME25.1', -par.pocket_1_upper_hole[i] + 80 + 3, 'XY.PLANE', 1, 94)
        mprog.add_offset_assembly('FRAME1.1', 'FRAME25.1', -50 - 35, 'XZ.PLANE', 0, 95)
        mprog.add_offset_assembly('GIB1.1', 'FRAME25.1', 0, 'YZ.PLANE', 1, 96)
        if i == 4:
            mprog.add_offset_assembly('GIB1.1', 'FRAME26.1', -34.5, 'XY.PLANE', 1, 97)
        else:
            mprog.add_offset_assembly('GIB1.1', 'FRAME26.1', 0, 'XY.PLANE', 1, 97)
        mprog.add_offset_assembly('FRAME1.1', 'FRAME26.1', -50 - 35, 'XZ.PLANE', 0, 98)
        mprog.add_offset_assembly('GIB1.1', 'FRAME26.1', 0, 'YZ.PLANE', 0, 99)
        if i == 4:
            mprog.add_offset_assembly('GIB1.1', 'FRAME28.1', -par.pocket_1_upper_hole[i] + 340 + 80, 'XY.PLANE', 1, 100)
        else:
            mprog.add_offset_assembly('GIB1.1', 'FRAME28.1', -par.pocket_1_upper_hole[i] + 340 + 80, 'XY.PLANE', 1, 100)
        mprog.add_offset_assembly('FRAME1.1', 'FRAME28.1', -50, 'XZ.PLANE', 0, 101)
        mprog.add_offset_assembly('GIB1.1', 'FRAME28.1', 0, 'YZ.PLANE', 1, 102)
        if i == 4:
            mprog.add_offset_assembly('GIB1.1', 'FRAME31.3', -par.pocket_1_upper_hole[i], 'XY.PLANE', 1, 103)
        else:
            mprog.add_offset_assembly('GIB1.1', 'FRAME31.3', -par.pocket_1_upper_hole[i], 'XY.PLANE', 1, 103)
        mprog.add_offset_assembly('FRAME25.1', 'FRAME31.3', -35, 'XZ.PLANE', 1, 104)
        mprog.add_offset_assembly('FRAME25.1', 'FRAME31.3', 130.35 + 22.5, 'YZ.PLANE', 1, 105)
        if i == 4:
            mprog.add_offset_assembly('GIB1.1', 'FRAME31.4', -150 - 34.5, 'XY.PLANE', 1, 106)
        else:
            mprog.add_offset_assembly('GIB1.1', 'FRAME31.4', -150, 'XY.PLANE', 1, 106)
        mprog.add_offset_assembly('FRAME31.3', 'FRAME31.4', 0, 'YZ.PLANE', 0, 107)
        mprog.add_offset_assembly('FRAME31.3', 'FRAME31.4', 0, 'XZ.PLANE', 0, 108)
        # 中間與平板鎖固方塊下板子
        mprog.add_offset_assembly('FRAME43.1', 'FRAME17.1', -48, 'XY.PLANE', 1, 109)
        mprog.add_offset_assembly('FRAME43.1', 'FRAME17.1', 0, 'XZ.PLANE', 0, 110)
        mprog.add_offset_assembly('FRAME43.1', 'FRAME17.1', 0, 'YZ.PLANE', 1, 111)
        # 後面安裝馬達下板子
        mprog.add_offset_assembly('FRAME41.1', 'FRAME20.1', -1340, 'XY.PLANE', 0, 112)
        mprog.add_offset_assembly('FRAME41.1', 'FRAME4.1', 0, 'XZ.PLANE', 0, 113)
        mprog.add_offset_assembly('FRAME41.1', 'FRAME4.1', 0, 'YZ.PLANE', 0, 114)
        # FRAME2中間下板子
        mprog.add_offset_assembly('FRAME32.1', 'BOLSTER1.1', 0, 'XZ.PLANE', 0, 115)
        mprog.add_offset_assembly('FRAME32.1', 'FRAME3.1', par.FRAME_32_XY[i] + 20, 'XY.PLANE', 0, 116)
        mprog.add_offset_assembly('FRAME32.1', 'FRAME30.1', 0, 'YZ.PLANE', 0, 117)
        # 後方大軸承支架
        mprog.add_offset_assembly('FRAME10.1', 'FRAME34.2', -485.543, 'YZ.PLANE', 1, 118)
        mprog.add_offset_assembly('FRAME34.2', 'FRAME10.1', -(par.FRAME_10_H[i] + 40), 'XY.PLANE', 0, 119)
        if l == 0:
            mprog.add_offset_assembly('FRAME34.2', 'FRAME3.1', -par.R_15[i] / 2 - 90, 'XZ.PLANE', 0, 120)
        else:
            mprog.add_offset_assembly('FRAME34.2', 'FRAME3.1', -par.R[i] / 2 - 90, 'XZ.PLANE', 0, 120)
        if l == 0:
            mprog.add_offset_assembly('FRAME34.1', 'FRAME34.2', -par.R_15[i] - 180, 'XZ.PLANE', 1, 121)
        else:
            mprog.add_offset_assembly('FRAME34.1', 'FRAME34.2', -par.R[i] - 180, 'XZ.PLANE', 1, 121)
        mprog.add_offset_assembly('FRAME34.1', 'FRAME34.2', 0, 'XY.PLANE', 1, 122)
        mprog.add_offset_assembly('FRAME34.1', 'FRAME34.2', 0, 'YZ.PLANE', 0, 123)
        # 後方大軸承
        mprog.add_offset_assembly('FRAME35.1', 'FRAME3.1', 0, 'XZ.PLANE', 1, 124)
        mprog.add_offset_assembly('FRAME35.1', 'FRAME20.1', -(par.H[i] - par.Z[i] - par.S[i] - 34 + 448), 'XY.PLANE', 0, 125)
        mprog.add_offset_assembly('FRAME35.1', 'FRAME20.1', -par.FRAME20_FRAME2_YZ[i] - 36, 'YZ.PLANE', 0, 126)
        # 後方馬達下支撐板
        mprog.add_offset_assembly('FRAME39.1', 'FRAME1.1', 50, 'XZ.PLANE', 0, 127)
        mprog.add_offset_assembly('FRAME39.1', 'FRAME20.1', -360, 'XY.PLANE', 0, 128)
        mprog.add_offset_assembly('FRAME39.1', 'FRAME7.1', 320, 'YZ.PLANE', 0, 129)
        mprog.add_offset_assembly('FRAME38.1', 'FRAME2.1', -50, 'XZ.PLANE', 0, 130)
        mprog.add_offset_assembly('FRAME38.1', 'FRAME20.1', 360, 'XY.PLANE', 1, 131)
        mprog.add_offset_assembly('FRAME38.1', 'FRAME6.1', -320, 'YZ.PLANE', 1, 132)
        mprog.add_offset_assembly('FRAME37.1', 'FRAME39.1', 0, 'XZ.PLANE', 0, 133)
        mprog.add_offset_assembly('FRAME37.1', 'FRAME39.1', 19, 'XY.PLANE', 1, 134)
        mprog.add_offset_assembly('FRAME37.1', 'FRAME39.1', 125, 'YZ.PLANE', 1, 135)
        mprog.add_offset_assembly('FRAME37.2', 'FRAME39.1', 0, 'XZ.PLANE', 0, 136)
        mprog.add_offset_assembly('FRAME37.2', 'FRAME39.1', 19, 'XY.PLANE', 1, 137)
        mprog.add_offset_assembly('FRAME37.2', 'FRAME39.1', -125, 'YZ.PLANE', 1, 138)
        mprog.add_offset_assembly('FRAME37.3', 'FRAME38.1', 0, 'XZ.PLANE', 1, 139)
        mprog.add_offset_assembly('FRAME37.3', 'FRAME38.1', 0, 'XY.PLANE', 0, 140)
        mprog.add_offset_assembly('FRAME37.3', 'FRAME38.1', -125, 'YZ.PLANE', 1, 141)
        mprog.add_offset_assembly('FRAME37.4', 'FRAME38.1', 0, 'XZ.PLANE', 1, 142)
        mprog.add_offset_assembly('FRAME37.4', 'FRAME38.1', 0, 'XY.PLANE', 0, 143)
        mprog.add_offset_assembly('FRAME37.4', 'FRAME38.1', 125, 'YZ.PLANE', 1, 144)
        # FRAME2右邊角鐵
        mprog.add_offset_assembly('FRAME33.1', 'FRAME2.1', 140, 'XZ.PLANE', 1, 145)
        mprog.add_offset_assembly('FRAME33.1', 'FRAME15.1', -708, 'XY.PLANE', 0, 146)
        mprog.add_offset_assembly('FRAME33.1', 'FRAME11.1', 313.984, 'YZ.PLANE', 0, 147)
        # FRAME2中間半圓形零件
        mprog.add_offset_assembly('FRAME40.1', 'FRAME20.1', 130.5, 'XZ.PLANE', 0, 148)
        mprog.add_offset_assembly('FRAME40.1', 'FRAME20.1', -320, 'XY.PLANE', 0, 149)
        mprog.add_offset_assembly('FRAME40.1', 'FRAME20.1', par.FRAME20_FRAME2_YZ[i], 'YZ.PLANE', 1, 150)
        # FRAME35上兩圓管
        mprog.add_offset_assembly('FRAME36.1', 'FRAME35.1', -272.236, 'XZ.PLANE', 1, 151)
        mprog.add_offset_assembly('FRAME36.1', 'FRAME35.1', 272.236, 'XY.PLANE', 0, 152)
        mprog.add_offset_assembly('FRAME36.1', 'FRAME35.1', 41, 'YZ.PLANE', 0, 153)
        mprog.add_offset_assembly('FRAME36.2', 'FRAME35.1', 272.236, 'XZ.PLANE', 1, 154)
        mprog.add_offset_assembly('FRAME36.2', 'FRAME35.1', 272.236, 'XY.PLANE', 0, 155)
        mprog.add_offset_assembly('FRAME36.2', 'FRAME35.1', 41, 'YZ.PLANE', 0, 156)
        # 後方馬達下支撐板上治具
        mprog.add_offset_assembly('Fixture.1', 'FRAME41.1', 95, 'XZ.PLANE', 0, 157)
        mprog.add_offset_assembly('Fixture.1', 'FRAME41.1', 22 + 18, 'XY.PLANE', 0, 158)
        mprog.add_offset_assembly('Fixture.1', 'FRAME41.1', 312.5, 'YZ.PLANE', 0, 159)
        # 馬達下方板與後方橫板
        mprog.add_offset_assembly('FRAME42.1', 'FRAME41.1', 0, 'XZ.PLANE', 1, 160)
        mprog.add_offset_assembly('FRAME42.1', 'FRAME41.1', 71.5, 'XY.PLANE', 1, 161)
        mprog.add_offset_assembly('FRAME42.1', 'FRAME41.1', -19, 'YZ.PLANE', 1, 162)
        # 平板2, 3組立
        mprog.add_offset_assembly('BOLSTER2.1', 'BOLSTER3.1', 0, 'XZ.PLANE', 0, 163)
        mprog.add_offset_assembly('BOLSTER2.1', 'BOLSTER3.1', 0, 'XY.PLANE', 0, 164)
        mprog.add_offset_assembly('BOLSTER2.1', 'BOLSTER3.1', 0, 'YZ.PLANE', 0, 165)
        # 平板1, 3組立
        mprog.add_offset_assembly('BOLSTER1.1', 'BOLSTER3.1', 0, 'XZ.PLANE', 0, 166)
        mprog.add_offset_assembly('BOLSTER1.1', 'BOLSTER3.1', 0, 'YZ.PLANE', 0, 167)
        # if k == 0:
        if h == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'BOLSTER3.1', par.DH_S[i], 'XY.PLANE', 0, 168)
        elif h == 1:
            mprog.add_offset_assembly('BOLSTER1.1', 'BOLSTER3.1', par.DH_H[i], 'XY.PLANE', 0, 168)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'BOLSTER3.1', par.DH_P[i], 'XY.PLANE', 0, 168)
        # SLIDE跟平板3組立
        if i == 4:
            mprog.add_offset_product_assembly('SLIDE_UNIT_All.1', 'Geometrical Set.1', 'BOLSTER3.1', 267.38, 'XY.PLANE',
                                              1, 169)
        else:
            mprog.add_offset_product_assembly('SLIDE_UNIT_All.1', 'Geometrical Set.1', 'BOLSTER3.1', 85, 'XY.PLANE', 1,
                                              170)
        mprog.add_offset_product_assembly('SLIDE_UNIT_All.1', 'Geometrical Set.1', 'BOLSTER3.1', 0, 'XZ.PLANE', 1, 171)
        if i == 4:
            mprog.add_offset_product_assembly('SLIDE_UNIT_All.1', 'Geometrical Set.1', 'BOLSTER3.1', 138.171,
                                              'YZ.PLANE', 0, 172)
        else:
            mprog.add_offset_product_assembly('SLIDE_UNIT_All.1', 'Geometrical Set.1', 'BOLSTER3.1', 0, 'YZ.PLANE', 0,
                                              172)
        # 時鐘跟FRAME20組立
        mprog.add_offset_product_assembly('CRANK_SHAFT_CLOCK.1', 'Geometrical Set.2', 'FRAME20.1', 0, 'XZ.PLANE', 0,
                                          173)
        mprog.add_offset_product_assembly('CRANK_SHAFT_CLOCK.1', 'Geometrical Set.2', 'FRAME20.1',  33, 'YZ.PLANE', 0,
                                          174)
        mprog.add_offset_product_assembly('CRANK_SHAFT_CLOCK.1', 'Geometrical Set.2', 'FRAME20.1', -(par.H[i] - 34 - par.S[i] - par.Z[i]), 'XY.PLANE',
                                          0, 175)
        # 氣壓缸跟FRAME20組立
        if i == 4:
            mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1', -87.5,
                                              'XY.PLANE', 1, 176)
        else:
            mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1', -32, 'XY.PLANE',
                                              1, 176)
        if i == 4:
            if l == 0:
                mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1', -500.5,
                                                  'XZ.PLANE', 0, 177)
            else:
                mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1', -347,
                                                  'XZ.PLANE', 0, 177)
        else:
            if l == 0:
                mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1',
                                                  -par.R_15[i] / 2 - 13, 'XZ.PLANE', 0, 178)
            else:
                mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1',
                                                  -par.R[i] / 2 - 13, 'XZ.PLANE', 0, 178)
        if i == 4:
            mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1', -261.248,
                                              'YZ.PLANE', 0, 179)
        else:
            mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1', -260,
                                              'YZ.PLANE', 0, 179)
        if i == 4:
            mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1', -21.8,
                                              'XY.PLANE', 1, 180)
        else:
            mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1', -32, 'XY.PLANE',
                                              1, 180)
        if i == 4:
            if l == 0:
                mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1', -500.5,
                                                  'XZ.PLANE', 1, 181)
            else:
                mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1', -347,
                                                  'XZ.PLANE', 1, 181)
        else:
            if l == 0:
                mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1',
                                                  -par.R_15[i] / 2 - 13, 'XZ.PLANE', 1, 182)
            else:
                mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1',
                                                  -par.R[i] / 2 - 13, 'XZ.PLANE', 1, 182)
        if i == 4:
            mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1', 261.248,
                                              'YZ.PLANE', 1, 183)
        else:
            mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1', 260, 'YZ.PLANE',
                                              1, 183)
        # 離合器與FRAME20結合
        mprog.add_offset_product_assembly('CLUCTH_ASSEMBLY_All.1', 'Geometrical Set.1', 'FRAME20.1', -(par.H[i] - par.Z[i] - par.S[i] - 34 - 5 + 448), 'XY.PLANE',
                                              0, 184)
        mprog.add_offset_product_assembly('CLUCTH_ASSEMBLY_All.1', 'Geometrical Set.1', 'FRAME20.1', 0, 'XZ.PLANE', 1,
                                          185)
        mprog.add_offset_product_assembly('CLUCTH_ASSEMBLY_All.1', 'Geometrical Set.1', 'FRAME20.1', 512, 'YZ.PLANE', 1,
                                          186)
        # 曲軸與時鐘結合
        mprog.add_offset_product_assembly('CRANK_SHAFT_CLOCK.1', 'Geometrical Set.2', 'CRANK_SHAFT.1.1', 0, 'XZ.PLANE',
                                          0, 187)
        mprog.add_offset_product_assembly('CRANK_SHAFT_CLOCK.1', 'Geometrical Set.2', 'CRANK_SHAFT.1.1', 3 , 'YZ.PLANE',
                                          1, 188)
        mprog.add_offset_product_assembly('CRANK_SHAFT_CLOCK.1', 'Geometrical Set.2', 'CRANK_SHAFT.1.1', 0, 'XY.PLANE',
                                          1, 189)
        # 大齒輪結合
        mprog.add_offset_assembly('MAIN_GEAR1.1', 'MAIN_GEAR4.1', 0, 'XZ.PLANE', 0, 190)
        mprog.add_offset_assembly('MAIN_GEAR1.1', 'MAIN_GEAR4.1', 77.172, 'XY.PLANE', 0, 191)
        mprog.add_offset_assembly('MAIN_GEAR1.1', 'MAIN_GEAR4.1', 0, 'YZ.PLANE', 1, 192)
        # 大齒輪結合曲軸
        mprog.add_offset_assembly('MAIN_GEAR1.1', 'CRANK_SHAFT.1.1', 0, 'XZ.PLANE', 1, 193)
        mprog.add_offset_assembly('MAIN_GEAR1.1', 'CRANK_SHAFT.1.1', 0, 'XY.PLANE', 0, 194)
        mprog.add_offset_assembly('MAIN_GEAR1.1', 'FRAME20.1', 592.5 + 150, 'YZ.PLANE', 1, 195)
        # 機架20結合曲軸旁圓管
        mprog.add_offset_assembly('MAIN_GEAR2.1', 'FRAME20.1', 0, 'XZ.PLANE', 0, 196)
        mprog.add_offset_assembly('MAIN_GEAR2.1', 'FRAME20.1', par.H[i] - par.Z[i] - par.S[i] - 34, 'XY.PLANE', 1, 197)
        mprog.add_offset_assembly('MAIN_GEAR2.1', 'FRAME20.1', 520.5, 'YZ.PLANE', 1, 198)
        # 大齒輪結合長棒
        mprog.add_offset_assembly('MAIN_GEAR1.1', 'JOINT1.1', 0, 'XZ.PLANE', 0, 199)
        mprog.add_offset_assembly('MAIN_GEAR1.1', 'JOINT1.1', 0, 'XY.PLANE', 1, 200)
        mprog.add_offset_assembly('MAIN_GEAR1.1', 'JOINT1.1', -84, 'YZ.PLANE', 1, 201)
        #
        mprog.add_offset_product_assembly('JOINT_All.1', 'Geometrical Set.1', 'FRAME20.1', 0, 'XZ.PLANE',
                                          0, 202)
        mprog.add_offset_product_assembly('JOINT_All.1', 'Geometrical Set.1', 'FRAME20.1', 1010, 'YZ.PLANE',
                                          1, 203)
        mprog.add_offset_product_assembly('JOINT_All.1', 'Geometrical Set.1', 'FRAME20.1', -(par.H[i] - par.Z[i] - par.S[i] - 34), 'XY.PLANE',
                                          0, 204)
        # 大齒輪內套環
        mprog.add_offset_assembly('MAIN_GEAR2.1', 'MAIN_GEAR3.1', 0, 'XZ.PLANE', 1, 205)
        mprog.add_offset_assembly('MAIN_GEAR2.1', 'MAIN_GEAR3.1', 0, 'XY.PLANE', 0, 206)
        mprog.add_offset_assembly('MAIN_GEAR2.1', 'MAIN_GEAR3.1', 11, 'YZ.PLANE', 1, 207)
        # BOLSTER後板
        mprog.add_offset_assembly('FRAME45.1', 'FRAME17.1', 0, 'XZ.PLANE', 0, 208)
        mprog.add_offset_assembly('FRAME45.1', 'FRAME17.1', -48, 'XY.PLANE', 1, 209)
        mprog.add_offset_assembly('FRAME45.1', 'FRAME17.1', 0, 'YZ.PLANE', 1, 210)
        #
        mprog.add_offset_assembly('FRAME46.1', 'FRAME17.1', -48, 'XY.PLANE', 1, 211)
        mprog.add_offset_assembly('FRAME46.1', 'FRAME17.1', 0, 'XZ.PLANE', 0, 212)
        mprog.add_offset_assembly('FRAME46.1', 'FRAME17.1', 0, 'YZ.PLANE', 1, 213)
        #
        mprog.add_offset_assembly('FRAME20.1', 'FRAME44.1', 5, 'XY.PLANE', 0, 214)
        mprog.add_offset_assembly('FRAME20.1', 'FRAME44.1', 0, 'XZ.PLANE', 0, 215)
        mprog.add_offset_assembly('FRAME20.1', 'FRAME44.1', 979, 'YZ.PLANE', 1, 216)
        #
        mprog.add_offset_assembly('FRAME26.1', 'FRAME47.1', 29, 'XY.PLANE', 1, 217)
        mprog.add_offset_assembly('FRAME26.1', 'FRAME47.1', 35, 'XZ.PLANE', 0, 218)
        mprog.add_offset_assembly('FRAME26.1', 'FRAME47.1', -99.35, 'YZ.PLANE', 1, 219)
        mprog.add_offset_assembly('FRAME24.1', 'FRAME47.2', 29, 'XY.PLANE', 0, 220)
        mprog.add_offset_assembly('FRAME24.1', 'FRAME47.2', -35, 'XZ.PLANE', 1, 221)
        mprog.add_offset_assembly('FRAME24.1', 'FRAME47.2', 099.35, 'YZ.PLANE', 0, 222)
        # 方管
        mprog.add_offset_assembly('FRAME44.1', 'FRAME48.1', 124, 'XY.PLANE', 1, 223)
        mprog.add_offset_assembly('FRAME44.1', 'FRAME48.1', 159 + 16, 'XZ.PLANE', 0, 224)
        mprog.add_offset_assembly('FRAME44.1', 'FRAME48.1', -32, 'YZ.PLANE', 0, 225)
        # 圓柱
        mprog.add_offset_assembly('FRAME44.1', 'FRAME49.1', 220, 'XY.PLANE', 0, 226)
        mprog.add_offset_assembly('FRAME44.1', 'FRAME49.1', 145, 'XZ.PLANE', 0, 227)
        mprog.add_offset_assembly('FRAME44.1', 'FRAME49.1', -32, 'YZ.PLANE', 0, 228)
        #2023/02/23更新 FRAME1左邊角鐵
        mprog.add_offset_assembly('FRAME1.1', 'FRAME33.2', -140, 'XZ.PLANE', 0, 229)
        mprog.add_offset_assembly('FRAME19.1', 'FRAME33.2', 708, 'XY.PLANE', 0, 230)
        mprog.add_offset_assembly('FRAME10.1', 'FRAME33.2', -313.984, 'YZ.PLANE', 1, 231)

        # 更新
        mprog.update()
        mprog.Close_All()
        # 儲存Product檔
        mprog.saveas(part_path, type, '.CATProduct')

        # --------------------------------------- 生成爆炸圖-------------------------------------------
        # 重新定義拘束尺寸
        BOLSTER1_Offset_value = 1750
        GIB_Offset_value = -3000
        Behide_GIB_Offset_value = GIB_Offset_value - 334.5 - 65 + 109.85
        CLOCK_Offset_value = 2850
        CLOCK_SHAFT_Offset_value = CLOCK_Offset_value - 1350
        BOLSTER1_XY_Z = -par.T[i]
        BOLSTER1_XY_T = par.Z[i] - par.T[i] + BOLSTER1_XY_Z
        BOLSTER1_XY = BOLSTER1_XY_T + par.DH_S[i]
        BALANCER = -1700
        if l == 1:
            mprog.constaint_value_change(3, -par.B[i] - BOLSTER1_Offset_value, 1)
            mprog.constaint_value_change(6, -par.B[i] - BOLSTER1_Offset_value, 1)
            mprog.constaint_value_change(36, -BOLSTER1_Offset_value - par.F[i] / 2 + 80, 1)
            mprog.constaint_value_change(39, -BOLSTER1_Offset_value - par.F[i] / 2 + 80, 1)
            mprog.constaint_value_change(63, -par.F[i] / 2 - 80 - 37.5 - BOLSTER1_Offset_value, 1)
            mprog.constaint_value_change(66, -par.F[i] / 2 - 80 - 37.5 - BOLSTER1_Offset_value, 0)
            mprog.constaint_value_change(60, -37.5 - BOLSTER1_Offset_value, 1)
            mprog.constaint_value_change(69, -37.5 - BOLSTER1_Offset_value, 0)
            mprog.constaint_value_change(75, GIB_Offset_value + 45, 0)
            mprog.constaint_value_change(78, GIB_Offset_value + 45, 0)
            mprog.constaint_value_change(84, Behide_GIB_Offset_value, 1)
            mprog.constaint_value_change(87, Behide_GIB_Offset_value, 1)
            mprog.constaint_value_change(81, Behide_GIB_Offset_value, 0)
            mprog.constaint_value_change(96, Behide_GIB_Offset_value, 1)
            mprog.constaint_value_change(102, Behide_GIB_Offset_value, 1)
            mprog.constaint_value_change(99, Behide_GIB_Offset_value, 0)
            mprog.constaint_value_change(174, CLOCK_Offset_value, 0)
            mprog.constaint_value_change(188, CLOCK_SHAFT_Offset_value, 1)
            mprog.constaint_value_change(176, BALANCER, 1)  # 右氣壓缸
            mprog.constaint_value_change(180, BALANCER, 1)  # 左氣壓缸
            mprog.constaint_value_change(2, BOLSTER1_XY_Z, 1)
            mprog.constaint_value_change(5, BOLSTER1_XY_Z, 0)
            mprog.constaint_value_change(8, BOLSTER1_XY_Z, 1)
            mprog.constaint_value_change(11, BOLSTER1_XY_Z, 0)
            mprog.constaint_value_change(26, BOLSTER1_XY_T, 1)
            mprog.constaint_value_change(29, BOLSTER1_XY_T, 0)
            mprog.constaint_value_change(35, BOLSTER1_XY_T, 0)
            mprog.constaint_value_change(38, BOLSTER1_XY_T, 0)
            mprog.constaint_value_change(56, BOLSTER1_XY_T, 0)
            mprog.constaint_value_change(59, BOLSTER1_XY_T, 0)
            mprog.constaint_value_change(62, BOLSTER1_XY_T, 0)
            mprog.constaint_value_change(65, BOLSTER1_XY_T, 0)
            mprog.constaint_value_change(68, BOLSTER1_XY_T, 0)
            mprog.constaint_value_change(71, BOLSTER1_XY_T, 0)
            mprog.constaint_value_change(168, BOLSTER1_XY, 0)
        else:
            mprog.constaint_value_change(3, -par.B_15[i] - BOLSTER1_Offset_value, 1)
            mprog.constaint_value_change(6, -par.B_15[i] - BOLSTER1_Offset_value, 1)
            mprog.constaint_value_change(36, -BOLSTER1_Offset_value - par.F[i] / 2 + 80, 1)
            mprog.constaint_value_change(39, -BOLSTER1_Offset_value - par.F[i] / 2 + 80, 1)
            mprog.constaint_value_change(63, -par.F[i] / 2 - 80 - 37.5 - BOLSTER1_Offset_value, 1)
            mprog.constaint_value_change(66, -par.F[i] / 2 - 80 - 37.5 - BOLSTER1_Offset_value, 0)
            mprog.constaint_value_change(60, -37.5 - BOLSTER1_Offset_value, 1)
            mprog.constaint_value_change(69, -37.5 - BOLSTER1_Offset_value, 0)
            mprog.constaint_value_change(75, GIB_Offset_value + 45, 0)
            mprog.constaint_value_change(78, GIB_Offset_value + 45, 0)
            mprog.constaint_value_change(84, Behide_GIB_Offset_value, 1)
            mprog.constaint_value_change(87, Behide_GIB_Offset_value, 1)
            mprog.constaint_value_change(81, Behide_GIB_Offset_value, 0)
            mprog.constaint_value_change(96, Behide_GIB_Offset_value, 1)
            mprog.constaint_value_change(102, Behide_GIB_Offset_value, 1)
            mprog.constaint_value_change(99, Behide_GIB_Offset_value, 0)
            mprog.constaint_value_change(174, CLOCK_Offset_value, 0)
            mprog.constaint_value_change(188, CLOCK_SHAFT_Offset_value, 1)
            mprog.constaint_value_change(176, BALANCER, 1)  # 右氣壓缸
            mprog.constaint_value_change(180, BALANCER, 1)  # 左氣壓缸
            mprog.constaint_value_change(2, BOLSTER1_XY_Z, 1)
            mprog.constaint_value_change(5, BOLSTER1_XY_Z, 0)
            mprog.constaint_value_change(8, BOLSTER1_XY_Z, 1)
            mprog.constaint_value_change(11, BOLSTER1_XY_Z, 0)
            mprog.constaint_value_change(26, BOLSTER1_XY_T, 1)
            mprog.constaint_value_change(29, BOLSTER1_XY_T, 0)
            mprog.constaint_value_change(35, BOLSTER1_XY_T, 0)
            mprog.constaint_value_change(38, BOLSTER1_XY_T, 0)
            mprog.constaint_value_change(56, BOLSTER1_XY_T, 0)
            mprog.constaint_value_change(59, BOLSTER1_XY_T, 0)
            mprog.constaint_value_change(62, BOLSTER1_XY_T, 0)
            mprog.constaint_value_change(65, BOLSTER1_XY_T, 0)
            mprog.constaint_value_change(68, BOLSTER1_XY_T, 0)
            mprog.constaint_value_change(71, BOLSTER1_XY_T, 0)
            mprog.constaint_value_change(168, BOLSTER1_XY, 0)

        mprog.update()  # 更新
        mprog.OPEN_Drawing()
        if l == 1:
            drafting_Coordinate_Position, drafting_isometric_Coordinate_Position, scale = draft.drafting_parameter_calculation(1, par.A[i], par.B[i], par.H[i], par.T[i])  # 計算爆炸圖比例及位置
        else :
            drafting_Coordinate_Position, drafting_isometric_Coordinate_Position, scale = draft.drafting_parameter_calculation(
                0, par.A_15[i], par.B_15[i], par.H[i], par.T[i])  # 計算爆炸圖比例及位置

        draft.change_Drawing_scale(1 / scale)  # 圖面比例
        draft.exploded_Drawing_1(type, drafting_isometric_Coordinate_Position['exploded_1'][0],
                                 drafting_isometric_Coordinate_Position['exploded_1'][1], scale)  # 爆炸圖1
        mprog.switch_window()  # 開啟3D圖視窗
        # # 還原零件初始位置
        BOLSTER1_Offset_value = 80 - par.F[i] / 2
        GIB_Offset_value = 334.5 - 45
        CLOCK_Offset_value = 35
        CLOCK_SHAFT_Offset_value = 45
        Behide_GIB_Offset_value = GIB_Offset_value - 334.5 - 65 + 109.85
        BOLSTER1_XY_Z = - par.Z[i]
        BOLSTER1_XY_T = -par.T[i]
        BOLSTER1_XY = par.DH_S[i]
        BALANCER = -32
        if l == 1:
            mprog.constaint_value_change(3, -par.B[i] - BOLSTER1_Offset_value, 1)
            mprog.constaint_value_change(6, -par.B[i] - BOLSTER1_Offset_value, 1)
            mprog.constaint_value_change(36, -BOLSTER1_Offset_value - par.F[i] / 2 + 80, 1)
            mprog.constaint_value_change(39, -BOLSTER1_Offset_value - par.F[i] / 2 + 80, 1)
            mprog.constaint_value_change(63, -par.F[i] / 2 - 80 - 37.5 - BOLSTER1_Offset_value, 1)
            mprog.constaint_value_change(66, -par.F[i] / 2 - 80 - 37.5 - BOLSTER1_Offset_value, 0)
            mprog.constaint_value_change(60, -37.5 - BOLSTER1_Offset_value, 1)
            mprog.constaint_value_change(69, -37.5 - BOLSTER1_Offset_value, 0)
            mprog.constaint_value_change(75, GIB_Offset_value + 45, 0)
            mprog.constaint_value_change(78, GIB_Offset_value + 45, 0)
            mprog.constaint_value_change(84, Behide_GIB_Offset_value, 1)
            mprog.constaint_value_change(87, Behide_GIB_Offset_value, 1)
            mprog.constaint_value_change(81, Behide_GIB_Offset_value, 0)
            mprog.constaint_value_change(96, Behide_GIB_Offset_value, 1)
            mprog.constaint_value_change(102, Behide_GIB_Offset_value, 1)
            mprog.constaint_value_change(99, Behide_GIB_Offset_value, 0)
            mprog.constaint_value_change(174, CLOCK_Offset_value, 0)
            mprog.constaint_value_change(188, CLOCK_SHAFT_Offset_value, 1)
            mprog.constaint_value_change(176, BALANCER, 1)  # 右氣壓缸
            mprog.constaint_value_change(180, BALANCER, 1)  # 左氣壓缸
            mprog.constaint_value_change(2, BOLSTER1_XY_Z, 1)
            mprog.constaint_value_change(5, BOLSTER1_XY_Z, 0)
            mprog.constaint_value_change(8, BOLSTER1_XY_Z, 1)
            mprog.constaint_value_change(11, BOLSTER1_XY_Z, 0)
            mprog.constaint_value_change(26, BOLSTER1_XY_T, 1)
            mprog.constaint_value_change(29, BOLSTER1_XY_T, 0)
            mprog.constaint_value_change(35, BOLSTER1_XY_T, 0)
            mprog.constaint_value_change(38, BOLSTER1_XY_T, 0)
            mprog.constaint_value_change(56, BOLSTER1_XY_T, 0)
            mprog.constaint_value_change(59, BOLSTER1_XY_T, 0)
            mprog.constaint_value_change(62, BOLSTER1_XY_T, 0)
            mprog.constaint_value_change(65, BOLSTER1_XY_T, 0)
            mprog.constaint_value_change(68, BOLSTER1_XY_T, 0)
            mprog.constaint_value_change(71, BOLSTER1_XY_T, 0)
            mprog.constaint_value_change(168, BOLSTER1_XY, 0)
        else :
            mprog.constaint_value_change(3, -par.B_15[i] - BOLSTER1_Offset_value, 1)
            mprog.constaint_value_change(6, -par.B_15[i] - BOLSTER1_Offset_value, 1)
            mprog.constaint_value_change(36, -BOLSTER1_Offset_value - par.F[i] / 2 + 80, 1)
            mprog.constaint_value_change(39, -BOLSTER1_Offset_value - par.F[i] / 2 + 80, 1)
            mprog.constaint_value_change(63, -par.F[i] / 2 - 80 - 37.5 - BOLSTER1_Offset_value, 1)
            mprog.constaint_value_change(66, -par.F[i] / 2 - 80 - 37.5 - BOLSTER1_Offset_value, 0)
            mprog.constaint_value_change(60, -37.5 - BOLSTER1_Offset_value, 1)
            mprog.constaint_value_change(69, -37.5 - BOLSTER1_Offset_value, 0)
            mprog.constaint_value_change(75, GIB_Offset_value + 45, 0)
            mprog.constaint_value_change(78, GIB_Offset_value + 45, 0)
            mprog.constaint_value_change(84, Behide_GIB_Offset_value, 1)
            mprog.constaint_value_change(87, Behide_GIB_Offset_value, 1)
            mprog.constaint_value_change(81, Behide_GIB_Offset_value, 0)
            mprog.constaint_value_change(96, Behide_GIB_Offset_value, 1)
            mprog.constaint_value_change(102, Behide_GIB_Offset_value, 1)
            mprog.constaint_value_change(99, Behide_GIB_Offset_value, 0)
            mprog.constaint_value_change(174, CLOCK_Offset_value, 0)
            mprog.constaint_value_change(188, CLOCK_SHAFT_Offset_value, 1)
            mprog.constaint_value_change(176, BALANCER, 1)  # 右氣壓缸
            mprog.constaint_value_change(180, BALANCER, 1)  # 左氣壓缸
            mprog.constaint_value_change(2, BOLSTER1_XY_Z, 1)
            mprog.constaint_value_change(5, BOLSTER1_XY_Z, 0)
            mprog.constaint_value_change(8, BOLSTER1_XY_Z, 1)
            mprog.constaint_value_change(11, BOLSTER1_XY_Z, 0)
            mprog.constaint_value_change(26, BOLSTER1_XY_T, 1)
            mprog.constaint_value_change(29, BOLSTER1_XY_T, 0)
            mprog.constaint_value_change(35, BOLSTER1_XY_T, 0)
            mprog.constaint_value_change(38, BOLSTER1_XY_T, 0)
            mprog.constaint_value_change(56, BOLSTER1_XY_T, 0)
            mprog.constaint_value_change(59, BOLSTER1_XY_T, 0)
            mprog.constaint_value_change(62, BOLSTER1_XY_T, 0)
            mprog.constaint_value_change(65, BOLSTER1_XY_T, 0)
            mprog.constaint_value_change(68, BOLSTER1_XY_T, 0)
            mprog.constaint_value_change(71, BOLSTER1_XY_T, 0)
            mprog.constaint_value_change(168, BOLSTER1_XY, 0)
        mprog.update()
        # # # --------------------爆炸圖右圖------------------

        Fixture_offset_value = 2850
        CLUCTH_ASSEMBLY_offset_value = par.B[i] + 1650
        CLUCTH_ASSEMBLY_offset_15_value = par.B_15[i] + 1650
        JOINT_ALL_offset_value = CLUCTH_ASSEMBLY_offset_value
        MAIN_GEAR_offset_value = par.B[i] + 850
        MAIN_GEAR_offset_15_value = par.B_15[i] + 850
        if l == 1:
            mprog.constaint_value_change(159, -312.5 - Fixture_offset_value, 0)  # 支架
            mprog.constaint_value_change(186, CLUCTH_ASSEMBLY_offset_value, 1)  # 離合器
            mprog.constaint_value_change(203, JOINT_ALL_offset_value, 1)  # JOINT_All.1
            mprog.constaint_value_change(195, MAIN_GEAR_offset_value, 1)  # 大齒輪MAIN_GEA1
            mprog.constaint_value_change(201, -375, 1)  # Joint1
        else:
            mprog.constaint_value_change(159, -312.5 - Fixture_offset_value, 0)  # 支架
            mprog.constaint_value_change(186, CLUCTH_ASSEMBLY_offset_15_value, 1)  # 離合器
            mprog.constaint_value_change(203, JOINT_ALL_offset_value, 1)  # JOINT_All.1
            mprog.constaint_value_change(195, MAIN_GEAR_offset_15_value, 1)  # 大齒輪MAIN_GEA1
            mprog.constaint_value_change(201, -375, 1)  # Joint1
        mprog.update()
        mprog.switch_window()
        draft.exploded_Drawing_2(type, drafting_isometric_Coordinate_Position['exploded_2'][0],
                                 drafting_isometric_Coordinate_Position['exploded_2'][1], scale)  # 爆炸圖2
        # 復原位置
        mprog.switch_window()
        mprog.constaint_value_change(159, 312.5, 0)  # 支架
        mprog.constaint_value_change(186, 510, 1)  # 離合器
        mprog.constaint_value_change(203, 1010, 1)  # JOINT_All.1
        mprog.constaint_value_change(195, 592.5 + 150, 1)  # 大齒輪MAIN_GEA1
        mprog.constaint_value_change(201, -84, 1)  # Joint1
        mprog.update()
        mprog.switch_window()
        # # --------------------爆炸圖下(前、左、右)----------------
        draft.Front_View_Drawing(type, drafting_Coordinate_Position['Front View'][0],
                                 drafting_Coordinate_Position['Front View'][1], scale)
        draft.Left_View_Drawing(type, drafting_Coordinate_Position['Left View'][0],
                                drafting_Coordinate_Position['Left View'][1], scale)
        draft.Right_View_Drawing(type, drafting_Coordinate_Position['Right View'][0],
                                 drafting_Coordinate_Position['Right View'][1], scale)

        # ------------圈碼圖------------
        sin45 = math.sin(math.radians(45))
        cos45 = math.cos(math.radians(45))
        sin30 = math.sin(math.radians(30))
        cos30 = math.cos(math.radians(30))
        sin60 = math.sin(math.radians(60))
        cos60 = math.cos(math.radians(60))
        cos35_267 = math.cos(math.radians(35.267))
        sin35_267 = math.sin(math.radians(35.267))

        BOLSTER1_Offset_value = 1750
        BLOSTER1_center_Offset_Value = 1775
        GIB_Offset_value = -3000
        Behide_GIB_Offset_value = GIB_Offset_value - 334.5 - 65 + 109.85
        CLOCK_Offset_value = 2850
        CLOCK_SHAFT_Offset_value = CLOCK_Offset_value - 1350
        BOLSTER1_XY_Z = -par.T[i]
        BOLSTER1_XY_T = par.Z[i] - par.T[i] + BOLSTER1_XY_Z
        BOLSTER1_XY = BOLSTER1_XY_T + par.DH_S[i]

        balloons_posotion = [750 * cos30, 750 * sin30]
        CLOCK_Pointer = 3325
        FRAME_TOP_LEFT_X = (par.R[i] + 90 + 50) * cos45 + 50
        FRAME_TOP_LEFT_X_1 = FRAME_TOP_LEFT_X * sin30 / cos30

        scale = 1 / scale

        # (將3D圖面尺寸轉換為2D尺寸, 將圖面x座標尺寸轉為R再將R轉為Y)
        # 引線點座標
        if l == 1:
            point_position = {'2': [-CLOCK_Offset_value * cos45,
                                    -CLOCK_Offset_value * cos45 * sin30 / cos30],
                              '3': [(-1250) * cos45,
                                    (-1250) * cos45 * sin30 / cos30],
                              '4': [(-BOLSTER1_Offset_value - par.R[i] / 2) * cos45,
                                    (-BOLSTER1_Offset_value - par.R[i] / 2) * cos45 * sin30 / cos30],
                              '5': [(-BOLSTER1_Offset_value - par.R[i]) * cos45,
                                    (-BOLSTER1_Offset_value - par.R[i]) * cos45 * sin30 / cos30 - par.S[i] + par.T[i]],
                              '6': [(GIB_Offset_value - (par.R[i] + 145) / 2) * cos45,
                                    (GIB_Offset_value - (par.R[i] + 145) / 2) * cos45 * sin30 / cos30]}
            # 圈圈座標`
            circle_position = {
                '2': ['2', point_position['2'][0] - FRAME_TOP_LEFT_X, point_position['2'][1] + FRAME_TOP_LEFT_X_1],
                '3': ['3', point_position['3'][0] - FRAME_TOP_LEFT_X, point_position['3'][1] + FRAME_TOP_LEFT_X_1],
                '4': ['4', point_position['4'][0] - FRAME_TOP_LEFT_X, point_position['4'][1] + FRAME_TOP_LEFT_X_1],
                '5': ['5', point_position['5'][0] * 2 * 0.9 - FRAME_TOP_LEFT_X,
                      point_position['5'][0] * 2 * 0.9 * sin30 / cos30 + FRAME_TOP_LEFT_X_1],
                '6': ['6', point_position['6'][0] - FRAME_TOP_LEFT_X, point_position['6'][1] + FRAME_TOP_LEFT_X_1]}

            draft.balloons('Isometric view1', circle_position['6'][0], circle_position['6'][1], circle_position['6'][2],
                           point_position['6'][0], point_position['6'][1], circle_position['6'][1])
            draft.balloons('Isometric view1', circle_position['2'][0], circle_position['2'][1], circle_position['2'][2],
                           point_position['2'][0], point_position['2'][1], circle_position['2'][1])
            draft.balloons('Isometric view1', circle_position['4'][0], circle_position['4'][1], circle_position['4'][2],
                           point_position['4'][0], point_position['4'][1], circle_position['4'][1])
            draft.balloons('Isometric view1', circle_position['5'][0], circle_position['5'][1], circle_position['5'][2],
                           point_position['5'][0], point_position['5'][1], circle_position['5'][1])
            draft.balloons('Isometric view1', circle_position['3'][0], circle_position['3'][1], circle_position['3'][2],
                           point_position['3'][0], point_position['3'][1], circle_position['3'][1])

            # ------------中心線-------------
            draft.create_center_line('Isometric view1', 0, 0, -CLOCK_Pointer * cos45,
                                     -CLOCK_Pointer * cos45 * sin30 / cos30)
            draft.create_center_line('Isometric view1', -BLOSTER1_center_Offset_Value * cos45,
                                     -BLOSTER1_center_Offset_Value * cos45 * sin30 / cos30,
                                     -BLOSTER1_center_Offset_Value * cos45, -par.S[i] - par.Z[i] - par.T[i])
            draft.create_center_line('Isometric view2', -742.5 * cos45, -742.5 * cos45 * sin30 / cos30,
                                     -JOINT_ALL_offset_value * cos45, -JOINT_ALL_offset_value * cos45 * sin30 / cos30)
            draft.create_center_line('Isometric view2', -742.5 * cos45, -742.5 * cos45 * sin30 / cos30 - 300,
                                     -(Fixture_offset_value + par.B[i]) * cos45,
                                     -(Fixture_offset_value + par.B[i]) * cos45 * sin30 / cos30 - 300)
        else:
            point_position = {'2': [-CLOCK_Offset_value * cos45,
                                    -CLOCK_Offset_value * cos45 * sin30 / cos30],
                              '3': [(-1250) * cos45,
                                    (-1250) * cos45 * sin30 / cos30],
                              '4': [(-BOLSTER1_Offset_value - par.R_15[i] / 2) * cos45,
                                    (-BOLSTER1_Offset_value - par.R_15[i] / 2) * cos45 * sin30 / cos30],
                              '5': [(-BOLSTER1_Offset_value - par.R_15[i]) * cos45,
                                    (-BOLSTER1_Offset_value - par.R_15[i]) * cos45 * sin30 / cos30 - par.S[i] + par.T[i]],
                              '6': [(GIB_Offset_value - (par.R_15[i] + 145) / 2) * cos45,
                                    (GIB_Offset_value - (par.R_15[i] + 145) / 2) * cos45 * sin30 / cos30]}
            # 圈圈座標`
            circle_position = {
                '2': ['2', point_position['2'][0] - FRAME_TOP_LEFT_X, point_position['2'][1] + FRAME_TOP_LEFT_X_1],
                '3': ['3', point_position['3'][0] - FRAME_TOP_LEFT_X, point_position['3'][1] + FRAME_TOP_LEFT_X_1],
                '4': ['4', point_position['4'][0] - FRAME_TOP_LEFT_X, point_position['4'][1] + FRAME_TOP_LEFT_X_1],
                '5': ['5', point_position['5'][0] * 2 * 0.9 - FRAME_TOP_LEFT_X,
                      point_position['5'][0] * 2 * 0.9 * sin30 / cos30 + FRAME_TOP_LEFT_X_1],
                '6': ['6', point_position['6'][0] - FRAME_TOP_LEFT_X, point_position['6'][1] + FRAME_TOP_LEFT_X_1]}

            draft.balloons('Isometric view1', circle_position['6'][0], circle_position['6'][1], circle_position['6'][2],
                           point_position['6'][0], point_position['6'][1], circle_position['6'][1])
            draft.balloons('Isometric view1', circle_position['2'][0], circle_position['2'][1], circle_position['2'][2],
                           point_position['2'][0], point_position['2'][1], circle_position['2'][1])
            draft.balloons('Isometric view1', circle_position['4'][0], circle_position['4'][1], circle_position['4'][2],
                           point_position['4'][0], point_position['4'][1], circle_position['4'][1])
            draft.balloons('Isometric view1', circle_position['5'][0], circle_position['5'][1], circle_position['5'][2],
                           point_position['5'][0], point_position['5'][1], circle_position['5'][1])
            draft.balloons('Isometric view1', circle_position['3'][0], circle_position['3'][1], circle_position['3'][2],
                           point_position['3'][0], point_position['3'][1], circle_position['3'][1])

            # ------------中心線-------------
            draft.create_center_line('Isometric view1', 0, 0, -CLOCK_Pointer * cos45,
                                     -CLOCK_Pointer * cos45 * sin30 / cos30)
            draft.create_center_line('Isometric view1', -BLOSTER1_center_Offset_Value * cos45,
                                     -BLOSTER1_center_Offset_Value * cos45 * sin30 / cos30,
                                     -BLOSTER1_center_Offset_Value * cos45, -par.S[i] - par.Z[i] - par.T[i])
            draft.create_center_line('Isometric view2', -742.5 * cos45, -742.5 * cos45 * sin30 / cos30,
                                     -JOINT_ALL_offset_value * cos45, -JOINT_ALL_offset_value * cos45 * sin30 / cos30)
            draft.create_center_line('Isometric view2', -742.5 * cos45, -742.5 * cos45 * sin30 / cos30 - 300,
                                     -(Fixture_offset_value + par.B_15[i]) * cos45,
                                     -(Fixture_offset_value + par.B_15[i]) * cos45 * sin30 / cos30 - 300)
        Left_view_list = []
        for x in Left_view_list:
            Left_view_list.append(x)

        for x in Left_view_list:
            draft.add_dimension_to_view('Left view', str(x), Left_view_list[x][0], Left_view_list[x][1], Left_view_list[x][2],
                                        Left_view_list[x][3], Left_view_list[x][4], Left_view_list[x][5])
        # -------------關閉虛框---------------
        # view_name = ['Isometric view1', 'Isometric view2', 'Front view', 'Left view', 'Right view']
        # for x in view_name:
        #     draft.close_broken_line_block_diagram(x)

        #--------------存檔------------------
        mprog.PDF_save(path, "Exploded_Views")
        mprog.save_detail_drawing(path , "Exploded_Views")

        #-------------
        # dp.Parts_drawing_generation(i , j , path)

        #-----------零件圖生成--------
        # mprog.OPEN_detail_drawing()
        # dp.Parts_drawing_generation(i, l, part_path)
        # dpc.bom_text_create()
        # mprog.PDF_save(path, "detail_drawing")
        # mprog.save_detail_drawing(path , "detail_drawing")

        # --------焊接圖---------
        # 隱藏機架外零件
        for part_name in par.hide_part_name:
            mprog.hide_show_part(part_name, 1)

        # 投影
        mprog.switch_window()
        mprog.OPEN_Welding_diagram()
        if l == 1:
            drafting_down_Coordinate_Position, drafting_up_Coordinate_Position, scale = draft.drafting_welding_view_parameter_calculation(
                par.A[i], par.B[i], par.H[i], par.S[i], par.Z[i], par.T[i])
        else:
            drafting_down_Coordinate_Position, drafting_up_Coordinate_Position, scale = draft.drafting_welding_view_parameter_calculation(
                par.A_15[i], par.B_15[i], par.H[i], par.S[i], par.Z[i], par.T[i])

        draft.change_Drawing_scale(1 / scale)
        # 副程式名稱須注意是否與立體圖相符
        draft.Front_View_Drawing(type, drafting_down_Coordinate_Position['Right View'][0],
                                 drafting_down_Coordinate_Position['Right View'][1], scale)  # 左側試圖
        draft.Right_View_Drawing(type, drafting_down_Coordinate_Position['Front View'][0],
                                 drafting_down_Coordinate_Position['Front View'][1], scale)  # 前視圖
        draft.Right_Top_View_Drawing(type, drafting_up_Coordinate_Position['Top View'][0],
                                     drafting_up_Coordinate_Position['Top View'][1], scale)  # 上視圖
        if l == 1:
            # 剖面圖座標
            Section_Coordinate = {'A-A': [
                [float(0), float(par.H[i] - par.S[i] - par.Z[i] + 150), float(0), float(-par.S[i] - par.Z[i] - 150)],
                0],
                                  'B-B': [[float(par.B[i] + 100), float(par.H[i] - par.S[i] - par.Z[i] + 150), float(par.B[i] + 100), float(-par.S[i] / 2 - 150)], 0],
                                  'C-C': [
                                      [float(-195), float(-par.S[i] + 100), float(par.B[i] + 100), float(-par.S[i] + 100)],
                                      0],
                                  'D-D': [[float(-par.A[i] / 2 - 100), float(-par.H[i] / 2), float(par.A[i] / 2 + 100),
                                           float(-par.H[i] / 2)], 1],
                                  'E-E': [[float(-100), float(-(par.S[i] + par.Z[i]) + 100), float(par.B[i] / 3),
                                           float(-(par.S[i] + par.Z[i]) + 100)], 0],
                                  'F-F': [[float(272), float(- 176 / 2), float(272), float(-176 * 3 / 2)], 0],
                                  'G-G': [
                                      [float(par.B[i] + 100), float(-(par.S[i] + par.Z[i]) + 200), float(par.B[i] + 100),
                                       float(-(par.S[i] + par.Z[i]) - 200)], 0]}
            Section_Coordinate_title = list(Section_Coordinate.keys())
            view_order = ['Front view', 'Right view', 'Section view A-A', 'Front view', 'Right view', 'Section view B-B',
                          'Section view A-A']
            for title in Section_Coordinate_title:
                for view_name in view_order:
                    if 'Section view ' + title in drafting_up_Coordinate_Position:  # 若此元素為此字典之鍵則執行
                        draft.Section(view_name, drafting_up_Coordinate_Position['Section view ' + title][0],
                                      drafting_up_Coordinate_Position['Section view ' + title][1], scale,
                                      Section_Coordinate[title][0],
                                      Section_Coordinate[title][1], 'Section view ' + title)
                    else:
                        draft.Section(view_name, drafting_down_Coordinate_Position['Section view ' + title][0],
                                      drafting_down_Coordinate_Position['Section view ' + title][1], scale,
                                      Section_Coordinate[title][0],
                                      Section_Coordinate[title][1], 'Section view ' + title)
                    view_order.pop(0)  # 刪除串列中第一個元素
                    break

            # # 裁切選定圖面範圍
            detail_view_coordinate = {
                'Section view D-D': [float(par.A[i] / 2), float(-270), float(-par.A[i] / 2), float(-270),
                                     float(-par.A[i] / 2),
                                     float(-630), float(par.A[i] / 2), float(-630)],
                'Section view F-F': [float(-979 + 42), float(-25), float(-979 + 42), float(-290), float(-54 - 979),
                                     float(-290),
                                     float(-979 - 54), float(-25)],
                'Section view G-G': [float(-par.A[i] / 2 - par.FRAME_6_7_width[i]), float(-par.S[i] - par.Z[i] + 150),
                                     float(-par.A[i] / 2 - par.FRAME_6_7_width[i]),
                                     float(-par.S[i] - par.Z[i] - 150), float(-par.A[i] / 2 + par.FRAME_6_7_width[i] + 170),
                                     float(-par.S[i] - par.Z[i] - 150), float(-par.A[i] / 2 + par.FRAME_6_7_width[i] + 170),
                                     float(-par.S[i] - par.Z[i] + 150)]}
            detail_view_coordinate_keys = list(detail_view_coordinate.keys())
            for coordinate_keys in detail_view_coordinate_keys:
                if coordinate_keys == 'Section view F-F' or coordinate_keys == 'Section view D-D':
                    mprog.switch_window()
                    if coordinate_keys == 'Section view F-F':
                        for name in par.part_name_Section_view_F_F:
                            mprog.hide_show_part(name, 1)
                    elif coordinate_keys == 'Section view D-D':
                        for name in par.part_name_Section_view_D_D:
                            mprog.hide_show_part(name, 1)
                    mprog.switch_window()
                    if coordinate_keys in drafting_up_Coordinate_Position:
                        draft.Define_Polygonal_Detail_View(coordinate_keys, detail_view_coordinate[coordinate_keys],
                                                           drafting_up_Coordinate_Position[coordinate_keys][0],
                                                           drafting_up_Coordinate_Position[coordinate_keys][1])
                    else:
                        draft.Define_Polygonal_Detail_View(coordinate_keys, detail_view_coordinate[coordinate_keys],
                                                           drafting_down_Coordinate_Position[coordinate_keys][0],
                                                           drafting_down_Coordinate_Position[coordinate_keys][1])
                    mprog.switch_window()
                    if coordinate_keys == 'Section view F-F':
                        for name in par.part_name_Section_view_F_F:
                            mprog.hide_show_part(name, 0)
                    elif coordinate_keys == 'Section view D-D':
                        for name in par.part_name_Section_view_D_D:
                            mprog.hide_show_part(name, 0)
                    mprog.switch_window()
                else:
                    if coordinate_keys in drafting_up_Coordinate_Position:
                        draft.Define_Polygonal_Detail_View(coordinate_keys, detail_view_coordinate[coordinate_keys],
                                                           drafting_up_Coordinate_Position[coordinate_keys][0],
                                                           drafting_up_Coordinate_Position[coordinate_keys][1])
                    else:
                        draft.Define_Polygonal_Detail_View(coordinate_keys, detail_view_coordinate[coordinate_keys],
                                                           drafting_down_Coordinate_Position[coordinate_keys][0],
                                                           drafting_down_Coordinate_Position[coordinate_keys][1])

            # 產生折斷線
            # 注意:折斷線若為水平線，方框四角座標須以⌈水平⌋建立；折斷線若為垂直線，方框四角座標須以⌈垂直⌋建立
            break_line_Coordinate = {'Section view E-E': [
                [float(0), float(par.A[i] / 6), float(par.B[i] / 3), float(par.A[i] / 6), float(0), float(-par.A[i] / 6),
                 float(par.B[i] / 3), float(-par.A[i] / 6)], 0, 1]}
            print(break_line_Coordinate['Section view E-E'][0][0:])
            draft.break_line('Section view E-E', break_line_Coordinate['Section view E-E'][0],
                             break_line_Coordinate['Section view E-E'][1], break_line_Coordinate['Section view E-E'][2])

            # 關閉所有剖線線所產生之虛線
            draft.close_all_Generated_Shape()

            # 產生防漏試驗之虛線
            leakproof_broken_line = {
                'Front view': [[float(-par.R[i] / 2 - 20), float(0),
                                float(-par.R[i] / 2 - 20), float(-par.S[i] / 2 - 200)], 0,
                               'Section view H-H']}
            draft.Section('Front view', drafting_down_Coordinate_Position['Section view A-A'][0],
                          drafting_down_Coordinate_Position['Section view A-A'][1], scale,
                          leakproof_broken_line['Front view'][0], leakproof_broken_line['Front view'][1],
                          leakproof_broken_line['Front view'][2])
            Detail_view_leak_broken_line = [float(550), float(par.FRAME_10_11_center_to_Y_1[i]), float(550), float(par.FRAME_10_11_center_to_Y_2[i]), float(978), float(par.FRAME_10_11_center_to_Y_2[i]),
                                            float(978), float(par.FRAME_10_11_center_to_Y_1[i])]
            draft.Define_Polygonal_Cipping_View('Section view H-H', Detail_view_leak_broken_line)
            draft.selection_Search_delete('Front view', "Name='Callout (Section View).3', all")

            # 焊接符號生成
            draft.symbol_of_weld('Front view', 3, -par.R[i] / 2 - 140, par.H[i] - par.S[i] - par.Z[i] - 32, 1,
                                 par.drafting_Welding_text['A-A Top'])
        else:
            # 剖面圖座標
            Section_Coordinate = {'A-A': [
                [float(0), float(par.H[i] - par.S[i] - par.Z[i] + 150), float(0), float(-par.S[i] - par.Z[i] - 150)],
                0],
                                  'B-B': [[float(par.B_15[i] + 100), float(520), float(par.B_15[i] + 100), float(-1200)], 0],
                                  'C-C': [
                                      [float(-195), float(-par.S[i] + 100), float(par.B_15[i] + 100),
                                       float(-par.S[i] + 100)],
                                      0],
                                  'D-D': [[float(-par.A_15[i] / 2 - 100), float(-par.H[i] / 2), float(par.A_15[i] / 2 + 100),
                                           float(-par.H[i] / 2)], 1],
                                  'E-E': [[float(-100), float(-(par.S[i] + par.Z[i]) + 100), float(par.B_15[i] / 3),
                                           float(-(par.S[i] + par.Z[i]) + 100)], 0],
                                  'F-F': [[float(272), float(- 176 / 2), float(272), float(-176 * 3 / 2)], 0],
                                  'G-G': [
                                      [float(par.B_15[i] + 100), float(-(par.S[i] + par.Z[i]) + 200),
                                       float(par.B_15[i] + 100),
                                       float(-(par.S[i] + par.Z[i]) - 200)], 0]}
            Section_Coordinate_title = list(Section_Coordinate.keys())
            view_order = ['Front view', 'Right view', 'Section view A-A', 'Front view', 'Right view',
                          'Section view B-B',
                          'Section view A-A']
            for title in Section_Coordinate_title:
                for view_name in view_order:
                    if 'Section view ' + title in drafting_up_Coordinate_Position:  # 若此元素為此字典之鍵則執行
                        draft.Section(view_name, drafting_up_Coordinate_Position['Section view ' + title][0],
                                      drafting_up_Coordinate_Position['Section view ' + title][1], scale,
                                      Section_Coordinate[title][0],
                                      Section_Coordinate[title][1], 'Section view ' + title)
                    else:
                        draft.Section(view_name, drafting_down_Coordinate_Position['Section view ' + title][0],
                                      drafting_down_Coordinate_Position['Section view ' + title][1], scale,
                                      Section_Coordinate[title][0],
                                      Section_Coordinate[title][1], 'Section view ' + title)
                    view_order.pop(0)  # 刪除串列中第一個元素
                    break

            # # 裁切選定圖面範圍
            detail_view_coordinate = {
                'Section view D-D': [float(par.A_15[i] / 2), float(-270), float(-par.A_15[i] / 2), float(-270),
                                     float(-par.A_15[i] / 2),
                                     float(-630), float(par.A_15[i] / 2), float(-630)],
                'Section view F-F': [float(-979 + 42), float(-25), float(-979 + 42), float(-290), float(-54 - 979),
                                     float(-290),
                                     float(-979 - 54), float(-25)],
                'Section view G-G': [float(-par.A_15[i] / 2 - par.FRAME_6_7_width[i]), float(-par.S[i] - par.Z[i] + 150),
                                     float(-par.A_15[i] / 2 - par.FRAME_6_7_width[i]),
                                     float(-par.S[i] - par.Z[i] - 150),
                                     float(-par.A_15[i] / 2 + par.FRAME_6_7_width[i] + 170),
                                     float(-par.S[i] - par.Z[i] - 150),
                                     float(-par.A_15[i] / 2 + par.FRAME_6_7_width[i] + 170),
                                     float(-par.S[i] - par.Z[i] + 150)]}
            detail_view_coordinate_keys = list(detail_view_coordinate.keys())
            for coordinate_keys in detail_view_coordinate_keys:
                if coordinate_keys == 'Section view F-F' or coordinate_keys == 'Section view D-D':
                    mprog.switch_window()
                    if coordinate_keys == 'Section view F-F':
                        for name in par.part_name_Section_view_F_F:
                            mprog.hide_show_part(name, 1)
                    elif coordinate_keys == 'Section view D-D':
                        for name in par.part_name_Section_view_D_D:
                            mprog.hide_show_part(name, 1)
                    mprog.switch_window()
                    if coordinate_keys in drafting_up_Coordinate_Position:
                        draft.Define_Polygonal_Detail_View(coordinate_keys, detail_view_coordinate[coordinate_keys],
                                                           drafting_up_Coordinate_Position[coordinate_keys][0],
                                                           drafting_up_Coordinate_Position[coordinate_keys][1])
                    else:
                        draft.Define_Polygonal_Detail_View(coordinate_keys, detail_view_coordinate[coordinate_keys],
                                                           drafting_down_Coordinate_Position[coordinate_keys][0],
                                                           drafting_down_Coordinate_Position[coordinate_keys][1])
                    mprog.switch_window()
                    if coordinate_keys == 'Section view F-F':
                        for name in par.part_name_Section_view_F_F:
                            mprog.hide_show_part(name, 0)
                    elif coordinate_keys == 'Section view D-D':
                        for name in par.part_name_Section_view_D_D:
                            mprog.hide_show_part(name, 0)
                    mprog.switch_window()
                else:
                    if coordinate_keys in drafting_up_Coordinate_Position:
                        draft.Define_Polygonal_Detail_View(coordinate_keys, detail_view_coordinate[coordinate_keys],
                                                           drafting_up_Coordinate_Position[coordinate_keys][0],
                                                           drafting_up_Coordinate_Position[coordinate_keys][1])
                    else:
                        draft.Define_Polygonal_Detail_View(coordinate_keys, detail_view_coordinate[coordinate_keys],
                                                           drafting_down_Coordinate_Position[coordinate_keys][0],
                                                           drafting_down_Coordinate_Position[coordinate_keys][1])

            # 產生折斷線
            # 注意:折斷線若為水平線，方框四角座標須以⌈水平⌋建立；折斷線若為垂直線，方框四角座標須以⌈垂直⌋建立
            break_line_Coordinate = {'Section view E-E': [
                [float(0), float(par.A_15[i] / 6), float(par.B_15[i] / 3), float(par.A_15[i] / 6), float(0),
                 float(-par.A_15[i] / 6),
                 float(par.B_15[i] / 3), float(-par.A_15[i] / 6)], 0, 1]}
            print(break_line_Coordinate['Section view E-E'][0][0:])
            draft.break_line('Section view E-E', break_line_Coordinate['Section view E-E'][0],
                             break_line_Coordinate['Section view E-E'][1], break_line_Coordinate['Section view E-E'][2])

            # 關閉所有剖線線所產生之虛線
            draft.close_all_Generated_Shape()

            # 產生防漏試驗之虛線
            leakproof_broken_line = {
                'Front view': [[float(-par.R_15[i] / 2 - 20), float(0),
                                float(-par.R_15[i] / 2 - 20), float(-par.S[i] / 2 - 200)], 0,
                               'Section view H-H']}
            draft.Section('Front view', drafting_down_Coordinate_Position['Section view A-A'][0],
                          drafting_down_Coordinate_Position['Section view A-A'][1], scale,
                          leakproof_broken_line['Front view'][0], leakproof_broken_line['Front view'][1],
                          leakproof_broken_line['Front view'][2])
            Detail_view_leak_broken_line = [float(550), float(par.FRAME_10_11_center_to_Y_1[i]), float(550), float(par.FRAME_10_11_center_to_Y_2[i]), float(978), float(par.FRAME_10_11_center_to_Y_2[i]),
                                                    float(978),
                                                    float(par.FRAME_10_11_center_to_Y_1[i])]
            draft.Define_Polygonal_Cipping_View('Section view H-H', Detail_view_leak_broken_line)
            draft.selection_Search_delete('Front view', "Name='Callout (Section View).3', all")

            # 焊接符號生成
            draft.symbol_of_weld('Front view', 3, -par.R_15[i] / 2 - 140, par.H[i] - par.S[i] - par.Z[i] - 32, 1,
                                 par.drafting_Welding_text['A-A Top'])
        mprog.PDF_save(path, "Welding_diagram")
        mprog.save_detail_drawing(path, "Welding_diagram")

if __name__ == "__main__":
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)  # 自適應屏幕分辨率
    app = QtWidgets.QApplication([])
    window = main()
    window.show()
    sys.exit(app.exec_())