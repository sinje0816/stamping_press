from PyQt5 import QtCore, QtGui, QtWidgets
import sys, datetime, os
from GUI import Ui_Dialog
import main_program as mprog
import para

change = ()
height = ()
type = ()
hole = ()
close = ()

# 電子型錄規格
A = [720, 830, 890, 940, 1050, 1160, 1300, 1480, 1560, 1760]
A_15 = [1080, 1245, 1335, 1410, 1575, 1740, 1950, 2220, 2340, 2640]
B = [1058, 1125, 1210, 1315, 1480, 1680, 1985, 2113, 2400, 2700]
B_15 = [1587, 1688, 1815, 1973, 2220, 2520, 2978, 3170, 3600, 4050]
H = [2060, 2185, 2290, 2540, 2755, 2990, 3270, 3725, 4005, 4285]
R = [388, 486, 516, 544, 614, 670, 730, 900, 970, 1040]
R_15 = [582, 729, 774, 816, 921, 1005, 1095, 1350, 1455, 1560]
E = [700, 780, 840, 900, 1050, 1150, 1250, 1400, 1500, 1600]
E_15 = [1051, 1170, 1260, 1350, 1575, 1725, 1875, 2100, 2250, 2400]
D_DH = [250, 280, 330, 350, 380, 430, 490, 550, 580, 610]
D_S = [80, 90, 110, 130, 150, 180, 200, 220, 250, 280]
D_H = [50, 60, 70, 80, 100, 110, 130, 150, 180, 210]
D_P = [35, 40, 45, 50, 60, 70, 80, 90, 100, 110]
DH_S = [230, 250, 270, 300, 330, 350, 400, 450, 450, 500]
DH_H = [200, 220, 240, 270, 300, 320, 360, 400, 400, 450]
DH_P = [200, 220, 240, 270, 300, 320, 360, 400, 400, 450]
S = [983, 1068, 1158, 1285, 1445, 1630, 1809, 2067, 2262, 2357]
H_Z = [1260, 1385, 1490, 1640, 1855, 2086.933, 2370, 2725, 3005, 3285]
O = [1045, 1075, 1125, 1145, 1175, 1225, 1285, 1345, 1375, 1405]
hole_type = [0, 1, 2]
P = [330, 380, 430, 480, 560, 650, 720, 860, 960, 1060]
P_15 = [495, 570, 645, 721, 840, 975, 1080, 1290, 1440, 1590]
Q = [250, 300, 350, 400, 460, 520, 580, 650, 720, 790]
Q_15 = [375, 450, 525, 600, 690, 780, 870, 975, 1080, 1185]
T = [85, 100, 115, 130, 140, 155, 165, 180, 180, 200]
Z = [800, 800, 800, 900, 900, 900, 900, 1000, 1000, 1100]
F = [320, 400, 440, 520, 600, 680, 760, 840, 900, 960]
FRAME20_H = [280, 355, 420, 437.5, 550, 680, 825, 990, 1120, 1170]
FRAME2_lower_depth = [166.016, 246.016, 286.016, 366.016, 446.016, 526.016, 606.016, 686.016, 746.016, 806.016]
FRAME2_lower_depth_15 = [331.016, 451.016, 511.016, 631.016, 751.016, 871.016, 991.016, 1111.016, 1201.016, 1291.016]
FRAME1_lower_high = [1330, 1335, 1340, 1445, 1455, 1460, 1470, 1575, 1595, 1695]
FRAME20_FRAME2_YZ = [805, 805, 979, 979, 979, 979, 979, 979, 979, 979]
BALANCER1_XZ = [204, 253, 268, 282, 317, 345, 375, 460, 575, 690]  # (R+180mm)/2+80mm
FRAME_10_H = [558, 620.5, 673, 798, 905.5, 1023, 1163, 1320.5, 1530.5, 1740.5]
FRAME_32_XY = [0, 0, 0, 1722, 1819.5, 1922, 2052, 2294.5, 2504.5, 2794.5]


class main(QtWidgets.QWidget, Ui_Dialog):
    def __init__(self):
        super(main, self).__init__()
        self.ui = Ui_Dialog()
        self.ui.setupUi(self)
        self.ui.pushButton.clicked.connect(self.start)
        self.add_item_for_comboBox()
        self.path = str()
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
        self.change_dir(self.l, self.h, self.i, self.j, self.path)
        self.ass_(self.l, self.h, self.i, self.j, self.path)

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

    def create_dir(self, type):
        time_now = datetime.datetime.now()
        dir_name = '{}_{}_{}_{}_{}'.format(type, time_now.day, time_now.hour, time_now.minute, time_now.second)
        desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
        path = desktop + '\\' + dir_name
        os.mkdir(path)
        self.path = path

    def change_dir(self, h, i, l, j, path):

        # 開啟CATIA
        env = mprog.set_CATIA_workbench_env()

        # 匯入零件檔
        file_name_list = ['BOLSTER1', 'BOLSTER2', 'BOLSTER3', 'Fixture', 'FRAME1', 'FRAME2', 'FRAME3', 'FRAME4',
                          'FRAME5',
                          'FRAME6', 'FRAME7', 'FRAME8', 'FRAME9', 'FRAME10', 'FRAME11', 'FRAME12', 'FRAME13', 'FRAME14',
                          'FRAME15', 'FRAME16', 'FRAME17', 'FRAME18', 'FRAME19', 'FRAME20', 'FRAME21', 'FRAME22',
                          'FRAME23',
                          'FRAME24', 'FRAME25', 'FRAME26', 'FRAME27', 'FRAME28', 'FRAME29', 'FRAME30', 'FRAME31',
                          'FRAME32',
                          'FRAME33', 'FRAME34', 'FRAME35', 'FRAME36', 'FRAME36', 'FRAME37', 'FRAME38', 'FRAME39',
                          'FRAME40',
                          'FRAME41', 'FRAME42', 'FRAME43', 'GIB1', 'GIB2', 'BALANCER_LEFT_All', 'BALANCER_RIGHT_ALL',
                          'CRANK_SHAFT_CLOCK', 'CLUCTH_ASSEMBLY_All', 'SLIDE_UNIT_All', 'CRANK_SHAFT', 'JOINT_All',
                          'MAIN_GEAR1', 'MAIN_GEAR2', 'MAIN_GEAR3', 'MAIN_GEAR4', 'JOINT1']

        # 開啟零件檔更改變數後儲存並關閉
        for y in file_name_list:
            # print('save' + str(y))
            mprog.import_part("C:\\Users\\USER\\Desktop\\stamping_press", y)
            # print('open' + str(y))
            if i == 4:
                if y == 'SLIDE_UNIT_All':
                    mprog.axis_system()
                    mprog.scaling(0.8)
                    if l == 0:
                        mprog.param_change(y, 'P', P_15[i])
                    else:
                        mprog.param_change(y, 'P', P[i])
                    mprog.save_file_part(path, y)
                elif y == 'BALANCER_LEFT_All' or y == 'BALANCER_RIGHT_ALL':
                    mprog.axis_system()
                    mprog.scaling(0.7)
                    mprog.save_file_part(path, y)
                elif y == 'GIB1' or y == 'GIB2':
                    mprog.axis_system()
                    mprog.scaling(0.9)
                    mprog.save_file_part(path, y)
                else:
                    if l == 0:
                        if y == 'FRAME1' or y == 'FRAME2':  # 更改零件變數B
                            mprog.param_change(y, 'B', B_15[i])
                            mprog.save_file_part(path, y)
                        elif y == 'FRAME3' or y == 'FRAME4' or y == 'FRAME9' or y == 'FRAME32' or y == 'FRAME41' or y == 'FRAME43' or y == 'FRAME20' or y == 'FRAME30' or y == 'FRAME29' or y == 'FRAME42':  # 更改零件變數R
                            mprog.param_change(y, 'R', R_15[i])
                            mprog.save_file_part(path, y)
                        elif y == 'FRAME10' or y == 'FRAME11' or y == 'FRAME12' or y == 'FRAME13' or y == 'BOLSTER1':  # 更改零件變數E
                            mprog.param_change(y, 'E', E_15[i])
                            if y == 'BOLSTER1':
                                mprog.param_change('BOLSTER1', "hole_type", hole_type[j])
                            mprog.save_file_part(path, y)
                        elif y == 'FRAME29' or y == 'FRAME8' or y == 'FRAME5' or y == 'FRAME6' or y == 'FRAME7':  # 更改零件變數A
                            mprog.param_change(y, 'A', A_15[i])
                            mprog.save_file_part(path, y)
                        elif y == 'SLIDE_UNIT_All':  # 更改零件變數P
                            mprog.param_change(y, 'P', P_15[i])
                            mprog.save_file_part(path, y)
                        elif y == 'BOLSTER3' or y == 'BOLSTER2':  # 更改零件變數Q
                            mprog.param_change(y, 'Q', Q_15[i])
                            mprog.save_file_part(path, y)
                        else:
                            mprog.save_file_part(path, y)
                    else:
                        if y == 'FRAME1' or y == 'FRAME2':  # 更改零件變數B
                            mprog.param_change(y, 'B', B[i])
                            mprog.save_file_part(path, y)
                        elif y == 'FRAME3' or y == 'FRAME4' or y == 'FRAME9' or y == 'FRAME32' or y == 'FRAME41' or y == 'FRAME43' or y == 'FRAME20' or y == 'FRAME30' or y == 'FRAME29' or y == 'FRAME42':  # 更改零件變數R
                            mprog.param_change(y, 'R', R[i])
                            mprog.save_file_part(path, y)
                        elif y == 'FRAME10' or y == 'FRAME11' or y == 'FRAME12' or y == 'FRAME13' or y == 'BOLSTER1':  # 更改零件變數E
                            mprog.param_change(y, 'E', E[i])
                            if y == 'BOLSTER1':
                                mprog.param_change('BOLSTER1', "hole_type", hole_type[j])
                            mprog.save_file_part(path, y)
                        elif y == 'FRAME29' or y == 'FRAME8' or y == 'FRAME5' or y == 'FRAME6' or y == 'FRAME7':  # 更改零件變數A
                            mprog.param_change(y, 'A', A[i])
                            mprog.save_file_part(path, y)
                        elif y == 'SLIDE_UNIT_All':  # 更改零件變數P
                            mprog.param_change(y, 'P', P[i])
                            mprog.save_file_part(path, y)
                        elif y == 'BOLSTER3' or y == 'BOLSTER2':  # 更改零件變數Q
                            mprog.param_change(y, 'Q', Q[i])
                            mprog.save_file_part(path, y)
                        else:
                            mprog.save_file_part(path, y)
            else:
                if l == 0:
                    if y == 'FRAME1' or y == 'FRAME2':  # 更改零件變數B
                        mprog.param_change(y, 'B', B_15[i])
                        mprog.save_file_part(path, y)
                    elif y == 'FRAME3' or y == 'FRAME4' or y == 'FRAME9' or y == 'FRAME32' or y == 'FRAME41' or y == 'FRAME43' or y == 'FRAME20' or y == 'FRAME30' or y == 'FRAME29' or y == 'FRAME42':  # 更改零件變數R
                        mprog.param_change(y, 'R', R_15[i])
                        mprog.save_file_part(path, y)
                    elif y == 'FRAME10' or y == 'FRAME11' or y == 'FRAME12' or y == 'FRAME13' or y == 'BOLSTER1':  # 更改零件變數E
                        mprog.param_change(y, 'E', E_15[i])
                        if y == 'BOLSTER1':
                            mprog.param_change('BOLSTER1', "hole_type", hole_type[j])
                        mprog.save_file_part(path, y)
                    elif y == 'FRAME29' or y == 'FRAME8' or y == 'FRAME5' or y == 'FRAME6' or y == 'FRAME7':  # 更改零件變數A
                        mprog.param_change(y, 'A', A_15[i])
                        mprog.save_file_part(path, y)
                    elif y == 'SLIDE_UNIT_All':  # 更改零件變數P
                        mprog.param_change(y, 'P', P_15[i])
                        mprog.save_file_part(path, y)
                    elif y == 'BOLSTER3' or y == 'BOLSTER2':  # 更改零件變數Q
                        mprog.param_change(y, 'Q', Q_15[i])
                        mprog.save_file_part(path, y)
                    else:
                        mprog.save_file_part(path, y)
                else:
                    if y == 'FRAME1' or y == 'FRAME2':  # 更改零件變數B
                        mprog.param_change(y, 'B', B[i])
                        mprog.save_file_part(path, y)
                    elif y == 'FRAME3' or y == 'FRAME4' or y == 'FRAME9' or y == 'FRAME32' or y == 'FRAME41' or y == 'FRAME43' or y == 'FRAME20' or y == 'FRAME30' or y == 'FRAME29' or y == 'FRAME42':  # 更改零件變數R
                        mprog.param_change(y, 'R', R[i])
                        mprog.save_file_part(path, y)
                    elif y == 'FRAME10' or y == 'FRAME11' or y == 'FRAME12' or y == 'FRAME13' or y == 'BOLSTER1':  # 更改零件變數E
                        mprog.param_change(y, 'E', E[i])
                        if y == 'BOLSTER1':
                            mprog.param_change('BOLSTER1', "hole_type", hole_type[j])
                        mprog.save_file_part(path, y)
                    elif y == 'FRAME29' or y == 'FRAME8' or y == 'FRAME5' or y == 'FRAME6' or y == 'FRAME7':  # 更改零件變數A
                        mprog.param_change(y, 'A', A[i])
                        mprog.save_file_part(path, y)
                    elif y == 'SLIDE_UNIT_All':  # 更改零件變數P
                        mprog.param_change(y, 'P', P[i])
                        mprog.save_file_part(path, y)
                    elif y == 'BOLSTER3' or y == 'BOLSTER2':  # 更改零件變數Q
                        mprog.param_change(y, 'Q', Q[i])
                        mprog.save_file_part(path, y)
                    else:
                        mprog.save_file_part(path, y)

    def ass_(self, h, i, l, j, path):

        # 開啟新組合檔
        mprog.assembly_create()

        # 匯入待組合零件檔
        file_Assembly_name_list = ['BOLSTER1', 'BOLSTER2', 'BOLSTER3', 'Fixture', 'FRAME1', 'FRAME2', 'FRAME3',
                                   'FRAME4', 'FRAME5', 'FRAME6', 'FRAME7', 'FRAME8', 'FRAME9', 'FRAME10', 'FRAME11',
                                   'FRAME12', 'FRAME13', 'FRAME14', 'FRAME15', 'FRAME16', 'FRAME17', 'FRAME18',
                                   'FRAME19', 'FRAME20', 'FRAME21', 'FRAME22', 'FRAME23', 'FRAME24', 'FRAME25',
                                   'FRAME26', 'FRAME27', 'FRAME28', 'FRAME29', 'FRAME30', 'FRAME31', 'FRAME31',
                                   'FRAME31', 'FRAME31', 'FRAME32', 'FRAME33', 'FRAME34', 'FRAME34', 'FRAME35',
                                   'FRAME36', 'FRAME36', 'FRAME37', 'FRAME37', 'FRAME37', 'FRAME37', 'FRAME38',
                                   'FRAME39', 'FRAME40', 'FRAME41', 'FRAME42', 'FRAME43', 'GIB1', 'GIB2',
                                   'BALANCER_LEFT_All', 'BALANCER_RIGHT_ALL', 'CRANK_SHAFT_CLOCK',
                                   'CLUCTH_ASSEMBLY_All', 'SLIDE_UNIT_ALL', 'CRANK_SHAFT', 'JOINT_All', 'MAIN_GEAR1',
                                   'MAIN_GEAR2', 'MAIN_GEAR3', 'MAIN_GEAR4', 'JOINT1']
        for x in file_Assembly_name_list:  # 讀取串列名稱並匯入檔案
            # print('import' + str(x) + 'part')
            mprog.import_file_Part(path, x)

        # 機架組合
        # mprog.base_lock('BOLSTER1.1', 'BOLSTER1.1', 0)  # 基準零件(定海神針)
        mprog.base_lock('FRAME20.1', 'FRAME20.1', 0)  # 基準零件(定海神針)

        # (0表示SAME, 1表示OPPOSITE)
        ##平板-四底座
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME6.1', A_15[i] / 2, 'XZ.PLANE', 1, 1)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME6.1', A[i] / 2, 'XZ.PLANE', 1, 1)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME6.1', -Z[i], 'XY.PLANE', 1, 2)
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME6.1', -80 - B_15[i] + F[i] / 2, 'YZ.PLANE', 1, 3)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME6.1', -80 - B[i] + F[i] / 2, 'YZ.PLANE', 1, 3)
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME7.1', -A_15[i] / 2, 'XZ.PLANE', 0, 4)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME7.1', -A[i] / 2, 'XZ.PLANE', 0, 4)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME7.1', -Z[i], 'XY.PLANE', 0, 5)
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME7.1', -80 - B_15[i] + F[i] / 2, 'YZ.PLANE', 1, 6)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME7.1', -80 - B[i] + F[i] / 2, 'YZ.PLANE', 1, 6)
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME5.1', -A_15[i] / 2, 'XZ.PLANE', 0, 7)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME5.1', -A[i] / 2, 'XZ.PLANE', 0, 7)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME5.1', -Z[i], 'XY.PLANE', 1, 8)
        if l == 0:
            mprog.add_offset_assembly('FRAME7.1', 'FRAME5.1', -B_15[i], 'YZ.PLANE', 1, 9)
        else:
            mprog.add_offset_assembly('FRAME7.1', 'FRAME5.1', -B[i], 'YZ.PLANE', 1, 9)
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME8.1', A_15[i] / 2, 'XZ.PLANE', 0, 10)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME8.1', A[i] / 2, 'XZ.PLANE', 0, 10)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME8.1', -Z[i], 'XY.PLANE', 0, 11)
        if l == 0:
            mprog.add_offset_assembly('FRAME6.1', 'FRAME8.1', -B_15[i], 'YZ.PLANE', 0, 12)
        else:
            mprog.add_offset_assembly('FRAME6.1', 'FRAME8.1', -B[i], 'YZ.PLANE', 0, 12)
        # 左右側板
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME2.1', R_15[i] / 2 + 140, 'XZ.PLANE', 1, 13)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME2.1', R[i] / 2 + 140, 'XZ.PLANE', 1, 13)
        mprog.add_offset_assembly('FRAME8.1', 'FRAME2.1', H[i] / 2, 'XY.PLANE', 0, 14)
        if l == 0:
            mprog.add_offset_assembly('FRAME8.1', 'FRAME2.1', B_15[i] / 2, 'YZ.PLANE', 0, 15)
        else:
            mprog.add_offset_assembly('FRAME8.1', 'FRAME2.1', B[i] / 2, 'YZ.PLANE', 0, 15)
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME1.1', -R_15[i] / 2 - 140, 'XZ.PLANE', 1, 16)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME1.1', -R[i] / 2 - 140, 'XZ.PLANE', 1, 16)
        mprog.add_offset_assembly('FRAME5.1', 'FRAME1.1', -H[i] / 2, 'XY.PLANE', 1, 17)
        if l == 0:
            mprog.add_offset_assembly('FRAME5.1', 'FRAME1.1', -B_15[i] / 2, 'YZ.PLANE', 1, 18)
        else:
            mprog.add_offset_assembly('FRAME5.1', 'FRAME1.1', -B[i] / 2, 'YZ.PLANE', 1, 18)
        # 底部前、中板
        if l == 0:
            mprog.add_offset_assembly('FRAME5.1', 'FRAME3.1', A_15[i] / 2, 'XZ.PLANE', 1, 19)
        else:
            mprog.add_offset_assembly('FRAME5.1', 'FRAME3.1', A[i] / 2, 'XZ.PLANE', 1, 19)
        mprog.add_offset_assembly('FRAME5.1', 'FRAME3.1', 0, 'XY.PLANE', 0, 20)
        mprog.add_offset_assembly('FRAME5.1', 'FRAME3.1', 0, 'YZ.PLANE', 1, 21)
        mprog.add_offset_assembly('FRAME9.1', 'FRAME3.1', 0, 'XZ.PLANE', 0, 22)
        mprog.add_offset_assembly('FRAME9.1', 'FRAME3.1', 0, 'XY.PLANE', 0, 23)
        if l == 0:
            mprog.add_offset_assembly('FRAME9.1', 'FRAME3.1', -FRAME2_lower_depth_15[i] + 5, 'YZ.PLANE', 0, 24)
        else:
            mprog.add_offset_assembly('FRAME9.1', 'FRAME3.1', -FRAME2_lower_depth[i] + 5, 'YZ.PLANE', 0, 24)
        # 中間左右側板
        if l == 0:
            mprog.add_offset_assembly('FRAME11.1', 'FRAME3.1', R_15[i] / 2, 'XZ.PLANE', 0, 25)
        else:
            mprog.add_offset_assembly('FRAME11.1', 'FRAME3.1', R[i] / 2, 'XZ.PLANE', 0, 25)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME11.1', -T[i], 'XY.PLANE', 1, 26)
        if l == 0:
            mprog.add_offset_assembly('FRAME3.1', 'FRAME11.1', FRAME2_lower_depth_15[i], 'YZ.PLANE', 1, 27)
        else:
            mprog.add_offset_assembly('FRAME3.1', 'FRAME11.1', FRAME2_lower_depth[i], 'YZ.PLANE', 1, 27)
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME10.1', -R_15[i] / 2 - 90, 'XZ.PLANE', 0, 28)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME10.1', -R[i] / 2 - 90, 'XZ.PLANE', 0, 28)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME10.1', -T[i], 'XY.PLANE', 0, 29)
        if l == 0:
            mprog.add_offset_assembly('FRAME3.1', 'FRAME10.1', FRAME2_lower_depth_15[i], 'YZ.PLANE', 1, 30)
        else:
            mprog.add_offset_assembly('FRAME3.1', 'FRAME10.1', FRAME2_lower_depth[i], 'YZ.PLANE', 1, 30)
        # 底部後面ㄇ形角鐵
        mprog.add_offset_assembly('FRAME3.1', 'FRAME4.1', 0, 'XZ.PLANE', 1, 31)
        mprog.add_offset_assembly('FRAME3.1', 'FRAME4.1', 0, 'XY.PLANE', 0, 32)
        if l == 0:
            mprog.add_offset_assembly('FRAME3.1', 'FRAME4.1', B_15[i], 'YZ.PLANE', 0, 33)
        else:
            mprog.add_offset_assembly('FRAME3.1', 'FRAME4.1', B[i], 'YZ.PLANE', 0, 33)
        # 平板底部零件
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME12.1', -R_15[i] / 2, 'XZ.PLANE', 1, 34)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME12.1', -R[i] / 2, 'XZ.PLANE', 1, 34)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME12.1', -T[i], 'XY.PLANE', 0, 35)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME12.1', 0, 'YZ.PLANE', 1, 36)
        # 平板底部零件
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME13.1', R_15[i] / 2, 'XZ.PLANE', 1, 37)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME13.1', R[i] / 2, 'XZ.PLANE', 1, 37)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME13.1', -T[i], 'XY.PLANE', 0, 38)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME13.1', 0, 'YZ.PLANE', 1, 39)
        # 前中上軸承板&角鐵
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME20.1', 0, 'XZ.PLANE', 1, 40)
        mprog.add_offset_assembly('FRAME3.1', 'FRAME20.1', -H[i] + 40, 'XY.PLANE', 0, 41)
        mprog.add_offset_assembly('FRAME3.1', 'FRAME20.1', 0, 'YZ.PLANE', 0, 42)
        mprog.add_offset_assembly('FRAME30.1', 'FRAME20.1', 0, 'XZ.PLANE', 0, 43)
        mprog.add_offset_assembly('FRAME30.1', 'FRAME20.1', -5, 'XY.PLANE', 0, 44)
        mprog.add_offset_assembly('FRAME30.1', 'FRAME20.1', -550, 'YZ.PLANE', 0, 45)
        mprog.add_offset_assembly('FRAME29.1', 'FRAME20.1', 0, 'XZ.PLANE', 1, 46)
        mprog.add_offset_assembly('FRAME29.1', 'FRAME20.1', -FRAME20_H[i], 'XY.PLANE', 0, 47)
        mprog.add_offset_assembly('FRAME29.1', 'FRAME20.1', 0, 'YZ.PLANE', 1, 48)
        # 氣壓缸鎖固板左右
        if i == 4:
            mprog.add_offset_assembly('FRAME2.1', 'FRAME21.1', 183, 'XZ.PLANE', 1, 49)
        else:
            mprog.add_offset_assembly('FRAME2.1', 'FRAME21.1', 242, 'XZ.PLANE', 1, 49)
        mprog.add_offset_assembly('FRAME3.1', 'FRAME21.1', -H[i] + 40, 'XY.PLANE', 1, 50)
        mprog.add_offset_assembly('FRAME20.1', 'FRAME21.1', 280, 'YZ.PLANE', 1, 51)
        if i == 4:
            mprog.add_offset_assembly('FRAME1.1', 'FRAME22.1', -183, 'XZ.PLANE', 1, 52)
        else:
            mprog.add_offset_assembly('FRAME1.1', 'FRAME22.1', -242, 'XZ.PLANE', 1, 52)
        mprog.add_offset_assembly('FRAME3.1', 'FRAME22.1', -H[i] + 40, 'XY.PLANE', 0, 53)
        mprog.add_offset_assembly('FRAME20.1', 'FRAME22.1', 280, 'YZ.PLANE', 1, 54)
        # 鎖固平板六兄弟
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME14.1', 0, 'XZ.PLANE', 0, 55)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME14.1', -T[i], 'XY.PLANE', 0, 56)
        mprog.add_offset_assembly('FRAME3.1', 'FRAME14.1', -75, 'YZ.PLANE', 1, 57)
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME15.1', R_15[i] / 2 + 75 + 140, 'XZ.PLANE', 0, 58)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME15.1', R[i] / 2 + 75 + 140, 'XZ.PLANE', 0, 58)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME15.1', -T[i], 'XY.PLANE', 0, 59)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME15.1', -F[i] / 2 + 80 + 37.5, 'YZ.PLANE', 1, 60)
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME16.1', R_15[i] / 2 + 75 + 140, 'XZ.PLANE', 0, 61)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME16.1', R[i] / 2 + 75 + 140, 'XZ.PLANE', 0, 61)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME16.1', -T[i], 'XY.PLANE', 0, 62)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME16.1', F[i] / 2 - 80 - 37.5, 'YZ.PLANE', 1, 63)
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME18.1', -R_15[i] / 2 - 75 - 140, 'XZ.PLANE', 1, 64)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME18.1', -R[i] / 2 - 75 - 140, 'XZ.PLANE', 1, 64)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME18.1', -T[i], 'XY.PLANE', 0, 65)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME18.1', F[i] / 2 - 80 - 37.5, 'YZ.PLANE', 0, 66)
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME19.1', -R_15[i] / 2 - 75 - 140, 'XZ.PLANE', 1, 67)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME19.1', -R[i] / 2 - 75 - 140, 'XZ.PLANE', 1, 67)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME19.1', -T[i], 'XY.PLANE', 0, 68)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME19.1', -F[i] / 2 + 80 + 37.5, 'YZ.PLANE', 0, 69)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME17.1', 0, 'XZ.PLANE', 1, 70)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME17.1', -T[i], 'XY.PLANE', 0, 71)
        if l == 0:
            mprog.add_offset_assembly('FRAME3.1', 'FRAME17.1', FRAME2_lower_depth_15[i], 'YZ.PLANE', 0, 72)
        else:
            mprog.add_offset_assembly('FRAME3.1', 'FRAME17.1', FRAME2_lower_depth[i], 'YZ.PLANE', 0, 72)
        # 左右側板前GIB
        if i == 4:
            mprog.add_offset_assembly('GIB1.1', 'FRAME3.1', FRAME1_lower_high[i] + 40 - 34.5, 'XY.PLANE', 0, 73)
        else:
            mprog.add_offset_assembly('GIB1.1', 'FRAME3.1', FRAME1_lower_high[i] + 40, 'XY.PLANE', 0, 73)
        mprog.add_offset_assembly('GIB1.1', 'FRAME1.1', 72.5, 'XZ.PLANE', 0, 74)
        mprog.add_offset_assembly('GIB1.1', 'FRAME5.1', 334.65, 'YZ.PLANE', 0, 75)
        if i == 4:
            mprog.add_offset_assembly('GIB2.1', 'FRAME3.1', FRAME1_lower_high[i] + 40 - 34.5, 'XY.PLANE', 0, 76)
        else:
            mprog.add_offset_assembly('GIB2.1', 'FRAME3.1', FRAME1_lower_high[i] + 40, 'XY.PLANE', 0, 76)
        mprog.add_offset_assembly('GIB2.1', 'FRAME2.1', -72.5, 'XZ.PLANE', 0, 77)
        mprog.add_offset_assembly('GIB2.1', 'FRAME5.1', 334.65, 'YZ.PLANE', 0, 78)
        # 左GIB後鎖固用方塊
        if i == 4:
            mprog.add_offset_assembly('GIB2.1', 'FRAME23.1', -690 - 34.5, 'XY.PLANE', 1, 79)
        else:
            mprog.add_offset_assembly('GIB2.1', 'FRAME23.1', -690, 'XY.PLANE', 1, 79)
        mprog.add_offset_assembly('FRAME2.1', 'FRAME23.1', 50 + 35, 'XZ.PLANE', 1, 80)
        mprog.add_offset_assembly('GIB2.1', 'FRAME23.1', 0, 'YZ.PLANE', 0, 81)
        if i == 4:
            mprog.add_offset_assembly('GIB2.1', 'FRAME24.1', -34.5, 'XY.PLANE', 1, 82)
        else:
            mprog.add_offset_assembly('GIB2.1', 'FRAME24.1', 0, 'XY.PLANE', 1, 82)
        mprog.add_offset_assembly('FRAME2.1', 'FRAME24.1', 50 + 35, 'XZ.PLANE', 0, 83)
        mprog.add_offset_assembly('GIB2.1', 'FRAME24.1', 0, 'YZ.PLANE', 1, 84)
        if i == 4:
            mprog.add_offset_assembly('GIB2.1', 'FRAME27.1', -690 / 2 - 34.5, 'XY.PLANE', 1, 85)
        else:
            mprog.add_offset_assembly('GIB2.1', 'FRAME27.1', -690 / 2, 'XY.PLANE', 1, 85)
        mprog.add_offset_assembly('FRAME2.1', 'FRAME27.1', 50, 'XZ.PLANE', 1, 86)
        mprog.add_offset_assembly('GIB2.1', 'FRAME27.1', 0, 'YZ.PLANE', 1, 87)
        if i == 4:
            mprog.add_offset_assembly('GIB2.1', 'FRAME31.1', -690 - 34.5, 'XY.PLANE', 1, 88)
        else:
            mprog.add_offset_assembly('GIB2.1', 'FRAME31.1', -690, 'XY.PLANE', 1, 88)
        mprog.add_offset_assembly('FRAME23.1', 'FRAME31.1', -50, 'XZ.PLANE', 0, 89)
        mprog.add_offset_assembly('FRAME23.1', 'FRAME31.1', -130.35, 'YZ.PLANE', 0, 90)
        if i == 4:
            mprog.add_offset_assembly('GIB2.1', 'FRAME31.2', -150 - 34.5, 'XY.PLANE', 1, 91)
        else:
            mprog.add_offset_assembly('GIB2.1', 'FRAME31.2', -150, 'XY.PLANE', 1, 91)
        mprog.add_offset_assembly('FRAME23.1', 'FRAME31.2', -50, 'YZ.PLANE', 0, 92)
        mprog.add_offset_assembly('FRAME23.1', 'FRAME31.2', -62.5, 'XZ.PLANE', 0, 93)
        # 右GIB後鎖固用方塊
        if i == 4:
            mprog.add_offset_assembly('GIB1.1', 'FRAME25.1', -690 - 34.5, 'XY.PLANE', 1, 94)
        else:
            mprog.add_offset_assembly('GIB1.1', 'FRAME25.1', -690, 'XY.PLANE', 1, 94)
        mprog.add_offset_assembly('FRAME1.1', 'FRAME25.1', -50 - 35, 'XZ.PLANE', 0, 95)
        mprog.add_offset_assembly('GIB1.1', 'FRAME25.1', 0, 'YZ.PLANE', 1, 96)
        if i == 4:
            mprog.add_offset_assembly('GIB1.1', 'FRAME26.1', -34.5, 'XY.PLANE', 1, 97)
        else:
            mprog.add_offset_assembly('GIB1.1', 'FRAME26.1', 0, 'XY.PLANE', 1, 97)
        mprog.add_offset_assembly('FRAME1.1', 'FRAME26.1', -50 - 35, 'XZ.PLANE', 0, 98)
        mprog.add_offset_assembly('GIB1.1', 'FRAME26.1', 0, 'YZ.PLANE', 0, 99)
        if i == 4:
            mprog.add_offset_assembly('GIB1.1', 'FRAME28.1', -690 / 2 - 34.5, 'XY.PLANE', 1, 100)
        else:
            mprog.add_offset_assembly('GIB1.1', 'FRAME28.1', -690 / 2, 'XY.PLANE', 1, 100)
        mprog.add_offset_assembly('FRAME1.1', 'FRAME28.1', -50, 'XZ.PLANE', 0, 101)
        mprog.add_offset_assembly('GIB1.1', 'FRAME28.1', 0, 'YZ.PLANE', 1, 102)
        if i == 4:
            mprog.add_offset_assembly('GIB1.1', 'FRAME31.3', -690 - 34.5, 'XY.PLANE', 1, 103)
        else:
            mprog.add_offset_assembly('GIB1.1', 'FRAME31.3', -690, 'XY.PLANE', 1, 103)
        mprog.add_offset_assembly('FRAME25.1', 'FRAME31.3', -50, 'XZ.PLANE', 1, 104)
        mprog.add_offset_assembly('FRAME25.1', 'FRAME31.3', 130.35, 'YZ.PLANE', 1, 105)
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
        mprog.add_offset_assembly('FRAME32.1', 'FRAME3.1', FRAME_32_XY[i] + 19, 'XY.PLANE', 0, 116)
        mprog.add_offset_assembly('FRAME32.1', 'FRAME30.1', 0, 'YZ.PLANE', 0, 117)
        # 後方大軸承支架
        mprog.add_offset_assembly('FRAME10.1', 'FRAME34.2', -485.543, 'YZ.PLANE', 0, 118)
        mprog.add_offset_assembly('FRAME34.2', 'FRAME10.1', FRAME_10_H[i] + 40, 'XY.PLANE', 1, 119)
        if l == 0:
            mprog.add_offset_assembly('FRAME34.2', 'FRAME3.1', -R_15[i] / 2 - 90, 'XZ.PLANE', 0, 120)
        else:
            mprog.add_offset_assembly('FRAME34.2', 'FRAME3.1', -R[i] / 2 - 90, 'XZ.PLANE', 0, 120)
        if l == 0:
            mprog.add_offset_assembly('FRAME34.1', 'FRAME34.2', -R_15[i] - 180, 'XZ.PLANE', 1, 121)
        else:
            mprog.add_offset_assembly('FRAME34.1', 'FRAME34.2', -R[i] - 180, 'XZ.PLANE', 1, 121)
        mprog.add_offset_assembly('FRAME34.1', 'FRAME34.2', 0, 'XY.PLANE', 1, 122)
        mprog.add_offset_assembly('FRAME34.1', 'FRAME34.2', 0, 'YZ.PLANE', 0, 123)
        # 後方大軸承
        mprog.add_offset_assembly('FRAME35.1', 'FRAME3.1', 0, 'XZ.PLANE', 1, 124)
        mprog.add_offset_assembly('FRAME35.1', 'FRAME20.1', -868, 'XY.PLANE', 0, 125)
        mprog.add_offset_assembly('FRAME35.1', 'FRAME20.1', -FRAME20_FRAME2_YZ[i] - 32, 'YZ.PLANE', 0, 126)
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
        mprog.add_offset_assembly('FRAME40.1', 'FRAME20.1', FRAME20_FRAME2_YZ[i], 'YZ.PLANE', 1, 150)
        # FRAME35上兩圓管
        mprog.add_offset_assembly('FRAME36.1', 'FRAME35.1', -272.236, 'XZ.PLANE', 1, 151)
        mprog.add_offset_assembly('FRAME36.1', 'FRAME35.1', -272.236, 'XY.PLANE', 1, 152)
        mprog.add_offset_assembly('FRAME36.1', 'FRAME35.1', 88, 'YZ.PLANE', 1, 153)
        mprog.add_offset_assembly('FRAME36.2', 'FRAME35.1', 272.236, 'XZ.PLANE', 1, 154)
        mprog.add_offset_assembly('FRAME36.2', 'FRAME35.1', -272.236, 'XY.PLANE', 1, 155)
        mprog.add_offset_assembly('FRAME36.2', 'FRAME35.1', 88, 'YZ.PLANE', 1, 156)
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
            mprog.add_offset_assembly('BOLSTER1.1', 'BOLSTER3.1', DH_S[i], 'XY.PLANE', 0, 168)
        elif h == 1:
            mprog.add_offset_assembly('BOLSTER1.1', 'BOLSTER3.1', DH_H[i], 'XY.PLANE', 0, 168)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'BOLSTER3.1', DH_P[i], 'XY.PLANE', 0, 168)
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
        mprog.add_offset_product_assembly('CRANK_SHAFT_CLOCK.1', 'Geometrical Set.2', 'FRAME20.1', 35, 'YZ.PLANE', 0,
                                          174)
        mprog.add_offset_product_assembly('CRANK_SHAFT_CLOCK.1', 'Geometrical Set.2', 'FRAME20.1', -422.052, 'XY.PLANE',
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
                                                  -R_15[i] / 2 - 13, 'XZ.PLANE', 0, 178)
            else:
                mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1',
                                                  -R[i] / 2 - 13, 'XZ.PLANE', 0, 178)
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
                                                  -R_15[i] / 2 - 13, 'XZ.PLANE', 1, 182)
            else:
                mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1',
                                                  -R[i] / 2 - 13, 'XZ.PLANE', 1, 182)
        if i == 4:
            mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1', 261.248,
                                              'YZ.PLANE', 1, 183)
        else:
            mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1', 260, 'YZ.PLANE',
                                              1, 183)
        # 離合器與FRAME20結合
        if i == 4:
            mprog.add_offset_product_assembly('CLUCTH_ASSEMBLY_All.1', 'Geometrical Set.1', 'FRAME20.1', 108.109,
                                              'XY.PLANE', 0, 184)
        else:
            mprog.add_offset_product_assembly('CLUCTH_ASSEMBLY_All.1', 'Geometrical Set.1', 'FRAME20.1', 0, 'XY.PLANE',
                                              0, 184)
        mprog.add_offset_product_assembly('CLUCTH_ASSEMBLY_All.1', 'Geometrical Set.1', 'FRAME20.1', 0, 'XZ.PLANE', 1,
                                          185)
        mprog.add_offset_product_assembly('CLUCTH_ASSEMBLY_All.1', 'Geometrical Set.1', 'FRAME20.1', 510, 'YZ.PLANE', 1,
                                          186)
        # 曲軸與時鐘結合
        mprog.add_offset_product_assembly('CRANK_SHAFT_CLOCK.1', 'Geometrical Set.2', 'CRANK_SHAFT.1.1', 0, 'XZ.PLANE', 0,
                                          187)
        mprog.add_offset_product_assembly('CRANK_SHAFT_CLOCK.1', 'Geometrical Set.2', 'CRANK_SHAFT.1.1', 35, 'YZ.PLANE',
                                          1, 188)
        mprog.add_offset_product_assembly('CRANK_SHAFT_CLOCK.1', 'Geometrical Set.2', 'CRANK_SHAFT.1.1', 0, 'XY.PLANE', 1,
                                          189)
        # 大齒輪結合
        mprog.add_offset_assembly('MAIN_GEAR1.1', 'MAIN_GEAR4.1', 0, 'XZ.PLANE', 0, 190)
        mprog.add_offset_assembly('MAIN_GEAR1.1', 'MAIN_GEAR4.1', 77.172, 'XY.PLANE', 0, 191)
        mprog.add_offset_assembly('MAIN_GEAR1.1', 'MAIN_GEAR4.1', 0, 'YZ.PLANE', 1, 192)
        # 大齒輪結合曲軸
        mprog.add_offset_assembly('MAIN_GEAR1.1', 'CRANK_SHAFT.1.1', 0, 'XZ.PLANE', 1, 193)
        mprog.add_offset_assembly('MAIN_GEAR1.1', 'CRANK_SHAFT.1.1', 0, 'XY.PLANE', 0, 194)
        mprog.add_offset_assembly('MAIN_GEAR1.1', 'CRANK_SHAFT.1.1', 765, 'YZ.PLANE', 0, 195)
        # 機架20結合曲軸旁圓管
        mprog.add_offset_assembly('MAIN_GEAR2.1', 'FRAME20.1', 0, 'XZ.PLANE', 0, 196)
        mprog.add_offset_assembly('MAIN_GEAR2.1', 'FRAME20.1', 420, 'XY.PLANE', 1, 197)
        mprog.add_offset_assembly('MAIN_GEAR2.1', 'FRAME20.1', 530, 'YZ.PLANE', 1, 198)
        # 大齒輪結合長棒
        mprog.add_offset_assembly('MAIN_GEAR1.1', 'JOINT1.1', 0, 'XZ.PLANE', 0, 199)
        mprog.add_offset_assembly('MAIN_GEAR1.1', 'JOINT1.1', 0, 'XY.PLANE', 1, 200)
        mprog.add_offset_assembly('MAIN_GEAR1.1', 'JOINT1.1', -84, 'YZ.PLANE', 1, 201)
        #
        mprog.add_offset_product_assembly('JOINT_All.1', 'Geometrical Set.1', 'FRAME20.1', 0, 'XZ.PLANE',
                                          0, 202)
        mprog.add_offset_product_assembly('JOINT_All.1', 'Geometrical Set.1', 'FRAME20.1', 1010, 'YZ.PLANE',
                                          1, 203)
        mprog.add_offset_product_assembly('JOINT_All.1', 'Geometrical Set.1', 'FRAME20.1', -422.052, 'XY.PLANE',
                                          0, 204)
        #大齒輪內套環
        mprog.add_offset_assembly('MAIN_GEAR2.1', 'MAIN_GEAR3.1', 0, 'XZ.PLANE', 1, 205)
        mprog.add_offset_assembly('MAIN_GEAR2.1', 'MAIN_GEAR3.1', 0, 'XY.PLANE', 0, 206)
        mprog.add_offset_assembly('MAIN_GEAR2.1', 'MAIN_GEAR3.1', 11, 'YZ.PLANE', 1, 207)

        # 更新
        mprog.update()
        mprog.Close_All()
        # 儲存Product檔
        type = str(self.ui.comboBox_4.currentText())
        mprog.saveas(path, type, '.CATProduct')

        # --------------------------------------- 生成爆炸圖--------------------------------------------
        # 重新定義拘束尺寸
        BOLSTER1_Offset_value = 2500
        GIB_Offset_value = -3200
        CLOCK_Offset_value = 2800
        CLOCK_SHAFT_Offset_value = 1500

        mprog.constaint_value_change(3, -80 - B[i] + F[i] / 2 + 353 - BOLSTER1_Offset_value, 1)
        mprog.constaint_value_change(6, -80 - B[i] + F[i] / 2 + 353 - BOLSTER1_Offset_value, 1)
        mprog.constaint_value_change(36, -BOLSTER1_Offset_value, 1)
        mprog.constaint_value_change(39, -BOLSTER1_Offset_value, 1)
        mprog.constaint_value_change(63, F[i] / 2 - 80 - 37.5 - BOLSTER1_Offset_value - 80, 1)
        mprog.constaint_value_change(66, F[i] / 2 - 80 - 37.5 - BOLSTER1_Offset_value - 80, 0)
        mprog.constaint_value_change(60, -F[i] / 2 + 80 + 37.5 - BOLSTER1_Offset_value + 80, 1)
        mprog.constaint_value_change(69, -F[i] / 2 + 80 + 37.5 - BOLSTER1_Offset_value + 80, 0)
        mprog.constaint_value_change(167, -250, 0)  # BLOSTER1 & 3
        mprog.constaint_value_change(75, GIB_Offset_value, 0)
        mprog.constaint_value_change(78, GIB_Offset_value, 0)
        mprog.constaint_value_change(84, GIB_Offset_value - 334.65, 1)
        mprog.constaint_value_change(87, GIB_Offset_value - 334.65, 1)
        mprog.constaint_value_change(81, GIB_Offset_value - 334.65, 0)
        mprog.constaint_value_change(96, GIB_Offset_value - 334.65, 1)
        mprog.constaint_value_change(102, GIB_Offset_value - 334.65, 1)
        mprog.constaint_value_change(99, GIB_Offset_value - 334.65, 0)
        mprog.constaint_value_change(174, CLOCK_Offset_value, 0)
        mprog.constaint_value_change(188, CLOCK_SHAFT_Offset_value, 1)
        mprog.constaint_value_change(195, 1892.5 + 150, 0)
        mprog.constaint_value_change(176, -1700, 1)
        mprog.constaint_value_change(180, -1700, 1)

if __name__ == "__main__":
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)  # 自適應屏幕分辨率
    app = QtWidgets.QApplication([])
    window = main()
    window.show()
    sys.exit(app.exec_())
