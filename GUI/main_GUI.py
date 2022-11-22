from PyQt5 import QtCore, QtGui, QtWidgets
import sys, datetime, os
from GUI import Ui_Dialog
import main_program as mprog

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
# FRAME1_cutout_bottom = [197, 277, 317, 397, 477, 557, 637, 717, 797, 877]
F = [320, 400, 440, 520, 600, 680, 760, 840, 900, 960]
# FRAME1_cutout = [655 + 370, 675 + 415, 695 + 450, 715 + 577.5, 735 + 670, 755 + 770, 775 + 895, 795 + 1080, 815 + 1210, ]
FRAME20_H = [280, 355, 420, 437.5, 550, 680, 825, 990, 1120, 1170]
FRAME2_lower_depth = [166.016, 246.016, 286.016, 366.016, 446.016, 526.016, 606.016, 686.016, 766.016, 846.016]
FRAME2_lower_depth_15 = [265.016, 385.016, 445.016, 565.016, 685.016, 925.016, 805.016, 1045.016, 1165.016, 1285.016]
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
        change = str(self.ui.comboBox_2.currentText())
        height = str(self.ui.comboBox_3.currentText())
        type = str(self.ui.comboBox_4.currentText())
        hole = str(self.ui.comboBox_5.currentText())
        print(type, change, height, close, hole)
        self.create_dir(type)
        self.l, self.h, self.i, self.j, self.k = self.choos(change, height, type, hole, close)
        self.change_dir(self.l, self.h, self.i, self.j, self.k, self.path)
        self.ass_(self.l, self.h, self.i, self.j, self.k, self.path)

    def choos(self, change, height, type, hole, close):
        # "輸入型號"
        type = input()
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
        print(i)

        # 高度類型選擇
        # '請輸入高度類型(S = 0 , H = 1 , P = 2)'
        height = input()
        if height == '0':
            h = 0
        elif height == '1':
            h = 1
        elif height == '2':
            h = 2
        print(h)

        # 判斷是否長寬
        # '請輸入是否變更長寬(0 = 1.5H , 1 = H)')
        l = input()
        if l == '0':
            l = 0
        elif l == '1':
            l = 1
        else:
            print('閉合輸入錯誤')
        print(l)

        # 輸入平板型號
        # print("請輸入平板型式(0 = 圓形平板 , 1 = 方形平板 , 2 = 模墊型平板)")
        hole = input()
        if hole == "0":
            j = 0
        elif hole == "1":
            j = 1
        elif hole == "2":
            j = 2
        else:
            print('平板型號輸入錯誤')
        print(j)
        return h, i, l, j, k

    def add_item_for_comboBox(self):
        print('insert')
        # data = ['sd1', 'sd2']
        # for item in data:
        #     self.ui.comboBox.addItem(item)

    def create_dir(self, type):
        time_now = datetime.datetime.now()
        dir_name = '{}_{}_{}_{}_{}'.format(type, time_now.day, time_now.hour, time_now.minute, time_now.second)
        # dir_name = '%s_%s_%s_%s_%s' % (type, time_now.day, time_now.hour, time_now.minute, time_now.second)
        desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
        path = desktop + '\\' + dir_name
        os.mkdir(path)
        self.path = path

    def change_dir(self, h, i, l, j, k, path):

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
                          'CRANK_SHAFT_All', 'CLUCTH_ASSEMBLY_All', 'SLIDE_UNIT_All']

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

    def ass_(self, h, i, l, j, k, path):

        # 開啟新組合檔
        mprog.assembly_create()

        # 匯入待組合零件檔
        file_Assembly_name_list = ['BOLSTER1', 'BOLSTER2', 'BOLSTER3', 'Fixture', 'FRAME1', 'FRAME2', 'FRAME3',
                                   'FRAME4', 'FRAME5',
                                   'FRAME6', 'FRAME7', 'FRAME8', 'FRAME9', 'FRAME10', 'FRAME11', 'FRAME12', 'FRAME13',
                                   'FRAME14',
                                   'FRAME15', 'FRAME16', 'FRAME17', 'FRAME18', 'FRAME19', 'FRAME20', 'FRAME21',
                                   'FRAME22',
                                   'FRAME23', 'FRAME24', 'FRAME25', 'FRAME26', 'FRAME27', 'FRAME28', 'FRAME29',
                                   'FRAME30',
                                   'FRAME31', 'FRAME31', 'FRAME31', 'FRAME31', 'FRAME32', 'FRAME33', 'FRAME34',
                                   'FRAME34',
                                   'FRAME35', 'FRAME36', 'FRAME36', 'FRAME37', 'FRAME37', 'FRAME37', 'FRAME37',
                                   'FRAME38',
                                   'FRAME39', 'FRAME40', 'FRAME41', 'FRAME42', 'FRAME43', 'GIB1', 'GIB2',
                                   'BALANCER_LEFT_All',
                                   'BALANCER_RIGHT_ALL', 'CRANK_SHAFT_All', 'CLUCTH_ASSEMBLY_All', 'SLIDE_UNIT_All']
        for x in file_Assembly_name_list:  # 讀取串列名稱並匯入檔案
            # print('import' + str(x) + 'part')
            mprog.import_file_Part(path, x)

        # 機架組合
        mprog.base_lock('BOLSTER1.1', 'BOLSTER1.1')  # 基準零件(定海神針)
        # (0表示SAME, 1表示OPPOSITE)
        ##平板-四底座
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME6.1', A_15[i] / 2, 'XZ.PLANE', 1)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME6.1', A[i] / 2, 'XZ.PLANE', 1)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME6.1', -Z[i], 'XY.PLANE', 1)
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME6.1', -80 - B_15[i] + F[i] / 2, 'YZ.PLANE', 1)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME6.1', -80 - B[i] + F[i] / 2, 'YZ.PLANE', 1)
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME7.1', -A_15[i] / 2, 'XZ.PLANE', 0)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME7.1', -A[i] / 2, 'XZ.PLANE', 0)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME7.1', -Z[i], 'XY.PLANE', 0)
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME7.1', -80 - B_15[i] + F[i] / 2, 'YZ.PLANE', 1)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME7.1', -80 - B[i] + F[i] / 2, 'YZ.PLANE', 1)
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME5.1', -A_15[i] / 2, 'XZ.PLANE', 0)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME5.1', -A[i] / 2, 'XZ.PLANE', 0)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME5.1', -Z[i], 'XY.PLANE', 1)
        if l == 0:
            mprog.add_offset_assembly('FRAME7.1', 'FRAME5.1', -B_15[i], 'YZ.PLANE', 1)
        else:
            mprog.add_offset_assembly('FRAME7.1', 'FRAME5.1', -B[i], 'YZ.PLANE', 1)
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME8.1', A_15[i] / 2, 'XZ.PLANE', 0)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME8.1', A[i] / 2, 'XZ.PLANE', 0)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME8.1', -Z[i], 'XY.PLANE', 0)
        if l == 0:
            mprog.add_offset_assembly('FRAME6.1', 'FRAME8.1', -B_15[i], 'YZ.PLANE', 0)
        else:
            mprog.add_offset_assembly('FRAME6.1', 'FRAME8.1', -B[i], 'YZ.PLANE', 0)
        # 左右側板
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME2.1', R_15[i] / 2 + 140, 'XZ.PLANE', 1)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME2.1', R[i] / 2 + 140, 'XZ.PLANE', 1)
        mprog.add_offset_assembly('FRAME8.1', 'FRAME2.1', H[i] / 2, 'XY.PLANE', 0)
        if l == 0:
            mprog.add_offset_assembly('FRAME8.1', 'FRAME2.1', B_15[i] / 2, 'YZ.PLANE', 0)
        else:
            mprog.add_offset_assembly('FRAME8.1', 'FRAME2.1', B[i] / 2, 'YZ.PLANE', 0)
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME1.1', -R_15[i] / 2 - 140, 'XZ.PLANE', 1)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME1.1', -R[i] / 2 - 140, 'XZ.PLANE', 1)
        mprog.add_offset_assembly('FRAME5.1', 'FRAME1.1', -H[i] / 2, 'XY.PLANE', 1)
        if l == 0:
            mprog.add_offset_assembly('FRAME5.1', 'FRAME1.1', -B_15[i] / 2, 'YZ.PLANE', 1)
        else:
            mprog.add_offset_assembly('FRAME5.1', 'FRAME1.1', -B[i] / 2, 'YZ.PLANE', 1)
        # 底部前、中板
        if l == 0:
            mprog.add_offset_assembly('FRAME5.1', 'FRAME3.1', A_15[i] / 2, 'XZ.PLANE', 1)
        else:
            mprog.add_offset_assembly('FRAME5.1', 'FRAME3.1', A[i] / 2, 'XZ.PLANE', 1)
        mprog.add_offset_assembly('FRAME5.1', 'FRAME3.1', 0, 'XY.PLANE', 0)
        mprog.add_offset_assembly('FRAME5.1', 'FRAME3.1', 0, 'YZ.PLANE', 1)
        mprog.add_offset_assembly('FRAME9.1', 'FRAME3.1', 0, 'XZ.PLANE', 0)
        mprog.add_offset_assembly('FRAME9.1', 'FRAME3.1', 0, 'XY.PLANE', 0)
        if l == 0:
            mprog.add_offset_assembly('FRAME9.1', 'FRAME3.1', -FRAME2_lower_depth_15[i] + 5, 'YZ.PLANE', 0)
        else:
            mprog.add_offset_assembly('FRAME9.1', 'FRAME3.1', -FRAME2_lower_depth[i] + 5, 'YZ.PLANE', 0)
        # 中間左右側板
        if l == 0:
            mprog.add_offset_assembly('FRAME11.1', 'FRAME3.1', R_15[i] / 2, 'XZ.PLANE', 0)
        else:
            mprog.add_offset_assembly('FRAME11.1', 'FRAME3.1', R[i] / 2, 'XZ.PLANE', 0)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME11.1', -T[i], 'XY.PLANE', 1)
        if l == 0:
            mprog.add_offset_assembly('FRAME3.1', 'FRAME11.1', FRAME2_lower_depth_15[i], 'YZ.PLANE', 1)
        else:
            mprog.add_offset_assembly('FRAME3.1', 'FRAME11.1', FRAME2_lower_depth[i], 'YZ.PLANE', 1)
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME10.1', -R_15[i] / 2 - 90, 'XZ.PLANE', 0)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME10.1', -R[i] / 2 - 90, 'XZ.PLANE', 0)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME10.1', -T[i], 'XY.PLANE', 0)
        if l == 0:
            mprog.add_offset_assembly('FRAME3.1', 'FRAME10.1', FRAME2_lower_depth_15[i], 'YZ.PLANE', 1)
        else:
            mprog.add_offset_assembly('FRAME3.1', 'FRAME10.1', FRAME2_lower_depth[i], 'YZ.PLANE', 1)
        # 底部後面ㄇ形角鐵
        mprog.add_offset_assembly('FRAME3.1', 'FRAME4.1', 0, 'XZ.PLANE', 1)
        mprog.add_offset_assembly('FRAME3.1', 'FRAME4.1', 0, 'XY.PLANE', 0)
        if l == 0:
            mprog.add_offset_assembly('FRAME3.1', 'FRAME4.1', B_15[i], 'YZ.PLANE', 0)
        else:
            mprog.add_offset_assembly('FRAME3.1', 'FRAME4.1', B[i], 'YZ.PLANE', 0)
        # 平板底部零件
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME12.1', -R_15[i] / 2, 'XZ.PLANE', 1)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME12.1', -R[i] / 2, 'XZ.PLANE', 1)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME12.1', -T[i], 'XY.PLANE', 0)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME12.1', 0, 'YZ.PLANE', 1)
        # 平板底部零件
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME13.1', R_15[i] / 2, 'XZ.PLANE', 1)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME13.1', R[i] / 2, 'XZ.PLANE', 1)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME13.1', -T[i], 'XY.PLANE', 0)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME13.1', 0, 'YZ.PLANE', 1)
        # 前中上軸承板&角鐵
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME20.1', 0, 'XZ.PLANE', 1)
        mprog.add_offset_assembly('FRAME3.1', 'FRAME20.1', -H[i] + 40, 'XY.PLANE', 0)
        mprog.add_offset_assembly('FRAME3.1', 'FRAME20.1', 0, 'YZ.PLANE', 0)
        mprog.add_offset_assembly('FRAME30.1', 'FRAME20.1', 0, 'XZ.PLANE', 0)
        mprog.add_offset_assembly('FRAME30.1', 'FRAME20.1', -5, 'XY.PLANE', 0)
        mprog.add_offset_assembly('FRAME30.1', 'FRAME20.1', -550, 'YZ.PLANE', 0)
        mprog.add_offset_assembly('FRAME29.1', 'FRAME20.1', 0, 'XZ.PLANE', 1)
        mprog.add_offset_assembly('FRAME29.1', 'FRAME20.1', -FRAME20_H[i], 'XY.PLANE', 0)
        mprog.add_offset_assembly('FRAME29.1', 'FRAME20.1', 0, 'YZ.PLANE', 1)
        # 氣壓缸鎖固板左右
        if i == 4:
            mprog.add_offset_assembly('FRAME2.1', 'FRAME21.1', 183, 'XZ.PLANE', 1)
        else:
            mprog.add_offset_assembly('FRAME2.1', 'FRAME21.1', 242, 'XZ.PLANE', 1)
        mprog.add_offset_assembly('FRAME3.1', 'FRAME21.1', -H[i] + 40, 'XY.PLANE', 1)
        mprog.add_offset_assembly('FRAME20.1', 'FRAME21.1', 280, 'YZ.PLANE', 1)
        if i == 4:
            mprog.add_offset_assembly('FRAME1.1', 'FRAME22.1', -183, 'XZ.PLANE', 1)
        else:
            mprog.add_offset_assembly('FRAME1.1', 'FRAME22.1', -242, 'XZ.PLANE', 1)
        mprog.add_offset_assembly('FRAME3.1', 'FRAME22.1', -H[i] + 40, 'XY.PLANE', 0)
        mprog.add_offset_assembly('FRAME20.1', 'FRAME22.1', 280, 'YZ.PLANE', 1)
        # 鎖固平板六兄弟
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME14.1', 0, 'XZ.PLANE', 0)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME14.1', -T[i], 'XY.PLANE', 0)
        mprog.add_offset_assembly('FRAME3.1', 'FRAME14.1', -75, 'YZ.PLANE', 1)
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME15.1', R_15[i] / 2 + 75 + 140, 'XZ.PLANE', 0)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME15.1', R[i] / 2 + 75 + 140, 'XZ.PLANE', 0)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME15.1', -T[i], 'XY.PLANE', 0)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME15.1', -F[i] / 2 + 80 + 37.5, 'YZ.PLANE', 1)
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME16.1', R_15[i] / 2 + 75 + 140, 'XZ.PLANE', 0)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME16.1', R[i] / 2 + 75 + 140, 'XZ.PLANE', 0)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME16.1', -T[i], 'XY.PLANE', 0)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME16.1', F[i] / 2 - 80 - 37.5, 'YZ.PLANE', 1)
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME18.1', -R_15[i] / 2 - 75 - 140, 'XZ.PLANE', 1)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME18.1', -R[i] / 2 - 75 - 140, 'XZ.PLANE', 1)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME18.1', -T[i], 'XY.PLANE', 0)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME18.1', F[i] / 2 - 80 - 37.5, 'YZ.PLANE', 0)
        if l == 0:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME19.1', -R_15[i] / 2 - 75 - 140, 'XZ.PLANE', 1)
        else:
            mprog.add_offset_assembly('BOLSTER1.1', 'FRAME19.1', -R[i] / 2 - 75 - 140, 'XZ.PLANE', 1)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME19.1', -T[i], 'XY.PLANE', 0)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME19.1', -F[i] / 2 + 80 + 37.5, 'YZ.PLANE', 0)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME17.1', 0, 'XZ.PLANE', 1)
        mprog.add_offset_assembly('BOLSTER1.1', 'FRAME17.1', -T[i], 'XY.PLANE', 0)
        if l == 0:
            mprog.add_offset_assembly('FRAME3.1', 'FRAME17.1', FRAME2_lower_depth_15[i], 'YZ.PLANE', 0)
        else:
            mprog.add_offset_assembly('FRAME3.1', 'FRAME17.1', FRAME2_lower_depth[i], 'YZ.PLANE', 0)
        # 左右側板前GIB
        if i == 4:
            mprog.add_offset_assembly('GIB1.1', 'FRAME3.1', FRAME1_lower_high[i] + 40 - 34.5, 'XY.PLANE', 0)
        else:
            mprog.add_offset_assembly('GIB1.1', 'FRAME3.1', FRAME1_lower_high[i] + 40, 'XY.PLANE', 0)
        mprog.add_offset_assembly('GIB1.1', 'FRAME1.1', 72.5, 'XZ.PLANE', 0)
        mprog.add_offset_assembly('GIB1.1', 'FRAME5.1', 334.65, 'YZ.PLANE', 0)
        if i == 4:
            mprog.add_offset_assembly('GIB2.1', 'FRAME3.1', FRAME1_lower_high[i] + 40 - 34.5, 'XY.PLANE', 0)
        else:
            mprog.add_offset_assembly('GIB2.1', 'FRAME3.1', FRAME1_lower_high[i] + 40, 'XY.PLANE', 0)
        mprog.add_offset_assembly('GIB2.1', 'FRAME2.1', -72.5, 'XZ.PLANE', 0)
        mprog.add_offset_assembly('GIB2.1', 'FRAME5.1', 334.65, 'YZ.PLANE', 0)
        # 左GIB後鎖固用方塊
        if i == 4:
            mprog.add_offset_assembly('GIB2.1', 'FRAME23.1', -690 - 34.5, 'XY.PLANE', 1)
        else:
            mprog.add_offset_assembly('GIB2.1', 'FRAME23.1', -690, 'XY.PLANE', 1)
        mprog.add_offset_assembly('FRAME2.1', 'FRAME23.1', 50 + 35, 'XZ.PLANE', 1)
        mprog.add_offset_assembly('GIB2.1', 'FRAME23.1', 0, 'YZ.PLANE', 0)
        if i == 4:
            mprog.add_offset_assembly('GIB2.1', 'FRAME24.1', -34.5, 'XY.PLANE', 1)
        else:
            mprog.add_offset_assembly('GIB2.1', 'FRAME24.1', 0, 'XY.PLANE', 1)
        mprog.add_offset_assembly('FRAME2.1', 'FRAME24.1', 50 + 35, 'XZ.PLANE', 0)
        mprog.add_offset_assembly('GIB2.1', 'FRAME24.1', 0, 'YZ.PLANE', 1)
        if i == 4:
            mprog.add_offset_assembly('GIB2.1', 'FRAME27.1', -690 / 2 - 34.5, 'XY.PLANE', 1)
        else:
            mprog.add_offset_assembly('GIB2.1', 'FRAME27.1', -690 / 2, 'XY.PLANE', 1)
        mprog.add_offset_assembly('FRAME2.1', 'FRAME27.1', 50, 'XZ.PLANE', 1)
        mprog.add_offset_assembly('GIB2.1', 'FRAME27.1', 0, 'YZ.PLANE', 1)
        if i == 4:
            mprog.add_offset_assembly('GIB2.1', 'FRAME31.1', -690 - 34.5, 'XY.PLANE', 1)
        else:
            mprog.add_offset_assembly('GIB2.1', 'FRAME31.1', -690, 'XY.PLANE', 1)
        mprog.add_offset_assembly('FRAME23.1', 'FRAME31.1', -50, 'XZ.PLANE', 0)
        mprog.add_offset_assembly('FRAME23.1', 'FRAME31.1', -130.35, 'YZ.PLANE', 0)
        if i == 4:
            mprog.add_offset_assembly('GIB2.1', 'FRAME31.2', -150 - 34.5, 'XY.PLANE', 1)
        else:
            mprog.add_offset_assembly('GIB2.1', 'FRAME31.2', -150, 'XY.PLANE', 1)
        mprog.add_offset_assembly('FRAME23.1', 'FRAME31.2', -50, 'YZ.PLANE', 0)
        mprog.add_offset_assembly('FRAME23.1', 'FRAME31.2', -62.5, 'XZ.PLANE', 0)
        # 右GIB後鎖固用方塊
        if i == 4:
            mprog.add_offset_assembly('GIB1.1', 'FRAME25.1', -690 - 34.5, 'XY.PLANE', 1)
        else:
            mprog.add_offset_assembly('GIB1.1', 'FRAME25.1', -690, 'XY.PLANE', 1)
        mprog.add_offset_assembly('FRAME1.1', 'FRAME25.1', -50 - 35, 'XZ.PLANE', 0)
        mprog.add_offset_assembly('GIB1.1', 'FRAME25.1', 0, 'YZ.PLANE', 1)
        if i == 4:
            mprog.add_offset_assembly('GIB1.1', 'FRAME26.1', -34.5, 'XY.PLANE', 1)
        else:
            mprog.add_offset_assembly('GIB1.1', 'FRAME26.1', 0, 'XY.PLANE', 1)
        mprog.add_offset_assembly('FRAME1.1', 'FRAME26.1', -50 - 35, 'XZ.PLANE', 0)
        mprog.add_offset_assembly('GIB1.1', 'FRAME26.1', 0, 'YZ.PLANE', 0)
        if i == 4:
            mprog.add_offset_assembly('GIB1.1', 'FRAME28.1', -690 / 2 - 34.5, 'XY.PLANE', 1)
        else:
            mprog.add_offset_assembly('GIB1.1', 'FRAME28.1', -690 / 2, 'XY.PLANE', 1)
        mprog.add_offset_assembly('FRAME1.1', 'FRAME28.1', -50, 'XZ.PLANE', 0)
        mprog.add_offset_assembly('GIB1.1', 'FRAME28.1', 0, 'YZ.PLANE', 1)
        if i == 4:
            mprog.add_offset_assembly('GIB1.1', 'FRAME31.3', -690 - 34.5, 'XY.PLANE', 1)
        else:
            mprog.add_offset_assembly('GIB1.1', 'FRAME31.3', -690, 'XY.PLANE', 1)
        mprog.add_offset_assembly('FRAME25.1', 'FRAME31.3', -50, 'XZ.PLANE', 1)
        mprog.add_offset_assembly('FRAME25.1', 'FRAME31.3', 130.35, 'YZ.PLANE', 1)
        if i == 4:
            mprog.add_offset_assembly('GIB1.1', 'FRAME31.4', -150 - 34.5, 'XY.PLANE', 1)
        else:
            mprog.add_offset_assembly('GIB1.1', 'FRAME31.4', -150, 'XY.PLANE', 1)
        mprog.add_offset_assembly('FRAME31.3', 'FRAME31.4', 0, 'YZ.PLANE', 0)
        mprog.add_offset_assembly('FRAME31.3', 'FRAME31.4', 0, 'XZ.PLANE', 0)
        # 中間與平板鎖固方塊下板子
        mprog.add_offset_assembly('FRAME43.1', 'FRAME17.1', -48, 'XY.PLANE', 1)
        mprog.add_offset_assembly('FRAME43.1', 'FRAME17.1', 0, 'XZ.PLANE', 0)
        mprog.add_offset_assembly('FRAME43.1', 'FRAME17.1', 0, 'YZ.PLANE', 1)
        # 後面安裝馬達下板子
        mprog.add_offset_assembly('FRAME41.1', 'FRAME20.1', -1340, 'XY.PLANE', 0)  # 要找他所有變數
        mprog.add_offset_assembly('FRAME41.1', 'FRAME4.1', 0, 'XZ.PLANE', 0)
        mprog.add_offset_assembly('FRAME41.1', 'FRAME4.1', 0, 'YZ.PLANE', 0)
        # FRAME2中間下板子
        mprog.add_offset_assembly('FRAME32.1', 'BOLSTER1.1', 0, 'XZ.PLANE', 0)
        mprog.add_offset_assembly('FRAME32.1', 'FRAME3.1', FRAME_32_XY[i] + 19, 'XY.PLANE', 0)  # 要找他所有變數
        mprog.add_offset_assembly('FRAME32.1', 'FRAME30.1', 0, 'YZ.PLANE', 0)
        # 後方大軸承支架
        mprog.add_offset_assembly('FRAME10.1', 'FRAME34.2', -485.543, 'YZ.PLANE', 0)
        mprog.add_offset_assembly('FRAME34.2', 'FRAME10.1', FRAME_10_H[i] + 40, 'XY.PLANE', 1)
        if l == 0:
            mprog.add_offset_assembly('FRAME34.2', 'FRAME3.1', -R_15[i] / 2 - 90, 'XZ.PLANE', 0)
        else:
            mprog.add_offset_assembly('FRAME34.2', 'FRAME3.1', -R[i] / 2 - 90, 'XZ.PLANE', 0)
        if l == 0:
            mprog.add_offset_assembly('FRAME34.1', 'FRAME34.2', -R_15[i] - 180, 'XZ.PLANE', 1)
        else:
            mprog.add_offset_assembly('FRAME34.1', 'FRAME34.2', -R[i] - 180, 'XZ.PLANE', 1)
        mprog.add_offset_assembly('FRAME34.1', 'FRAME34.2', 0, 'XY.PLANE', 1)
        mprog.add_offset_assembly('FRAME34.1', 'FRAME34.2', 0, 'YZ.PLANE', 0)
        # 後方大軸承
        mprog.add_offset_assembly('FRAME35.1', 'FRAME3.1', 0, 'XZ.PLANE', 1)
        mprog.add_offset_assembly('FRAME35.1', 'FRAME20.1', -868, 'XY.PLANE', 0)
        mprog.add_offset_assembly('FRAME35.1', 'FRAME20.1', -FRAME20_FRAME2_YZ[i] - 32, 'YZ.PLANE', 0)
        # 後方馬達下支撐板
        mprog.add_offset_assembly('FRAME39.1', 'FRAME1.1', 50, 'XZ.PLANE', 0)
        mprog.add_offset_assembly('FRAME39.1', 'FRAME20.1', -360, 'XY.PLANE', 0)  # 要找他所有變數
        mprog.add_offset_assembly('FRAME39.1', 'FRAME7.1', 320, 'YZ.PLANE', 0)  # 要找他所有變數
        mprog.add_offset_assembly('FRAME38.1', 'FRAME2.1', -50, 'XZ.PLANE', 0)
        mprog.add_offset_assembly('FRAME38.1', 'FRAME20.1', 360, 'XY.PLANE', 1)  # 要找他所有變數
        mprog.add_offset_assembly('FRAME38.1', 'FRAME6.1', -320, 'YZ.PLANE', 1)  # 要找他所有變數
        mprog.add_offset_assembly('FRAME37.1', 'FRAME39.1', 0, 'XZ.PLANE', 0)
        mprog.add_offset_assembly('FRAME37.1', 'FRAME39.1', 19, 'XY.PLANE', 1)
        mprog.add_offset_assembly('FRAME37.1', 'FRAME39.1', 125, 'YZ.PLANE', 1)
        mprog.add_offset_assembly('FRAME37.2', 'FRAME39.1', 0, 'XZ.PLANE', 0)
        mprog.add_offset_assembly('FRAME37.2', 'FRAME39.1', 19, 'XY.PLANE', 1)
        mprog.add_offset_assembly('FRAME37.2', 'FRAME39.1', -125, 'YZ.PLANE', 1)
        mprog.add_offset_assembly('FRAME37.3', 'FRAME38.1', 0, 'XZ.PLANE', 1)
        mprog.add_offset_assembly('FRAME37.3', 'FRAME38.1', 0, 'XY.PLANE', 0)
        mprog.add_offset_assembly('FRAME37.3', 'FRAME38.1', -125, 'YZ.PLANE', 1)
        mprog.add_offset_assembly('FRAME37.4', 'FRAME38.1', 0, 'XZ.PLANE', 1)
        mprog.add_offset_assembly('FRAME37.4', 'FRAME38.1', 0, 'XY.PLANE', 0)
        mprog.add_offset_assembly('FRAME37.4', 'FRAME38.1', 125, 'YZ.PLANE', 1)
        # FRAME2右邊角鐵
        mprog.add_offset_assembly('FRAME33.1', 'FRAME2.1', 140, 'XZ.PLANE', 1)
        mprog.add_offset_assembly('FRAME33.1', 'FRAME15.1', -708, 'XY.PLANE', 0)  # 要找他所有變數
        mprog.add_offset_assembly('FRAME33.1', 'FRAME11.1', 313.984, 'YZ.PLANE', 0)  # 要找他所有變數
        # FRAME2中間半圓形零件
        mprog.add_offset_assembly('FRAME40.1', 'FRAME20.1', 130.5, 'XZ.PLANE', 0)
        mprog.add_offset_assembly('FRAME40.1', 'FRAME20.1', -320, 'XY.PLANE', 0)
        mprog.add_offset_assembly('FRAME40.1', 'FRAME20.1', FRAME20_FRAME2_YZ[i], 'YZ.PLANE', 1)
        # FRAME35上兩圓管
        mprog.add_offset_assembly('FRAME36.1', 'FRAME35.1', -272.236, 'XZ.PLANE', 1)
        mprog.add_offset_assembly('FRAME36.1', 'FRAME35.1', -272.236, 'XY.PLANE', 1)
        mprog.add_offset_assembly('FRAME36.1', 'FRAME35.1', 88, 'YZ.PLANE', 1)
        mprog.add_offset_assembly('FRAME36.2', 'FRAME35.1', 272.236, 'XZ.PLANE', 1)
        mprog.add_offset_assembly('FRAME36.2', 'FRAME35.1', -272.236, 'XY.PLANE', 1)
        mprog.add_offset_assembly('FRAME36.2', 'FRAME35.1', 88, 'YZ.PLANE', 1)
        # 後方馬達下支撐板上治具
        mprog.add_offset_assembly('Fixture.1', 'FRAME41.1', 95, 'XZ.PLANE', 0)
        mprog.add_offset_assembly('Fixture.1', 'FRAME41.1', 22 + 18, 'XY.PLANE', 0)
        mprog.add_offset_assembly('Fixture.1', 'FRAME41.1', 312.5, 'YZ.PLANE', 0)
        # 馬達下方板與後方橫板
        mprog.add_offset_assembly('FRAME42.1', 'FRAME41.1', 0, 'XZ.PLANE', 1)
        mprog.add_offset_assembly('FRAME42.1', 'FRAME41.1', 71.5, 'XY.PLANE', 1)
        mprog.add_offset_assembly('FRAME42.1', 'FRAME41.1', -19, 'YZ.PLANE', 1)
        # 平板2, 3組立
        mprog.add_offset_assembly('BOLSTER2.1', 'BOLSTER3.1', 0, 'XZ.PLANE', 0)
        mprog.add_offset_assembly('BOLSTER2.1', 'BOLSTER3.1', 0, 'XY.PLANE', 0)
        mprog.add_offset_assembly('BOLSTER2.1', 'BOLSTER3.1', 0, 'YZ.PLANE', 0)
        # 平板1, 3組立
        mprog.add_offset_assembly('BOLSTER1.1', 'BOLSTER3.1', 0, 'XZ.PLANE', 0)
        mprog.add_offset_assembly('BOLSTER1.1', 'BOLSTER3.1', 0, 'YZ.PLANE', 0)
        if k == 0:
            if h == 0:
                mprog.add_offset_assembly('BOLSTER1.1', 'BOLSTER3.1', DH_S[i], 'XY.PLANE', 0)
            elif h == 1:
                mprog.add_offset_assembly('BOLSTER1.1', 'BOLSTER3.1', DH_H[i], 'XY.PLANE', 0)
            else:
                mprog.add_offset_assembly('BOLSTER1.1', 'BOLSTER3.1', DH_P[i], 'XY.PLANE', 0)
        else:
            if h == 0:
                mprog.add_offset_assembly('BOLSTER1.1', 'BOLSTER3.1', DH_S[i] + D_S[i], 'XY.PLANE', 0)
            elif h == 1:
                mprog.add_offset_assembly('BOLSTER1.1', 'BOLSTER3.1', DH_H[i] + D_H[i], 'XY.PLANE', 0)
            else:
                mprog.add_offset_assembly('BOLSTER1.1', 'BOLSTER3.1', DH_P[i] + D_P[i], 'XY.PLANE', 0)
        # SLIDE跟平板3組立
        if i == 4:
            mprog.add_offset_product_assembly('SLIDE_UNIT_All.1', 'Geometrical Set.1', 'BOLSTER3.1', 267.38, 'XY.PLANE',
                                              1)
        else:
            mprog.add_offset_product_assembly('SLIDE_UNIT_All.1', 'Geometrical Set.1', 'BOLSTER3.1', 85, 'XY.PLANE', 1)
        mprog.add_offset_product_assembly('SLIDE_UNIT_All.1', 'Geometrical Set.1', 'BOLSTER3.1', 0, 'XZ.PLANE', 1)
        if i == 4:
            mprog.add_offset_product_assembly('SLIDE_UNIT_All.1', 'Geometrical Set.1', 'BOLSTER3.1', 138.171,
                                              'YZ.PLANE', 0)
        else:
            mprog.add_offset_product_assembly('SLIDE_UNIT_All.1', 'Geometrical Set.1', 'BOLSTER3.1', 0, 'YZ.PLANE', 0)
        # 離合器跟FRAME20組立
        if i == 4:
            mprog.add_offset_product_assembly('CRANK_SHAFT_All.1', 'Geometrical Set.2', 'FRAME20.1', -348.943,
                                              'XY.PLANE', 0)
        else:
            mprog.add_offset_product_assembly('CRANK_SHAFT_All.1', 'Geometrical Set.2', 'FRAME20.1', -457.052,
                                              'XY.PLANE', 0)
        mprog.add_offset_product_assembly('CRANK_SHAFT_All.1', 'Geometrical Set.2', 'FRAME20.1', 0, 'XZ.PLANE', 0)
        mprog.add_offset_product_assembly('CRANK_SHAFT_All.1', 'Geometrical Set.2', 'FRAME20.1', 0, 'YZ.PLANE', 0)
        # 氣壓缸跟FRAME20組立
        if i == 4:
            mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1', -87.5,
                                              'XY.PLANE', 1)
        else:
            mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1', -32, 'XY.PLANE',
                                              1)
        if i == 4:
            if l == 0:
                mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1', -500.5,
                                                  'XZ.PLANE', 0)
            else:
                mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1', -347,
                                                  'XZ.PLANE', 0)
        else:
            if l == 0:
                mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1',
                                                  -R_15[i] / 2 - 13, 'XZ.PLANE', 0)
            else:
                mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1',
                                                  -R[i] / 2 - 13, 'XZ.PLANE', 0)
        if i == 4:
            mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1', -261.248,
                                              'YZ.PLANE', 0)
        else:
            mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1', -260,
                                              'YZ.PLANE', 0)
        if i == 4:
            mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1', -21.8,
                                              'XY.PLANE', 1)
        else:
            mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1', -32, 'XY.PLANE',
                                              1)
        if i == 4:
            if l == 0:
                mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1', -500.5,
                                                  'XZ.PLANE', 1)
            else:
                mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1', -347,
                                                  'XZ.PLANE', 1)
        else:
            if l == 0:
                mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1',
                                                  -R_15[i] / 2 - 13, 'XZ.PLANE', 1)
            else:
                mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1',
                                                  -R[i] / 2 - 13, 'XZ.PLANE', 1)
        if i == 4:
            mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1', 261.248,
                                              'YZ.PLANE', 1)
        else:
            mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1', 260, 'YZ.PLANE',
                                              1)
        # 離合器與FRAME20結合
        if i == 4:
            mprog.add_offset_product_assembly('CLUCTH_ASSEMBLY_All.1', 'Geometrical Set.1', 'FRAME20.1', 108.109,
                                              'XY.PLANE', 0)
        else:
            mprog.add_offset_product_assembly('CLUCTH_ASSEMBLY_All.1', 'Geometrical Set.1', 'FRAME20.1', 0, 'XY.PLANE',
                                              0)
        mprog.add_offset_product_assembly('CLUCTH_ASSEMBLY_All.1', 'Geometrical Set.1', 'FRAME20.1', 0, 'XZ.PLANE', 1)
        mprog.add_offset_product_assembly('CLUCTH_ASSEMBLY_All.1', 'Geometrical Set.1', 'FRAME20.1', 510, 'YZ.PLANE', 1)

        # 更新
        mprog.update()
        mprog.Close_All()


if __name__ == "__main__":
    app = QtWidgets.QApplication([])
    window = main()
    window.show()
    sys.exit(app.exec_())
