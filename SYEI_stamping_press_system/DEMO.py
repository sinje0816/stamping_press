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
import excel_parameter_change_1 as epc




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
        processing = str(self.ui.comboBox.currentText())
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
        self.i , self.p= self.choos(type , processing)
        self.change_dir(self.i, self.p, self.aplha, self.beta, self.zeta, self.delta, self.machining, self.welding)

    def choos(self, type, prossing):
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
        if prossing == '是':
            p = 0
        elif prossing == '否':
            p = 1

        return i , p

    def create_dir(self, type):  # 創建資料夾
        time_now = datetime.datetime.now()
        dir_name = '{}_{}_{}_{}_{}'.format(type, time_now.day, time_now.hour, time_now.minute, time_now.second)
        desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
        test = os.path.join("Z:")
        path = test + '\\' + dir_name
        os.mkdir(path)
        machining = path + "\\" + "machining"
        os.mkdir(machining)
        welding = path + "\\" + "welding"
        os.mkdir(welding)
        self.path = path
        self.machining = machining
        self.welding = welding

    def change_dir(self, i, p, alpha, beta, zeta, delta, machining, welding):
        # 開啟CATIA
        env = mprog.set_CATIA_workbench_env()
        # 開啟零件檔更改變數後儲存並關閉
        for name in par.test_list:
            mprog.import_part(fp.system_root + fp.DEMO_part, name)
            if name == "FRAME3":
                excel = epc.ExcelOp('FRAME3')
                try:
                    excel.part_parameter('FRAME3', i)
                    print('FRAME3 Parameter change success')
                except:
                    print('FRAME3 Parameter change error')
                try:
                    if i == 0:
                        mprog.activatefeature('SN1_25250_Body', 2)
                        mprog.activatefeature('Hole_1', 1)
                        mprog.activatefeature('Hole_2', 0)
                        mprog.partbodyfeatureactivate('FRAME_SN1_25250_PQ')
                        mprog.partbodyfeatureactivate('FRAME_SN1_2560_S')
                        mprog.partbodyfeatureactivate('FRAME_SN1_2560_U')
                        mprog.activatefeature('M10X20L(不可貫穿)', 0)
                    elif i == 1:
                        mprog.activatefeature('SN1_25250_Body', 2)
                        mprog.activatefeature('Hole_1', 1)
                        mprog.activatefeature('Hole_2', 0)
                        mprog.partbodyfeatureactivate('FRAME_SN1_25250_PQ')
                        mprog.partbodyfeatureactivate('FRAME_SN1_2560_S')
                        mprog.partbodyfeatureactivate('FRAME_SN1_2560_U')
                        mprog.activatefeature('M10X20L(不可貫穿)', 0)

                    elif i == 2:
                        mprog.activatefeature('SN1_25250_Body', 2)
                        mprog.activatefeature('Hole_1', 1)
                        mprog.activatefeature('Hole_2', 0)
                        mprog.partbodyfeatureactivate('FRAME_SN1_25250_PQ')
                        mprog.partbodyfeatureactivate('FRAME_SN1_2560_S')
                        mprog.partbodyfeatureactivate('FRAME_SN1_2560_U')
                        mprog.activatefeature('M10X20L(不可貫穿)', 0)
                    elif i == 3:
                        mprog.activatefeature('SN1_25250_Body', 2)
                        mprog.activatefeature('Hole_1', 1)
                        mprog.activatefeature('Hole_2', 0)
                        mprog.partbodyfeatureactivate('FRAME_SN1_25250_PQ')
                        mprog.partbodyfeatureactivate('FRAME_SN1_2560_S')
                        mprog.partbodyfeatureactivate('FRAME_SN1_2560_U')
                        mprog.activatefeature('M10X20L(不可貫穿)', 0)
                    elif i == 4:
                        mprog.activatefeature('SN1_25250_Body', 2)
                        mprog.activatefeature('Hole_1', 1)
                        mprog.activatefeature('Hole_2', 0)
                        mprog.partbodyfeatureactivate('FRAME_SN1_25250_PQ')
                        mprog.partbodyfeatureactivate('FRAME_SN1_80250_R')
                        mprog.activatefeature('M10X20L(不可貫穿)', 0)
                    elif i == 5:
                        mprog.activatefeature('SN1_25250_Body', 2)
                        mprog.activatefeature('Hole_1', 1)
                        mprog.activatefeature('Hole_2', 0)
                        mprog.partbodyfeatureactivate('FRAME_SN1_25250_PQ')
                        mprog.partbodyfeatureactivate('FRAME_SN1_80250_R')
                        mprog.activatefeature('M10X20L(不可貫穿)', 0)
                    elif i == 6:
                        mprog.activatefeature('SN1_25250_Body', 2)
                        mprog.activatefeature('Hole_1', 1)
                        mprog.activatefeature('Hole_2', 0)
                        mprog.partbodyfeatureactivate('FRAME_SN1_25250_PQ')
                        mprog.partbodyfeatureactivate('FRAME_SN1_80250_R')
                        mprog.activatefeature('M10X20L(不可貫穿)', 0)
                    elif i == 7:
                        mprog.activatefeature('SN1_25250_Body', 2)
                        mprog.activatefeature('Hole_1', 1)
                        mprog.activatefeature('Hole_2', 0)
                        mprog.partbodyfeatureactivate('FRAME_SN1_25250_PQ')
                        mprog.partbodyfeatureactivate('FRAME_SN1_80250_R')
                        mprog.activatefeature('M10X20L(不可貫穿)', 0)
                    elif i == 8:
                        mprog.activatefeature('SN1_25250_Body', 2)
                        mprog.activatefeature('Hole_1', 1)
                        mprog.activatefeature('Hole_2', 0)
                        mprog.partbodyfeatureactivate('FRAME_SN1_25250_PQ')
                        mprog.partbodyfeatureactivate('FRAME_SN1_80250_R')
                        mprog.activatefeature('M10X20L(不可貫穿)', 0)
                except:
                    print('FRAME3 Parameter activate error')
                finally:
                    try:
                        mprog.Update()
                        print('FRAME3 Update success')
                        mprog.save_file_stp(welding, name)
                        mprog.save_stpfile_part(welding, name)
                        mprog.bodydeactivate('Hole_1', 0)
                        mprog.bodydeactivate('Hole_2', 0)
                    except:
                        print('FRAME3 Update error')
                try:
                    if i == 0:
                        mprog.activatefeature('SN1_25250_Body', 0)
                        mprog.activatefeature('SN1_25250_M', 0)
                        mprog.activatefeature('Hole_1', 0)
                        mprog.activatefeature('Hole_2', 1)
                        mprog.partdeactivate('FRAME_SN1_25250_PQ')
                        mprog.partdeactivate('FRAME_SN1_2560_S')
                        mprog.partdeactivate('FRAME_SN1_2560_U')
                        mprog.bodydeactivate('M10X20L(不可貫穿)', 0)
                    elif i == 1:
                        mprog.activatefeature('SN1_25250_Body', 0)
                        mprog.activatefeature('SN1_25250_M', 0)
                        mprog.activatefeature('Hole_1', 0)
                        mprog.activatefeature('Hole_2', 1)
                        mprog.partdeactivate('FRAME_SN1_25250_PQ')
                        mprog.partdeactivate('FRAME_SN1_2560_S')
                        mprog.partdeactivate('FRAME_SN1_2560_U')
                        mprog.bodydeactivate('M10X20L(不可貫穿)', 0)
                    elif i == 2:
                        mprog.activatefeature('SN1_25250_Body', 0)
                        mprog.activatefeature('SN1_25250_M', 0)
                        mprog.activatefeature('Hole_1', 0)
                        mprog.activatefeature('Hole_2', 1)
                        mprog.partdeactivate('FRAME_SN1_25250_PQ')
                        mprog.partdeactivate('FRAME_SN1_2560_S')
                        mprog.partdeactivate('FRAME_SN1_2560_U')
                        mprog.bodydeactivate('M10X20L(不可貫穿)', 0)
                    elif i == 3:
                        mprog.activatefeature('SN1_25250_Body', 0)
                        mprog.activatefeature('Hole_1', 0)
                        mprog.activatefeature('Hole_2', 1)
                        mprog.partdeactivate('FRAME_SN1_25250_PQ')
                        mprog.partdeactivate('FRAME_SN1_2560_S')
                        mprog.partdeactivate('FRAME_SN1_2560_U')
                        mprog.bodydeactivate('M10X20L(不可貫穿)', 0)
                    elif i == 4:
                        mprog.activatefeature('SN1_25250_Body', 0)
                        mprog.activatefeature('SN1_25250_M', 0)
                        mprog.activatefeature('Hole_1', 0)
                        mprog.activatefeature('Hole_2', 1)
                        mprog.partdeactivate('FRAME_SN1_25250_PQ')
                        mprog.partdeactivate('FRAME_SN1_80250_R')
                        mprog.bodydeactivate('M10X20L(不可貫穿)', 0)
                    elif i == 5:
                        mprog.activatefeature('SN1_25250_Body', 0)
                        mprog.activatefeature('Hole_1', 0)
                        mprog.activatefeature('Hole_2', 1)
                        mprog.partdeactivate('FRAME_SN1_25250_PQ')
                        mprog.partdeactivate('FRAME_SN1_80250_R')
                        mprog.bodydeactivate('M10X20L(不可貫穿)', 0)
                    elif i == 6:
                        mprog.activatefeature('SN1_25250_Body', 0)
                        mprog.activatefeature('SN1_25250_M', 0)
                        mprog.activatefeature('Hole_1', 0)
                        mprog.activatefeature('Hole_2', 1)
                        mprog.partdeactivate('FRAME_SN1_25250_PQ')
                        mprog.partdeactivate('FRAME_SN1_80250_R')
                        mprog.bodydeactivate('M10X20L(不可貫穿)', 0)
                    elif i == 7:
                        mprog.activatefeature('SN1_25250_Body', 0)
                        mprog.activatefeature('SN1_25250_M', 0)
                        mprog.activatefeature('Hole_1', 0)
                        mprog.activatefeature('Hole_2', 1)
                        mprog.partdeactivate('FRAME_SN1_25250_PQ')
                        mprog.partdeactivate('FRAME_SN1_80250_R')
                        mprog.bodydeactivate('M10X20L(不可貫穿)', 0)
                    elif i == 8:
                        mprog.activatefeature('SN1_25250_Body', 4)
                        mprog.activatefeature('SN1_25250_M', 0)
                        mprog.activatefeature('Hole_1', 0)
                        mprog.activatefeature('Hole_2', 1)
                        mprog.partdeactivate('FRAME_SN1_25250_PQ')
                        mprog.partdeactivate('FRAME_SN1_80250_R')
                        mprog.bodydeactivate('M10X20L(不可貫穿)', 0)
                except:
                    print('FRAME3 Parameter activate error')
                finally:
                    try:
                        mprog.Update()
                        print('FRAME3 Update success')
                        mprog.save_file_stp(machining, name)
                        mprog.save_stpfile_part(machining, name)
                    except:
                        print('FRAME3 Update error')
            elif name == "FRAME4":
                excel = epc.ExcelOp('FRAME4')
                try:
                    excel.part_parameter('FRAME4', i)
                    print('FRAME4 Parameter change success')
                except:
                    print('FRAME4 Parameter change error')
                try:
                    if i == 0:
                        mprog.activatefeature('SN1_25250_Body', 1)
                        mprog.activatefeature('Hole_1', 1)
                        mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                    elif i == 1:
                        mprog.activatefeature('SN1_25250_Body', 1)
                        mprog.activatefeature('Hole_1', 1)
                        mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                    elif i == 2:
                        mprog.activatefeature('SN1_25250_Body', 1)
                        mprog.activatefeature('Hole_1', 1)
                        mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                    elif i == 3:
                        mprog.activatefeature('SN1_25250_Body', 1)
                        mprog.activatefeature('Hole_1', 1)
                        mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                        mprog.activatefeature('Hole_3', 2)
                    elif i == 4:
                        mprog.activatefeature('SN1_25250_Body', 1)
                        mprog.activatefeature('Hole_1', 1)
                        mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                        mprog.activatefeature('Hole_3', 2)
                        mprog.activatefeature('Hole_4', 0)
                    elif i == 5:
                        mprog.activatefeature('SN1_25250_Body', 1)
                        mprog.activatefeature('Hole_1', 1)
                        mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                        mprog.activatefeature('Hole_3', 2)
                        mprog.activatefeature('Hole_4', 0)
                    elif i == 6:
                        mprog.activatefeature('SN1_25250_Body', 1)
                        mprog.activatefeature('Hole_1', 1)
                        mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                        mprog.activatefeature('Hole_3', 2)
                        mprog.activatefeature('Hole_4', 0)
                    elif i == 7:
                        mprog.activatefeature('SN1_25250_Body', 1)
                        mprog.activatefeature('Hole_1', 1)
                        mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                        mprog.activatefeature('Hole_3', 2)
                        mprog.activatefeature('Hole_4', 0)
                    elif i == 8:
                        mprog.activatefeature('SN1_25250_Body', 1)
                        mprog.activatefeature('Hole_1', 1)
                        mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                        mprog.activatefeature('Hole_3', 4)
                        mprog.activatefeature('Hole_4', 0)
                except:
                    print('FRAME4 Parameter activate error')
                finally:
                    try:
                        mprog.Update()
                        mprog.save_file_stp(welding, name)
                        mprog.save_stpfile_part(welding, name)
                        print('FRAME4 Update success')
                    except:
                        print('FRAME4 Update error')
                try:
                    if i == 0:
                        mprog.activatefeature('SN1_25250_Body', 0)
                        mprog.activatefeature('Hole_1', 3)
                        mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                    elif i == 1:
                        mprog.activatefeature('SN1_25250_Body', 0)
                        mprog.activatefeature('Hole_1', 2)
                        mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                    elif i == 2:
                        mprog.activatefeature('SN1_25250_Body', 0)
                        mprog.activatefeature('Hole_1', 1)
                        mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                    elif i == 3:
                        mprog.activatefeature('SN1_25250_Body', 0)
                        mprog.activatefeature('Hole_1', 1)
                        mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                        mprog.activatefeature('Hole_3', 2)
                    elif i == 4:
                        mprog.activatefeature('SN1_25250_Body', 0)
                        mprog.activatefeature('Hole_1', 1)
                        mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                        mprog.activatefeature('Hole_3', 2)
                        mprog.activatefeature('Hole_4', 0)
                    elif i == 5:
                        mprog.activatefeature('SN1_25250_Body', 0)
                        mprog.activatefeature('Hole_1', 1)
                        mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                        mprog.activatefeature('Hole_3', 2)
                        mprog.activatefeature('Hole_4', 0)
                    elif i == 6:
                        mprog.activatefeature('SN1_25250_Body', 0)
                        mprog.activatefeature('Hole_1', 1)
                        mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                        mprog.activatefeature('Hole_3', 2)
                        mprog.activatefeature('Hole_4', 0)
                    elif i == 7:
                        mprog.activatefeature('SN1_25250_Body', 0)
                        mprog.activatefeature('Hole_1', 1)
                        mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                        mprog.activatefeature('Hole_3', 2)
                        mprog.activatefeature('Hole_4', 0)
                    elif i == 8:
                        mprog.activatefeature('SN1_25250_Body', 0)
                        mprog.activatefeature('Hole_1', 1)
                        mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                        mprog.activatefeature('Hole_3', 4)
                        mprog.activatefeature('Hole_4', 0)
                except:
                    print('FRAME4 Parameter activate error')
                finally:
                    try:
                        mprog.Update()
                        mprog.save_file_stp(welding, name)
                        mprog.save_stpfile_part(welding, name)
                        print('FRAME4 Update success')
                    except:
                        print('FRAME4 Update error')
            elif name == "FRAME6":
                excel = epc.ExcelOp('FRAME4')
                try:
                    excel.part_parameter('FRAME4', i)
                    print('FRAME4 Parameter change success')
                except:
                    print('FRAME4 Parameter change error')
                try:
                    if i == 0:
                        mprog.activatefeature('FRAME_SN1_2545_Body', 0)
                        mprog.activatefeature('M16_i', 0)
                    elif i == 1:
                        mprog.activatefeature('FRAME_SN1_2545_Body', 0)
                        mprog.activatefeature('M16_i', 0)
                    elif i == 2:
                        mprog.activatefeature('FRAME_SN1_2545_Body', 0)
                        mprog.activatefeature('M16_i', 0)
                    elif i == 3:
                        mprog.activatefeature('FRAME_SN1_60_Body', 0)
                        mprog.activatefeature('M16_h FRAME9_L', 0)
                    elif i == 4:
                        mprog.activatefeature('FRAME_SN1_80110_Body', 0)
                        mprog.activatefeature('M16_h FRAME9_L', 0)
                    elif i == 5:
                        mprog.activatefeature('FRAME_SN1_80110_Body', 0)
                        mprog.activatefeature('M16_h FRAME9_L', 0)
                    elif i == 7:
                        mprog.activatefeature('FRAME_SN1_200_Body', 0)
                        mprog.activatefeature('M16_h FRAME9_L', 0)
                except:
                    print('FRAME6 Parameter activate error')
                finally:
                    try:
                        mprog.Update()
                        mprog.save_file_stp(welding, name)
                        mprog.save_stpfile_part(welding, name)
                        print('FRAME6 Update success')
                    except:
                        print('FRAME6 Update error')
                try:
                    if i == 0:
                        mprog.bodydeactivate('FRAME_SN1_2545_Body', 0)
                        mprog.bodydeactivate('M16_i', 0)
                        mprog.activatefeature('SN1_2545_Body', 0)
                    elif i == 1:
                        mprog.bodydeactivate('FRAME_SN1_2545_Body', 0)
                        mprog.bodydeactivate('M16_i', 0)
                        mprog.activatefeature('SN1_2545_Body', 0)
                    elif i == 2:
                        mprog.bodydeactivate('FRAME_SN1_2545_Body', 0)
                        mprog.bodydeactivate('M16_i', 0)
                        mprog.activatefeature('SN1_2545_Body', 0)
                    elif i == 3:
                        mprog.bodydeactivate('FRAME_SN1_60_Body', 0)
                        mprog.bodydeactivate('M16_h FRAME9_L', 0)
                        mprog.activatefeature('SN1_60_Body', 0)
                    elif i == 4:
                        mprog.bodydeactivate('FRAME_SN1_80110_Body', 0)
                        mprog.bodydeactivate('M16_h FRAME9_L', 0)
                        mprog.activatefeature('SN1_80110_Body', 0)
                    elif i == 5:
                        mprog.bodydeactivate('FRAME_SN1_80110_Body', 0)
                        mprog.bodydeactivate('M16_h FRAME9_L', 0)
                        mprog.activatefeature('SN1_80110_Body', 0)
                    elif i == 7:
                        mprog.bodydeactivate('FRAME_SN1_200_Body', 0)
                        mprog.bodydeactivate('M16_h FRAME9_L', 0)
                        mprog.activatefeature('SN1_200_Body', 0)
                except:
                    print('FRAME6 Parameter activate error')
                finally:
                    try:
                        mprog.Update()
                        mprog.save_file_stp(welding, name)
                        mprog.save_stpfile_part(welding, name)
                        print('FRAME6 Update success')
                    except:
                        print('FRAME6 Update error')










if __name__ == "__main__":
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)  # 自適應屏幕分辨率
    app = QtWidgets.QApplication([])
    window = main()
    window.show()
    sys.exit(app.exec_())