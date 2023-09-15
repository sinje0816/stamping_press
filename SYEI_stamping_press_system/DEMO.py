from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget, QPushButton, QLineEdit, QLabel, QTableWidget, QTableWidgetItem
from PyQt5.QtCore import QObject
from DEMOGUI import Ui_Dialog
from PAD_main import Ui_Form as pad_main_Form
from PAD_MACHINING import Ui_Form as pad_machining_Form
from PAD_dimension import Ui_Form as pad_dimension_Form
from pad_feeding_hole import Ui_Form as pad_feeding_hole_Form
from cutout_hole_GUI import Ui_Form as cutout_hole_machining_form
from io import StringIO
import main_program as mprog
import file_path as fp
import parameter as par
import Assmebly_design as Ad
import machining_part_TYPE_change as mptc
import welding_part_TYPE_change as wptc
import excel_parameter_change as epc
import parameter_design_part as pdp
import interference as itf
import test_T as tT
import sys
import datetime
import os
import time
import traceback



test_stop = False
class main(QtWidgets.QWidget, Ui_Dialog):
    def __init__(self):
        super(main, self).__init__()
        self.ui = Ui_Dialog()
        self.ui.setupUi(self)
        self.ui.pushButton.clicked.connect(self.start)
        self.ui.pushButton_2.clicked.connect(self.showPadwindows)
        self.ui.comboBox_4.currentIndexChanged.connect(lambda :self.change_label_7())
        self.ui.comboBox_2.currentIndexChanged.connect(lambda :self.change_label_7())
        self.ui.comboBox_4.currentIndexChanged.connect(lambda :self.change_label_9())
        self.ui.comboBox_2.currentIndexChanged.connect(lambda :self.change_label_9())


    def start(self):
        global test_stop
        global type
        type = str(self.ui.comboBox_4.currentText())
        travel_type = str(self.ui.comboBox_2.currentText())
        travel = str(self.ui.label_7.text())
        specifications_travel_value = str(self.ui.lineEdit_5.text())
        specifications_close_working_height_value = str(self.ui.lineEdit_2.text())
        close_working_height = str(self.ui.label_9.text())
        # delta = str(self.ui.lineEdit_4.text())
        processing = str(self.ui.comboBox.currentText())
        print(type, travel_type, travel, specifications_travel_value, specifications_close_working_height_value,
              close_working_height)
        self.create_dir(type)
        if specifications_travel_value == "":
            self.specifications_travel_value = 0
        else:
            self.specifications_travel_value = int(specifications_travel_value)
        if specifications_close_working_height_value == "":
            self.specifications_close_working_height_value = 0
        else:
            self.specifications_close_working_height_value = int(specifications_close_working_height_value)
        self.i, self.p, self.travel_type = self.choos(type, processing, travel_type)
        global i

        self.alpha, self.beta, self.zeta, self.epsilon = self.frame_calculate(self.i, self.specifications_travel_value,
                                                                                  self.specifications_close_working_height_value,
                                                                                  self.travel_type)
        if test_stop == False:
            self.create_txt(self.path, type, travel_type, self.specifications_travel_value, self.specifications_close_working_height_value, self.alpha, self.beta, self.zeta, self.epsilon)
            self.change_dir(self.i, self.p, self.alpha, self.beta,self.zeta, self.epsilon , self.machining, self.welding)

    def showPadwindows(self):
        self.hide()
        self.nw = padwindows()
        self.nw.show()


    def choos(self, type, prossing, travel_type):
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
        # 確認加工方式
        if prossing == '是':
            p = 0
        elif prossing == '否':
            p = 1
        # 確認行程類型
        if travel_type == 'S':
            travel_type = 1
        elif travel_type == 'H':
            travel_type = 2
        elif travel_type == 'P':
            travel_type = 3
        return i, p, travel_type

    def frame_calculate(self, i, specifications_travel_value, specifications_close_working_height_value, travel_type):
        Form = QtWidgets.QWidget()
        # Form.setWindowTitle('警告')
        Form.resize(400, 300)
        mbox = QtWidgets.QMessageBox(Form)

        # 讀取標準資料
        excel = epc.ExcelOp('尺寸整理表', '標準資料')
        type_name, travel_value, close_working_height_value, specifications_travel_min_value, specifications_travel_max_value, specifications_close_working_height_min_value, specifications_close_working_height_max_value = excel.get_standard_parts(
            i * 3)

        # 噸數&行程類型
        all_type = i * 3 + travel_type - 1
        # 行程
        alpha = int(specifications_travel_value) - travel_value[all_type]
        # 牙球伸長量
        epsilon_1 = (-alpha / 2) + (close_working_height_value[all_type] - int(specifications_close_working_height_value))
        elsilon_2 = (alpha / 2)
        epsilon_3 = 0
        epsilon = max(epsilon_1, elsilon_2, epsilon_3)
        # 喉部拉高量
        zeta = (alpha / 2) + epsilon + specifications_close_working_height_value - close_working_height_value[all_type]
        print(zeta)
        # 閉合工作高度驗證公式
        verification_specifications_close_working_height_value = close_working_height_value[all_type] - (
                alpha / 2) - epsilon + zeta
        print(verification_specifications_close_working_height_value)

        error = False

        # 驗證行程
        if specifications_travel_value >= specifications_travel_min_value[all_type] and specifications_travel_value <= \
                specifications_travel_max_value[all_type]:
            print('行程尺寸在允許範圍: %s <= %s <= %s' % (
                specifications_travel_min_value[all_type], specifications_travel_value,
                specifications_travel_max_value[all_type]))
        else:
            print('行程尺寸不再允許範圍: %s <= %s <= %s' % (
                specifications_travel_min_value[all_type], specifications_travel_value,
                specifications_travel_max_value[all_type]))
            mbox.warning(Form, '警告', '行程超出範圍')
            error = True
            self.ui.lineEdit_5.clear()

        # 驗證閉合工作高度
        if specifications_close_working_height_value >= specifications_close_working_height_min_value[
            all_type] and specifications_close_working_height_value <= \
                specifications_close_working_height_max_value[all_type]:
            print('閉合工作高度在允許範圍: %s <= %s <= %s' % (
                specifications_close_working_height_min_value[all_type], specifications_close_working_height_value,
                specifications_close_working_height_max_value[all_type]))
        else:
            print('閉合工作高度不在允許範圍: %s <= %s <= %s' % (
                specifications_close_working_height_min_value[all_type], specifications_close_working_height_value,
                specifications_close_working_height_max_value[all_type]))
            mbox.warning(Form, '警告', '閉合工作高度超出範圍')
            error = True
            self.ui.lineEdit_2.clear()

        # 驗證閉合工作高度是否符合客戶需求
        if verification_specifications_close_working_height_value == specifications_close_working_height_value:
            print('閉合工作高度符合要求: %s = %s' % (verification_specifications_close_working_height_value, specifications_close_working_height_value))
        else:
            print(
                'verification_specifications_close_working_height_value: %s；specifications_close_working_height_value: %s' % (
                verification_specifications_close_working_height_value, specifications_close_working_height_value))
            mbox.warning(Form, '警告', '閉合工作高度不符合客戶需求')
            error = True
            self.ui.lineEdit_2.clear()
            self.ui.lineEdit_5.clear()

        # 驗證牙球伸長量
        if epsilon >= 0:
            print('牙球伸長量: %s >= 0' % epsilon)
        else:
            print('牙球伸長量: %s <= 0' % epsilon)
            mbox.warning(Form, '警告', '牙球伸長量小於0')
            error = True
            self.ui.lineEdit_2.clear()
            self.ui.lineEdit_5.clear()

        # 驗證機架拉高量
        if zeta >= 0:
            print('機架拉高量: %s >= 0' % zeta)
        else:
            print('機架拉高量: %s <= 0' % zeta)
            mbox.warning(Form, '警告', '機架拉高量小於0')
            error = True
            self.ui.lineEdit_2.clear()
            self.ui.lineEdit_5.clear()

        # 驗證喉部拉高量
        if epsilon >= (alpha / 2):
            print('喉部拉高量: %s >= %s' % (epsilon, (alpha / 2)))
        else:
            mbox.warning(Form, '警告', '衝頭縮進導軌內部')
            error = True
            self.ui.lineEdit_2.clear()
            self.ui.lineEdit_5.clear()

        if error == True:
            global test_stop
            test_stop = True

        alpha = alpha  # 行程差
        beta = specifications_close_working_height_value - close_working_height_value[all_type]  # 閉合工作高度差
        zeta = zeta  # 喉部拉高量
        epsilon = epsilon  # 牙球伸長量
        return alpha, beta, zeta, epsilon

    # 建立txt檔
    def create_txt(self, path, travel_type, specifications_travel_value, specifications_close_working_height_value, type, alpha, beta, zeta, epsilon):
        file_txt = path
        txt_name = "生成參數.txt"
        with open(file_txt + "\\" + txt_name, "w") as f:
            f.write("噸數=%s\n" % type)
            f.write("型式:%s\n" % travel_type)
            f.write("本次行程=%s\n" % specifications_travel_value)
            f.write("本次閉合工作高度=%s\n" % specifications_close_working_height_value)
            f.write("行程=%s\n" % alpha)
            f.write("閉合工作高度=%s\n" % beta)
            f.write("喉部拉高量=%s\n" % zeta)
            f.write("牙球伸長量=%s\n" % epsilon)

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

    def finish(self, machining_file_change_error, welding_file_change_error):
        Form = QtWidgets.QWidget()
        Form.setWindowTitle('oxxo.studio')
        Form.resize(500, 200)
        mbox = QtWidgets.QMessageBox(Form)
        mbox.information(Form, '完成', '生成完成\nmachining_file_change_error:%s\nwelding_file_change_error:%s\n' %(machining_file_change_error, welding_file_change_error))
        self.ui.lineEdit_5.clear()
        self.ui.lineEdit_2.clear()

    def label_7_change_data(self):
        label_7_data = {250: {"S": "標準:80", "H": ("標準:50"), "P": ("標準:35")},
                350: {"S": ("標準:90"), "H": ("標準:60"), "P": ("標準:40")},
                450: {"S": ("標準:110"), "H": ("標準:70"), "P": ("標準:45")},
                600: {"S": ("標準:130"), "H": ("標準:80"), "P": ("標準:50")},
                800: {"S": ("標準:150"), "H": ("標準:100"), "P": ("標準:60")},
                1100: {"S": ("標準:180"), "H": ("標準:110"), "P": ("標準:70")},
                1600: {"S": ("標準:200"), "H": ("標準:130"), "P": ("標準:80")},
                2000: {"S": ("標準:220"), "H": ("標準:150"), "P": ("標準:90")},
                2500: {"S": ("標準:250"), "H": ("標準:180"), "P": ("標準:100")},
                }
        return label_7_data

    def change_label_7(self):
        label_7_data = self.label_7_change_data()
        type = str(self.ui.comboBox_4.currentText())
        travel_type = str(self.ui.comboBox_2.currentText())
        ton = int(type.split('-')[-1] + '0')
        travel_standard = label_7_data[ton][travel_type]
        travel_standard = str(travel_standard)

        self.ui.label_7.clear()
        self.ui.label_7.setText(travel_standard)

    def label_9_data(self):
        label_9_data = {250: {"S": ("標準:230"), "H": ("標準:200"), "P": ("標準:200")},
                350: {"S": ("標準:250"), "H": ("標準:220"), "P": ("標準:220")},
                450: {"S": ("標準:270"), "H": ("標準:240"), "P": ("標準:240")},
                600: {"S": ("標準:300"), "H": ("標準:270"), "P": ("標準:270")},
                800: {"S": ("標準:330"), "H": ("標準:300"), "P": ("標準:300")},
                1100: {"S": ("標準:350"), "H": ("標準:320"), "P": ("標準:320")},
                1600: {"S": ("標準:400"), "H": ("標準:360"), "P": ("標準:360")},
                2000: {"S": ("標準:450"), "H": ("標準:400"), "P": ("標準:400")},
                2500: {"S": ("標準:450"), "H": ("標準:400"), "P": ("標準:400")},}

        return label_9_data

    def change_label_9(self):
        label_9_data = self.label_9_data()
        type = str(self.ui.comboBox_4.currentText())
        travel_type = str(self.ui.comboBox_2.currentText())
        ton = int(type.split('-')[-1] + '0')
        close_h = label_9_data[ton][travel_type]
        close_h = str(close_h)

        self.ui.label_9.clear()
        self.ui.label_9.setText(close_h)

    def change_dir(self, i, p, alpha, beta, zeta, epsilon, machining, welding):
        start_time = time.time()
        all_part_name = {}
        all_part_value = {}
        all_parameter_save = {}
        all_parameter_value = {}
        # 開啟CATIA
        env = mprog.set_CATIA_workbench_env()
        machining_file_change_error = []
        machining_file_change_pass = []
        welding_file_change_error = []
        welding_file_change_pass = []
        # 開啟零件檔更改變數後儲存並關閉
        for name in epc.ExcelOp('尺寸整理表', '沖床機架零件清單').get_col_cell(1):
            print(name)
            file_list_name, file_list_value = epc.ExcelOp('尺寸整理表', '沖床機架零件清單').get_sheet_par('沖床機架零件清單', i)
            file_list_name_index = file_list_name.index(name)
            if file_list_value[file_list_name_index] == 0:
                pass
            else:
                try:
                    # 保存原始的sys.stdout
                    original_stdout = sys.stdout
                    # 创建一个新的StringIO对象来捕获输出
                    captured_output = StringIO()
                    sys.stdout = captured_output
                    mprog.import_part(fp.system_root + fp.DEMO_part, name)
                    if name == 'FRAME52' and p == 0:
                        try:
                            mprog.param_change(name, "alpha", alpha)
                            mprog.param_change(name, "beta", beta)
                            mprog.param_change(name, 'zeta', zeta)
                        except:
                            pass
                        # 加工圖零件
                        parameter_name, parameter_value = mptc.change_machining_parameter(name, i, 0)
                        all_part_name[name] = parameter_name
                        all_part_value[name] = parameter_value
                        for x in range(len(parameter_name)):
                            all_parameter_list.setdefault(parameter_name[x], parameter_value[x])
                            all_parameter_value[name] = all_parameter_list
                            apv = all_parameter_value

                        # 恢复原始的sys.stdout
                        sys.stdout = original_stdout
                        # 从捕获的输出中获取文本
                        output_text = captured_output.getvalue()

                        # 判断文本内容
                        if "error" in output_text:
                            machining_file_change_error.append(name)
                        else:
                            machining_file_change_pass.append(name)
                            mprog.close_file(name)
                    else:
                        try:
                            mprog.param_change(name, "alpha", alpha)
                            mprog.param_change(name, "beta", beta)
                            mprog.param_change(name, 'zeta', zeta)
                        except:
                            pass

                        # 加工圖零件
                        parameter_name, parameter_value = mptc.change_machining_parameter(name, i, 1)
                        all_part_name[name] = parameter_name
                        all_part_value[name] = parameter_value
                        for x in range(len(parameter_name)):
                            all_parameter_save.setdefault(parameter_name[x], parameter_value[x])
                            all_parameter_list = all_parameter_save.copy()
                            all_parameter_value[name] = all_parameter_list
                            apv = all_parameter_value
                        all_parameter_save.clear()

                        # 恢复原始的sys.stdout
                        sys.stdout = original_stdout
                        output_text = captured_output.getvalue()
                        # 判断文本内容
                        if "error" in output_text:
                            machining_file_change_error.append(name)
                        else:
                            machining_file_change_pass.append(name)
                    # 儲存加工圖零件檔
                    mprog.save_file_stp(machining, name)
                    mprog.save_stpfile_part(machining, name)
                    # 進行裁料圖特徵變更
                    wptc.change_welding_feature(name, i)
                    print(output_text)
                    if "error" in output_text:
                        welding_file_change_error.append(name)
                    else:
                        welding_file_change_pass.append(name)
                    # 儲存裁料圖零件檔
                    mprog.save_file_stp(welding, name)
                    mprog.save_stpfile_part(welding, name)
                    mprog.close_file(name)
                except:
                    pass
        print(i)
        print('all_part_name', all_part_name)
        print('all_part_value', all_part_value)
        print('machining_file_change_error', machining_file_change_error)
        print('machining_file_change_pass', machining_file_change_pass)
        print('welding_file_change_error', welding_file_change_error)
        print('welding_file_change_pass', welding_file_change_pass)
        print('總用時%s' % (time.time() - start_time))#建立3D組立
        Ad.assembly(i, apv, self.path)

        return machining_file_change_error, welding_file_change_error

# 平板主頁面
class padwindows(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.ui = pad_main_Form()
        self.ui.setupUi(self)
        self.setWindowTitle('平板')
        self.ui.t_dimension.clicked.connect(self.showpaddimensionwindows)
        self.ui.t_machining.clicked.connect(self.showpadmachiningwindows)
        self.ui.plate_start.clicked.connect(self.start)
        self.ui.remove_machining.clicked.connect(self.showcutoutmachiningwindows)
        self.ui.remove_machining.clicked.connect(self.showremovemachiningwindows)

    def showpaddimensionwindows(self):
        self.hide()
        self.nw = pad_dimension()
        self.nw.show()

    def showpadmachiningwindows(self):
        self.hide()
        self.nw = pad_machining()
        self.nw.show()

    def showcutoutmachiningwindows(self):
        par.plate_hole_type = [self.ui.remove_type.currentText()]
        self.hide()
        self.nw = cutout_hole_machining()
        self.nw.show()


    def showremovemachiningwindows(self):
        self.hide()
        self.nw = remove_machining()
        self.nw.show()


    def start(self, i):
        pad_type = str(self.ui.pad_type.currentText())
        # 對平板進行變數變換
        try:
            self.create_dir('plate')
        except:
            print('資料夾生成路徑連接失敗')
        mprog.import_part(fp.system_root + fp.DEMO_part, 'plate')
        plate_name, plate_value = pdp.padchange(i)
        for name in plate_name:
            par.plate_all_parameter[name] = plate_value[plate_name.index(name)]
        print(par.plate_all_parameter)

        try:
            if pad_type == '標準':
                par.lv = [0]
            elif pad_type == '加大一型':
                par.lv = [epc.ExcelOp('平板', '平板尺寸').get_single_data_sheet_par('平板尺寸', 'LV1', i)]
            elif pad_type == '加大二型':
                par.lv = [epc.ExcelOp('平板', '平板尺寸').get_single_data_sheet_par('平板尺寸', 'LV2', i)]
            print('plate type LV update successfully')
        except:
            print('plate type LV update error')

        # 對T型槽進行變數變換
        mprog.import_part(fp.system_root + fp.DEMO_part, 'T')
        pdp.padchange(i)
        try:
            for t in par.t_all_dimension:
                for t_name in range(len(par.t_all_dimension_name)+1):
                    print(par.t_all_dimension_name[t_name], t)
                    mprog.param_change('T', par.t_all_dimension_name[t_name], t)
                    break
            print('T-slot parameter change successfully')
        except:
            print('T-slot parameter change error')
        mprog.Update()
        mprog.save_file_stp(self.path, 'T')
        mprog.save_stpfile_part(self.path, 'T')
        mprog.close_window()
        # 將T型槽移至平板上進行除料
        try:
            for turn in range(0, len(par.total_pierce)):
                print("T形槽方向:", par.total_t_direction[turn])
                print("貫穿:", par.total_pierce[turn])
                print("間隙孔(讓孔):", par.total_clearance_hole[turn])
                print("T形槽尺寸:", par.total_t_dimension[turn])
                print("非貫穿尺寸:", par.total_unpierce[turn])
                mprog.import_part(self.path, 'T')
                if par.total_pierce[turn] == '是':
                    if par.total_t_direction[turn] == '前後':
                        mprog.param_change('T', 'Depth', par.plate_all_parameter['B'])
                    elif par.total_t_direction[turn] == '左右':
                        mprog.param_change('T', 'Depth', (par.plate_all_parameter['A']+par.lv[0]))
                else:
                    mprog.param_change('T', 'Depth', par.total_unpierce[turn])
                    if par.total_t_direction[turn] == '前後':
                        mprog.param_change('T', 'mirror', par.plate_all_parameter['B'])
                        print(par.plate_all_parameter['B']/2)
                    elif par.total_t_direction[turn] == '左右':
                        mprog.param_change('T', 'mirror', (par.plate_all_parameter['A']+par.lv[0]))
                        print((par.plate_all_parameter['A']+par.lv[0])/2)
                    mprog.partbodyfeatureactivate('Mirror.3')

                if par.total_clearance_hole[turn] == '是':
                    mprog.partbodyfeatureactivate('讓孔')
                    mprog.partbodyfeatureactivate('讓孔倒圓角')
                else:
                    pass

                if par.total_t_direction[turn] == '前後':
                    tT.changerotate(-90)
                elif par.total_t_direction[turn] == '左右':
                    tT.changerotate(0)
                else:
                    tT.changerotate(45)

                mprog.Update()
                if par.total_t_direction[turn] == '前後':
                    tT.create_t_solt(-par.total_t_dimension[turn], turn, par.total_pierce[turn], par.total_clearance_hole[turn], par.total_t_direction[turn], par.lv[0])
                elif par.total_t_direction[turn] == '左右':
                    tT.create_t_solt(par.total_t_dimension[turn], turn, par.total_pierce[turn], par.total_clearance_hole[turn], par.total_t_direction[turn], par.lv[0])
            print('T-slot create successfully')
        except Exception as e:
            print('T-slot create error:', e)
            error_class = e.__class__.__name__  # 取得錯誤類型
            detail = e.args[0]  # 取得詳細內容
            cl, exc, tb = sys.exc_info()  # 取得Call Stack
            lastCallStack = traceback.extract_tb(tb)[-1]  # 取得Call Stack的最後一筆資料
            fileName = lastCallStack[0]  # 取得發生的檔案名稱
            lineNum = lastCallStack[1]  # 取得發生的行號
            funcName = lastCallStack[2]  # 取得發生的函數名稱
            errMsg = "File \"{}\", line {}, in {}: [{}] {}".format(fileName, lineNum, funcName, error_class, detail)
            print(errMsg)

        # remove_type = self.ui.remove_type.currnetText()
        # # 下料口
        # if remove_type != '無孔':
        #     try:
        #         if remove_type == '圓形':
        #             mprog.param_change('cutout_hole_circle', 'C', par.plate_all_parameter['C'])
        #             mprog.param_change('cutout_hole_circle', 'HD', HD)
        #         elif remove_type == '方孔':
        #
        #             mprog.param_change('cutout_hole_square', 'C', par.plate_all_parameter['C'])
        #             mprog.param_change('cutout_hole_square', 'HFB', HFB)
        #             mprog.param_change('cutout_hole_square', 'HLR', HLR)
        #         elif remove_type == '漏斗型':
        #             HULR = self.ui.HULR.text()
        #             HDLR = self.ui.HDLR.text()
        #             HH = self.ui.HH.text()
        #             HUFB = self.ui.HUFB.text()
        #             HDFB = self.ui.HDFB.text()
        #             mprog.param_change('cutout_funnel', 'C', par.plate_all_parameter['C'])
        #             mprog.param_change('cutout_funnel', 'HULR', HULR)
        #             mprog.param_change('cutout_funnel', 'HDLR', HDLR)
        #             mprog.param_change('cutout_funnel', 'HH', HH)
        #             mprog.param_change('cutout_funnel', 'HUFB', HUFB)
        #             mprog.param_change('cutout_funnel', 'HDFB', HDFB)
        #
        #     except:
        #         print('remove_type error')
        # mprog.save_file_stp(self.path, 'plate')
        # mprog.save_stpfile_part(self.path, 'plate')

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

# T型槽外型尺寸
class pad_dimension(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.ui = pad_dimension_Form()
        self.ui.setupUi(self)
        self.setWindowTitle('平板')
        self.ui.setup_value.clicked.connect(self.setup)
        self.ui.escape.clicked.connect(self.showpadwindows)
        self.ui.reset_value.clicked.connect(self.reset)

    def setup(self):
        t_a = str(self.ui.A.text())
        t_b = str(self.ui.B.text())
        t_c = str(self.ui.C.text())
        t_d = str(self.ui.D.text())
        par.t_all_dimension = [t_a, t_b, t_c, t_d]
        print("T型槽外型尺寸:", par.t_all_dimension)
        self.hide()
        self.nw = padwindows()
        self.nw.show()

    def reset(self):
        self.ui.A.clear()
        self.ui.B.clear()
        self.ui.C.clear()
        self.ui.D.clear()

    def showpadwindows(self):
        self.hide()
        self.nw = padwindows()
        self.nw.show()

# T型槽加工
class pad_machining(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.ui = pad_machining_Form()
        self.ui.setupUi(self)
        self.setWindowTitle('平板加工設定')
        self.ui.escape.clicked.connect(self.showpadwindows)
        self.ui.new_order.clicked.connect(self.inputneworder)
        self.ui.setup.clicked.connect(self.start)
        self.ui.reset.clicked.connect(self.removeAction)

    def showpadwindows(self):
        data_entries = []
        self.hide()
        self.nw = padwindows()
        self.nw.show()

    def inputneworder(self):
        pierce = str(self.ui.pierce.currentText())  # 貫穿
        clearance_hole = str(self.ui.clearance_hole.currentText())  # 間隙孔(讓孔)
        t_direction = str(self.ui.direction.currentText())  # T形槽方向
        t_dimension = str(self.ui.position_dimension.text())  # T形槽尺寸
        unpierce = str(self.ui.unpierce.text())  # 非貫穿尺寸
        if t_dimension == "":
            t_dimension = 0
        data_entries = [pierce, clearance_hole, t_direction, t_dimension, unpierce]
        self.add_data(data_entries)

    def add_data(self, data_entries):
        # 取得目前表單的行數
        row_position = self.ui.tableWidget.rowCount()
        # 在表單中插入新的一行
        self.ui.tableWidget.insertRow(row_position)
        # 將手動輸入的四個資料新增到表單的下一列中
        for col, entry in enumerate(data_entries):
            item = QTableWidgetItem(entry)
            self.ui.tableWidget.setItem(row_position, col, item)
            self.ui.position_dimension.clear()
            self.ui.unpierce.clear()

    def start(self):
        # 將表單資料存至parameter.py指定串列
        for row in range(self.ui.tableWidget.rowCount()):
            pierce = self.ui.tableWidget.item(row, 0).text()
            clearance_hole = self.ui.tableWidget.item(row, 1).text()
            t_direction = self.ui.tableWidget.item(row, 2).text()
            t_dimension = self.ui.tableWidget.item(row, 3).text()
            unpierce = self.ui.tableWidget.item(row, 4).text()
            par.total_pierce.append(pierce)
            par.total_clearance_hole.append(clearance_hole)
            par.total_t_direction.append(t_direction)
            par.total_t_dimension.append(int(t_dimension))
            par.total_unpierce.append(unpierce)

        print("T形槽方向:", par.total_t_direction)
        print("貫穿:", par.total_pierce)
        print("間隙孔(讓孔):", par.total_clearance_hole)
        print("T形槽尺寸:", par.total_t_dimension)
        print("非貫穿尺寸:", par.total_unpierce)
        # itf.interference(i, par.lv[0], par.total_t_dimension, par.total_t_direction, par.plate_all_parameter['tw1'])

        self.hide()
        self.nw = padwindows()
        self.nw.show()

    def removeAction(self):
        # 取得目前表單的行數
        row_position = self.ui.tableWidget.rowCount()
        # 在表單中刪除最後一行
        self.ui.tableWidget.removeRow(row_position)



class cutout_hole_machining(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.ui = cutout_hole_machining_form()
        self.ui.setupUi(self)
        self.setWindowTitle('下料口位置')
        if par.plate_hole_type[0] != '無孔':
            if par.plate_hole_type[0] == '圓孔':
                self.ui.picture.setPixmap(QtGui.QPixmap("feeding_hole_A.png"))
            elif par.plate_hole_type[0] == '方孔':
                self.ui.picture.setPixmap(QtGui.QPixmap("feeding_hole_B.png"))
            elif par.plate_hole_type[0] == '漏斗型':
                self.ui.picture.setPixmap(QtGui.QPixmap("feeding_hole_C.png"))
        self.ui.setup.clicked.connect(self.setup)
        self.ui.esc.clicked.connect(self.esc)

    def setup(self):
        par.cutout_hole_machining_X = self.ui.X.text()
        par.cutout_hole_machining_Y = self.ui.Y.text()
        self.esc()

    def esc(self):
        self.hide()
        self.nw = padwindows()
        self.nw.show()




class remove_machining(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.ui = pad_feeding_hole_Form()
        self.ui.setupUi(self)
        self.setWindowTitle('下料口加工設定')
        self.ui.escape.clicked.connect(self.showpadwindows)
        self.ui.setup.clicked.connect(self.setup)
        self.ui.reset.clicked.connect(self.clean_data)

    def setup(self):
        X = self.ui.X.text()
        Y = self.ui.Y.text()
        par.feeding_hole_position = [X, Y]
        print("下料口位置:", par.feeding_hole_position)
        self.hide()
        self.nw = padwindows()
        self.nw.show()

    def clean_data(self):
        self.ui.X.clear()
        self.ui.Y.clear()

    def showpadwindows(self):
        data_entries = []
        self.hide()
        self.nw = padwindows()
        self.nw.show()

if __name__ == "__main__":
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)  # 自適應屏幕分辨率
    app = QtWidgets.QApplication([])
    window = main()
    window.show()
    sys.exit(app.exec_())
