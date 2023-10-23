from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget, QPushButton, QLineEdit, QLabel, \
    QTableWidget, QTableWidgetItem, QHeaderView, QSpinBox, QComboBox, QAbstractItemView
from PyQt5.QtCore import QObject, QTimer, Qt
from PyQt5.QtGui import QColor, QBrush
from DEMOGUI import Ui_Dialog
from PAD_main import Ui_Form as pad_main_Form
from PAD_MACHINING import Ui_Form as pad_machining_Form
from PAD_dimension import Ui_Form as pad_dimension_Form
from pad_feeding_hole import Ui_Form as pad_feeding_hole_Form
from cutout_hole_GUI import Ui_Form as cutout_hole_machining_form
from plate_main_first import Ui_Form as plate_main_first_form
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
import check as ch

test_stop = False


class main(QtWidgets.QWidget, Ui_Dialog):
    def __init__(self):
        super(main, self).__init__()
        self.ui = Ui_Dialog()
        self.ui.setupUi(self)
        self.ui.pushButton.clicked.connect(self.start)
        self.ui.pushButton_2.clicked.connect(self.showPadwindows)
        self.ui.comboBox_4.currentIndexChanged.connect(lambda: self.change_label_7())
        self.ui.comboBox_2.currentIndexChanged.connect(lambda: self.change_label_7())
        self.ui.comboBox_4.currentIndexChanged.connect(lambda: self.change_label_9())
        self.ui.comboBox_2.currentIndexChanged.connect(lambda: self.change_label_9())

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

        self.alpha, self.beta, self.zeta, self.epsilon = self.frame_calculate(self.i, self.specifications_travel_value,
                                                                              self.specifications_close_working_height_value,
                                                                              self.travel_type)
        if test_stop == False:
            self.create_txt(self.path, type, travel_type, self.specifications_travel_value,
                            self.specifications_close_working_height_value, self.alpha, self.beta, self.zeta,
                            self.epsilon)
            self.change_dir(self.i, self.p, self.alpha, self.beta, self.zeta, self.epsilon, self.machining,
                            self.welding)

    def showPadwindows(self):
        type = str(self.ui.comboBox_4.currentText())
        travel_type = str(self.ui.comboBox_2.currentText())
        processing = str(self.ui.comboBox.currentText())
        self.i, self.p, self.travel_type = self.choos(type, processing, travel_type)
        global i
        # i = self.i
        # print(i)
        self.hide()
        self.nw = padwindows(self.i)
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
        epsilon_1 = (-alpha / 2) + (
                close_working_height_value[all_type] - int(specifications_close_working_height_value))
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
            print('閉合工作高度符合要求: %s = %s' % (
                verification_specifications_close_working_height_value, specifications_close_working_height_value))
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
    def create_txt(self, path, travel_type, specifications_travel_value, specifications_close_working_height_value,
                   type, alpha, beta, zeta, epsilon):
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
        mbox.information(Form, '完成', '生成完成\nmachining_file_change_error:%s\nwelding_file_change_error:%s\n' % (
            machining_file_change_error, welding_file_change_error))
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
                        2500: {"S": ("標準:450"), "H": ("標準:400"), "P": ("標準:400")}, }

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
                            all_parameter_save.setdefault(parameter_name[x], parameter_value[x])
                            all_parameter_list = all_parameter_save.copy()
                            all_parameter_value[name] = all_parameter_list
                            apv = all_parameter_value
                        all_parameter_save.clear()

                        # 恢复原始的sys.stdout
                        sys.stdout = original_stdout
                        # 从捕获的输出中获取文本
                        output_text = captured_output.getvalue()

                        # 判断文本内容
                        if "error" in output_text:
                            machining_file_change_error.append(name)
                        else:
                            machining_file_change_pass.append(name)
                            # mprog.close_file(name)
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
        print('總用時%s' % (time.time() - start_time))  # 建立3D組立
        Ad.assembly(i, apv, self.path, alpha, beta, zeta)

        return machining_file_change_error, welding_file_change_error

class plate_first_windows(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.ui = plate_main_first_form()
        self.ui.setupUi(self)
        self.setWindowTitle('平板')
        self.ui.pad.clicked.connect(self.showPadwindows)
        self.ui.pushButton_2.clicked.connect(self.showcutoutwindows)

# 平板主頁面
class padwindows(QtWidgets.QWidget):
    def __init__(self, i):
        super().__init__()
        self.ui = pad_main_Form()
        self.ui.setupUi(self)
        self.setWindowTitle('平板')
        self.ui.remove_type.currentIndexChanged.connect(self.cutout_parameter_change)
        # 平板長寬尺寸
        self.plate_type(i)
        self.ui.pad_extrasize.currentIndexChanged.connect(lambda: self.plate_area_dimension(i))
        # 刪除垂直表頭
        self.ui.removetable.verticalHeader().setVisible(False)
        for number in range(0, 4):
            self.ui.t_solttable.setItem(number, 0, QTableWidgetItem(par.t_table_dimension_parameter[number]))
        # 刪除垂直及水平表頭
        self.ui.t_solttable.verticalHeader().setVisible(False)
        self.ui.t_solttable.horizontalHeader().setVisible(False)
        # 設定T_solt表格內容
        self.T_solt_table_normel_setup()
        for number in range(1, 10):
            newItem = QTableWidgetItem("-")
            self.ui.t_solttable.setItem(number, 1, newItem)
        self.ui.t_solt_type.currentIndexChanged.connect(self.T_solt_combobox_change)
        self.ui.plate_start.clicked.connect(lambda: self.start(i))
        self.ui.t_machining.clicked.connect(lambda: self.showpadmachiningwindows(i))

        self.ui.remove_machining.clicked.connect(self.showcutoutmachiningwindows)
        self.chack_plate_table()

    def plate_area_dimension(self, i):
        get_pad_select_name = self.ui.pad_select.currentText()
        if get_pad_select_name == "特殊平板":
            if get_pad_select_name == '標準':
                LR_value = str(par.plate_length[i])
                FB_value = str(par.plate_width[i])
            elif get_pad_select_name == '加大I型':
                LR_value = str(par.plate_length[i] + par.plate_lv1[i])
                FB_value = str(par.plate_width[i])
            elif get_pad_select_name == '加大II型':
                LR_value = str(par.plate_length[i] + par.plate_lv2[i])
                FB_value = str(par.plate_width[i])
            # 設定最終的 LR 和 FB 值
            self.ui.LR.setText(LR_value)
            self.ui.FB.setText(FB_value)

    def plate_type(self, i):
        for number in range(0, 10):
            self.ui.pad_select.setItemText(number, str(par.plate_type[number])+'('+str(par.plate_length[i])+'x'+str(par.plate_width[i])+")")



    def T_solt_table_normel_setup(self):
        # 設定T_solt表格內容
        # 第一行
        self.ui.t_solttable.setSpan(0, 0, 1, 1)
        newItem = QTableWidgetItem("尺寸代號")
        self.ui.t_solttable.setItem(0, 0, newItem)
        self.ui.t_solttable.setSpan(0, 0, 1, 1)
        newItem = QTableWidgetItem("尺寸")
        self.ui.t_solttable.setItem(0, 1, newItem)
        self.ui.t_solttable.setSpan(0, 0, 1, 1)
        newItem = QTableWidgetItem("公差")
        self.ui.t_solttable.setItem(0, 2, newItem)
        # 第一列
        self.ui.t_solttable.setSpan(1, 0, 2, 1)
        newItem = QTableWidgetItem("A")
        self.ui.t_solttable.setItem(1, 0, newItem)
        self.ui.t_solttable.setSpan(3, 0, 2, 1)
        newItem = QTableWidgetItem("B")
        self.ui.t_solttable.setItem(3, 0, newItem)
        self.ui.t_solttable.setSpan(5, 0, 2, 1)
        newItem = QTableWidgetItem("C")
        self.ui.t_solttable.setItem(5, 0, newItem)
        self.ui.t_solttable.setSpan(7, 0, 2, 1)
        newItem = QTableWidgetItem("D")
        self.ui.t_solttable.setItem(7, 0, newItem)
        # 第二列
        self.ui.t_solttable.setSpan(1, 1, 2, 1)
        self.ui.t_solttable.setSpan(3, 1, 2, 1)
        self.ui.t_solttable.setSpan(5, 1, 2, 1)
        self.ui.t_solttable.setSpan(7, 1, 2, 1)
        # 第三列
        newItem = QTableWidgetItem("+1")
        self.ui.t_solttable.setItem(1, 2, newItem)
        newItem = QTableWidgetItem("0")
        self.ui.t_solttable.setItem(2, 2, newItem)
        newItem = QTableWidgetItem("+1")
        self.ui.t_solttable.setItem(3, 2, newItem)
        newItem = QTableWidgetItem("-1")
        self.ui.t_solttable.setItem(4, 2, newItem)
        newItem = QTableWidgetItem("+0.25")
        self.ui.t_solttable.setItem(5, 2, newItem)
        newItem = QTableWidgetItem("-0.25")
        self.ui.t_solttable.setItem(6, 2, newItem)
        newItem = QTableWidgetItem("+2")
        self.ui.t_solttable.setItem(7, 2, newItem)
        newItem = QTableWidgetItem("0")
        self.ui.t_solttable.setItem(8, 2, newItem)


    def T_solt_combobox_change(self):
        t_solt_type = self.ui.t_solt_type.currentText()
        if t_solt_type == "T型槽代號:F(SN1-25~60標準)" or t_solt_type == "T型槽代號:G(SN180~250標準)":
            for number in range(1, 10):
                newItem = QTableWidgetItem("-")
                self.ui.t_solttable.setItem(number, 1, newItem)
        elif t_solt_type == "特殊T型槽":
            for number in range(1, 10):
                newItem = QTableWidgetItem("")
                self.ui.t_solttable.setItem(number, 1, newItem)

    def chack_plate_table(self):
        if par.plate_hole_type[0] != '':
            if par.plate_hole_type[0] == '圓孔':
                self.ui.remove_type.setCurrentIndex(1)
                self.ui.removetable.setItem(0, 0, QTableWidgetItem(par.cutout_parameter_circle[0]))
                self.ui.removetable.setItem(0, 1, QTableWidgetItem(par.cutout_part_dimension[0]))
            if par.plate_hole_type[0] == '方孔':
                self.ui.remove_type.setCurrentIndex(2)
                for i in range(0, 2):
                    self.ui.removetable.setItem(i, 0, QTableWidgetItem(par.cutout_parameter_square[i]))
                    self.ui.removetable.setItem(i, 1, QTableWidgetItem(par.cutout_part_dimension[i]))
            if par.plate_hole_type[0] == '漏斗型':
                self.ui.remove_type.setCurrentIndex(3)
                for i in range(0, 5):
                    self.ui.removetable.setItem(i, 0, QTableWidgetItem(par.cutout_parameter_funnel[i]))
                    self.ui.removetable.setItem(i, 1, QTableWidgetItem(par.cutout_part_dimension[i]))

    def showpaddimensionwindows(self):
        self.hide()
        self.nw = pad_dimension()
        self.nw.show()

    def showpadmachiningwindows(self, i):
        for value in range(1, 9, 2):
            par.t_all_dimension.append(self.ui.t_solttable.item(value, 1).text())
        print("par.t_all_dimension:", par.t_all_dimension)
        self.hide()
        self.nw = t_machining(i)
        self.nw.show()

    def showcutoutmachiningwindows(self):
        par.plate_hole_type = [self.ui.remove_type.currentText()]
        for i in range(0, 5):
            par.cutout_part_dimension[i] = (self.ui.removetable.item(i, 1).text())
        print(par.cutout_part_dimension)
        if par.plate_hole_type == '漏斗型':
            if int(par.cutout_part_dimension[0]) < int(par.cutout_part_dimension[1]) or int(
                par.cutout_part_dimension[2]) < int(par.cutout_part_dimension[3]):
                print('error')
        self.hide()
        self.nw = cutout_hole_machining()
        self.nw.show()

    def showremovemachiningwindows(self):
        self.hide()
        self.nw = remove_machining()
        self.nw.show()

    # 下料孔形狀選單連動功能
    def cutout_parameter_change(self):
        translate = QtCore.QCoreApplication.translate
        cutout_type = self.ui.remove_type.currentText()
        for i in range(0, 5):
            self.ui.removetable.setItem(i, 0, QTableWidgetItem(''))
            self.ui.removetable.setItem(i, 1, QTableWidgetItem(''))
        if cutout_type == '無孔':
            self.ui.removepicture.setPixmap(QtGui.QPixmap("cutout_shape_A.png"))
            self.ui.remove_machining.setDisabled(True)
        if cutout_type == '圓孔':
            self.ui.removepicture.setPixmap(QtGui.QPixmap("cutout_shape_B.png"))
            self.ui.remove_machining.setDisabled(False)
            self.ui.removetable.setItem(0, 0, QTableWidgetItem('HD'))
        if cutout_type == '方孔':
            self.ui.removepicture.setPixmap(QtGui.QPixmap("cutout_shape_C.png"))
            self.ui.remove_machining.setDisabled(False)
            self.ui.removetable.setItem(0, 0, QTableWidgetItem('HLR'))
            self.ui.removetable.setItem(1, 0, QTableWidgetItem('HFB'))
        if cutout_type == '漏斗型':
            self.ui.removepicture.setPixmap(QtGui.QPixmap("cutout_shape_D.png"))
            self.ui.remove_machining.setDisabled(False)
            self.ui.removetable.setItem(0, 0, QTableWidgetItem('HULR'))
            self.ui.removetable.setItem(1, 0, QTableWidgetItem('HDLR'))
            self.ui.removetable.setItem(2, 0, QTableWidgetItem('HUFB'))
            self.ui.removetable.setItem(3, 0, QTableWidgetItem('HDFB'))
            self.ui.removetable.setItem(4, 0, QTableWidgetItem('HH'))
        if cutout_type == '模墊型':
            self.ui.removepicture.setPixmap(QtGui.QPixmap("cutout_shape_E.png"))
            self.ui.remove_machining.setDisabled(True)

    def create_dir(self, type):  # 創建資料夾
        time_now = datetime.datetime.now()
        dir_name = '{}_{}_{}_{}_{}'.format(type, time_now.day, time_now.hour, time_now.minute, time_now.second)
        # desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
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

    def start(self, i):
        pad_type = str(self.ui.pad_type.currentText())
        # 對平板進行變數變換
        try:
            self.create_dir('plate')
        except:
            print('資料夾生成路徑連接失敗')
        mprog.set_CATIA_workbench_env()
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

        # T溝程式
        self.T_solt()


        #下料孔生成
        if par.plate_hole_type[0] != '無孔':
            par.plate_hole_type = [self.ui.remove_type.currentText()]
            print(i)
            for x in range(0, 5):
                par.cutout_part_dimension[x] = (self.ui.removetable.item(x, 1).text())
            print(par.cutout_part_dimension)
            print(i)
            ch.edge_test(i)
            # 開下料孔檔案+變數變換
            if par.plate_hole_type[0] == '圓孔':
                mprog.import_part(fp.system_root + fp.DEMO_part, 'cutout_hole_circle')
                mprog.param_change('cutout_hole_circle', par.cutout_parameter_circle[0], int(par.cutout_part_dimension[0]))
                mprog.param_change('cutout_hole_circle', 'C', par.plate_all_parameter['C'])
                mprog.param_change('cutout_hole_circle', 'X', par.cutout_hole_machining_X - par.plate_all_parameter['A']/2 - par.lv[0]/2)
                mprog.param_change('cutout_hole_circle', 'Y', par.plate_all_parameter['B']/2 + par.cutout_hole_machining_Y)
                par.plate_hole_type[0] = 'cutout_hole_circle'
            elif par.plate_hole_type[0] == '方孔':
                mprog.import_part(fp.system_root + fp.DEMO_part, 'cutout_hole_square')
                for n in range(0, 2):
                    mprog.param_change('cutout_hole_square', par.cutout_parameter_square[n],
                                       int(par.cutout_part_dimension[n]))
                mprog.param_change('cutout_hole_square', 'C', par.plate_all_parameter['C'])
                mprog.param_change('cutout_hole_square', 'X', par.cutout_hole_machining_X - par.plate_all_parameter['A']/2 - par.lv[0]/2)
                mprog.param_change('cutout_hole_square', 'Y', par.plate_all_parameter['B']/2 + par.cutout_hole_machining_Y)
                par.plate_hole_type[0] = 'cutout_hole_square'
            elif par.plate_hole_type[0] == '漏斗型':
                mprog.import_part(fp.system_root + fp.DEMO_part, 'cutout_funnel')
                for n in range(0, 5):
                    mprog.param_change('cutout_funnel', par.cutout_parameter_funnel[n], int(par.cutout_part_dimension[n]))
                mprog.param_change('cutout_funnel', 'C', par.plate_all_parameter['C'])
                mprog.param_change('cutout_funnel', 'X',
                                   par.cutout_hole_machining_X - par.plate_all_parameter['A'] / 2 - par.lv[0] / 2)
                mprog.param_change('cutout_funnel', 'Y',
                                   par.plate_all_parameter['B'] / 2 + par.cutout_hole_machining_Y)
                par.plate_hole_type[0] = 'cutout_funnel'
            elif par.plate_hole_type[0] == '模墊型':
                mprog.import_part(fp.system_root + fp.DEMO_part, 'cutout_molded_cushion')
                mprog.param_change('cutout_molded_cushion', 'C', par.plate_all_parameter['C'])
                if i < 5:
                    mprog.param_change('cutout_molded_cushion', 'A', par.cutout_molded_cushion_A[0])
                    mprog.param_change('cutout_molded_cushion', 'B', par.cutout_molded_cushion_B[0])
                    mprog.param_change('cutout_molded_cushion', 'D', par.cutout_molded_cushion_L[0])
                else:
                    mprog.param_change('cutout_molded_cushion', 'A', par.cutout_molded_cushion_A[1])
                    mprog.param_change('cutout_molded_cushion', 'B', par.cutout_molded_cushion_B[1])
                    mprog.param_change('cutout_molded_cushion', 'D', par.cutout_molded_cushion_L[1])
                mprog.param_change('cutout_molded_cushion', 'i', par.cutout_molded_cushion_i[i])
                mprog.param_change('cutout_molded_cushion', 'j', par.cutout_molded_cushion_j[i])
                mprog.param_change('cutout_molded_cushion', 'X', 65 * (par.cutout_molded_cushion_i[i] - 1) / 2 - par.plate_all_parameter['A'] / 2 - par.lv[0] / 2)
                mprog.param_change('cutout_molded_cushion', 'Y', par.plate_all_parameter['B'] / 2 - 60 * (par.cutout_molded_cushion_j[i] - 1) / 2)
                par.plate_hole_type[0] = 'cutout_molded_cushion'
            mprog.Update()
            mprog.save_file_stp(self.path, par.plate_hole_type[0]) #存檔(stp)
            mprog.save_stpfile_part(self.path, par.plate_hole_type[0]) #存檔(part)
            tT.copybody()
            tT.switch_to_window_by_name('plate.CATPart')
            tT.pastebody(0, par.plate_hole_type[0]) #count??
            tT.removebody(0, par.plate_hole_type[0]) #count??
            mprog.Update()
            tT.switch_to_window_by_name(par.plate_hole_type[0] + ".CATPart")
            mprog.close_window()

    def T_solt(self):
        # 對T型槽進行變數變換
        mprog.import_part(fp.system_root + fp.DEMO_part, 'T')
        pdp.padchange(i)
        try:
            for t in par.t_all_dimension:
                for t_name in range(len(par.t_all_dimension_name) + 1):
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
                        mprog.param_change('T', 'Depth', (par.plate_all_parameter['A'] + par.lv[0]))
                else:
                    mprog.param_change('T', 'Depth', par.total_unpierce[turn])
                    if par.total_t_direction[turn] == '前後':
                        mprog.param_change('T', 'mirror', par.plate_all_parameter['B'])
                        print(par.plate_all_parameter['B'] / 2)
                    elif par.total_t_direction[turn] == '左右':
                        mprog.param_change('T', 'mirror', (par.plate_all_parameter['A'] + par.lv[0]))
                        print((par.plate_all_parameter['A'] + par.lv[0]) / 2)
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
                    tT.create_t_solt(-par.total_t_dimension[turn], turn, par.total_pierce[turn],
                                     par.total_clearance_hole[turn], par.total_t_direction[turn], par.lv[0])
                elif par.total_t_direction[turn] == '左右':
                    tT.create_t_solt(par.total_t_dimension[turn], turn, par.total_pierce[turn],
                                     par.total_clearance_hole[turn], par.total_t_direction[turn], par.lv[0])
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
        self.nw = padwindows(i)
        self.nw.show()

    def reset(self):
        self.ui.A.clear()
        self.ui.B.clear()
        self.ui.C.clear()
        self.ui.D.clear()

    def showpadwindows(self):
        self.hide()
        self.nw = padwindows(i)
        self.nw.show()


class t_machining(QWidget):
    def __init__(self, i):
        super().__init__()
        self.ui = pad_machining_Form()
        self.ui.setupUi(self)
        self.setWindowTitle('平板加工設定')
        # 橫向T型槽
        self.ui.t_slot_table_h.verticalHeader().setVisible(False)
        self.ui.t_slot_table_h.horizontalHeader().setVisible(False)
        self.ui.t_slot_table_h.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.ui.t_slot_table_h.setEditTriggers(QAbstractItemView.AllEditTriggers)
        self.t_slot_table_h_setup()
        self.ui.t_slot_h_number.textChanged.connect(self.check_slot_number)
        self.table_h_combo_boxes = {}
        # 縱向T型槽
        self.ui.t_slot_table_v.verticalHeader().setVisible(False)
        self.ui.t_slot_table_v.horizontalHeader().setVisible(False)
        self.ui.t_slot_table_v.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.ui.t_slot_table_v.setEditTriggers(QAbstractItemView.AllEditTriggers)
        self.t_slot_table_v_setup()
        self.ui.t_slot_v_number.textChanged.connect(self.check_slot_v_number)
        self.table_v_combo_boxes = {}
        # 確定按鈕
        self.ui.setup.clicked.connect(self.setup)
        # 重製按鈕
        self.ui.reset.clicked.connect(self.reset)
        # 重新排列
        self.ui.rearrange_the_order.clicked.connect(self.rearrange_the_order)


    def change_table_h_clear_table(self):
        while self.ui.t_slot_table_h.rowCount() > 1:
            self.ui.t_slot_table_h.removeRow(1)

    def check_slot_number(self, current_value):
        try:
            if current_value == '':
                current_value = 0
            current_value = int(current_value)  # 尝试将文本转换为整数
            self.add_rows_to_table(current_value)  # 插入新行
        except ValueError:
            print("Invalid input. Please enter a valid number")

    def add_rows_to_table(self, num_rows):
        self.change_table_h_clear_table()  # 清除表格内容
        col_count = self.ui.t_slot_table_h.columnCount()
        for row in range(num_rows):
            row_position = self.ui.t_slot_table_h.rowCount()
            self.ui.t_slot_table_h.insertRow(row_position)
            for col in range(col_count):
                item = QTableWidgetItem("H" + str(row_position) if (row_position == 0 or col == 0) else "")
                if col == 0 or row_position == 0:
                    # 如果是第一行或第一列，设置单元格不可编辑
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                else:
                    # 其他单元格默认不可编辑，会在后续根据条件进行修改
                    item.setFlags(item.flags() | Qt.ItemIsEditable)
                self.ui.t_slot_table_h.setItem(row_position, col, item)
            # 创建 QComboBox 并添加到新行的第三列
            combo_box = QComboBox()
            combo_box.addItem("")
            combo_box.addItem("貫穿")
            combo_box.addItem("分段")
            combo_box.currentIndexChanged.connect(lambda index, row=row_position: self.combo_box_changed(row, index))
            self.ui.t_slot_table_h.setCellWidget(row_position, 2, combo_box)
            self.table_h_combo_boxes[row_position] = combo_box  # 存储QComboBox
            # 初始化选项状态为不可编辑
            for col in range(3, 5):
                item = QTableWidgetItem("")
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                color = QColor(192, 192, 192)
                brush = QBrush(color)
                item.setBackground(brush)
                self.ui.t_slot_table_h.setItem(row_position, col, item)

    def combo_box_changed(self, row, index):
        combo_box = self.table_h_combo_boxes.get(row)
        if combo_box:
            if index == 2:  # 如果选择了"分段"
                self.set_editable_cells(row, is_editable=True)
            else:
                self.set_editable_cells(row, is_editable=False)

    def set_editable_cells(self, row, is_editable=True):
        for col in range(3, 5):
            item = self.ui.t_slot_table_h.item(row, col)
            if item:
                if is_editable and row != 0:  # 如果选择了"分段"且不是第一行
                    item.setFlags(item.flags() | Qt.ItemIsEditable)
                    item.setBackground(QBrush(QColor(255, 255, 255)))
                else:
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    item.setBackground(QBrush(QColor(192, 192, 192)))

    def t_slot_table_h_setup(self):
        # 第一行
        self.ui.t_slot_table_h.setSpan(0, 0, 1, 1)
        newItem = QTableWidgetItem("編號")
        self.ui.t_slot_table_h.setItem(0, 0, newItem)
        self.ui.t_slot_table_h.setSpan(0, 0, 1, 1)
        newItem = QTableWidgetItem("位置(Y)")
        self.ui.t_slot_table_h.setItem(0, 1, newItem)
        self.ui.t_slot_table_h.setSpan(0, 0, 1, 1)
        newItem = QTableWidgetItem("形式")
        self.ui.t_slot_table_h.setItem(0, 2, newItem)
        self.ui.t_slot_table_h.setSpan(0, 0, 1, 1)
        newItem = QTableWidgetItem("LL")
        self.ui.t_slot_table_h.setItem(0, 3, newItem)
        self.ui.t_slot_table_h.setSpan(0, 0, 1, 1)
        newItem = QTableWidgetItem("LR")
        self.ui.t_slot_table_h.setItem(0, 4, newItem)
        self.ui.t_slot_table_h.setSpan(0, 0, 1, 1)
        newItem = QTableWidgetItem("SL")
        self.ui.t_slot_table_h.setItem(0, 5, newItem)
        self.ui.t_slot_table_h.setSpan(0, 0, 1, 1)
        newItem = QTableWidgetItem("SR")
        self.ui.t_slot_table_h.setItem(0, 6, newItem)

    def combo_box_changed_v(self, row, index):
        combo_box = self.table_v_combo_boxes.get(row)
        if combo_box:
            if index == 2:  # 如果选择了"分段"
                self.set_editable_v_cells(row, is_editable=True)
            else:
                self.set_editable_v_cells(row, is_editable=False)

    def change_table_v_clear_table(self):
        while self.ui.t_slot_table_v.rowCount() > 1:
            self.ui.t_slot_table_v.removeRow(1)

    def check_slot_v_number(self, current_value):
        try:
            if current_value == '':
                current_value = 0
            current_value = int(current_value)  # 尝试将文本转换为整数
            self.add_rows_to_table_v(current_value)  # 插入新行
        except ValueError:
            print("Invalid input. Please enter a valid number")

    def add_rows_to_table_v(self, num_rows):
        self.change_table_v_clear_table()  # 清除表格内容
        col_count = self.ui.t_slot_table_v.columnCount()
        for row in range(num_rows):
            row_position = self.ui.t_slot_table_v.rowCount()
            self.ui.t_slot_table_v.insertRow(row_position)
            for col in range(col_count):
                item = QTableWidgetItem("V" + str(row_position) if (row_position == 0 or col == 0) else "")
                if col == 0 or row_position == 0:
                    # 如果是第一行或第一列，设置单元格不可编辑
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                else:
                    # 其他单元格默认不可编辑，会在后续根据条件进行修改
                    item.setFlags(item.flags() | Qt.ItemIsEditable)
                self.ui.t_slot_table_v.setItem(row_position, col, item)

            # 创建 QComboBox 并添加到新行的第三列
            combo_box = QComboBox()
            combo_box.addItem("")
            combo_box.addItem("貫穿")
            combo_box.addItem("分段")
            combo_box.currentIndexChanged.connect(lambda index, row=row_position: self.combo_box_changed_v(row, index))
            self.ui.t_slot_table_v.setCellWidget(row_position, 2, combo_box)
            self.table_v_combo_boxes[row_position] = combo_box  # 存储QComboBox
            # 初始化选项状态为不可编辑
            for col in range(3, 5):
                item = QTableWidgetItem("")
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                color = QColor(192, 192, 192)
                brush = QBrush(color)
                item.setBackground(brush)
                self.ui.t_slot_table_v.setItem(row_position, col, item)

    def set_editable_v_cells(self, row, is_editable=True):
        for col in range(3, 5):
            item = self.ui.t_slot_table_v.item(row, col)
            if item:
                if is_editable and row != 0:  # 如果选择了"分段"且不是第一行
                    item.setFlags(item.flags() | Qt.ItemIsEditable)
                    item.setBackground(QBrush(QColor(255, 255, 255)))
                else:
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    item.setBackground(QBrush(QColor(192, 192, 192)))

    def t_slot_table_v_setup(self):
        # 第一行
        self.ui.t_slot_table_v.setSpan(0, 0, 1, 1)
        newItem = QTableWidgetItem("編號")
        self.ui.t_slot_table_v.setItem(0, 0, newItem)
        self.ui.t_slot_table_v.setSpan(0, 0, 1, 1)
        newItem = QTableWidgetItem("位置(Y)")
        self.ui.t_slot_table_v.setItem(0, 1, newItem)
        self.ui.t_slot_table_v.setSpan(0, 0, 1, 1)
        newItem = QTableWidgetItem("形式")
        self.ui.t_slot_table_v.setItem(0, 2, newItem)
        self.ui.t_slot_table_v.setSpan(0, 0, 1, 1)
        newItem = QTableWidgetItem("LF")
        self.ui.t_slot_table_v.setItem(0, 3, newItem)
        self.ui.t_slot_table_v.setSpan(0, 0, 1, 1)
        newItem = QTableWidgetItem("LB")
        self.ui.t_slot_table_v.setItem(0, 4, newItem)
        self.ui.t_slot_table_v.setSpan(0, 0, 1, 1)
        newItem = QTableWidgetItem("SF")
        self.ui.t_slot_table_v.setItem(0, 5, newItem)
        self.ui.t_slot_table_v.setSpan(0, 0, 1, 1)
        newItem = QTableWidgetItem("SB")
        self.ui.t_slot_table_v.setItem(0, 6, newItem)

    def showpadwindows(self):
        data_entries = []
        self.hide()
        self.nw = padwindows(i)
        self.nw.show()

    def setup(self, i):
        self.hide()
        self.nw = padwindows(i)
        self.nw.show()

    def rearrange_the_order(self):
        par.total_position_y.clear()
        par.total_t_slot_h_type.clear()
        par.total_LL.clear()
        par.total_LR.clear()
        par.total_SL.clear()
        par.total_SR.clear()
        par.total_position_x.clear()
        par.total_t_slot_v_type.clear()
        par.total_LF.clear()
        par.total_LB.clear()
        par.total_SF.clear()
        par.total_SB.clear()

        if self.ui.t_slot_table_h.rowCount() != 0:
            for row in range(1, self.ui.t_slot_table_h.rowCount()):
                position_y = self.ui.t_slot_table_h.item(row, 1).text()
                t_slot_type = self.ui.t_slot_table_h.cellWidget(row, 2).currentText()
                LL = self.ui.t_slot_table_h.item(row, 3).text()
                LR = self.ui.t_slot_table_h.item(row, 4).text()
                SL = self.ui.t_slot_table_h.item(row, 5).text()
                SR = self.ui.t_slot_table_h.item(row, 6).text()
                par.total_position_y.append(position_y)
                par.total_t_slot_h_type.append(t_slot_type)
                par.total_LL.append(LL)
                par.total_LR.append(LR)
                par.total_SL.append(SL)
                par.total_SR.append(SR)
            par.total_position_y = [int(x) for x in par.total_position_y]
            position_y_change_position = sorted(enumerate(par.total_position_y), key=lambda x: x[1], reverse=True)
            print('position_y_change_position:', position_y_change_position)
            rearrange = [position[0] for position in position_y_change_position]
            print('rearrange:', rearrange)
            par.total_position_y = [par.total_position_y[order_position] for order_position in rearrange]
            par.total_LL = [par.total_LL[order_position] for order_position in rearrange]
            par.total_LR = [par.total_LR[order_position] for order_position in rearrange]
            par.total_SL = [par.total_SL[order_position] for order_position in rearrange]
            par.total_SR = [par.total_SR[order_position] for order_position in rearrange]
            par.total_t_slot_h_type = [par.total_t_slot_h_type[order_position] for order_position in rearrange]

            print('total_position_y:', par.total_position_y)
            print('total_t_slot_h_type:', par.total_t_slot_h_type)
            print('total_LL:', par.total_LL)
            print('total_LR:', par.total_LR)
            print('total_SL:', par.total_SL)
            print('total_SR:', par.total_SR)
            for position, item in enumerate(par.total_position_y):
                # 將整數轉換為字串，然後設定為表格的項目文本
                item_text = str(item)
                table_item = QtWidgets.QTableWidgetItem(item_text)
                self.ui.t_slot_table_h.setItem(position + 1, 1, table_item)
            for position, item in enumerate(par.total_LL):
                item = QtWidgets.QTableWidgetItem(item)
                self.ui.t_slot_table_h.setItem(position + 1, 3, item)
            for position, item in enumerate(par.total_LR):
                item = QtWidgets.QTableWidgetItem(item)
                self.ui.t_slot_table_h.setItem(position + 1, 4, item)
            for position, item in enumerate(par.total_SL):
                item = QtWidgets.QTableWidgetItem(item)
                self.ui.t_slot_table_h.setItem(position + 1, 5, item)
            for position, item in enumerate(par.total_SR):
                item = QtWidgets.QTableWidgetItem(item)
                self.ui.t_slot_table_h.setItem(position + 1, 6, item)
            for position, item in enumerate(par.total_t_slot_h_type):
                combo_box = self.ui.t_slot_table_h.cellWidget(position + 1, 2)  # 從表格中獲取 ComboBox
                combo_box.setCurrentText(item)
                self.combo_box_changed(position + 1, combo_box.currentIndex())

        if self.ui.t_slot_table_v.rowCount() != 0:
            for row in range(1, self.ui.t_slot_table_v.rowCount()):
                position_x = self.ui.t_slot_table_v.item(row, 1).text()
                t_slot_type = self.ui.t_slot_table_v.cellWidget(row, 2).currentText()
                LF = self.ui.t_slot_table_v.item(row, 3).text()
                LB = self.ui.t_slot_table_v.item(row, 4).text()
                SF = self.ui.t_slot_table_v.item(row, 5).text()
                SB = self.ui.t_slot_table_v.item(row, 6).text()
                par.total_position_x.append(position_x)
                par.total_t_slot_v_type.append(t_slot_type)
                par.total_LF.append(LF)
                par.total_LB.append(LB)
                par.total_SF.append(SF)
                par.total_SB.append(SB)
            par.total_position_x = [int(x) for x in par.total_position_x]
            position_x_change_position = sorted(enumerate(par.total_position_x), key=lambda x: x[1], reverse=True)
            print('position_x_change_position:', position_x_change_position)
            rearrange = [position[0] for position in position_x_change_position]
            print('rearrange:', rearrange)
            par.total_position_x = [par.total_position_x[order_position] for order_position in rearrange]
            par.total_LF = [par.total_LF[order_position] for order_position in rearrange]
            par.total_LB = [par.total_LB[order_position] for order_position in rearrange]
            par.total_SF = [par.total_SF[order_position] for order_position in rearrange]
            par.total_SB = [par.total_SB[order_position] for order_position in rearrange]
            par.total_t_slot_v_type = [par.total_t_slot_v_type[order_position] for order_position in rearrange]
            print('total_position_x:', par.total_position_x)
            print('total_t_slot_v_type:', par.total_t_slot_v_type)
            print('total_LF:', par.total_LF)
            print('total_LB:', par.total_LB)
            print('total_SF:', par.total_SF)
            print('total_SB:', par.total_SB)
            for position, item in enumerate(par.total_position_x):
                # 將整數轉換為字串，然後設定為表格的項目文本
                item_text = str(item)
                table_item = QtWidgets.QTableWidgetItem(item_text)
                self.ui.t_slot_table_v.setItem(position + 1, 1, table_item)
            for position, item in enumerate(par.total_LF):
                item = QtWidgets.QTableWidgetItem(item)
                self.ui.t_slot_table_v.setItem(position + 1, 3, item)
            for position, item in enumerate(par.total_LB):
                item = QtWidgets.QTableWidgetItem(item)
                self.ui.t_slot_table_v.setItem(position + 1, 4, item)
            for position, item in enumerate(par.total_SF):
                item = QtWidgets.QTableWidgetItem(item)
                self.ui.t_slot_table_v.setItem(position + 1, 5, item)
            for position, item in enumerate(par.total_SB):
                item = QtWidgets.QTableWidgetItem(item)
                self.ui.t_slot_table_v.setItem(position + 1, 6, item)
            for position, item in enumerate(par.total_t_slot_v_type):
                combo_box = self.ui.t_slot_table_v.cellWidget(position + 1, 2)
                combo_box.setCurrentText(item)
                self.combo_box_changed_v(position + 1, combo_box.currentIndex())

    def reset(self):
        self.ui.t_slot_h_number.clear()
        self.ui.t_slot_v_number.clear()
        par.total_position_y.clear()
        par.total_t_slot_h_type.clear()
        par.total_LL.clear()
        par.total_LR.clear()
        par.total_SL.clear()
        par.total_SR.clear()
        par.total_position_x.clear()
        par.total_t_slot_v_type.clear()
        par.total_LF.clear()
        par.total_LB.clear()
        par.total_SF.clear()
        par.total_SB.clear()


# 下料孔設定介面
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

    def setup(self, i):
        par.cutout_hole_machining_X = self.ui.X.text()
        par.cutout_hole_machining_Y = self.ui.Y.text()
        if par.cutout_hole_machining_X == '':
            par.cutout_hole_machining_X = 0
        if par.cutout_hole_machining_Y == '':
            par.cutout_hole_machining_Y = 0
        par.cutout_hole_machining_X = int(par.cutout_hole_machining_X)
        par.cutout_hole_machining_Y = int(par.cutout_hole_machining_Y)

        excel_cutout_limit = epc.ExcelOp('平板', '下料孔界線')
        excel_cutout_limit.get_cutout_limit(i)
        print(par.cutout_hole_machining_X, par.cutout_hole_machining_Y, par.cutout_all_limit['X'],
              par.cutout_all_limit['Y'])
        if par.plate_hole_type[0] == '圓孔':
            if abs(par.cutout_hole_machining_X) + int(par.cutout_part_dimension[0]) / 2 + 10 >= par.cutout_all_limit[
                'X'] / 2 or abs(par.cutout_hole_machining_Y) + int(par.cutout_part_dimension[0]) / 2 + 10 >= \
                    par.cutout_all_limit['Y'] / 2:
                print('error')
            else:
                self.esc()
        elif par.plate_hole_type[0] == '方孔':
            if abs(par.cutout_hole_machining_X) + int(par.cutout_part_dimension[0]) / 2 + 10 >= par.cutout_all_limit[
                'X'] / 2 or abs(par.cutout_hole_machining_Y) + int(par.cutout_part_dimension[0]) / 2 + 10 >= \
                    par.cutout_all_limit['Y'] / 2:
                print('error')
            else:
                self.esc()
        elif par.plate_hole_type[0] == '漏斗型':
            if abs(par.cutout_hole_machining_X) + int(par.cutout_part_dimension[0]) / 2 + 10 >= par.cutout_all_limit[
                'X'] / 2 or abs(par.cutout_hole_machining_Y) + int(par.cutout_part_dimension[0]) / 2 + 10 >= \
                    par.cutout_all_limit['Y'] / 2:
                print('error')
            else:
                self.esc()
        else:
            self.esc()

    def esc(self):
        self.hide()
        self.nw = padwindows(i)
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
        self.nw = padwindows(i)
        self.nw.show()

    def clean_data(self):
        self.ui.X.clear()
        self.ui.Y.clear()

    def showpadwindows(self):
        data_entries = []
        self.hide()
        self.nw = padwindows(i)
        self.nw.show()


if __name__ == "__main__":
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)  # 自適應屏幕分辨率
    app = QtWidgets.QApplication([])
    window = main()
    window.show()
    sys.exit(app.exec_())
