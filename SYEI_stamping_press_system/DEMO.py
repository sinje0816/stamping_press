from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QObject
from DEMOGUI import Ui_Dialog
from io import StringIO
import main_program as mprog
import file_path as fp
import parameter as par
import machining_part_TYPE_change as mptc
import welding_part_TYPE_change as wptc
import excel_parameter_change as epc
import sys
import datetime
import os
import time

test_stop = False
class main(QtWidgets.QWidget, Ui_Dialog):
    def __init__(self):
        super(main, self).__init__()
        self.ui = Ui_Dialog()
        self.ui.setupUi(self)
        self.ui.pushButton.clicked.connect(self.start)
        self.ui.comboBox_4.currentIndexChanged.connect(lambda :self.change_label_7())
        self.ui.comboBox_2.currentIndexChanged.connect(lambda :self.change_label_7())
        self.ui.comboBox_4.currentIndexChanged.connect(lambda :self.change_label_9())
        self.ui.comboBox_2.currentIndexChanged.connect(lambda :self.change_label_9())

    def start(self):
        global test_stop
        type = str(self.ui.comboBox_4.currentText())
        travel_type = str(self.ui.comboBox_2.currentText())
        travel = str(self.ui.label_7.text())
        specifications_travel_value = str(self.ui.lineEdit_5.text())
        specifications_close_working_height_value = str(self.ui.lineEdit_2.text())
        close_working_height = str(self.ui.label_9.text())
        delta = str(self.ui.lineEdit_4.text())
        processing = str(self.ui.comboBox.currentText())
        print(type, travel_type, travel, specifications_travel_value, specifications_close_working_height_value,
              close_working_height, delta)
        self.create_dir(type)
        if specifications_travel_value == "":
            self.specifications_travel_value = 0
        else:
            self.specifications_travel_value = int(specifications_travel_value)
        if specifications_close_working_height_value == "":
            self.specifications_close_working_height_value = 0
        else:
            self.specifications_close_working_height_value = int(specifications_close_working_height_value)
        if delta == "":
            self.delta = 0
        else:
            self.delta = int(delta)
        self.i, self.p, self.travel_type = self.choos(type, processing, travel_type)
        self.alpha, self.beta, self.zeta, self.epsilon = self.frame_calculate(self.i, self.specifications_travel_value,
                                                                                  self.specifications_close_working_height_value,
                                                                                  self.travel_type)
        if test_stop == False:
            self.create_txt(self.path, type, travel_type, self.specifications_travel_value, self.specifications_close_working_height_value, self.alpha, self.beta, self.delta, self.zeta, self.epsilon)
            self.change_dir(self.i, self.p, self.alpha, self.beta, self.delta,self.zeta, self.epsilon , self.machining, self.welding)

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
        Form.setWindowTitle('oxxo.studio')
        Form.resize(400, 300)
        mbox = QtWidgets.QMessageBox(Form)

        # 讀取標準資料
        excel = epc.ExcelOp('標準資料')
        type_name, travel_value, close_working_height_value, specifications_travel_min_value, specifications_travel_max_value, specifications_close_working_height_min_value, specifications_close_working_height_max_value = excel.get_standard_parts(
            i * 3)

        # 噸數&行程類型
        all_type = i * 3 + travel_type - 1
        # 行程
        alpha = int(specifications_travel_value) - travel_value[all_type]
        # 牙球伸長量
        epsilon = (-alpha / 2) + (close_working_height_value[all_type] - int(specifications_close_working_height_value))
        if epsilon <= 0:
            epsilon = 0
        else:
            pass
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
    def create_txt(self, path, travel_type, specifications_travel_value, specifications_close_working_height_value, type, alpha, beta, delta, zeta, epsilon):
        file_txt = path
        txt_name = "生成參數.txt"
        with open(file_txt + "\\" + txt_name, "w") as f:
            f.write("噸數=%s\n" % type)
            f.write("型式:%s\n" % travel_type)
            f.write("本次行程=%s\n" % specifications_travel_value)
            f.write("本次閉合工作高度=%s\n" % specifications_close_working_height_value)
            f.write("行程=%s\n" % alpha)
            f.write("閉合工作高度=%s\n" % beta)
            f.write("平板前後=%s\n" % delta)
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
        self.ui.lineEdit_4.clear()

    def label_7_change_data(self):
        label_7_data = {250: {"S": ("標準:80"), "H": ("標準:50"), "P": ("標準:35")},
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

    def change_dir(self, i, p, alpha, beta, delta, zeta, epsilon, machining, welding):
        start_time = time.time()
        all_part_name = {}
        all_part_value = {}
        # 開啟CATIA
        env = mprog.set_CATIA_workbench_env()
        machining_file_change_error = []
        machining_file_change_pass = []
        welding_file_change_error = []
        welding_file_change_pass = []
        # 開啟零件檔更改變數後儲存並關閉
        for name in epc.ExcelOp('沖床機架零件清單').get_col_cell(1):
            print(name)
            file_list_name, file_list_value = epc.ExcelOp('沖床機架零件清單').get_sheet_par('沖床機架零件清單', i)
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
                            mprog.param_change(name, "delta", delta)
                            mprog.param_change(name, 'zeta', zeta)
                        except:
                            pass
                        # 加工圖零件
                        parameter_name, parameter_value = mptc.change_machining_parameter(name, i, 0)
                        all_part_name[name] = parameter_name
                        all_part_value[name] = parameter_value

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
                            mprog.param_change(name, "delta", delta)
                            mprog.param_change(name, 'zeta', zeta)
                        except:
                            pass

                        # 加工圖零件
                        parameter_name, parameter_value = mptc.change_machining_parameter(name, i, 1)
                        all_part_name[name] = parameter_name
                        all_part_value[name] = parameter_value

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
        print('總用時%s' % (time.time() - start_time))

        return machining_file_change_error, welding_file_change_error


if __name__ == "__main__":
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)  # 自適應屏幕分辨率
    app = QtWidgets.QApplication([])
    window = main()
    window.show()
    sys.exit(app.exec_())
