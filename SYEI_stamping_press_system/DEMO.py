from PyQt5 import QtCore, QtGui, QtWidgets
from DEMOGUI import Ui_Dialog
from io import StringIO
import main_program as mprog
import file_path as fp
import parameter as par
import machining_part_TYPE_change as mptc
import welding_part_TYPE_change1 as wptc
import excel_parameter_change as epc
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
        travel_type = str(self.ui.comboBox_2.currentText())
        travel = str(self.ui.label_7.text())
        specifications_travel_value = str(self.ui.lineEdit_5.text())
        specifications_close_working_height_value = str(self.ui.lineEdit_2.text())
        close_working_height = str(self.ui.label_9.text())
        delta = str(self.ui.lineEdit_4.text())
        processing = str(self.ui.comboBox.currentText())
        print(type, travel_type, travel, specifications_travel_value, specifications_close_working_height_value, close_working_height, delta)
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
        self.i, self.p = self.choos(type, processing)
        self.create_txt(self.path, type, travel_type, self.aplha, self.beta, self.delta)
        self.change_dir(self.i, self.p, self.aplha, self.beta, self.delta, self.machining, self.welding)

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

        return i, p

    def frame_calculate(self, i, specifications_travel_value, specifications_close_working_height_value, zeta, delta):
        Form = QtWidgets.QWidget()
        Form.setWindowTitle('warning.studio')
        Form.resize(400, 300)
        mbox = QtWidgets.QMessageBox(Form)

        excel = epc.ExcelOp('標準資料')
        type_name, travel_value, close_working_height_value, specifications_travel_min_value, specifications_travel_max_value, specifications_close_working_height_min_value, specifications_close_working_height_max_value = excel.get_standard_parts(
            i)

        # 行程
        alpha = specifications_travel_value - travel_value
        # 牙球伸長量
        epsilon = (alpha / 2) + (close_working_height_value - specifications_close_working_height_value)
        if epsilon <= 0:
            epsilon = 0
        else:
            pass
        # 喉部拉高量
        zeta = (alpha / 2) + epsilon + specifications_close_working_height_value - close_working_height_value
        # 閉合工作高度驗證公式
        verification_specifications_close_working_height_value = close_working_height_value - (
                alpha / 2) - epsilon + specifications_close_working_height_value - close_working_height_value
        try:
            # 驗證行程
            if specifications_travel_value < specifications_travel_min_value or specifications_travel_value > specifications_travel_max_value:
                mbox.warning(Form, '警告', '行程超出範圍')
                self.ui.lineEdit.clear()
            # 驗證閉合工作高度
            if specifications_close_working_height_value >= specifications_close_working_height_min_value and specifications_close_working_height_value <= specifications_close_working_height_max_value:
                pass
            else:
                mbox.warning(Form, '警告', '閉合工作高度超出範圍')
                self.ui.lineEdit_2.clear()
            # 驗證閉合工作高度是否符合客戶需求
            if verification_specifications_close_working_height_value == specifications_close_working_height_value:
                pass
            else:
                mbox.warning(Form, '警告', '閉合工作高度不符合客戶需求')
                self.ui.lineEdit_2.clear()
            # 驗證牙球伸長量
            if epsilon < 0:
                mbox.warning(Form, '警告', '牙球伸長量小於0')
                self.ui.lineEdit_3.clear()
            else:
                pass
            # 驗證喉部拉高量
            if epsilon >= (alpha/2):
                pass
            else:
                mbox.warning(Form, '警告', '衝頭縮進導軌內部')
                self.ui.lineEdit_4.clear()
        except:
            alpha = specifications_travel_value - travel_value  # 行程差

    def create_txt(self, path, travel_type, type, alpha, beta, delta):
        file_txt = path
        txt_name = "生成參數.txt"
        with open(file_txt + "\\" + txt_name, "w") as f:
            f.write("噸數=%s\n" % type)
            f.write("型式:%s\n" % travel_type)
            f.write("行程=%s\n" % alpha)
            f.write("閉合工作高度=%s\n" % beta)
            f.write("平板前後=%s\n" % delta)

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

    def change_dir(self, i, p, alpha, beta, delta, machining, welding):
        start_time = time.time()
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
                        except:
                            pass
                        mptc.change_machining_parameter(name, i, 0)
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
                        except:
                            pass
                        mptc.change_machining_parameter(name, i, 1)

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
        print('machining_file_change_error', machining_file_change_error)
        print('machining_file_change_pass', machining_file_change_pass)
        print('welding_file_change_error', welding_file_change_error)
        print('welding_file_change_pass', welding_file_change_pass)
        print('總用時%s' % (time.time() - start_time))


if __name__ == "__main__":
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)  # 自適應屏幕分辨率
    app = QtWidgets.QApplication([])
    window = main()
    window.show()
    sys.exit(app.exec_())
