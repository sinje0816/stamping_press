from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QVBoxLayout, QWidget, QPushButton, QLabel, QTableWidgetItem, QHeaderView, QComboBox, \
    QAbstractItemView, QMessageBox, QHBoxLayout
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor, QBrush, QPixmap, QFont
from DEMOGUI import Ui_Dialog
from PAD_main import Ui_Form as pad_main_Form
from PAD_MACHINING import Ui_Form as pad_machining_Form
from PAD_dimension import Ui_Form as pad_dimension_Form
from pad_feeding_hole import Ui_Form as pad_feeding_hole_Form
from cutout_hole_GUI import Ui_Form as cutout_hole_machining_form
from plate_main_first import Ui_Form as plate_main_first_form
from pad_main_simple import Ui_Form as pad_main_simple_form
from punch_main_simple import Ui_Form as punch_main_form
from punch_main import Ui_Form as punch_secend_form
from window_main import Ui_Form
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
import STP_input as S_i

test_stop = False


class main(QtWidgets.QWidget, Ui_Form):
    def __init__(self):
        super(main, self).__init__()
        self.ui = Ui_Form()
        self.ui.setupUi(self)
        self.setting()
        self.ui.start.clicked.connect(self.start)

    def setting(self):
        # 设置额外的宽度和高度来容纳其他界面元素
        extra_width = 800
        extra_height = 150
        # 获取表格的大小
        table_size = self.ui.window_main_table.size()
        # 设置窗口的大小为表格大小加上额外的宽度和高度
        self.setFixedSize(table_size.width() + extra_width, table_size.height() + extra_height)
        window_size = self.size()
        # 调整表格的大小以填充整个窗口
        self.ui.window_main_table.setGeometry(0, 0, window_size.width(), window_size.height())
        for x in range(0, 2):
            self.ui.window_main_table.setSpan(x, 0, 1, 5) #以某格為基準向下向左合併儲存格
            newItem = QTableWidgetItem(par.series1[x]) #儲存格內文字
            newItem.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter) #設定置中
            newItem.setFlags(newItem.flags() & ~Qt.ItemIsEditable) #不可編輯化
            self.ui.window_main_table.setItem(x, 0, newItem) #指定文字放置位置
        for x in range(2, 8):
            self.ui.window_main_table.setSpan(x, 0, 1, 3)
            newItem = QTableWidgetItem(par.series2[x-2])
            newItem.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
            newItem.setFlags(newItem.flags() & ~Qt.ItemIsEditable)
            self.ui.window_main_table.setItem(x, 0, newItem)
        for x in range(2, 6):
            self.ui.window_main_table.setSpan(x, 3, 1, 2)
        combo_box_unit = QComboBox()
        combo_box_unit.setEditable(False)
        combo_box_unit.addItem('公制')
        combo_box_unit.addItem('英制')
        self.ui.window_main_table.setCellWidget(2, 3, combo_box_unit)
        combo_box_law = QComboBox()
        combo_box_law.setEditable(False)
        combo_box_law.addItem('一般內銷')
        combo_box_law.addItem('一般外銷')
        combo_box_law.addItem('美國(ASME)')
        combo_box_law.addItem('加拿大')
        combo_box_law.addItem('巴西')
        combo_box_law.addItem('歐洲(CE)')
        self.ui.window_main_table.setCellWidget(3, 3, combo_box_law)
        combo_box_machine = QComboBox()
        combo_box_machine.setEditable(False)
        combo_box_machine.addItem('SN1-25')
        combo_box_machine.addItem('SN1-35')
        combo_box_machine.addItem('SN1-45')
        combo_box_machine.addItem('SN1-60')
        combo_box_machine.addItem('SN1-80')
        combo_box_machine.addItem('SN1-110')
        combo_box_machine.addItem('SN1-160')
        combo_box_machine.addItem('SN1-200')
        combo_box_machine.addItem('SN1-250')
        self.ui.window_main_table.setCellWidget(4, 3, combo_box_machine)
        combo_box_type = QComboBox()
        combo_box_type.setEditable(False)
        combo_box_type.addItem('S')
        combo_box_type.addItem('H')
        combo_box_type.addItem('P')
        self.ui.window_main_table.setCellWidget(5, 3, combo_box_type)
        plate_setup = QPushButton('設定') #括號內為按鈕名稱
        self.ui.window_main_table.setCellWidget(6, 3, plate_setup)
        punch_setup = QPushButton('設定')
        self.ui.window_main_table.setCellWidget(7, 3, punch_setup)
        pad_finish = QtWidgets.QLabel('未設定')
        pad_finish.setAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        self.ui.window_main_table.setCellWidget(6, 4, pad_finish)
        punch_finish = QtWidgets.QLabel('未設定')
        punch_finish.setAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        self.ui.window_main_table.setCellWidget(7, 4, punch_finish)
        for x in range(8, 12):
            self.ui.window_main_table.setSpan(x, 0, 1, 2)
            newItem = QTableWidgetItem(par.series3[x - 8])
            newItem.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
            newItem.setFlags(newItem.flags() & ~Qt.ItemIsEditable)
            self.ui.window_main_table.setItem(x, 0, newItem)
        self.ui.window_main_table.setSpan(12, 0, 2, 1)
        newItem = QTableWidgetItem('主馬達')
        newItem.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        newItem.setFlags(newItem.flags() & ~Qt.ItemIsEditable)
        self.ui.window_main_table.setItem(12, 0, newItem)
        self.ui.window_main_table.setSpan(14, 0, 2, 1)
        newItem = QTableWidgetItem('變頻器')
        newItem.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        newItem.setFlags(newItem.flags() & ~Qt.ItemIsEditable)
        self.ui.window_main_table.setItem(14, 0, newItem)
        for x in range(12, 16):
            newItem = QTableWidgetItem(par.series4[x - 12])
            newItem.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
            newItem.setFlags(newItem.flags() & ~Qt.ItemIsEditable)
            self.ui.window_main_table.setItem(x, 1, newItem)
        self.ui.window_main_table.setSpan(16, 0, 2, 2)
        newItem = QTableWidgetItem('上模吊重')
        newItem.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        newItem.setFlags(newItem.flags() & ~Qt.ItemIsEditable)
        self.ui.window_main_table.setItem(16, 0, newItem)
        self.ui.window_main_table.setSpan(18, 0, 2, 1)
        newItem = QTableWidgetItem('防震腳')
        newItem.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        newItem.setFlags(newItem.flags() & ~Qt.ItemIsEditable)
        self.ui.window_main_table.setItem(18, 0, newItem)
        newItem = QTableWidgetItem('廠牌')
        newItem.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        newItem.setFlags(newItem.flags() & ~Qt.ItemIsEditable)
        self.ui.window_main_table.setItem(18, 1, newItem)
        newItem = QTableWidgetItem('規格')
        newItem.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        newItem.setFlags(newItem.flags() & ~Qt.ItemIsEditable)
        self.ui.window_main_table.setItem(19, 1, newItem)
        for x in range(20, 22):
            self.ui.window_main_table.setSpan(x, 0, 1, 2)
            newItem = QTableWidgetItem(par.series5[x - 20])
            newItem.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
            newItem.setFlags(newItem.flags() & ~Qt.ItemIsEditable)
            self.ui.window_main_table.setItem(x, 0, newItem)
        for x in range(22, 24):
            self.ui.window_main_table.setSpan(x, 0, 1, 3)
            newItem = QTableWidgetItem(par.series6[x - 22])
            newItem.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
            newItem.setFlags(newItem.flags() & ~Qt.ItemIsEditable)
            self.ui.window_main_table.setItem(x, 0, newItem)
        self.ui.window_main_table.setSpan(16, 2, 2, 1)
        for x in range(8, 17):
            newItem = QTableWidgetItem(par.series7[x - 8])
            newItem.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
            newItem.setFlags(newItem.flags() & ~Qt.ItemIsEditable)
            self.ui.window_main_table.setItem(x, 2, newItem)
        for x in range(18, 22):
            newItem = QTableWidgetItem(par.series8[x - 18])
            newItem.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
            newItem.setFlags(newItem.flags() & ~Qt.ItemIsEditable)
            self.ui.window_main_table.setItem(x, 2, newItem)
        newItem = QTableWidgetItem('標準')
        newItem.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        newItem.setFlags(newItem.flags() & ~Qt.ItemIsEditable)
        self.ui.window_main_table.setItem(8, 3, newItem)
        newItem = QTableWidgetItem('客戶自訂')
        newItem.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        newItem.setFlags(newItem.flags() & ~Qt.ItemIsEditable)
        self.ui.window_main_table.setItem(8, 4, newItem)
        punch_stroke = QTableWidgetItem('80')
        punch_stroke.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        punch_stroke.setFlags(punch_stroke.flags() & ~Qt.ItemIsEditable)
        self.ui.window_main_table.setItem(9, 3, punch_stroke)
        punch_cycle = QTableWidgetItem('70~110')
        punch_cycle.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        punch_cycle.setFlags(punch_cycle.flags() & ~Qt.ItemIsEditable)
        self.ui.window_main_table.setItem(10, 3, punch_cycle)
        punch_DH = QTableWidgetItem('230')
        punch_DH.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        punch_DH.setFlags(punch_DH.flags() & ~Qt.ItemIsEditable)
        self.ui.window_main_table.setItem(11, 3, punch_DH)
        motor_company = QTableWidgetItem('東元')
        motor_company.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        motor_company.setFlags(motor_company.flags() & ~Qt.ItemIsEditable)
        self.ui.window_main_table.setItem(12, 3, motor_company)
        motor_power = QTableWidgetItem('3.7×4')
        motor_power.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        motor_power.setFlags(motor_power.flags() & ~Qt.ItemIsEditable)
        self.ui.window_main_table.setItem(13, 3, motor_power)
        frequency_company = QTableWidgetItem('東元')
        frequency_company.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        frequency_company.setFlags(frequency_company.flags() & ~Qt.ItemIsEditable)
        self.ui.window_main_table.setItem(14, 3, frequency_company)
        frequency_power = QTableWidgetItem('3.7×4')
        frequency_power.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        frequency_power.setFlags(frequency_power.flags() & ~Qt.ItemIsEditable)
        self.ui.window_main_table.setItem(15, 3, frequency_power)
        upper_kg = QTableWidgetItem('54')
        upper_kg.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        upper_kg.setFlags(upper_kg.flags() & ~Qt.ItemIsEditable)
        self.ui.window_main_table.setItem(16, 3, upper_kg)
        for x in range(17, 22):
            self.ui.window_main_table.setSpan(x, 3, 1, 2)
        upper_extra = QTableWidgetItem('PC加大型平衡汽缸')
        upper_extra.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        upper_extra.setFlags(upper_extra.flags() & ~Qt.ItemIsEditable)
        self.ui.window_main_table.setItem(17, 3, upper_extra)
        shockproof_company = QComboBox()
        shockproof_company.setEditable(False)
        shockproof_company.addItem('無')
        shockproof_company.addItem('穎益')
        shockproof_company.addItem('UNISORB')
        shockproof_company.addItem('商定')
        self.ui.window_main_table.setCellWidget(18, 3, shockproof_company)
        shockproof_size = QTableWidgetItem('_')
        shockproof_size.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        shockproof_size.setFlags(shockproof_size.flags() & ~Qt.ItemIsEditable)
        self.ui.window_main_table.setItem(19, 3, shockproof_size)
        work_height = QTableWidgetItem('800')
        work_height.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        work_height.setFlags(work_height.flags() & ~Qt.ItemIsEditable)
        self.ui.window_main_table.setItem(20, 3, work_height)
        power = QComboBox()
        power.setEditable(False)
        power.addItem('200V×50Hz')
        power.addItem('200V×60Hz')
        power.addItem('220V×50Hz')
        power.addItem('220V×60Hz')
        power.addItem('380V×50Hz')
        power.addItem('380V×60Hz')
        power.addItem('415V×50Hz')
        power.addItem('440V×60Hz')
        power.addItem('460V×60Hz')
        power.addItem('480V×60Hz')
        power.addItem('575V×60Hz')
        self.ui.window_main_table.setCellWidget(21, 3, power)
        select_setup = QPushButton('設定')
        self.ui.window_main_table.setCellWidget(22, 3, select_setup)
        spare_parts_setup = QPushButton('設定')
        self.ui.window_main_table.setCellWidget(23, 3, spare_parts_setup)
        select_finish = QtWidgets.QLabel('未設定')
        select_finish.setAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        self.ui.window_main_table.setCellWidget(22, 4, select_finish)
        spare_parts_finish = QtWidgets.QLabel('未設定')
        spare_parts_finish.setAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        self.ui.window_main_table.setCellWidget(23, 4, spare_parts_finish)
        punch_stroke_customize = QTableWidgetItem('')
        punch_stroke_customize.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        # punch_stroke_customize.setBackground(QBrush(QColor(255, 191, 0)))  # 背景色
        self.ui.window_main_table.setItem(9, 4, punch_stroke_customize)
        punch_cycle_customize = QTableWidgetItem('')
        punch_cycle_customize.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        # punch_cycle_customize.setBackground(QBrush(QColor(255, 191, 0)))
        self.ui.window_main_table.setItem(10, 4, punch_cycle_customize)
        punch_DH_customize = QTableWidgetItem('')
        punch_DH_customize.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        # punch_DH_customize.setBackground(QBrush(QColor(255, 191, 0)))
        self.ui.window_main_table.setItem(11, 4, punch_DH_customize)
        motor_company_customize = QTableWidgetItem('')
        motor_company_customize.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        # motor_company_customize.setBackground(QBrush(QColor(255, 191, 0)))
        self.ui.window_main_table.setItem(12, 4, motor_company_customize)
        motor_power_customize = QTableWidgetItem('')
        motor_power_customize.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        # motor_power_customize.setBackground(QBrush(QColor(255, 191, 0)))
        self.ui.window_main_table.setItem(13, 4, motor_power_customize)
        frequency_company_customize = QTableWidgetItem('')
        frequency_company_customize.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        # frequency_company_customize.setBackground(QBrush(QColor(255, 191, 0)))
        self.ui.window_main_table.setItem(14, 4, frequency_company_customize)
        frequency_power_customize = QTableWidgetItem('')
        frequency_power_customize.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        # frequency_power_customize.setBackground(QBrush(QColor(255, 191, 0)))
        self.ui.window_main_table.setItem(15, 4, frequency_power_customize)
        upper_kg_customize = QTableWidgetItem('')
        upper_kg_customize.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        # upper_kg_customize.setBackground(QBrush(QColor(255, 191, 0)))
        self.ui.window_main_table.setItem(16, 4, upper_kg_customize)

        self.ui.window_main_table.setSpan(0, 5, 24, 9)

        # 在初始化函数中创建 QWidget 和 QLabel
        machine_picture_widget = QWidget()
        machine_picture = QLabel(machine_picture_widget)

        # 加载图像
        pixmap = QPixmap('machine_picture.png')
        machine_picture.setPixmap(pixmap)

        scaled_pixmap = pixmap.scaled(pixmap.width()-100,pixmap.height()-110)
        machine_picture.setPixmap(scaled_pixmap)

        # # 设置 QLabel 居中
        machine_picture.setAlignment(Qt.AlignCenter)

        # 添加 QLabel 到 QWidget
        machine_picture_layout = QVBoxLayout()
        machine_picture_layout.addWidget(machine_picture)
        machine_picture_widget.setLayout(machine_picture_layout)

        # 将 QWidget 添加到表格单元格
        self.ui.window_main_table.setCellWidget(0, 5, machine_picture_widget)
        # 将 QLabel 添加到表格单元格
        self.ui.window_main_table.setCellWidget(0, 5, machine_picture)
        # 创建一个 QWidget 作为容器
        button_container = QWidget()
        button_layout = QHBoxLayout()  # 创建一个水平布局对象
        button_container.setLayout(button_layout)  # 将布局附加到button_container上

        self.ui.window_main_table.cellWidget(2, 3).currentIndexChanged.connect(lambda :self.unit_change())
        self.ui.window_main_table.cellWidget(4, 3).currentIndexChanged.connect(lambda :self.choose_stamping_press_type())
        self.ui.window_main_table.cellWidget(5, 3).currentIndexChanged.connect(lambda :self.choose_stamping_press_type())
        # self.ui.window_main_table.item(9, 4).currentChanged.connect(lambda :self.customize_typing('stroke'))
        plate_setup.clicked.connect(lambda:self.switch_to_plate(par.stamping_press_type))
        # punch_setup.clicked.connect()
        # select_setup.clicked.connect()
        # spare_parts_setup.clicked.connect()

    # 切換至平板介面
    def switch_to_plate(self, stamping_press_type):
        self.hide()
        self.nw = plate_first_windows(stamping_press_type)
        self.nw.show()

    # 切換至沖頭介面
    def switch_to_punch(self, stamping_press_type):
        self.hide()
        self.nw = punch_first_windows(stamping_press_type)
        self.nw.show()

    def unit_change(self):
        unit_type = self.ui.window_main_table.cellWidget(2, 3).currentText()
        if unit_type == '公制':
            for x in range(9, 22):
                if x != 17:
                    newItem = QTableWidgetItem(par.unit_metric[x - 9])
                    newItem.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
                    newItem.setFlags(newItem.flags() & ~Qt.ItemIsEditable)
                    self.ui.window_main_table.setItem(x, 2, newItem)
                else:
                    pass
        elif unit_type == '英制':
            for x in range(9, 22):
                if x != 17:
                    newItem = QTableWidgetItem(par.unit_english[x - 9])
                    newItem.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
                    newItem.setFlags(newItem.flags() & ~Qt.ItemIsEditable)
                    self.ui.window_main_table.setItem(x, 2, newItem)
                else:
                    pass

    def customize_typing(self, parameter_type):
        if parameter_type == 'stroke':
            item = self.ui.window_main_table.cellWidget(9, 3).text()
            punch_stroke = QTableWidgetItem(item)
            punch_stroke.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
            punch_stroke.setFlags(punch_stroke.flags() & ~Qt.ItemIsEditable)
            self.ui.window_main_table.setItem(9, 3, punch_stroke)

    def start(self):
        type = self.ui.window_main_table.cellWidget(4, 3).currentText()
        travel_type = str(self.ui.window_main_table.cellWidget(4, 5).currentText())
        specifications_travel_value = str(self.ui.lineEdit_5.text())
        specifications_close_working_height_value = str(self.ui.lineEdit_2.text())
        close_working_height = str(self.ui.label_9.text())
        # delta = str(self.ui.lineEdit_4.text())
        processing = str(self.ui.comboBox.currentText())
        print(type, travel_type, specifications_travel_value, specifications_close_working_height_value,
              close_working_height)
        travel_type = str(self.ui.window_main_table.cellWidget(5, 3).currentText())
        specifications_travel_value = str(self.ui.window_main_table.item(9, 3).text())
        specifications_close_working_height_value = str(self.ui.window_main_table.item(11, 3).text())
        processing = '是'
        print(type, travel_type, specifications_travel_value, specifications_close_working_height_value)
        self.create_dir(type)
        if specifications_travel_value == "":
            self.specifications_travel_value = 0
        else:
            self.specifications_travel_value = int(specifications_travel_value)
        if specifications_close_working_height_value == "":
            self.specifications_close_working_height_value = 0
        else:
            self.specifications_close_working_height_value = int(specifications_close_working_height_value)
        self.stamping_press_type = self.choose_stamping_press_type()
        self.p, self.travel_type = self.choos(processing, travel_type)

        self.alpha, self.beta, self.zeta, self.epsilon = self.frame_calculate(self.stamping_press_type, self.specifications_travel_value,
                                                                              self.specifications_close_working_height_value,
                                                                              self.travel_type)
        if test_stop == False:
            self.create_txt(self.path, type, travel_type, self.specifications_travel_value,
                            self.specifications_close_working_height_value, self.alpha, self.beta, self.zeta,
                            self.epsilon)
            self.change_dir(self.stamping_press_type, self.p, self.alpha, self.beta, self.zeta, self.epsilon, self.machining,
                            self.welding, self.travel_type)

    def choose_stamping_press_type(self):
        type = self.ui.window_main_table.cellWidget(4, 3).currentText()
        style = self.ui.window_main_table.cellWidget(5, 3).currentText()
        if type == "SN1-25":
            stamping_press_type = 0
        elif type == "SN1-35":
            stamping_press_type = 1
        elif type == "SN1-45":
            stamping_press_type = 2
        elif type == "SN1-60":
            stamping_press_type = 3
        elif type == "SN1-80":
            stamping_press_type = 4
        elif type == "SN1-110":
            stamping_press_type = 5
        elif type == "SN1-160":
            stamping_press_type = 6
        elif type == "SN1-200":
            stamping_press_type = 7
        elif type == "SN1-250":
            stamping_press_type = 8
        punch_stroke = QTableWidgetItem(par.stamping_press_stroke[style][stamping_press_type])
        punch_stroke.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        punch_stroke.setFlags(punch_stroke.flags() & ~Qt.ItemIsEditable)
        self.ui.window_main_table.setItem(9, 3, punch_stroke)
        punch_cycle = QTableWidgetItem(par.stamping_press_cycle[style][stamping_press_type])
        punch_cycle.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        punch_cycle.setFlags(punch_cycle.flags() & ~Qt.ItemIsEditable)
        self.ui.window_main_table.setItem(10, 3, punch_cycle)
        punch_DH = QTableWidgetItem(par.stamping_press_DH[style][stamping_press_type])
        punch_DH.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        punch_DH.setFlags(punch_DH.flags() & ~Qt.ItemIsEditable)
        self.ui.window_main_table.setItem(11, 3, punch_DH)

        par.stamping_press_type = stamping_press_type
        par.stamping_press_style = style

        return stamping_press_type


    def choos(self,prossing, travel_type):
        # 確認型號"輸入型號"
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
        return p, travel_type

    def frame_calculate(self, stamping_press_type, specifications_travel_value, specifications_close_working_height_value, travel_type):
        Form = QtWidgets.QWidget()
        # Form.setWindowTitle('警告')
        Form.resize(400, 300)
        mbox = QtWidgets.QMessageBox(Form)

        # 讀取標準資料
        excel = epc.ExcelOp('尺寸整理表', '標準資料')
        type_name, travel_value, close_working_height_value, specifications_travel_min_value, specifications_travel_max_value, specifications_close_working_height_min_value, specifications_close_working_height_max_value = excel.get_standard_parts(
            stamping_press_type * 3)

        # 噸數&行程類型
        all_type = stamping_press_type * 3 + travel_type - 1
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
        test = os.path.join('Z:')
        path = desktop + '\\' + dir_name
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

    def change_dir(self, stamping_press_type, p, alpha, beta, zeta, epsilon, machining, welding, travel_type):
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
            file_list_name, file_list_value = epc.ExcelOp('尺寸整理表', '沖床機架零件清單').get_sheet_par('沖床機架零件清單', stamping_press_type)
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
                    if name == 'PANEL' or name == 'CON_ROD' or name == 'CON_ROD_BASE' or name == 'CON_ROD_CAP' or name == 'INVERTERBRACKET' \
                            or name == 'POINTER' or name == 'COVER' or name == 'PLUG' or name == 'feeding_shaft_cover' or name == 'OIL_LEVEL_GAUGE' \
                            or name == 'slide_gib' or name == 'ELECTRIC_BOX_PLATE' or name == 'MOUNT_FILTER'or name == 'CONTROL_PANEL' or name == 'PANEL_BOX'\
                            or name == 'PANEL_BOX_BRACKET' or name == 'ELECTRIC_BOX' or name == 'GUARD_FLYWHEEL' or name == 'NAME_PLATE'\
                            or name == 'TRADEMARK_NAMEPLATE'or name == 'OPERATION_BOX' or name == 'PORTABLE_STAND' or name == 'OPERATION_BOX'\
                            or name == 'BEARING_HOUSING':
                        # 讀取其餘STP檔
                        S_i.STP(name, stamping_press_type, machining)
                        continue
                    else:
                        # 讀取機架零件
                        mprog.import_part(fp.system_root + fp.DEMO_part, name)
                    if name == 'FRAME52' and p == 0:
                        try:
                            mprog.param_change(name, "alpha", alpha)
                            mprog.param_change(name, "beta", beta)
                            mprog.param_change(name, 'zeta', zeta)
                            mprog.param_change(name, 'epsilon', epsilon)
                        except:
                            pass
                        # 加工圖零件
                        parameter_name, parameter_value = mptc.change_machining_parameter(name, stamping_press_type, 0, travel_type)
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
                            mprog.param_change(name, 'epsilon', epsilon)
                        except:
                            pass

                        # 加工圖零件
                        parameter_name, parameter_value = mptc.change_machining_parameter(name, stamping_press_type, 1, travel_type)
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
                    wptc.change_welding_feature(name, stamping_press_type)
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
        print(stamping_press_type)
        print('all_part_name', all_part_name)
        print('all_part_value', all_part_value)
        print('machining_file_change_error', machining_file_change_error)
        print('machining_file_change_pass', machining_file_change_pass)
        print('welding_file_change_error', welding_file_change_error)
        print('welding_file_change_pass', welding_file_change_pass)
        print('總用時%s' % (time.time() - start_time))  # 建立3D組立
        Ad.assembly(stamping_press_type, apv, self.path, alpha, beta, zeta, epsilon)

        return machining_file_change_error, welding_file_change_error


class punch_first_windows(QtWidgets.QWidget):
    def __init__(self, stamping_press_type):
        super().__init__()
        self.ui = punch_main_form()
        self.ui.setupUi(self)
        self.setWindowTitle('衝頭')
        self.ui.punch_select.currentIndexChanged.connect(lambda: self.select_punch_type_name(stamping_press_type))
        self.ui.punch_save.clicked.connect(self.switch_to_stamping_press_main_windows)

    def switch_to_stamping_press_main_windows(self):
        self.hide()
        self.nw = main()
        self.nw.show()

    def select_punch_type_name(self, stamping_press_type):
        get_punch_select_name = self.ui.punch_select.currentText()
        if get_punch_select_name == "特殊衝頭":
            self.hide()
            self.nw = punch_secend_windows(stamping_press_type)
            self.nw.show()


class punch_secend_windows(QtWidgets.QWidget):
    def __init__(self, stamping_press_type):
        super().__init__()
        self.ui = punch_secend_form()
        self.ui.setupUi(self)
        self.setWindowTitle('衝頭')
        # 切換衝頭第一頁
        self.ui.punch_select.currentIndexChanged.connect(
            lambda: self.switch_to_first_punch_windows(stamping_press_type))
        self.ui.punch_select.setCurrentText('特殊衝頭')
        # 沖頭平板面積
        self.punch_type(stamping_press_type)
        self.ui.punch_extrasize.currentIndexChanged.connect(lambda: self.punch_area_dimension(stamping_press_type))
        self.ui.punch_LR.setText(str(par.punch_length[stamping_press_type]))
        self.ui.punch_FB.setText(str(par.punch_width[stamping_press_type]))
        # T溝
        for number in range(0, 4):
            self.ui.t_solttable.setItem(number, 0, QTableWidgetItem(par.t_table_dimension_parameter[number]))
        self.ui.t_solttable.verticalHeader().setVisible(False)
        self.ui.t_solttable.horizontalHeader().setVisible(False)
        # 設定T_solt表格內容
        self.T_solt_table_normel_setup()
        for number in range(1, 10):
            newItem = QTableWidgetItem("-")
            self.ui.t_solttable.setItem(number, 1, newItem)
        self.ui.t_solttable.setSizeAdjustPolicy(QtWidgets.QAbstractScrollArea.AdjustToContents)
        self.ui.t_solttable.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.ui.t_solttable.verticalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.ui.t_solt_type.currentIndexChanged.connect(self.T_solt_combobox_change)
        self.ui.punch_start.clicked.connect(lambda: self.start(stamping_press_type))
        self.ui.t_machining.clicked.connect(lambda: self.showpadmachiningwindows(stamping_press_type))
        self.ui.punch_save.clicked.connect(self.switch_to_stamping_press_main_windows)

    def switch_to_stamping_press_main_windows(self):
        self.hide()
        self.nw = main()
        self.nw.show()

    def switch_to_first_punch_windows(self, stamping_press_type):
        if self.ui.punch_select.currentText() != "特殊衝頭":
            self.hide()
            self.nw = punch_first_windows(stamping_press_type)
            self.nw.show()

    def punch_area_dimension(self, stamping_press_type):
        get_pad_select_name = self.ui.punch_select.currentText()
        get_pad_area_name = self.ui.punch_extrasize.currentText()
        if get_pad_select_name == "特殊衝頭":
            if get_pad_area_name == '標準':
                LR_value = str(par.punch_length[stamping_press_type])
                FB_value = str(par.punch_width[stamping_press_type])
            elif get_pad_area_name == '加大I型':
                LR_value = str(par.punch_length[stamping_press_type] + par.punch_lv1[stamping_press_type])
                FB_value = str(par.punch_width[stamping_press_type])
            elif get_pad_area_name == '加大II型':
                LR_value = str(par.punch_length[stamping_press_type] + par.punch_lv2[stamping_press_type])
                FB_value = str(par.punch_width[stamping_press_type])
            par.plate_special_type = [get_pad_area_name]
            # 設定最終的 LR 和 FB 值
            self.ui.punch_LR.setText(LR_value)
            self.ui.punch_FB.setText(FB_value)
        else:
            self.ui.punch_LR.clear()
            self.ui.punch_FB.clear()

    def punch_type(self, stamping_press_type):
        self.ui.punch_select.setItemText(0, '標準' + '(' + str(
            par.punch_length[stamping_press_type]) + 'x' + str(par.punch_width[stamping_press_type]) + ")")
        self.ui.punch_select.setItemText(1, '加大I型' + '(' + str(
            par.punch_length[stamping_press_type] + par.punch_lv1[stamping_press_type]) + 'x' + str(
            par.punch_width[stamping_press_type]) + ")")
        self.ui.punch_select.setItemText(2, '加大II型' + '(' + str(
            par.punch_length[stamping_press_type] + par.punch_lv2[stamping_press_type]) + 'x' + str(
            par.punch_width[stamping_press_type]) + ")")

    def T_solt_table_normel_setup(self):
        # 設定T_solt表格內容
        # 第一行
        newItem = QTableWidgetItem("尺寸代號")
        self.ui.t_solttable.setItem(0, 0, newItem)
        newItem = QTableWidgetItem("尺寸")
        self.ui.t_solttable.setItem(0, 1, newItem)
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
        if t_solt_type == "T溝代號:F(SN1-25~60標準)" or t_solt_type == "T溝代號:G(SN1-80~250標準)":
            if t_solt_type == "T溝代號:F(SN1-25~60標準)":
                self.ui.removepicture_2.setPixmap(QtGui.QPixmap("T溝詳圖_F.png"))
            elif t_solt_type == "T溝代號:G(SN1-80~250標準)":
                self.ui.removepicture_2.setPixmap(QtGui.QPixmap("T溝詳圖_G.png"))
            for number in range(1, 10):
                newItem = QTableWidgetItem("-")
                self.ui.t_solttable.setItem(number, 1, newItem)
        elif t_solt_type == "特殊T溝":
            self.ui.removepicture_2.setPixmap(QtGui.QPixmap("T溝詳圖.png"))
            for number in range(1, 10):
                newItem = QTableWidgetItem("")
                self.ui.t_solttable.setItem(number, 1, newItem)
        for number in range(0, 10):
            for number2 in range(0, 3):
                item = self.ui.t_solttable.item(number, number2)
                if item is not None:
                    item.setFlags(item.flags() & ~QtCore.Qt.ItemIsEditable)
        if self.ui.t_solt_type.currentText() == '特殊T溝':
            for number in range(1, 10):
                item = self.ui.t_solttable.item(number, 1)
                if item is not None:
                    item.setFlags(item.flags() | QtCore.Qt.ItemIsEditable)
                    self.ui.t_solttable.setEditTriggers(QtWidgets.QAbstractItemView.AllEditTriggers)

    def showpadmachiningwindows(self, stamping_press_type):
        t_solt_type = self.ui.t_solt_type.currentText()
        if t_solt_type == "T溝代號:F(SN1-25~60標準)":
            par.t_all_dimension = [22, 38, 23, 16]
        elif t_solt_type == "T溝代號:G(SN1-80~250標準)":
            par.t_all_dimension = [28, 48, 30, 20]
        elif t_solt_type == "特殊T溝":
            for value in range(1, 9, 2):
                par.t_all_dimension.append(self.ui.t_solttable.item(value, 1).text())
        punch_type_select = self.ui.punch_extrasize.currentText()
        if punch_type_select == '標準':
            punch_length = par.punch_length[stamping_press_type]
            punch_width = par.punch_width[stamping_press_type]
        elif punch_type_select == '加大I型':
            punch_length = par.punch_length[stamping_press_type] + par.punch_lv1[stamping_press_type]
            punch_width = par.punch_width[stamping_press_type]
        elif punch_type_select == '加大II型':
            punch_length = par.punch_length[stamping_press_type] + par.punch_lv2[stamping_press_type]
            punch_width = par.punch_width[stamping_press_type]

        self.hide()
        self.nw = t_machining(stamping_press_type, punch_length, punch_width, 'punch_secend_windows')
        self.nw.show()

    def start(self, stamping_press_type):
        # 對平板進行變數
        path = FolderManager('plate').path
        path = str(path)
        mprog.set_CATIA_workbench_env()
        mprog.import_part(fp.system_root + fp.DEMO_part, 'plate')
        punch_name, punch_value = pdp.padchange(stamping_press_type)
        for name in punch_name:
            par.punch_all_parameter[name] = punch_value[punch_name.index(name)]
        print(par.punch_all_parameter)
        if len(par.punch_normal_name) != 0:
            if '加大I型' in par.punch_normal_name[0]:
                par.lv = [par.punch_lv1[stamping_press_type]]
                mprog.param_change('punch', 'LV', par.lv[0])
            elif '加大II型' in par.punch_normal_name[0]:
                par.lv = [par.punch_lv2[stamping_press_type]]
                mprog.param_change('punch', 'LV', par.lv[0])
            else:
                par.lv = [0]
            if '模墊型' in par.punch_normal_name[0]:
                if stamping_press_type <= 4:
                    mprog.activatefeature('cutout_molded_cushion', 4)
                    mprog.param_change('punch', 'cutout_molded_cushion_A',
                                       par.normal_cutout_molded_cushion_A[stamping_press_type])
                    mprog.param_change('punch', 'cutout_molded_cushion_B',
                                       par.normal_cutout_molded_cushion_B[stamping_press_type])
                    mprog.param_change('punch', 'cutout_molded_cushion_D',
                                       par.normal_cutout_molded_cushion_D[stamping_press_type])
                    mprog.param_change('punch', 'cutout_molded_cushion_length',
                                       par.normal_cutout_molded_cushion_length[stamping_press_type])
                    mprog.param_change('punch', 'cutout_molded_cushion_width',
                                       par.normal_cutout_molded_cushion_width[stamping_press_type])
            mprog.Update()
        else:
            punch_type = self.ui.punch_extrasize.currentText()
            par.punch_special_type = [punch_type]
            if par.punch_special_type[0] == '標準':
                par.lv = [0]
            elif par.punch_special_type[0] == '加大I型':
                par.lv.append(par.punch_lv1[stamping_press_type])
                mprog.param_change('punch', 'LV', par.lv[0])
            elif par.punch_special_type[0] == '加大II型':
                par.lv.append(par.punch_lv2[stamping_press_type])
                mprog.param_change('punch', 'LV', par.lv[0])
            # T溝程式
            self.t_solt(path)
            # 關閉實體外所有東西
            mprog.Close_All()
            # 平板存檔
            mprog.save_file_stp(path, 'punch')
            mprog.save_stpfile_part(path, 'punch')
        return path

    # T形槽
    def t_solt(self, path):
        # 對T溝進行變數變換
        mprog.import_part(fp.system_root + fp.DEMO_part, 'T')
        for t in par.t_all_dimension:
            for t_name in range(len(par.t_all_dimension_name) + 1):
                mprog.param_change('T', par.t_all_dimension_name[t_name], t)
                break
        mprog.param_change('T', 'LB', 0)
        mprog.param_change('T', 'LF', 0)
        mprog.param_change('T', 'SB', 0)
        mprog.param_change('T', 'SF', 0)
        mprog.Update()
        mprog.save_file_stp(path, 'T')
        mprog.save_stpfile_part(path, 'T')
        mprog.close_window()

        # 橫向T溝
        if len(par.total_position_y) != 0:
            for turn in range(0, len(par.total_position_y)):
                mprog.import_part(path, 'T')
                tT.changerotate(0)
                mprog.param_change('T', 'A', (par.plate_all_parameter['A'] + par.lv[0]))
                mprog.param_change('T', 'C', par.plate_all_parameter['C'])
                if par.total_t_slot_h_type[turn] == '貫穿' or par.total_t_slot_h_type[turn] == '':
                    mprog.param_change('T', 'LL', (par.plate_all_parameter['A'] + par.lv[0]))
                    mprog.param_change('T', 'LR', (par.plate_all_parameter['A'] + par.lv[0]))
                elif par.total_t_slot_h_type[turn] == '分段':
                    mprog.param_change('T', 'LL', par.total_LL[turn])
                    mprog.param_change('T', 'LR', par.total_LR[turn])
                # 判斷SL和SR是否為0或空值
                if len(par.total_SL) != 0:
                    if par.total_SL[turn] == '' or par.total_SL[turn] == '0':
                        mprog.partdeactivate('讓孔1')
                        mprog.partdeactivate('讓孔倒圓角1')
                    elif par.total_SL[turn] != '':
                        mprog.param_change('T', 'SL', par.total_SL[turn])
                else:
                    mprog.partdeactivate('讓孔1')
                    mprog.partdeactivate('讓孔倒圓角1')
                if len(par.total_SR) != 0:
                    if par.total_SR[turn] == '' or par.total_SR[turn] == '0' or len(par.total_SR) == 0:
                        mprog.partdeactivate('讓孔2')
                        mprog.partdeactivate('讓孔倒圓角2')
                    elif par.total_SR[turn] != '':
                        mprog.param_change('T', 'SR', par.total_SR[turn])
                else:
                    mprog.partdeactivate('讓孔2')
                    mprog.partdeactivate('讓孔倒圓角2')
                tT.create_t_solt((par.plate_all_parameter['B'] / 2) + par.total_position_y[turn], turn)
                if len(par.total_SL) != 0:
                    if par.total_SL[turn] == '' or par.total_SL[turn] == '0':
                        mprog.partbodyfeatureactivate('讓孔1')
                        mprog.partbodyfeatureactivate('讓孔倒圓角1')
                else:
                    mprog.partbodyfeatureactivate('讓孔1')
                    mprog.partbodyfeatureactivate('讓孔倒圓角1')
                if len(par.total_SR) != 0:
                    if par.total_SR[turn] == '' or par.total_SR[turn] == '0':
                        mprog.partbodyfeatureactivate('讓孔2')
                        mprog.partbodyfeatureactivate('讓孔倒圓角2')
                else:
                    mprog.partbodyfeatureactivate('讓孔2')
                    mprog.partbodyfeatureactivate('讓孔倒圓角2')
                mprog.Update()
                mprog.close_window()
                # time.sleep(0.1)
        else:
            turn = 0

        # 縱向T溝
        if len(par.total_position_x) != 0:
            if turn != 0:
                mprog.param_change('T', 'LL', 0)
                mprog.param_change('T', 'LR', 0)
                mprog.param_change('T', 'SL', 0)
                mprog.param_change('T', 'SR', 0)
            for turn_2 in range(0, len(par.total_position_x)):
                mprog.import_part(path, 'T')
                tT.changerotate(-90)
                mprog.param_change('T', 'A', par.plate_all_parameter['B'])
                mprog.param_change('T', 'C', par.plate_all_parameter['C'])
                if par.total_t_slot_v_type[turn_2] == '貫穿' or par.total_t_slot_v_type[turn_2] == '':
                    mprog.param_change('T', 'LF', par.plate_all_parameter['B'])
                    mprog.param_change('T', 'LB', par.plate_all_parameter['B'])
                elif par.total_t_slot_v_type[turn_2] == '分段':
                    mprog.param_change('T', 'LF', par.total_LF[turn_2])
                    mprog.param_change('T', 'LB', par.total_LB[turn_2])
                # 判斷SL和SR是否為0或空值
                if par.total_SF[turn_2] == '' or par.total_SF[turn_2] == '0':
                    mprog.partdeactivate('讓孔1')
                    mprog.partdeactivate('讓孔倒圓角1')
                elif par.total_SF[turn_2] != '':
                    mprog.param_change('T', 'SF', par.total_SF[turn_2])
                if par.total_SB[turn_2] == '' or par.total_SB[turn_2] == '0':
                    mprog.partdeactivate('讓孔2')
                    mprog.partdeactivate('讓孔倒圓角2')
                elif par.total_SB[turn_2] != '':
                    mprog.param_change('T', 'SB', par.total_SB[turn_2])
                if turn == 0:
                    tT.create_t_solt(
                        -((par.plate_all_parameter['A'] / 2) + (par.lv[0] / 2) + par.total_position_x[turn_2]), turn_2)
                else:
                    tT.create_t_solt(
                        -((par.plate_all_parameter['A'] / 2) + (par.lv[0] / 2) + par.total_position_x[turn_2]),
                        turn_2 + turn + 1)
                if par.total_SF[turn_2] == '' or par.total_SF[turn_2] == '0':
                    mprog.param_change('T', 'SF', 10)
                    mprog.partbodyfeatureactivate('讓孔1')
                    mprog.partbodyfeatureactivate('讓孔倒圓角1')
                if par.total_SB[turn_2] == '' or par.total_SB[turn_2] == '0':
                    mprog.param_change('T', 'SB', 10)
                    mprog.partbodyfeatureactivate('讓孔2')
                    mprog.partbodyfeatureactivate('讓孔倒圓角2')
                mprog.Update()
                mprog.close_window()
        print('T-slot create successfully')


# 平板第一頁
class plate_first_windows(QtWidgets.QWidget):
    def __init__(self, stamping_press_type):
        super().__init__()
        self.ui = pad_main_simple_form()
        self.ui.setupUi(self)
        self.setWindowTitle('平板')
        self.ui.pad_select.currentIndexChanged.connect(lambda: self.select_plate_type_name(stamping_press_type))
        self.ui.plate_start.clicked.connect(lambda: self.start(stamping_press_type))
        self.ui.plate_escape.clicked.connect(self.switch_to_stamping_press_main_windows)

    # 選擇平板類型
    def select_plate_type_name(self, stamping_press_type):
        get_pad_select_name = self.ui.pad_select.currentText()
        if get_pad_select_name == "特殊平板":
            self.hide()
            self.nw = pad_secend_windows(stamping_press_type)
            self.nw.show()

    def switch_to_stamping_press_main_windows(self):
        self.hide()
        self.nw = main()
        self.nw.show()

    def start(self, stamping_press_type):
        par.plate_normal_name = [self.ui.pad_select.currentText()]
        # 判斷模墊型
        if par.plate_normal_name[0] == "標準模墊型":
            if stamping_press_type == 0:
                par.total_t_slot_h_type = ['貫穿', '分段', '貫穿']
                par.total_position_y = [100, 0, -100]
                par.total_LL = ['', 185, '']
                par.total_LR = ['', 185, '']
                par.t_all_dimension = [22, 38, 23, 16]
            elif stamping_press_type == 1:
                par.total_t_slot_h_type = ['貫穿', '分段', '貫穿']
                par.total_position_y = [125, 0, -125]
                par.total_LL = ['', 210, '']
                par.total_LR = ['', 210, '']
                par.t_all_dimension = [22, 38, 23, 16]
            elif stamping_press_type == 2:
                par.total_t_slot_h_type = ['貫穿', '分段', '貫穿']
                par.total_position_y = [140, 0, -140]
                par.total_LL = ['', 220, '']
                par.total_LR = ['', 220, '']
                par.t_all_dimension = [22, 38, 23, 16]
            elif stamping_press_type == 3:
                par.total_t_slot_h_type = ['貫穿', '分段', '貫穿']
                par.total_position_y = [150, 0, -150]
                par.total_LL = ['', 240, '']
                par.total_LR = ['', 240, '']
                par.t_all_dimension = [22, 38, 23, 16]
            elif stamping_press_type == 4:
                par.total_t_slot_h_type = ['貫穿', '分段', '貫穿']
                par.total_position_y = [180, 0, -180]
                par.total_LL = ['', 270, '']
                par.total_LR = ['', 270, '']
                par.t_all_dimension = [28, 48, 30, 20]
            elif stamping_press_type == 5:
                par.total_t_slot_h_type = ['貫穿', '分段', '貫穿']
                par.total_position_y = [210, 0, -210]
                par.total_LL = ['', 280, '']
                par.total_LR = ['', 280, '']
                par.t_all_dimension = [28, 48, 30, 20]
            elif stamping_press_type == 6:
                par.total_t_slot_h_type = ['貫穿', '分段', '分段', '分段', '貫穿']
                par.total_position_y = [290, 165, 0, -165, -290]
                par.total_LL = ['', 290, 290, 290, '']
                par.total_LR = ['', 290, 290, 290, '']
                par.t_all_dimension = [28, 48, 30, 20]
            elif stamping_press_type == 7:
                par.total_t_slot_h_type = ['貫穿', '分段', '分段', '分段', '貫穿']
                par.total_position_y = [320, 180, 0, -180, -320]
                par.total_LL = ['', 320, 320, 320, '']
                par.total_LR = ['', 320, 320, 320, '']
                par.t_all_dimension = [28, 48, 30, 20]
            elif stamping_press_type == 8:
                par.total_t_slot_h_type = ['貫穿', '分段', '分段', '分段', '貫穿']
                par.total_position_y = [350, 200, 0, -200, -350]
                par.total_LL = ['', 320, 320, 320, '']
                par.total_LR = ['', 320, 320, 320, '']
                par.t_all_dimension = [28, 48, 30, 20]
        else:
            # 判斷下料孔
            if '圓孔' in par.plate_normal_name[0]:
                if stamping_press_type == 0:
                    par.cutout_part_dimension = [100, '', '', '', '']
                elif stamping_press_type == 1:
                    par.cutout_part_dimension = [110, '', '', '', '']
                elif stamping_press_type == 2:
                    par.cutout_part_dimension = [130, '', '', '', '']
                elif stamping_press_type == 3:
                    par.cutout_part_dimension = [150, '', '', '', '']
                elif stamping_press_type == 4:
                    par.cutout_part_dimension = [180, '', '', '', '']
                elif stamping_press_type == 5:
                    par.cutout_part_dimension = [200, '', '', '', '']
                elif stamping_press_type == 6:
                    par.cutout_part_dimension = [220, '', '', '', '']
                elif stamping_press_type == 7:
                    par.cutout_part_dimension = [250, '', '', '', '']
                elif stamping_press_type == 8:
                    par.cutout_part_dimension = [270, '', '', '', '']
            elif '方孔' in par.plate_normal_name[0]:
                if stamping_press_type == 0:
                    par.cutout_part_dimension = [200, 100, '', '', '']
                    par.cutout_spuare_R = [20]
                elif stamping_press_type == 1:
                    par.cutout_part_dimension = [220, 110, '', '', '']
                    par.cutout_spuare_R = [25]
                elif stamping_press_type == 2:
                    par.cutout_part_dimension = [260, 130, '', '', '']
                    par.cutout_spuare_R = [25]
                elif stamping_press_type == 3:
                    par.cutout_part_dimension = [300, 150, '', '', '']
                    par.cutout_spuare_R = [30]
                elif stamping_press_type == 4:
                    par.cutout_part_dimension = [360, 180, '', '', '']
                    par.cutout_spuare_R = [35]
                elif stamping_press_type == 5:
                    par.cutout_part_dimension = [400, 200, '', '', '']
                    par.cutout_spuare_R = [40]
                elif stamping_press_type == 6:
                    par.cutout_part_dimension = [440, 220, '', '', '']
                    par.cutout_spuare_R = [45]
                elif stamping_press_type == 7:
                    par.cutout_part_dimension = [500, 250, '', '', '']
                    par.cutout_spuare_R = [50]
                elif stamping_press_type == 8:
                    par.cutout_part_dimension = [540, 270, '', '', '']
                    par.cutout_spuare_R = [55]
            # 判斷T溝
            if stamping_press_type == 0:
                par.total_t_slot_h_type = ['貫穿', '貫穿', '貫穿']
                par.total_position_y = [100, 0, -100]
                par.t_all_dimension = [22, 38, 23, 16]
            elif stamping_press_type == 1:
                par.total_t_slot_h_type = ['貫穿', '貫穿', '貫穿']
                par.total_position_y = [125, 0, -125]
                par.t_all_dimension = [22, 38, 23, 16]
            elif stamping_press_type == 2:
                par.total_t_slot_h_type = ['貫穿', '貫穿', '貫穿']
                par.total_position_y = [140, 0, -140]
                par.t_all_dimension = [22, 38, 23, 16]
            elif stamping_press_type == 3:
                par.total_t_slot_h_type = ['貫穿', '貫穿', '貫穿']
                par.total_position_y = [150, 0, -150]
                par.t_all_dimension = [22, 38, 23, 16]
            elif stamping_press_type == 4:
                par.total_t_slot_h_type = ['貫穿', '貫穿', '貫穿']
                par.total_position_y = [180, 0, -180]
                par.t_all_dimension = [28, 48, 30, 20]
            elif stamping_press_type == 5:
                par.total_t_slot_h_type = ['貫穿', '貫穿', '貫穿']
                par.total_position_y = [210, 0, -210]
                par.t_all_dimension = [28, 48, 30, 20]
            elif stamping_press_type == 6:
                par.total_t_slot_h_type = ['貫穿', '貫穿', '貫穿', '貫穿', '貫穿']
                par.total_position_y = [290, 165, 0, -165, -290]
                par.t_all_dimension = [28, 48, 30, 20]
            elif stamping_press_type == 7:
                par.total_t_slot_h_type = ['貫穿', '貫穿', '貫穿', '貫穿', '貫穿']
                par.total_position_y = [320, 180, 0, -180, -320]
                par.t_all_dimension = [28, 48, 30, 20]
            elif stamping_press_type == 8:
                par.total_t_slot_h_type = ['貫穿', '貫穿', '貫穿', '貫穿', '貫穿']
                par.total_position_y = [350, 200, 0, -200, -350]
                par.t_all_dimension = [28, 48, 30, 20]
        # 連結到平板第二頁的函式進行生成
        path = pad_secend_windows.start(stamping_press_type, stamping_press_type)
        # T溝
        pad_secend_windows.t_solt(stamping_press_type, path)
        # 下料孔
        pad_secend_windows.plate_hole(stamping_press_type, stamping_press_type, path)
        # 關閉實體外所有東西
        mprog.Close_All()
        # 平板存檔
        mprog.save_file_stp(path, 'plate')
        mprog.save_stpfile_part(path, 'plate')


# 平板第二頁
class pad_secend_windows(QtWidgets.QWidget):
    def __init__(self, stamping_press_type):
        super().__init__()
        self.ui = pad_main_Form()
        self.ui.setupUi(self)
        self.setWindowTitle('平板')
        # 切換平板第一頁
        self.ui.pad_select.currentIndexChanged.connect(lambda: self.switch_to_first_plate_windows(stamping_press_type))
        self.ui.pad_select.setCurrentText('特殊平板')
        # 下料孔
        self.ui.remove_type.currentIndexChanged.connect(self.cutout_parameter_change)
        self.ui.removetable.verticalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.ui.removetable.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.ui.removetable.setSizeAdjustPolicy(QtWidgets.QAbstractScrollArea.AdjustToContents)
        self.ui.removetable.verticalHeader().setVisible(False)
        # 平板面積
        self.plate_type(stamping_press_type)
        self.ui.pad_extrasize.currentIndexChanged.connect(lambda: self.plate_area_dimension(stamping_press_type))
        if len(par.plate_special_type) != 0:
            self.ui.pad_extrasize.setCurrentText(par.plate_special_type[0])
            self.plate_area_dimension(stamping_press_type)
        else:
            self.ui.LR.setText(str(par.plate_length[stamping_press_type]))
            self.ui.FB.setText(str(par.plate_width[stamping_press_type]))
        # T溝
        for number in range(0, 4):
            self.ui.t_solttable.setItem(number, 0, QTableWidgetItem(par.t_table_dimension_parameter[number]))
        self.ui.t_solttable.verticalHeader().setVisible(False)
        self.ui.t_solttable.horizontalHeader().setVisible(False)
        # 設定T_solt表格內容
        self.T_solt_table_normel_setup()
        for number in range(1, 10):
            newItem = QTableWidgetItem("-")
            self.ui.t_solttable.setItem(number, 1, newItem)
        self.ui.t_solttable.setSizeAdjustPolicy(QtWidgets.QAbstractScrollArea.AdjustToContents)
        self.ui.t_solttable.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.ui.t_solttable.verticalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.ui.t_solt_type.currentIndexChanged.connect(self.T_solt_combobox_change)
        if stamping_press_type <= 3:
            self.ui.t_solt_type.setCurrentText('T溝代號:F(SN1-25~60標準)')
        else:
            self.ui.t_solt_type.setCurrentText('T溝代號:G(SN1-80~250標準)')
        self.ui.plate_start.clicked.connect(lambda: self.start(stamping_press_type))
        self.ui.t_machining.clicked.connect(lambda: self.showpadmachiningwindows(stamping_press_type))
        self.ui.remove_machining.clicked.connect(lambda: self.showcutoutmachiningwindows(stamping_press_type))
        self.chack_plate_table()
        self.ui.plate_escape.clicked.connect(self.switch_to_stamping_press_main_windows)

    def switch_to_stamping_press_main_windows(self):
        self.hide()
        self.nw = main()
        self.nw.show()

    def switch_to_first_plate_windows(self, stamping_press_type):
        if self.ui.pad_select.currentText() != "特殊平板":
            self.hide()
            self.nw = plate_first_windows(stamping_press_type)
            self.nw.show()

    def plate_area_dimension(self, stamping_press_type):
        get_pad_select_name = self.ui.pad_select.currentText()
        get_pad_area_name = self.ui.pad_extrasize.currentText()
        if get_pad_select_name == "特殊平板":
            if get_pad_area_name == '標準':
                LR_value = str(par.plate_length[stamping_press_type])
                FB_value = str(par.plate_width[stamping_press_type])
            elif get_pad_area_name == '加大I型':
                LR_value = str(par.plate_length[stamping_press_type] + par.plate_lv1[stamping_press_type])
                FB_value = str(par.plate_width[stamping_press_type])
            elif get_pad_area_name == '加大II型':
                LR_value = str(par.plate_length[stamping_press_type] + par.plate_lv2[stamping_press_type])
                FB_value = str(par.plate_width[stamping_press_type])
            par.plate_special_type = [get_pad_area_name]
            # 設定最終的 LR 和 FB 值
            self.ui.LR.setText(LR_value)
            self.ui.FB.setText(FB_value)
        else:
            self.ui.LR.clear()
            self.ui.FB.clear()

    def plate_type(self, stamping_press_type):
        for number in range(0, 10):
            self.ui.pad_select.setItemText(number, str(par.plate_type[number]) + '(' + str(
                par.plate_length[stamping_press_type]) + 'x' + str(par.plate_width[stamping_press_type]) + ")")

    def T_solt_table_normel_setup(self):
        # 設定T_solt表格內容
        # 第一行
        newItem = QTableWidgetItem("尺寸代號")
        self.ui.t_solttable.setItem(0, 0, newItem)
        newItem = QTableWidgetItem("尺寸")
        self.ui.t_solttable.setItem(0, 1, newItem)
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
        if t_solt_type == "T溝代號:F(SN1-25~60標準)" or t_solt_type == "T溝代號:G(SN1-80~250標準)":
            if t_solt_type == "T溝代號:F(SN1-25~60標準)":
                self.ui.removepicture_2.setPixmap(QtGui.QPixmap("T溝詳圖_F.png"))
            elif t_solt_type == "T溝代號:G(SN1-80~250標準)":
                self.ui.removepicture_2.setPixmap(QtGui.QPixmap("T溝詳圖_G.png"))
            for number in range(1, 10):
                newItem = QTableWidgetItem("-")
                self.ui.t_solttable.setItem(number, 1, newItem)
        elif t_solt_type == "特殊T溝":
            self.ui.removepicture_2.setPixmap(QtGui.QPixmap("T溝詳圖.png"))
            for number in range(1, 10):
                newItem = QTableWidgetItem("")
                self.ui.t_solttable.setItem(number, 1, newItem)
        for number in range(0, 10):
            for number2 in range(0, 3):
                item = self.ui.t_solttable.item(number, number2)
                if item is not None:
                    item.setFlags(item.flags() & ~QtCore.Qt.ItemIsEditable)
        if self.ui.t_solt_type.currentText() == '特殊T溝':
            for number in range(1, 10):
                item = self.ui.t_solttable.item(number, 1)
                if item is not None:
                    item.setFlags(item.flags() | QtCore.Qt.ItemIsEditable)
                    self.ui.t_solttable.setEditTriggers(QtWidgets.QAbstractItemView.AllEditTriggers)

    def chack_plate_table(self):
        if len(par.plate_hole_type) != 0:
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

    # def showpaddimensionwindows(self, stamping_press_type):
    #     self.hide()
    #     self.nw = pad_dimension(stamping_press_type)
    #     self.nw.show()

    def showpadmachiningwindows(self, stamping_press_type):
        t_solt_type = self.ui.t_solt_type.currentText()
        if t_solt_type == "T溝代號:F(SN1-25~60標準)":
            par.t_all_dimension = [22, 38, 23, 16]
        elif t_solt_type == "T溝代號:G(SN1-80~250標準)":
            par.t_all_dimension = [28, 48, 30, 20]
        elif t_solt_type == "特殊T溝":
            for value in range(1, 9, 2):
                par.t_all_dimension.append(self.ui.t_solttable.item(value, 1).text())
        plate_type_select = self.ui.pad_extrasize.currentText()
        if plate_type_select == '標準':
            plate_lenght = par.plate_length[stamping_press_type]
            plate_width = par.plate_width[stamping_press_type]
        elif plate_type_select == '加大I型':
            plate_lenght = par.plate_length[stamping_press_type] + par.plate_lv1[stamping_press_type]
            plate_width = par.plate_width[stamping_press_type]
        elif plate_type_select == '加大II型':
            plate_lenght = par.plate_length[stamping_press_type] + par.plate_lv2[stamping_press_type]
            plate_width = par.plate_width[stamping_press_type]

        self.hide()
        self.nw = t_machining(stamping_press_type, plate_lenght, plate_width, 'pad_secend_windows')
        self.nw.show()

    def showcutoutmachiningwindows(self, i):
        for turn in range(0, 5):
            par.cutout_part_dimension[turn] = (self.ui.removetable.item(turn, 1).text())
        print(par.cutout_part_dimension)
        if par.plate_hole_type == '漏斗型':
            if int(par.cutout_part_dimension[0]) < int(par.cutout_part_dimension[1]) or int(
                    par.cutout_part_dimension[2]) < int(par.cutout_part_dimension[3]):
                print('error')
        self.hide()
        self.nw = cutout_hole_machining(i)
        self.nw.show()

    # 下料孔形狀選單連動功能
    def cutout_parameter_change(self):
        translate = QtCore.QCoreApplication.translate
        cutout_type = self.ui.remove_type.currentText()
        for number in range(0, 5):
            self.ui.removetable.setItem(number, 0, QTableWidgetItem(''))
            self.ui.removetable.setItem(number, 1, QTableWidgetItem(''))
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
        for number in range(0, 5):
            item = self.ui.removetable.item(number, 0)
            if item is not None:
                item.setFlags(item.flags() & ~QtCore.Qt.ItemIsEditable)
        par.plate_hole_type = [self.ui.remove_type.currentText()]

    def start(self, stamping_press_type):
        # 對平板進行變數
        path = FolderManager('plate').path
        path = str(path)
        mprog.set_CATIA_workbench_env()
        mprog.import_part(fp.system_root + fp.DEMO_part, 'plate')
        plate_name, plate_value = pdp.padchange(stamping_press_type)
        for name in plate_name:
            par.plate_all_parameter[name] = plate_value[plate_name.index(name)]
        # print(par.plate_all_parameter)
        if len(par.plate_normal_name) != 0:
            if '加大I型' in par.plate_normal_name[0]:
                par.lv = [par.plate_lv1[stamping_press_type]]
                mprog.param_change('plate', 'LV', par.lv[0])
            elif '加大II型' in par.plate_normal_name[0]:
                par.lv = [par.plate_lv2[stamping_press_type]]
                mprog.param_change('plate', 'LV', par.lv[0])
            else:
                par.lv = [0]
            if '模墊型' in par.plate_normal_name[0]:
                if stamping_press_type <= 4:
                    mprog.activatefeature('cutout_molded_cushion', 4)
                    mprog.param_change('plate', 'cutout_molded_cushion_A',
                                       par.normal_cutout_molded_cushion_A[stamping_press_type])
                    mprog.param_change('plate', 'cutout_molded_cushion_B',
                                       par.normal_cutout_molded_cushion_B[stamping_press_type])
                    mprog.param_change('plate', 'cutout_molded_cushion_D',
                                       par.normal_cutout_molded_cushion_D[stamping_press_type])
                    mprog.param_change('plate', 'cutout_molded_cushion_length',
                                       par.normal_cutout_molded_cushion_length[stamping_press_type])
                    mprog.param_change('plate', 'cutout_molded_cushion_width',
                                       par.normal_cutout_molded_cushion_width[stamping_press_type])
            mprog.Update()
        else:
            plate_type = self.ui.pad_extrasize.currentText()
            par.plate_special_type = [plate_type]
            if par.plate_special_type[0] == '標準':
                par.lv = [0]
            elif par.plate_special_type[0] == '加大I型':
                par.lv.append(par.plate_lv1[stamping_press_type])
                mprog.param_change('plate', 'LV', par.lv[0])
            elif par.plate_special_type[0] == '加大II型':
                par.lv.append(par.plate_lv2[stamping_press_type])
                mprog.param_change('plate', 'LV', par.lv[0])
            # T溝程式
            self.t_solt(path)
            # 下料孔程式
            self.plate_hole(stamping_press_type, path)
            # 關閉實體外所有東西
            mprog.Close_All()
            # 平板存檔
            mprog.save_file_stp(path, 'plate')
            mprog.save_stpfile_part(path, 'plate')
        return path

    # T形槽
    def t_solt(self, path):
        # 對T溝進行變數變換
        mprog.import_part(fp.system_root + fp.DEMO_part, 'T')
        for t in par.t_all_dimension:
            for t_name in range(len(par.t_all_dimension_name) + 1):
                mprog.param_change('T', par.t_all_dimension_name[t_name], t)
                break
        mprog.param_change('T', 'LB', 0)
        mprog.param_change('T', 'LF', 0)
        mprog.param_change('T', 'SB', 0)
        mprog.param_change('T', 'SF', 0)
        mprog.Update()
        mprog.save_file_stp(path, 'T')
        mprog.save_stpfile_part(path, 'T')
        mprog.close_window()

        # 橫向T溝
        if len(par.total_position_y) != 0:
            for turn in range(0, len(par.total_position_y)):
                mprog.import_part(path, 'T')
                tT.changerotate(0)
                mprog.param_change('T', 'A', (par.plate_all_parameter['A'] + par.lv[0]))
                mprog.param_change('T', 'C', par.plate_all_parameter['C'])
                if par.total_t_slot_h_type[turn] == '貫穿' or par.total_t_slot_h_type[turn] == '':
                    mprog.param_change('T', 'LL', (par.plate_all_parameter['A'] + par.lv[0]))
                    mprog.param_change('T', 'LR', (par.plate_all_parameter['A'] + par.lv[0]))
                elif par.total_t_slot_h_type[turn] == '分段':
                    mprog.param_change('T', 'LL', par.total_LL[turn])
                    mprog.param_change('T', 'LR', par.total_LR[turn])
                # 判斷SL和SR是否為0或空值
                if len(par.total_SL) != 0:
                    if par.total_SL[turn] == '' or par.total_SL[turn] == '0':
                        mprog.partdeactivate('讓孔1')
                        mprog.partdeactivate('讓孔倒圓角1')
                    elif par.total_SL[turn] != '':
                        mprog.param_change('T', 'SL', par.total_SL[turn])
                else:
                    mprog.partdeactivate('讓孔1')
                    mprog.partdeactivate('讓孔倒圓角1')
                if len(par.total_SR) != 0:
                    if par.total_SR[turn] == '' or par.total_SR[turn] == '0' or len(par.total_SR) == 0:
                        mprog.partdeactivate('讓孔2')
                        mprog.partdeactivate('讓孔倒圓角2')
                    elif par.total_SR[turn] != '':
                        mprog.param_change('T', 'SR', par.total_SR[turn])
                else:
                    mprog.partdeactivate('讓孔2')
                    mprog.partdeactivate('讓孔倒圓角2')
                tT.create_t_solt((par.plate_all_parameter['B'] / 2) + par.total_position_y[turn], turn)
                if len(par.total_SL) != 0:
                    if par.total_SL[turn] == '' or par.total_SL[turn] == '0':
                        mprog.partbodyfeatureactivate('讓孔1')
                        mprog.partbodyfeatureactivate('讓孔倒圓角1')
                else:
                    mprog.partbodyfeatureactivate('讓孔1')
                    mprog.partbodyfeatureactivate('讓孔倒圓角1')
                if len(par.total_SR) != 0:
                    if par.total_SR[turn] == '' or par.total_SR[turn] == '0':
                        mprog.partbodyfeatureactivate('讓孔2')
                        mprog.partbodyfeatureactivate('讓孔倒圓角2')
                else:
                    mprog.partbodyfeatureactivate('讓孔2')
                    mprog.partbodyfeatureactivate('讓孔倒圓角2')
                mprog.Update()
                mprog.close_window()
                # time.sleep(0.1)
        else:
            turn = 0

        # 縱向T溝
        if len(par.total_position_x) != 0:
            if turn != 0:
                mprog.param_change('T', 'LL', 0)
                mprog.param_change('T', 'LR', 0)
                mprog.param_change('T', 'SL', 0)
                mprog.param_change('T', 'SR', 0)
            for turn_2 in range(0, len(par.total_position_x)):
                mprog.import_part(path, 'T')
                tT.changerotate(-90)
                mprog.param_change('T', 'A', par.plate_all_parameter['B'])
                mprog.param_change('T', 'C', par.plate_all_parameter['C'])
                if par.total_t_slot_v_type[turn_2] == '貫穿' or par.total_t_slot_v_type[turn_2] == '':
                    mprog.param_change('T', 'LF', par.plate_all_parameter['B'])
                    mprog.param_change('T', 'LB', par.plate_all_parameter['B'])
                elif par.total_t_slot_v_type[turn_2] == '分段':
                    mprog.param_change('T', 'LF', par.total_LF[turn_2])
                    mprog.param_change('T', 'LB', par.total_LB[turn_2])
                # 判斷SL和SR是否為0或空值
                if par.total_SF[turn_2] == '' or par.total_SF[turn_2] == '0':
                    mprog.partdeactivate('讓孔1')
                    mprog.partdeactivate('讓孔倒圓角1')
                elif par.total_SF[turn_2] != '':
                    mprog.param_change('T', 'SF', par.total_SF[turn_2])
                if par.total_SB[turn_2] == '' or par.total_SB[turn_2] == '0':
                    mprog.partdeactivate('讓孔2')
                    mprog.partdeactivate('讓孔倒圓角2')
                elif par.total_SB[turn_2] != '':
                    mprog.param_change('T', 'SB', par.total_SB[turn_2])
                if turn == 0:
                    tT.create_t_solt(
                        -((par.plate_all_parameter['A'] / 2) + (par.lv[0] / 2) + par.total_position_x[turn_2]), turn_2)
                else:
                    tT.create_t_solt(
                        -((par.plate_all_parameter['A'] / 2) + (par.lv[0] / 2) + par.total_position_x[turn_2]),
                        turn_2 + turn + 1)
                if par.total_SF[turn_2] == '' or par.total_SF[turn_2] == '0':
                    mprog.param_change('T', 'SF', 10)
                    mprog.partbodyfeatureactivate('讓孔1')
                    mprog.partbodyfeatureactivate('讓孔倒圓角1')
                if par.total_SB[turn_2] == '' or par.total_SB[turn_2] == '0':
                    mprog.param_change('T', 'SB', 10)
                    mprog.partbodyfeatureactivate('讓孔2')
                    mprog.partbodyfeatureactivate('讓孔倒圓角2')
                mprog.Update()
                mprog.close_window()
        print('T-slot create successfully')

    # 下料孔
    def plate_hole(self, stamping_press_type, path):
        try:
            par.plate_hole_type = [self.ui.remove_type.currentText()]
        except:
            pass
        if len(par.plate_hole_type) != 0:
            if par.plate_hole_type[0] != '無孔' or len(par.plate_hole_type) != 0:
                par.plate_hole_type = [self.ui.remove_type.currentText()]
                print(stamping_press_type)
                for x in range(0, 5):
                    par.cutout_part_dimension[x] = (self.ui.removetable.item(x, 1).text())
                print(par.cutout_part_dimension)
                print(stamping_press_type)
                ch.edge_test(stamping_press_type)
                # 開下料孔檔案+變數變換
                if par.plate_hole_type[0] == '圓孔':
                    mprog.import_part(fp.system_root + fp.DEMO_part, 'cutout_hole_circle')
                    mprog.param_change('cutout_hole_circle', par.cutout_parameter_circle[0],
                                       int(par.cutout_part_dimension[0]))
                    mprog.param_change('cutout_hole_circle', 'C', par.plate_all_parameter['C'])
                    mprog.param_change('cutout_hole_circle', 'X',
                                       par.cutout_hole_machining_X - par.plate_all_parameter['A'] / 2 - par.lv[0] / 2)
                    mprog.param_change('cutout_hole_circle', 'Y',
                                       par.plate_all_parameter['B'] / 2 + par.cutout_hole_machining_Y)
                    par.plate_hole_type[0] = 'cutout_hole_circle'
                elif par.plate_hole_type[0] == '方孔':
                    mprog.import_part(fp.system_root + fp.DEMO_part, 'cutout_hole_square')
                    for n in range(0, 2):
                        mprog.param_change('cutout_hole_square', par.cutout_parameter_square[n],
                                           int(par.cutout_part_dimension[n]))
                    mprog.param_change('cutout_hole_square', 'C', par.plate_all_parameter['C'])
                    mprog.param_change('cutout_hole_square', 'X',
                                       par.cutout_hole_machining_X - par.plate_all_parameter['A'] / 2 - par.lv[0] / 2)
                    mprog.param_change('cutout_hole_square', 'Y',
                                       par.plate_all_parameter['B'] / 2 + par.cutout_hole_machining_Y)
                    par.plate_hole_type[0] = 'cutout_hole_square'
                elif par.plate_hole_type[0] == '漏斗型':
                    mprog.import_part(fp.system_root + fp.DEMO_part, 'cutout_funnel')
                    for n in range(0, 5):
                        mprog.param_change('cutout_funnel', par.cutout_parameter_funnel[n],
                                           int(par.cutout_part_dimension[n]))
                    mprog.param_change('cutout_funnel', 'C', par.plate_all_parameter['C'])
                    mprog.param_change('cutout_funnel', 'X',
                                       par.cutout_hole_machining_X - par.plate_all_parameter['A'] / 2 - par.lv[0] / 2)
                    mprog.param_change('cutout_funnel', 'Y',
                                       par.plate_all_parameter['B'] / 2 + par.cutout_hole_machining_Y)
                    par.plate_hole_type[0] = 'cutout_funnel'
                elif par.plate_hole_type[0] == '模墊型':
                    mprog.import_part(fp.system_root + fp.DEMO_part, 'cutout_molded_cushion')
                    mprog.param_change('cutout_molded_cushion', 'C', par.plate_all_parameter['C'])
                    if stamping_press_type < 5:
                        mprog.param_change('cutout_molded_cushion', 'A', par.cutout_molded_cushion_A[0])
                        mprog.param_change('cutout_molded_cushion', 'B', par.cutout_molded_cushion_B[0])
                        mprog.param_change('cutout_molded_cushion', 'D', par.cutout_molded_cushion_L[0])
                    else:
                        mprog.param_change('cutout_molded_cushion', 'A', par.cutout_molded_cushion_A[1])
                        mprog.param_change('cutout_molded_cushion', 'B', par.cutout_molded_cushion_B[1])
                        mprog.param_change('cutout_molded_cushion', 'D', par.cutout_molded_cushion_L[1])
                    mprog.param_change('cutout_molded_cushion', 'width_gap', 60)
                    mprog.param_change('cutout_molded_cushion', 'length_gap', 65)
                    mprog.param_change('cutout_molded_cushion', 'i', par.cutout_molded_cushion_i[stamping_press_type])
                    mprog.param_change('cutout_molded_cushion', 'j', par.cutout_molded_cushion_j[stamping_press_type])
                    mprog.param_change('cutout_molded_cushion', 'X',
                                       65 * (par.cutout_molded_cushion_i[stamping_press_type] - 1) / 2 -
                                       par.plate_all_parameter[
                                           'A'] / 2 - par.lv[0] / 2)
                    mprog.param_change('cutout_molded_cushion', 'Y',
                                       par.plate_all_parameter['B'] / 2 - 60 * (
                                                   par.cutout_molded_cushion_j[stamping_press_type] - 1) / 2)
                    par.plate_hole_type[0] = 'cutout_molded_cushion'
                mprog.Update()
                mprog.save_file_stp(path, par.plate_hole_type[0])
                mprog.save_stpfile_part(path, par.plate_hole_type[0])
                tT.copybody()
                tT.switch_to_window_by_name('plate.CATPart')
                # 計算下料孔變數名稱
                if len(par.total_position_y) != 0:
                    position_y = len(par.total_position_y)
                else:
                    position_y = 0
                if len(par.total_position_x) != 0:
                    position_x = len(par.total_position_x)
                else:
                    position_x = 0
                total_turn = position_y + position_x
                print(total_turn)
                # 移除下料孔
                tT.pastebody(total_turn, par.plate_hole_type[0])
                tT.removebody(total_turn, par.plate_hole_type[0])
                mprog.Update()
                tT.switch_to_window_by_name(par.plate_hole_type[0] + ".CATPart")
                mprog.close_window()
        if len(par.plate_normal_name) != 0:
            if '無孔' not in par.plate_normal_name[0]:
                # 開下料孔檔案+變數變換
                if '圓孔' in par.plate_normal_name[0]:
                    mprog.import_part(fp.system_root + fp.DEMO_part, 'cutout_hole_circle')
                    mprog.param_change('cutout_hole_circle', par.cutout_parameter_circle[0],
                                       int(par.cutout_part_dimension[0]))
                    mprog.param_change('cutout_hole_circle', 'C', par.plate_all_parameter['C'])
                    mprog.param_change('cutout_hole_circle', 'X',
                                       par.cutout_hole_machining_X - par.plate_all_parameter['A'] / 2 - par.lv[
                                           0] / 2)
                    mprog.param_change('cutout_hole_circle', 'Y',
                                       par.plate_all_parameter['B'] / 2 + par.cutout_hole_machining_Y)
                    par.plate_normal_name[0] = 'cutout_hole_circle'
                elif '方孔' in par.plate_normal_name[0]:
                    mprog.import_part(fp.system_root + fp.DEMO_part, 'cutout_hole_square')
                    for n in range(0, 2):
                        mprog.param_change('cutout_hole_square', par.cutout_parameter_square[n],
                                           int(par.cutout_part_dimension[n]))
                    mprog.param_change('cutout_hole_square', 'C', par.plate_all_parameter['C'])
                    mprog.param_change('cutout_hole_square', 'X',
                                       par.cutout_hole_machining_X - par.plate_all_parameter['A'] / 2 - par.lv[
                                           0] / 2)
                    mprog.param_change('cutout_hole_square', 'Y',
                                       par.plate_all_parameter['B'] / 2 + par.cutout_hole_machining_Y)
                    mprog.param_change('cutout_hole_square', 'edge_fillet', par.cutout_spuare_R[0])
                    par.plate_normal_name[0] = 'cutout_hole_square'
                elif '漏斗型' in par.plate_normal_name[0]:
                    mprog.import_part(fp.system_root + fp.DEMO_part, 'cutout_funnel')
                    for n in range(0, 5):
                        mprog.param_change('cutout_funnel', par.cutout_parameter_funnel[n],
                                           int(par.cutout_part_dimension[n]))
                    mprog.param_change('cutout_funnel', 'C', par.plate_all_parameter['C'])
                    mprog.param_change('cutout_funnel', 'X',
                                       par.cutout_hole_machining_X - par.plate_all_parameter['A'] / 2 - par.lv[
                                           0] / 2)
                    mprog.param_change('cutout_funnel', 'Y',
                                       par.plate_all_parameter['B'] / 2 + par.cutout_hole_machining_Y)
                    par.plate_normal_name[0] = 'cutout_funnel'
                elif '模墊型' in par.plate_normal_name[0]:
                    mprog.import_part(fp.system_root + fp.DEMO_part, 'cutout_molded_cushion')
                    mprog.param_change('cutout_molded_cushion', 'C', par.plate_all_parameter['C'])
                    if stamping_press_type < 5:
                        mprog.param_change('cutout_molded_cushion', 'A', par.cutout_molded_cushion_A[0])
                        mprog.param_change('cutout_molded_cushion', 'B', par.cutout_molded_cushion_B[0])
                        mprog.param_change('cutout_molded_cushion', 'D', par.cutout_molded_cushion_L[0])
                    else:
                        mprog.param_change('cutout_molded_cushion', 'A', par.cutout_molded_cushion_A[1])
                        mprog.param_change('cutout_molded_cushion', 'B', par.cutout_molded_cushion_B[1])
                        mprog.param_change('cutout_molded_cushion', 'D', par.cutout_molded_cushion_L[1])
                    mprog.param_change('cutout_molded_cushion', 'i',
                                       par.normal_cutout_molded_cushion_length_quantity[stamping_press_type])
                    mprog.param_change('cutout_molded_cushion', 'j',
                                       par.normal_cutout_molded_cushion_width_quantity[stamping_press_type])
                    mprog.param_change('cutout_molded_cushion', 'width_gap',
                                       par.normal_cutout_molded_cushion_length_gap[stamping_press_type])
                    mprog.param_change('cutout_molded_cushion', 'length_gap',
                                       par.normal_cutout_molded_cushion_width_gap[stamping_press_type])
                    mprog.param_change('cutout_molded_cushion', 'X', -(
                                (par.plate_all_parameter['A'] + par.lv[0]) / 2 - (
                                    par.normal_cutout_molded_cushion_length_quantity[stamping_press_type] - 1) *
                                par.normal_cutout_molded_cushion_width_gap[stamping_press_type] / 2))
                    print((par.plate_all_parameter['A'] + par.lv[0]) / 2 - (
                                par.normal_cutout_molded_cushion_length_quantity[stamping_press_type] - 1) *
                          par.normal_cutout_molded_cushion_width_gap[stamping_press_type] / 2)

                    mprog.param_change('cutout_molded_cushion', 'Y',
                                       (par.plate_all_parameter['B'] / 2 -
                                        (par.normal_cutout_molded_cushion_length_gap[stamping_press_type] *
                                         (par.normal_cutout_molded_cushion_width_quantity[
                                              stamping_press_type] - 1)) / 2))
                    par.plate_normal_name[0] = 'cutout_molded_cushion'

                mprog.Update()
                mprog.save_file_stp(path, par.plate_normal_name[0])
                mprog.save_stpfile_part(path, par.plate_normal_name[0])
                tT.copybody()
                tT.switch_to_window_by_name('plate.CATPart')
                # 計算下料孔變數名稱
                if len(par.total_position_y) != 0:
                    position_y = len(par.total_position_y)
                else:
                    position_y = 0
                if len(par.total_position_x) != 0:
                    position_x = len(par.total_position_x)
                else:
                    position_x = 0
                total_turn = position_y + position_x
                print(total_turn)
                # 移除下料孔
                tT.pastebody(total_turn, par.plate_normal_name[0])
                tT.removebody(total_turn, par.plate_normal_name[0])
                mprog.Update()
                tT.switch_to_window_by_name(par.plate_normal_name[0] + ".CATPart")
                mprog.close_window()

class t_machining(QWidget):
    def __init__(self, stamping_press_type, lenght, width, parent_page):
        super().__init__()
        self.ui = pad_machining_Form()
        self.ui.setupUi(self)
        self.setWindowTitle('T溝加工設定')
        # 橫向T溝
        self.ui.t_slot_table_h.verticalHeader().setVisible(False)
        self.ui.t_slot_table_h.horizontalHeader().setVisible(False)
        self.ui.t_slot_table_h.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.ui.t_slot_table_h.setEditTriggers(QAbstractItemView.AllEditTriggers)
        self.t_slot_table_h_setup()
        self.ui.t_slot_h_number.textChanged.connect(self.check_slot_number)
        self.table_h_combo_boxes = {}
        # 縱向T溝
        self.ui.t_slot_table_v.verticalHeader().setVisible(False)
        self.ui.t_slot_table_v.horizontalHeader().setVisible(False)
        self.ui.t_slot_table_v.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.ui.t_slot_table_v.setEditTriggers(QAbstractItemView.AllEditTriggers)
        self.t_slot_table_v_setup()
        self.ui.t_slot_v_number.textChanged.connect(self.check_slot_v_number)
        self.table_v_combo_boxes = {}
        # 確定按鈕
        self.ui.setup.clicked.connect(lambda: self.setup(stamping_press_type, parent_page, lenght, width))
        # 重製按鈕
        self.ui.reset.clicked.connect(lambda :self.reset(parent_page))
        # 重新排列
        self.ui.rearrange_the_order.clicked.connect(lambda: self.rearrange_the_order(lenght, width, parent_page))
        # 禁用setup
        # self.ui.setup.setEnabled(False)
        # self.ui.setup.setStyleSheet("background-color: gray; color: white;")
        # 將上一輪資料輸入表格
        self.re_paste_the_data(parent_page)

    def re_paste_the_data(self, parent_page):
        if parent_page == 'pad_secend_windows':
            if len(par.total_position_y) != 0:
                self.ui.t_slot_h_number.setText(str(len(par.total_position_y)))
                for row, item_text in enumerate(par.total_position_y):
                    self.ui.t_slot_table_h.setItem(row + 1, 1, QTableWidgetItem(str(item_text)))
                    combo_box = QComboBox()
                    combo_box.addItem("")
                    combo_box.addItem("貫穿")
                    combo_box.addItem("分段")
                    cell_widget = self.ui.t_slot_table_h.cellWidget(row + 1, 2)
                    cell_widget.setCurrentText(par.total_t_slot_h_type[row])
                    if par.total_t_slot_h_type[row] != '分段':
                        item = self.ui.t_slot_table_h.item(row+1, 3)
                        item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                        item = self.ui.t_slot_table_h.item(row + 1, 4)
                        item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    else:
                        self.ui.t_slot_table_h.setItem(row + 1, 3, QTableWidgetItem(str(par.total_LL[row])))
                        self.ui.t_slot_table_h.setItem(row + 1, 4, QTableWidgetItem(str(par.total_LR[row])))
                    self.ui.t_slot_table_h.setItem(row + 1, 5, QTableWidgetItem(str(par.total_SL[row])))
                    self.ui.t_slot_table_h.setItem(row + 1, 6, QTableWidgetItem(str(par.total_SR[row])))
            if len(par.total_position_x) != 0:
                self.ui.t_slot_v_number.setText(str(len(par.total_position_x)))
                for row, item_text in enumerate(par.total_position_x):
                    self.ui.t_slot_table_v.setItem(row + 1, 1, QTableWidgetItem(str(item_text)))
                    combo_box = QComboBox()
                    combo_box.addItem(par.total_t_slot_v_type[row])
                    self.ui.t_slot_table_v.setCellWidget(row + 1, 2, combo_box)
                    if par.total_t_slot_v_type[row] != '分段':
                        item = self.ui.t_slot_table_v.item(row+1, 3)
                        item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                        item = self.ui.t_slot_table_v.item(row + 1, 4)
                        item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    else:
                        self.ui.t_slot_table_v.setItem(row + 1, 3, QTableWidgetItem(str(par.total_LF[row])))
                        self.ui.t_slot_table_v.setItem(row + 1, 4, QTableWidgetItem(str(par.total_LB[row])))
                    self.ui.t_slot_table_v.setItem(row + 1, 5, QTableWidgetItem(str(par.total_SF[row])))
                    self.ui.t_slot_table_v.setItem(row + 1, 6, QTableWidgetItem(str(par.total_SB[row])))



    def restore_setup_button(self):
        # 啟用
        self.ui.setup.setEnabled(True)
        # 還原樣式
        self.ui.setup.setStyleSheet("")

    def change_table_h_clear_table(self):
        while self.ui.t_slot_table_h.rowCount() > 1:
            self.ui.t_slot_table_h.removeRow(1)

    def check_slot_number(self, current_value):
        try:
            if current_value == '':
                current_value = 0
            current_value = int(current_value)
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
                if is_editable and row != 0:
                    item.setFlags(item.flags() | Qt.ItemIsEditable)
                    item.setBackground(QBrush(QColor(255, 255, 255)))
                else:
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    item.setBackground(QBrush(QColor(192, 192, 192)))

    def t_slot_table_h_setup(self):
        # 第一行
        for col in range(7):  # 假设有7列
            newItem = QTableWidgetItem(
                "編號" if col == 0 else "位置(Y)" if col == 1 else "形式" if col == 2 else "LL" if col == 3 else "LR" if col == 4 else "SL" if col == 5 else "SR")
            newItem.setTextAlignment(Qt.AlignCenter)  # 可以根据需要设置文本对齐方式
            self.ui.t_slot_table_h.setItem(0, col, newItem)
            newItem.setFlags(newItem.flags() & ~Qt.ItemIsEditable)  # 移除可编辑标志

    def combo_box_changed_v(self, row, index):
        combo_box = self.table_v_combo_boxes.get(row)
        if combo_box:
            if index == 2:
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
            current_value = int(current_value)
            self.add_rows_to_table_v(current_value)
        except ValueError:
            print("Invalid input. Please enter a valid number")

    def add_rows_to_table_v(self, num_rows):
        self.change_table_v_clear_table()
        col_count = self.ui.t_slot_table_v.columnCount()
        for row in range(num_rows):
            row_position = self.ui.t_slot_table_v.rowCount()
            self.ui.t_slot_table_v.insertRow(row_position)
            for col in range(col_count):
                item = QTableWidgetItem("V" + str(row_position) if (row_position == 0 or col == 0) else "")
                if col == 0 or row_position == 0:
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                else:
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
                if is_editable and row != 0:
                    item.setFlags(item.flags() | Qt.ItemIsEditable)
                    item.setBackground(QBrush(QColor(255, 255, 255)))
                else:
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    item.setBackground(QBrush(QColor(192, 192, 192)))

    def t_slot_table_v_setup(self):
        # 第一行
        for col in range(7):  # 假設有7列
            newItem = QTableWidgetItem(
                "編號" if col == 0 else "位置(X)" if col == 1 else "形式" if col == 2 else "LF" if col == 3 else "LB" if col == 4 else "SF" if col == 5 else "SB")
            newItem.setTextAlignment(Qt.AlignCenter)  # 可以根據需要設定文本對齊方式
            self.ui.t_slot_table_v.setItem(0, col, newItem)
            newItem.setFlags(newItem.flags() & ~Qt.ItemIsEditable)  # 移除可編輯標誌

    def setup(self, stamping_press_type, parent_page, lenght, width):
        self.rearrange_the_order(lenght, width, parent_page)
        if parent_page == 'pad_secend_windows':
            self.hide()
            self.nw = pad_secend_windows(stamping_press_type)
            self.nw.show()
        elif parent_page == 'punch_secend_windows':
            self.hide()
            self.nw = punch_secend_windows(stamping_press_type)
            self.nw.show()

    def rearrange_the_order(self, length, width, parent_page):
        if parent_page == 'pad_secend_windows':
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

                # 檢查T溝是否超出界線
                # T溝尺寸與平板邊緣>=50mm
                if len(par.total_position_y) != 0:
                    if par.total_position_y[0] + (par.t_all_dimension[1] / 2) > (width / 2) - 50:
                        self.show_alert('橫向T溝位置(Y) H1 與平板邊緣過近')
                    if par.total_position_y[-1] - (par.t_all_dimension[1] / 2) <= (-width / 2) + 50:
                        self.show_alert('橫向T溝位置(Y) H' + str(len(par.total_position_y)) + ' 與平板邊緣過近')

                    # T溝與T溝之間間距需>=50mm
                    for turn in range(len(par.total_position_y) - 1):
                        third_value = abs(par.total_position_y[turn] - par.total_position_y[turn + 1]) - \
                                      par.t_all_dimension[1]
                        if third_value <= 50:
                            self.show_alert('橫向T溝位置(Y) H' + str(turn + 1) + ' 與 H' + str(turn + 2) + ' 間距過近')
                        print('third_value:', third_value)

                if len(par.total_position_x) != 0:
                    if par.total_position_x[0] + (par.t_all_dimension[0] / 2) > (length / 2) - 50:
                        self.show_alert('縱向T溝位置(X) V1 與平板邊緣過近')
                    if par.total_position_x[-1] - (par.t_all_dimension[0] / 2) <= (-length / 2) + 50:
                        self.show_alert('縱向T溝位置(X) V' + str(len(par.total_position_x)) + ' 與平板邊緣過近')

                    # T溝與T溝之間間距需>=50mm
                    for turn in range(len(par.total_position_x) - 1):
                        third_value = abs(par.total_position_x[turn] - par.total_position_x[turn + 1]) - \
                                      par.t_all_dimension[0]
                        if third_value <= 50:
                            self.show_alert('縱向T溝位置(X) V' + str(turn + 1) + ' 與 V' + str(turn + 2) + ' 間距過近')
                        print('third_value:', third_value)

        elif parent_page == 'punch_secend_windows':
            par.ph_total_position_y.clear()
            par.ph_total_t_slot_h_type.clear()
            par.ph_total_LL.clear()
            par.ph_total_LR.clear()
            par.ph_total_SL.clear()
            par.ph_total_SR.clear()
            par.ph_total_position_x.clear()
            par.ph_total_t_slot_v_type.clear()
            par.ph_total_LF.clear()
            par.ph_total_LB.clear()
            par.ph_total_SF.clear()
            par.ph_total_SB.clear()

            if self.ui.t_slot_table_h.rowCount() != 0:
                for row in range(1, self.ui.t_slot_table_h.rowCount()):
                    position_y = self.ui.t_slot_table_h.item(row, 1).text()
                    t_slot_type = self.ui.t_slot_table_h.cellWidget(row, 2).currentText()
                    LL = self.ui.t_slot_table_h.item(row, 3).text()
                    LR = self.ui.t_slot_table_h.item(row, 4).text()
                    SL = self.ui.t_slot_table_h.item(row, 5).text()
                    SR = self.ui.t_slot_table_h.item(row, 6).text()
                    par.ph_total_position_y.append(position_y)
                    par.ph_total_t_slot_h_type.append(t_slot_type)
                    par.ph_total_LL.append(LL)
                    par.ph_total_LR.append(LR)
                    par.ph_total_SL.append(SL)
                    par.ph_total_SR.append(SR)
                par.ph_total_position_y = [int(x) for x in par.ph_total_position_y]
                position_y_change_position = sorted(enumerate(par.ph_total_position_y), key=lambda x: x[1], reverse=True)
                print('position_y_change_position:', position_y_change_position)
                rearrange = [position[0] for position in position_y_change_position]
                print('rearrange:', rearrange)
                par.ph_total_position_y = [par.ph_total_position_y[order_position] for order_position in rearrange]
                par.ph_total_LL = [par.ph_total_LL[order_position] for order_position in rearrange]
                par.ph_total_LR = [par.ph_total_LR[order_position] for order_position in rearrange]
                par.ph_total_SL = [par.ph_total_SL[order_position] for order_position in rearrange]
                par.ph_total_SR = [par.ph_total_SR[order_position] for order_position in rearrange]
                par.ph_total_t_slot_h_type = [par.ph_total_t_slot_h_type[order_position] for order_position in rearrange]

                print('ph_total_position_y:', par.ph_total_position_y)
                print('ph_total_t_slot_h_type:', par.ph_total_t_slot_h_type)
                print('ph_total_LL:', par.ph_total_LL)
                print('ph_total_LR:', par.ph_total_LR)
                print('ph_total_SL:', par.ph_total_SL)
                print('ph_total_SR:', par.ph_total_SR)
                for position, item in enumerate(par.ph_total_position_y):
                    # 將整數轉換為字串，然後設定為表格的項目文本
                    item_text = str(item)
                    table_item = QtWidgets.QTableWidgetItem(item_text)
                    self.ui.t_slot_table_h.setItem(position + 1, 1, table_item)
                for position, item in enumerate(par.ph_total_LL):
                    item = QtWidgets.QTableWidgetItem(item)
                    self.ui.t_slot_table_h.setItem(position + 1, 3, item)
                for position, item in enumerate(par.ph_total_LR):
                    item = QtWidgets.QTableWidgetItem(item)
                    self.ui.t_slot_table_h.setItem(position + 1, 4, item)
                for position, item in enumerate(par.ph_total_SL):
                    item = QtWidgets.QTableWidgetItem(item)
                    self.ui.t_slot_table_h.setItem(position + 1, 5, item)
                for position, item in enumerate(par.ph_total_SR):
                    item = QtWidgets.QTableWidgetItem(item)
                    self.ui.t_slot_table_h.setItem(position + 1, 6, item)
                for position, item in enumerate(par.ph_total_t_slot_h_type):
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
                    par.ph_total_position_x.append(position_x)
                    par.ph_total_t_slot_v_type.append(t_slot_type)
                    par.ph_total_LF.append(LF)
                    par.ph_total_LB.append(LB)
                    par.ph_total_SF.append(SF)
                    par.ph_total_SB.append(SB)
                par.ph_total_position_x = [int(x) for x in par.ph_total_position_x]
                position_x_change_position = sorted(enumerate(par.ph_total_position_x), key=lambda x: x[1], reverse=True)
                print('position_x_change_position:', position_x_change_position)
                rearrange = [position[0] for position in position_x_change_position]
                print('rearrange:', rearrange)
                par.ph_total_position_x = [par.ph_total_position_x[order_position] for order_position in rearrange]
                par.ph_total_LF = [par.ph_total_LF[order_position] for order_position in rearrange]
                par.ph_total_LB = [par.ph_total_LB[order_position] for order_position in rearrange]
                par.ph_total_SF = [par.ph_total_SF[order_position] for order_position in rearrange]
                par.ph_total_SB = [par.ph_total_SB[order_position] for order_position in rearrange]
                par.ph_total_t_slot_v_type = [par.ph_total_t_slot_v_type[order_position] for order_position in rearrange]
                print('ph_total_position_x:', par.ph_total_position_x)
                print('ph_total_t_slot_v_type:', par.ph_total_t_slot_v_type)
                print('ph_total_LF:', par.ph_total_LF)
                print('ph_total_LB:', par.ph_total_LB)
                print('ph_total_SF:', par.ph_total_SF)
                print('ph_total_SB:', par.ph_total_SB)
                for position, item in enumerate(par.ph_total_position_x):
                    # 將整數轉換為字串，然後設定為表格的項目文本
                    item_text = str(item)
                    table_item = QtWidgets.QTableWidgetItem(item_text)
                    self.ui.t_slot_table_v.setItem(position + 1, 1, table_item)
                for position, item in enumerate(par.ph_total_LF):
                    item = QtWidgets.QTableWidgetItem(item)
                    self.ui.t_slot_table_v.setItem(position + 1, 3, item)
                for position, item in enumerate(par.ph_total_LB):
                    item = QtWidgets.QTableWidgetItem(item)
                    self.ui.t_slot_table_v.setItem(position + 1, 4, item)
                for position, item in enumerate(par.ph_total_SF):
                    item = QtWidgets.QTableWidgetItem(item)
                    self.ui.t_slot_table_v.setItem(position + 1, 5, item)
                for position, item in enumerate(par.ph_total_SB):
                    item = QtWidgets.QTableWidgetItem(item)
                    self.ui.t_slot_table_v.setItem(position + 1, 6, item)
                for position, item in enumerate(par.ph_total_t_slot_v_type):
                    combo_box = self.ui.t_slot_table_v.cellWidget(position + 1, 2)
                    combo_box.setCurrentText(item)
                    self.combo_box_changed_v(position + 1, combo_box.currentIndex())

                # 檢查T溝是否超出界線
                # T溝尺寸與平板邊緣>=50mm
                if len(par.ph_total_position_y) != 0:
                    if par.ph_total_position_y[0] + (par.t_all_dimension[1] / 2) > (width / 2) - 50:
                        self.show_alert('橫向T溝位置(Y) H1 與平板邊緣過近')
                    if par.ph_total_position_y[-1] - (par.t_all_dimension[1] / 2) <= (-width / 2) + 50:
                        self.show_alert('橫向T溝位置(Y) H' + str(len(par.ph_total_position_y)) + ' 與平板邊緣過近')

                    # T溝與T溝之間間距需>=50mm
                    for turn in range(len(par.ph_total_position_y) - 1):
                        third_value = abs(par.ph_total_position_y[turn] - par.ph_total_position_y[turn + 1]) - \
                                      par.t_all_dimension[1]
                        if third_value <= 50:
                            self.show_alert('橫向T溝位置(Y) H' + str(turn + 1) + ' 與 H' + str(turn + 2) + ' 間距過近')
                        print('third_value:', third_value)

                if len(par.ph_total_position_x) != 0:
                    if par.ph_total_position_x[0] + (par.t_all_dimension[0] / 2) > (length / 2) - 50:
                        self.show_alert('縱向T溝位置(X) V1 與平板邊緣過近')
                    if par.ph_total_position_x[-1] - (par.t_all_dimension[0] / 2) <= (-length / 2) + 50:
                        self.show_alert('縱向T溝位置(X) V' + str(len(par.ph_total_position_x)) + ' 與平板邊緣過近')

                    # T溝與T溝之間間距需>=50mm
                    for turn in range(len(par.ph_total_position_x) - 1):
                        third_value = abs(par.ph_total_position_x[turn] - par.ph_total_position_x[turn + 1]) - \
                                      par.t_all_dimension[0]
                        if third_value <= 50:
                            self.show_alert('縱向T溝位置(X) V' + str(turn + 1) + ' 與 V' + str(turn + 2) + ' 間距過近')
                        print('third_value:', third_value)
        self.restore_setup_button()


    def show_alert(self, alert):
        QMessageBox.about(self, "警告", alert)

    def reset(self, parent_page):
        self.ui.t_slot_h_number.clear()
        self.ui.t_slot_v_number.clear()
        if parent_page == 'pad_secend_windows':
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
        elif parent_page == 'punch_secend_windows':
            par.ph_total_position_y.clear()
            par.ph_total_t_slot_h_type.clear()
            par.ph_total_LL.clear()
            par.ph_total_LR.clear()
            par.ph_total_SL.clear()
            par.ph_total_SR.clear()
            par.ph_total_position_x.clear()
            par.ph_total_t_slot_v_type.clear()
            par.ph_total_LF.clear()
            par.ph_total_LB.clear()
            par.ph_total_SF.clear()
            par.ph_total_SB.clear()


# 下料孔設定介面
class cutout_hole_machining(QtWidgets.QWidget):
    def __init__(self, stamping_press_type):
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
        self.ui.setup.clicked.connect(lambda: self.setup(stamping_press_type))
        self.ui.esc.clicked.connect(lambda: self.esc(stamping_press_type))

    def setup(self, stamping_press_type):
        par.cutout_hole_machining_X = self.ui.X.text()
        par.cutout_hole_machining_Y = self.ui.Y.text()
        if par.cutout_hole_machining_X == '':
            par.cutout_hole_machining_X = 0
        if par.cutout_hole_machining_Y == '':
            par.cutout_hole_machining_Y = 0
        par.cutout_hole_machining_X = int(par.cutout_hole_machining_X)
        par.cutout_hole_machining_Y = int(par.cutout_hole_machining_Y)

        excel_cutout_limit = epc.ExcelOp('平板', '下料孔界線')
        excel_cutout_limit.get_cutout_limit(stamping_press_type)
        print(par.cutout_hole_machining_X, par.cutout_hole_machining_Y, par.cutout_all_limit['X'],
              par.cutout_all_limit['Y'])
        if par.plate_hole_type[0] == '圓孔':
            if abs(par.cutout_hole_machining_X) + int(par.cutout_part_dimension[0]) / 2 + 10 >= par.cutout_all_limit[
                'X'] / 2 or abs(par.cutout_hole_machining_Y) + int(par.cutout_part_dimension[0]) / 2 + 10 >= \
                    par.cutout_all_limit['Y'] / 2:
                print('error')
            else:
                self.esc(stamping_press_type)
        elif par.plate_hole_type[0] == '方孔':
            if abs(par.cutout_hole_machining_X) + int(par.cutout_part_dimension[0]) / 2 + 10 >= par.cutout_all_limit[
                'X'] / 2 or abs(par.cutout_hole_machining_Y) + int(par.cutout_part_dimension[0]) / 2 + 10 >= \
                    par.cutout_all_limit['Y'] / 2:
                print('error')
            else:
                self.esc(stamping_press_type)
        elif par.plate_hole_type[0] == '漏斗型':
            if abs(par.cutout_hole_machining_X) + int(par.cutout_part_dimension[0]) / 2 + 10 >= par.cutout_all_limit[
                'X'] / 2 or abs(par.cutout_hole_machining_Y) + int(par.cutout_part_dimension[0]) / 2 + 10 >= \
                    par.cutout_all_limit['Y'] / 2:
                print('error')
            else:
                self.esc(stamping_press_type)
        else:
            self.esc(stamping_press_type)

    def esc(self, stamping_press_type):
        self.hide()
        self.nw = pad_secend_windows(stamping_press_type)
        self.nw.show()


# class remove_machining(QtWidgets.QWidget):
#     def __init__(self, stamping_press_type):
#         super().__init__()
#         self.ui = pad_feeding_hole_Form()
#         self.ui.setupUi(self)
#         self.setWindowTitle('下料口加工設定')
#         self.ui.escape.clicked.connect(self.show_pad_secend_windows)
#         self.ui.setup.clicked.connect(lambda :self.setup(stamping_press_type))
#         self.ui.reset.clicked.connect(self.clean_data)
#
#     def setup(self, stamping_press_type):
#         X = self.ui.X.text()
#         Y = self.ui.Y.text()
#         par.feeding_hole_position = [X, Y]
#         print("下料口位置:", par.feeding_hole_position)
#         self.hide()
#         self.nw = pad_secend_windows(stamping_press_type)
#         self.nw.show()
#
#     def clean_data(self):
#         self.ui.X.clear()
#         self.ui.Y.clear()
#
#     def show_pad_secend_windows(self, stamping_press_type):
#         self.hide()
#         self.nw = pad_secend_windows(stamping_press_type)
#         self.nw.show()


class FolderManager:
    _instance = None

    def __new__(cls, create_name):
        if cls._instance is None:
            cls._instance = super(FolderManager, cls).__new__(cls)
            cls._instance.init_folders(create_name)
        return cls._instance

    def init_folders(self, create_name):
        time_now = datetime.datetime.now()
        dir_name = '{}_{}_{}_{}_{}'.format(create_name, time_now.day, time_now.hour, time_now.minute, time_now.second)
        desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
        path = desktop + '\\' + dir_name
        os.mkdir(path)
        self.path = path

    # 添加其他管理資料夾的方法，如machining、welding等


if __name__ == "__main__":
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)  # 自適應屏幕分辨率
    app = QtWidgets.QApplication([])
    window = main()
    window.show()
    sys.exit(app.exec_())
