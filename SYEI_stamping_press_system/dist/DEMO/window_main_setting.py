from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget, QPushButton, QLineEdit, QLabel, \
    QTableWidget, QTableWidgetItem, QHeaderView, QSpinBox, QComboBox, QAbstractItemView
from PyQt5.QtCore import QObject, QTimer, Qt
from PyQt5.QtGui import QColor, QBrush, QPixmap
from window_main import Ui_Form
import sys
import parameter as par
from PyQt5.QtWidgets import QHBoxLayout


class main(QtWidgets.QWidget, Ui_Form):
    def __init__(self):
        super(main, self).__init__()
        self.ui = Ui_Form()
        self.ui.setupUi(self)
        self.setting()


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
        # combo_box_unit.setEditable(True) #先將選單可編輯化
        # combo_box_unit.lineEdit().setAlignment(Qt.AlignHCenter | Qt.AlignVCenter) #接著設定置中
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
        pad_setup = QPushButton('設定') #括號內為按鈕名稱
        self.ui.window_main_table.setCellWidget(6, 3, pad_setup)
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




        # plane_setup.clicked.connect()
        # punch_setup.clicked.connect()
        # select_setup.clicked.connect()
        # spare_parts_setup.clicked.connect()


if __name__ == "__main__":
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)  # 自適應屏幕分辨率
    app = QtWidgets.QApplication([])
    window = main()
    window.show()
    sys.exit(app.exec_())