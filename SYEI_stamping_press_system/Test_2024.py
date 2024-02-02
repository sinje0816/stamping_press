from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QVBoxLayout, QWidget, QPushButton, QLabel, QTableWidgetItem, QHeaderView, QComboBox, QAbstractItemView, QMessageBox, QHBoxLayout, QLineEdit, QCheckBox
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor, QBrush, QPixmap, QFont
import sys
from T_test import Ui_Form as T_test_form

type = 0
default_posision_Y = {0: ['50', '0', '-50'], 1: ['75', '0', '-75'], 2: ['100', '0', '-100']}
total_position_y = []
total_t_slot_h_type = []
total_LL = []
total_LR = []
total_SL = []
total_SR = []

class main(QtWidgets.QWidget):
    def __init__(self):
        super(main, self).__init__()
        self.ui = T_test_form()
        self.ui.setupUi(self)
        self.setWindowTitle('測試')
        self.increasing = False
        self.decreasing = False
        self.item_setting = False
        self.item_rearrange = False
        self.error = False
        self.table_setup()
        self.ui.AddNewRow.clicked.connect(self.add_new_row)
        # self.ui.T_solt.itemChanged.connect(self.table_item_change) # 當表格改變時相關功能(暫無作用)
        self.default() # 給予預設值
        self.ui.rearrange.clicked.connect(self.rearrange)

    def default(self):
        self.item_setting = True  # 標籤:正在設定單位格
        if type == 0:
            for x in range(0, len(default_posision_Y[type])):
                self.add_new_row()
                item = QTableWidgetItem(default_posision_Y[type][x])
                self.ui.T_solt.setItem(x+1, 1, item)
        else:
            pass
        self.item_setting = False  # 標籤:完成設定單位格

    def table_setup(self):
        self.ui.T_solt.verticalHeader().setVisible(False)
        self.ui.T_solt.horizontalHeader().setVisible(False)
        self.ui.T_solt.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.ui.T_solt.setEditTriggers(QAbstractItemView.AllEditTriggers)
        # 第一行
        for col in range(8):  # 假设有8列
            newItem = QTableWidgetItem(
                "編號" if col == 0 else "位置(Y)" if col == 1 else "形式" if col == 2 else "LL" if col == 3 else "LR" if col == 4 else "SL" if col == 5 else "SR" if col == 6 else "")
            newItem.setTextAlignment(Qt.AlignCenter)  # 可以根据需要设置文本对齐方式
            self.ui.T_solt.setItem(0, col, newItem)
            newItem.setFlags(newItem.flags() & ~Qt.ItemIsEditable)  # 設定為不可編輯

    def add_new_row(self):
        self.increasing = True # 標籤:正在新增行列
        new_row = self.ui.T_solt.rowCount()
        col_count = self.ui.T_solt.columnCount()
        self.ui.T_solt.insertRow(new_row)
        for col in range(col_count):
            item = QTableWidgetItem("H" + str(new_row) if (new_row == 0 or col == 0) else "")
            if col == 0 or new_row == 0:
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            else:
                # 其他单元格默认可编辑，会在后续根据条件进行修改
                item.setFlags(item.flags() | Qt.ItemIsEditable)
            self.ui.T_solt.setItem(new_row, col, item)
        # 创建 QComboBox 并添加到新行的第三列
        combo_box = QComboBox()
        combo_box.addItem("")
        combo_box.addItem("貫穿")
        combo_box.addItem("分段")
        combo_box.currentIndexChanged.connect(lambda index, row=new_row: self.combo_box_changed(row, index))
        self.ui.T_solt.setCellWidget(new_row, 2, combo_box)
        # 初始化选项状态为不可编辑
        for col in range(3, 5):
            item = QTableWidgetItem("")
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            color = QColor(192, 192, 192)
            brush = QBrush(color)
            item.setBackground(brush)
            self.ui.T_solt.setItem(new_row, col, item)
        # 加入刪除按鈕
        delete_button = QPushButton("刪除")
        delete_button.clicked.connect(self.delete_current_row)
        self.ui.T_solt.setCellWidget(new_row, 7, delete_button)
        self.increasing = False # 標籤:完成新增行列

    def combo_box_changed(self, row, index):
        if self.item_rearrange:
            current_row = row
            if index == 2:  # 如果选择了"分段"
                self.set_editable_cells(current_row, is_editable=True)
            else:
                self.set_editable_cells(current_row, is_editable=False)
        else:
            current_row = self.ui.T_solt.currentRow()
            if index == 2:  # 如果选择了"分段"
                self.set_editable_cells(current_row, is_editable=True)
            else:
                self.set_editable_cells(current_row, is_editable=False)

    def set_editable_cells(self, row, is_editable=True):
        self.item_setting = True # 標籤:正在設定單位格
        for col in range(3, 5):
            # 清空單位格內文
            if self.item_rearrange:
                pass
            else:
                context = QTableWidgetItem("")
                self.ui.T_solt.setItem(row, col, context)
            # 調整編輯權限
            item = self.ui.T_solt.item(row, col)
            if item:
                if is_editable and row != 0:
                    item.setFlags(item.flags() | Qt.ItemIsEditable)
                    item.setBackground(QBrush(QColor(255, 255, 255)))
                else:
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    item.setBackground(QBrush(QColor(192, 192, 192)))
        self.item_setting = False # 標籤:完成設定單位格

    def delete_current_row(self):
        self.decreasing = True # 標籤:正在刪除行列
        current_row = self.ui.T_solt.currentRow()
        self.ui.T_solt.removeRow(current_row)
        self.rename_T_number()
        self.decreasing = False # 標籤:完成刪除行列

    # 重新命名T溝編號
    def rename_T_number(self):
        self.item_setting = True # 標籤:正在設定單位格
        total_row = self.ui.T_solt.rowCount()
        for row in range(0, total_row):
            if row == 0:
                pass
            else:
                item = QTableWidgetItem("H" + str(row))
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.ui.T_solt.setItem(row, 0, item)
        self.item_setting = False # 標籤:完成設定單位格

    def rearrange(self):
        self.item_rearrange = True  # 標籤:正在重新排列單位格
        total_row = self.ui.T_solt.rowCount()
        global total_position_y, total_t_slot_h_type, total_LL, total_LR, total_SL, total_SR
        total_position_y.clear()
        total_t_slot_h_type.clear()
        total_LL.clear()
        total_LR.clear()
        total_SL.clear()
        total_SR.clear()
        for row in range(1, total_row):
            position_y = self.ui.T_solt.item(row, 1).text()
            t_slot_h_type = self.ui.T_solt.cellWidget(row, 2).currentText()
            LL = self.ui.T_solt.item(row, 3).text()
            LR = self.ui.T_solt.item(row, 4).text()
            SL = self.ui.T_solt.item(row, 5).text()
            SR = self.ui.T_solt.item(row, 6).text()
            total_position_y.append(position_y)
            total_t_slot_h_type.append(t_slot_h_type)
            total_LL.append(LL)
            total_LR.append(LR)
            total_SL.append(SL)
            total_SR.append(SR)
        for x in range(0, len(total_position_y)):
            if total_position_y[x] == '':
                self.show_alert('T溝編號H'+str(x+1)+'位置(Y)未輸入')
                self.error = True
            else:
                self.error = False
        for x in range(0, len(total_position_y)):
            for y in range(x+1, len(total_position_y)):
                if total_position_y[x] == total_position_y[y]:
                    self.show_alert('T溝編號H'+str(x+1)+'與T溝編號H'+str(y+1)+'位置(Y)重複')
                    self.error = True
                else:
                    self.error = False
        if self.error:
            pass
        else:
            total_position_y = [int(x) for x in total_position_y]
            position_y_change_position = sorted(enumerate(total_position_y), key=lambda x: x[1], reverse=True)
            rearrange = [position[0] for position in position_y_change_position]
            total_position_y = [total_position_y[order_position] for order_position in rearrange]
            total_LL = [total_LL[order_position] for order_position in rearrange]
            total_LR = [total_LR[order_position] for order_position in rearrange]
            total_SL = [total_SL[order_position] for order_position in rearrange]
            total_SR = [total_SR[order_position] for order_position in rearrange]
            total_t_slot_h_type = [total_t_slot_h_type[order_position] for order_position in rearrange]
            for position, item in enumerate(total_position_y):
                # 將整數轉換為字串，然後設定為表格的項目文本
                item_text = str(item)
                table_item = QtWidgets.QTableWidgetItem(item_text)
                self.ui.T_solt.setItem(position + 1, 1, table_item)
            for position, item in enumerate(total_LL):
                item = QtWidgets.QTableWidgetItem(item)
                self.ui.T_solt.setItem(position + 1, 3, item)
            for position, item in enumerate(total_LR):
                item = QtWidgets.QTableWidgetItem(item)
                self.ui.T_solt.setItem(position + 1, 4, item)
            for position, item in enumerate(total_SL):
                item = QtWidgets.QTableWidgetItem(item)
                self.ui.T_solt.setItem(position + 1, 5, item)
            for position, item in enumerate(total_SR):
                item = QtWidgets.QTableWidgetItem(item)
                self.ui.T_solt.setItem(position + 1, 6, item)
            for position, item in enumerate(total_t_slot_h_type):
                combo_box = self.ui.T_solt.cellWidget(position + 1, 2)  # 從表格中獲取 ComboBox
                combo_box.setCurrentText(item)
                self.combo_box_changed(position+1, combo_box.currentIndex())
        self.item_rearrange = False  # 標籤:完成重新排列單位格

    # 當表格改變時相關功能(暫無作用)
    def table_item_change(self, item):
        change_row = item.row()
        change_col = item.column()
        current_context = item.text()
        if self.increasing:
            pass
        elif self.decreasing:
            pass
        elif self.item_setting:
            pass
        elif self.item_rearrange:
            if type == 0:
                for row in range(0, len(default_posision_Y[type])):
                    if change_row == row+1:
                        if change_col ==1:
                            if current_context != default_posision_Y[type][row]:
                                self.change_context_color(change_row, change_col, current_context, turn_red=True)
                            else:
                                self.change_context_color(change_row, change_col, current_context, turn_red=False)
                        else:
                            pass
                    else:
                        pass
            else:
                pass
        else:
            print('change('+str(change_row)+','+str(change_col)+")")
            # 如果修改預設值，要換色
            if type == 0:
                for row in range(0, len(default_posision_Y[type])):
                    if change_row == row+1:
                        if change_col ==1:
                            if current_context != default_posision_Y[type][row]:
                                self.change_context_color(change_row, change_col, current_context, turn_red=True)
                            else:
                                self.change_context_color(change_row, change_col, current_context, turn_red=False)
                        else:
                            pass
                    else:
                        pass
            else:
                pass

    def change_context_color(self, change_row, change_col, current_context, turn_red=False):
        self.item_setting = True  # 標籤:正在設定單位格
        if turn_red:
            item = QTableWidgetItem(current_context)
            item.setForeground(QBrush(QColor(255, 0, 0)))
            self.ui.T_solt.setItem(change_row, change_col, item)
        else:
            item = QTableWidgetItem(current_context)
            item.setForeground(QBrush(QColor(0, 0, 0)))
            self.ui.T_solt.setItem(change_row, change_col, item)
        self.item_setting = False  # 標籤:完成設定單位格

    def show_alert(self, alert):
        QMessageBox.about(self, "警告", alert)

if __name__ == "__main__":
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)  # 自適應屏幕分辨率
    app = QtWidgets.QApplication([])
    window = main()
    window.show()
    sys.exit(app.exec_())