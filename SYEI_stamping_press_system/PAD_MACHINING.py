# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'pad_machining.ui'
#
# Created by: PyQt5 UI code generator 5.15.4
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(1116, 507)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(Form.sizePolicy().hasHeightForWidth())
        Form.setSizePolicy(sizePolicy)
        self.tableWidget = QtWidgets.QTableWidget(Form)
        self.tableWidget.setGeometry(QtCore.QRect(560, 20, 501, 391))
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(12)
        self.tableWidget.setFont(font)
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.setColumnCount(5)
        self.tableWidget.setRowCount(0)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(3, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(4, item)
        self.layoutWidget = QtWidgets.QWidget(Form)
        self.layoutWidget.setGeometry(QtCore.QRect(40, 330, 441, 111))
        self.layoutWidget.setObjectName("layoutWidget")
        self.gridLayout_2 = QtWidgets.QGridLayout(self.layoutWidget)
        self.gridLayout_2.setContentsMargins(0, 0, 0, 0)
        self.gridLayout_2.setObjectName("gridLayout_2")
        self.new_order = QtWidgets.QPushButton(self.layoutWidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Maximum, QtWidgets.QSizePolicy.Maximum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.new_order.sizePolicy().hasHeightForWidth())
        self.new_order.setSizePolicy(sizePolicy)
        self.new_order.setMaximumSize(QtCore.QSize(200, 16777215))
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(18)
        self.new_order.setFont(font)
        self.new_order.setObjectName("new_order")
        self.gridLayout_2.addWidget(self.new_order, 0, 0, 1, 1)
        self.setup = QtWidgets.QPushButton(self.layoutWidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.setup.sizePolicy().hasHeightForWidth())
        self.setup.setSizePolicy(sizePolicy)
        self.setup.setMaximumSize(QtCore.QSize(200, 16777215))
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(18)
        self.setup.setFont(font)
        self.setup.setObjectName("setup")
        self.gridLayout_2.addWidget(self.setup, 0, 1, 1, 1)
        self.reset = QtWidgets.QPushButton(self.layoutWidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Maximum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.reset.sizePolicy().hasHeightForWidth())
        self.reset.setSizePolicy(sizePolicy)
        self.reset.setMaximumSize(QtCore.QSize(200, 16777215))
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(18)
        self.reset.setFont(font)
        self.reset.setObjectName("reset")
        self.gridLayout_2.addWidget(self.reset, 1, 0, 1, 1)
        self.escape = QtWidgets.QPushButton(self.layoutWidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.escape.sizePolicy().hasHeightForWidth())
        self.escape.setSizePolicy(sizePolicy)
        self.escape.setMaximumSize(QtCore.QSize(200, 16777215))
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(18)
        self.escape.setFont(font)
        self.escape.setObjectName("escape")
        self.gridLayout_2.addWidget(self.escape, 1, 1, 1, 1)
        self.widget = QtWidgets.QWidget(Form)
        self.widget.setGeometry(QtCore.QRect(40, 70, 491, 214))
        self.widget.setObjectName("widget")
        self.gridLayout = QtWidgets.QGridLayout(self.widget)
        self.gridLayout.setContentsMargins(0, 0, 0, 0)
        self.gridLayout.setObjectName("gridLayout")
        self.label_23 = QtWidgets.QLabel(self.widget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Maximum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_23.sizePolicy().hasHeightForWidth())
        self.label_23.setSizePolicy(sizePolicy)
        self.label_23.setMaximumSize(QtCore.QSize(150, 40))
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(18)
        self.label_23.setFont(font)
        self.label_23.setAlignment(QtCore.Qt.AlignLeading | QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
        self.label_23.setObjectName("label_23")
        self.gridLayout.addWidget(self.label_23, 1, 0, 1, 1)
        self.direction = QtWidgets.QComboBox(self.widget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.direction.sizePolicy().hasHeightForWidth())
        self.direction.setSizePolicy(sizePolicy)
        self.direction.setMaximumSize(QtCore.QSize(250, 40))
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(18)
        self.direction.setFont(font)
        self.direction.setContextMenuPolicy(QtCore.Qt.ActionsContextMenu)
        self.direction.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.direction.setMaxVisibleItems(10)
        self.direction.setObjectName("direction")
        self.direction.addItem("")
        self.direction.addItem("")
        self.gridLayout.addWidget(self.direction, 1, 1, 1, 1)
        self.label_25 = QtWidgets.QLabel(self.widget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Maximum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_25.sizePolicy().hasHeightForWidth())
        self.label_25.setSizePolicy(sizePolicy)
        self.label_25.setMaximumSize(QtCore.QSize(150, 40))
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(18)
        self.label_25.setFont(font)
        self.label_25.setAlignment(QtCore.Qt.AlignLeading | QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
        self.label_25.setObjectName("label_25")
        self.gridLayout.addWidget(self.label_25, 2, 0, 1, 1)
        self.pierce = QtWidgets.QComboBox(self.widget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.pierce.sizePolicy().hasHeightForWidth())
        self.pierce.setSizePolicy(sizePolicy)
        self.pierce.setMaximumSize(QtCore.QSize(250, 40))
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(18)
        self.pierce.setFont(font)
        self.pierce.setContextMenuPolicy(QtCore.Qt.ActionsContextMenu)
        self.pierce.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.pierce.setMaxVisibleItems(10)
        self.pierce.setObjectName("pierce")
        self.pierce.addItem("")
        self.pierce.addItem("")
        self.gridLayout.addWidget(self.pierce, 2, 1, 1, 1)
        self.label_27 = QtWidgets.QLabel(self.widget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Maximum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_27.sizePolicy().hasHeightForWidth())
        self.label_27.setSizePolicy(sizePolicy)
        self.label_27.setMaximumSize(QtCore.QSize(150, 40))
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(18)
        self.label_27.setFont(font)
        self.label_27.setAlignment(QtCore.Qt.AlignLeading | QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
        self.label_27.setObjectName("label_27")
        self.gridLayout.addWidget(self.label_27, 3, 0, 1, 1)
        self.clearance_hole = QtWidgets.QComboBox(self.widget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.clearance_hole.sizePolicy().hasHeightForWidth())
        self.clearance_hole.setSizePolicy(sizePolicy)
        self.clearance_hole.setMaximumSize(QtCore.QSize(250, 40))
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(18)
        self.clearance_hole.setFont(font)
        self.clearance_hole.setContextMenuPolicy(QtCore.Qt.ActionsContextMenu)
        self.clearance_hole.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.clearance_hole.setMaxVisibleItems(10)
        self.clearance_hole.setObjectName("clearance_hole")
        self.clearance_hole.addItem("")
        self.clearance_hole.addItem("")
        self.gridLayout.addWidget(self.clearance_hole, 3, 1, 1, 1)
        self.label_28 = QtWidgets.QLabel(self.widget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Maximum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_28.sizePolicy().hasHeightForWidth())
        self.label_28.setSizePolicy(sizePolicy)
        self.label_28.setMaximumSize(QtCore.QSize(150, 40))
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(18)
        self.label_28.setFont(font)
        self.label_28.setAlignment(QtCore.Qt.AlignLeading | QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
        self.label_28.setObjectName("label_28")
        self.gridLayout.addWidget(self.label_28, 4, 0, 1, 1)
        self.position_dimension = QtWidgets.QLineEdit(self.widget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.position_dimension.sizePolicy().hasHeightForWidth())
        self.position_dimension.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(18)
        self.position_dimension.setFont(font)
        self.position_dimension.setObjectName("position_dimension")
        self.gridLayout.addWidget(self.position_dimension, 4, 1, 1, 1)
        self.label_29 = QtWidgets.QLabel(self.widget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Maximum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_29.sizePolicy().hasHeightForWidth())
        self.label_29.setSizePolicy(sizePolicy)
        self.label_29.setMaximumSize(QtCore.QSize(150, 40))
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(18)
        self.label_29.setFont(font)
        self.label_29.setAlignment(QtCore.Qt.AlignLeading | QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
        self.label_29.setObjectName("label_29")
        self.gridLayout.addWidget(self.label_29, 5, 0, 1, 1)
        self.unpierce = QtWidgets.QLineEdit(self.widget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.unpierce.sizePolicy().hasHeightForWidth())
        self.unpierce.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(18)
        self.unpierce.setFont(font)
        self.unpierce.setObjectName("unpierce")
        self.gridLayout.addWidget(self.unpierce, 5, 1, 1, 1)
        self.label = QtWidgets.QLabel(self.widget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label.sizePolicy().hasHeightForWidth())
        self.label.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(18)
        font.setBold(True)
        font.setWeight(75)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.gridLayout.addWidget(self.label, 0, 0, 1, 3)
        self.picture = QtWidgets.QLabel(self.widget)
        self.picture.setText("")
        self.picture.setObjectName("picture")
        self.gridLayout.addWidget(self.picture, 1, 2, 5, 1)

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Form"))
        item = self.tableWidget.horizontalHeaderItem(0)
        item.setText(_translate("Form", "方向"))
        item = self.tableWidget.horizontalHeaderItem(1)
        item.setText(_translate("Form", "是否貫穿"))
        item = self.tableWidget.horizontalHeaderItem(2)
        item.setText(_translate("Form", "是否讓孔"))
        item = self.tableWidget.horizontalHeaderItem(3)
        item.setText(_translate("Form", "偏移距離"))
        item = self.tableWidget.horizontalHeaderItem(4)
        item.setText(_translate("Form", "非貫穿長度"))
        self.new_order.setText(_translate("Form", "新增一筆資料"))
        self.setup.setText(_translate("Form", "設定參數"))
        self.reset.setText(_translate("Form", "重設參數"))
        self.escape.setText(_translate("Form", "離開"))
        self.label_23.setText(_translate("Form", "1.方向"))
        self.direction.setItemText(0, _translate("Form", "前後"))
        self.direction.setItemText(1, _translate("Form", "左右"))
        self.label_25.setText(_translate("Form", "2.是否貫穿"))
        self.pierce.setItemText(0, _translate("Form", "是"))
        self.pierce.setItemText(1, _translate("Form", "否"))
        self.label_27.setText(_translate("Form", "3.是否讓孔"))
        self.clearance_hole.setItemText(0, _translate("Form", "是"))
        self.clearance_hole.setItemText(1, _translate("Form", "否"))
        self.label_28.setText(_translate("Form", "4.位置尺寸"))
        self.label_29.setText(_translate("Form", "5.非貫穿長度"))
        self.label.setText(_translate("Form", "b.加工設定"))


if __name__ == "__main__":
    import sys

    app = QtWidgets.QApplication(sys.argv)
    Form = QtWidgets.QWidget()
    ui = Ui_Form()
    ui.setupUi(Form)
    Form.show()
    sys.exit(app.exec_())
