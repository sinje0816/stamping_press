# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'pad_main.ui'
#
# Created by: PyQt5 UI code generator 5.15.4
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(1323, 596)
        self.layoutWidget = QtWidgets.QWidget(Form)
        self.layoutWidget.setGeometry(QtCore.QRect(50, 20, 901, 511))
        self.layoutWidget.setObjectName("layoutWidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.layoutWidget)
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout.setObjectName("verticalLayout")
        self.frame = QtWidgets.QFrame(self.layoutWidget)
        self.frame.setStyleSheet("#frame{\n"
"    border: 1px solid block；\n"
"}\n"
"")
        self.frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame.setObjectName("frame")
        self.picture = QtWidgets.QLabel(self.frame)
        self.picture.setGeometry(QtCore.QRect(480, 10, 321, 101))
        self.picture.setText("")
        self.picture.setObjectName("picture")
        self.layoutWidget1 = QtWidgets.QWidget(self.frame)
        self.layoutWidget1.setGeometry(QtCore.QRect(30, 20, 371, 105))
        self.layoutWidget1.setObjectName("layoutWidget1")
        self.gridLayout_3 = QtWidgets.QGridLayout(self.layoutWidget1)
        self.gridLayout_3.setContentsMargins(0, 0, 0, 0)
        self.gridLayout_3.setObjectName("gridLayout_3")
        self.pad_type = QtWidgets.QComboBox(self.layoutWidget1)
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(18)
        self.pad_type.setFont(font)
        self.pad_type.setObjectName("pad_type")
        self.pad_type.addItem("")
        self.pad_type.addItem("")
        self.pad_type.addItem("")
        self.gridLayout_3.addWidget(self.pad_type, 1, 1, 1, 1)
        self.label = QtWidgets.QLabel(self.layoutWidget1)
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(18)
        font.setBold(True)
        font.setWeight(75)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.gridLayout_3.addWidget(self.label, 0, 0, 1, 2)
        self.label_2 = QtWidgets.QLabel(self.layoutWidget1)
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(18)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.gridLayout_3.addWidget(self.label_2, 1, 0, 1, 1)
        self.t_machining_2 = QtWidgets.QPushButton(self.layoutWidget1)
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(18)
        self.t_machining_2.setFont(font)
        self.t_machining_2.setObjectName("t_machining_2")
        self.gridLayout_3.addWidget(self.t_machining_2, 2, 0, 1, 2)
        self.verticalLayout.addWidget(self.frame)
        self.frame_2 = QtWidgets.QFrame(self.layoutWidget)
        self.frame_2.setStyleSheet("#frame_2{\n"
"    border: 1px solid block；\n"
"}\n"
"\n"
"")
        self.frame_2.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_2.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_2.setObjectName("frame_2")
        self.layoutWidget2 = QtWidgets.QWidget(self.frame_2)
        self.layoutWidget2.setGeometry(QtCore.QRect(30, 20, 361, 81))
        self.layoutWidget2.setObjectName("layoutWidget2")
        self.gridLayout_2 = QtWidgets.QGridLayout(self.layoutWidget2)
        self.gridLayout_2.setContentsMargins(0, 0, 0, 0)
        self.gridLayout_2.setObjectName("gridLayout_2")
        self.label_5 = QtWidgets.QLabel(self.layoutWidget2)
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(18)
        font.setBold(True)
        font.setWeight(75)
        self.label_5.setFont(font)
        self.label_5.setObjectName("label_5")
        self.gridLayout_2.addWidget(self.label_5, 0, 0, 1, 2)
        self.t_dimension = QtWidgets.QPushButton(self.layoutWidget2)
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(18)
        self.t_dimension.setFont(font)
        self.t_dimension.setObjectName("t_dimension")
        self.gridLayout_2.addWidget(self.t_dimension, 1, 0, 1, 1)
        self.t_machining = QtWidgets.QPushButton(self.layoutWidget2)
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(18)
        self.t_machining.setFont(font)
        self.t_machining.setObjectName("t_machining")
        self.gridLayout_2.addWidget(self.t_machining, 1, 1, 1, 1)
        self.verticalLayout.addWidget(self.frame_2)
        self.frame_3 = QtWidgets.QFrame(self.layoutWidget)
        self.frame_3.setStyleSheet("#frame_3{\n"
"    border: 1px solid block；\n"
"}\n"
"\n"
"")
        self.frame_3.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_3.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_3.setObjectName("frame_3")
        self.layoutWidget3 = QtWidgets.QWidget(self.frame_3)
        self.layoutWidget3.setGeometry(QtCore.QRect(20, 10, 371, 139))
        self.layoutWidget3.setObjectName("layoutWidget3")
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.layoutWidget3)
        self.horizontalLayout.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.gridLayout = QtWidgets.QGridLayout()
        self.gridLayout.setObjectName("gridLayout")
        self.label_6 = QtWidgets.QLabel(self.layoutWidget3)
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(18)
        font.setBold(True)
        font.setWeight(75)
        self.label_6.setFont(font)
        self.label_6.setObjectName("label_6")
        self.gridLayout.addWidget(self.label_6, 0, 0, 1, 2)
        self.label_7 = QtWidgets.QLabel(self.layoutWidget3)
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(18)
        self.label_7.setFont(font)
        self.label_7.setObjectName("label_7")
        self.gridLayout.addWidget(self.label_7, 1, 0, 1, 1)
        self.remove_type = QtWidgets.QComboBox(self.layoutWidget3)
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(18)
        self.remove_type.setFont(font)
        self.remove_type.setObjectName("remove_type")
        self.remove_type.addItem("")
        self.remove_type.addItem("")
        self.remove_type.addItem("")
        self.remove_type.addItem("")
        self.remove_type.addItem("")
        self.gridLayout.addWidget(self.remove_type, 1, 1, 1, 1)
        self.label_8 = QtWidgets.QLabel(self.layoutWidget3)
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(18)
        self.label_8.setFont(font)
        self.label_8.setObjectName("label_8")
        self.gridLayout.addWidget(self.label_8, 2, 0, 1, 1)
        self.remove_dimension = QtWidgets.QLineEdit(self.layoutWidget3)
        font = QtGui.QFont()
        font.setPointSize(18)
        self.remove_dimension.setFont(font)
        self.remove_dimension.setObjectName("remove_dimension")
        self.gridLayout.addWidget(self.remove_dimension, 2, 1, 1, 1)
        self.remove_machining = QtWidgets.QPushButton(self.layoutWidget3)
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(18)
        self.remove_machining.setFont(font)
        self.remove_machining.setObjectName("remove_machining")
        self.gridLayout.addWidget(self.remove_machining, 3, 0, 1, 2)
        self.horizontalLayout.addLayout(self.gridLayout)
        self.picture_2 = QtWidgets.QLabel(self.layoutWidget3)
        self.picture_2.setText("")
        self.picture_2.setObjectName("picture_2")
        self.horizontalLayout.addWidget(self.picture_2)
        self.verticalLayout.addWidget(self.frame_3)
        self.layoutWidget4 = QtWidgets.QWidget(Form)
        self.layoutWidget4.setGeometry(QtCore.QRect(150, 540, 631, 27))
        self.layoutWidget4.setObjectName("layoutWidget4")
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout(self.layoutWidget4)
        self.horizontalLayout_2.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.plate_start = QtWidgets.QPushButton(self.layoutWidget4)
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(12)
        self.plate_start.setFont(font)
        self.plate_start.setObjectName("plate_start")
        self.horizontalLayout_2.addWidget(self.plate_start)
        self.plate_reset = QtWidgets.QPushButton(self.layoutWidget4)
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(12)
        self.plate_reset.setFont(font)
        self.plate_reset.setObjectName("plate_reset")
        self.horizontalLayout_2.addWidget(self.plate_reset)
        self.plate_escape = QtWidgets.QPushButton(self.layoutWidget4)
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(12)
        self.plate_escape.setFont(font)
        self.plate_escape.setObjectName("plate_escape")
        self.horizontalLayout_2.addWidget(self.plate_escape)

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Form"))
        self.pad_type.setItemText(0, _translate("Form", "標準"))
        self.pad_type.setItemText(1, _translate("Form", "加大一型"))
        self.pad_type.setItemText(2, _translate("Form", "加大二型"))
        self.label.setText(_translate("Form", "1.平板外型"))
        self.label_2.setText(_translate("Form", "a.尺寸"))
        self.t_machining_2.setText(_translate("Form", "開始生成平板"))
        self.label_5.setText(_translate("Form", "2.T形槽"))
        self.t_dimension.setText(_translate("Form", "a.外型尺寸"))
        self.t_machining.setText(_translate("Form", "b.加工設定"))
        self.label_6.setText(_translate("Form", "3.除料口"))
        self.label_7.setText(_translate("Form", "a.類型"))
        self.remove_type.setItemText(0, _translate("Form", "無孔"))
        self.remove_type.setItemText(1, _translate("Form", "圓孔"))
        self.remove_type.setItemText(2, _translate("Form", "方孔"))
        self.remove_type.setItemText(3, _translate("Form", "漏斗型"))
        self.remove_type.setItemText(4, _translate("Form", "模墊型"))
        self.label_8.setText(_translate("Form", "b.尺寸"))
        self.remove_machining.setText(_translate("Form", "c.加工設定"))
        self.plate_start.setText(_translate("Form", "開始生成"))
        self.plate_reset.setText(_translate("Form", "重設參數"))
        self.plate_escape.setText(_translate("Form", "離開"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Form = QtWidgets.QWidget()
    ui = Ui_Form()
    ui.setupUi(Form)
    Form.show()
    sys.exit(app.exec_())