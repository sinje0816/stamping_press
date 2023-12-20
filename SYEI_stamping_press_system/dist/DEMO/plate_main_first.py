# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'plate_main_first.ui'
#
# Created by: PyQt5 UI code generator 5.15.4
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(530, 55)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Maximum, QtWidgets.QSizePolicy.Maximum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(Form.sizePolicy().hasHeightForWidth())
        Form.setSizePolicy(sizePolicy)
        Form.setMinimumSize(QtCore.QSize(530, 55))
        Form.setMaximumSize(QtCore.QSize(530, 55))
        self.plate_start = QtWidgets.QPushButton(Form)
        self.plate_start.setGeometry(QtCore.QRect(310, 30, 219, 25))
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(12)
        self.plate_start.setFont(font)
        self.plate_start.setObjectName("plate_start")
        self.layoutWidget = QtWidgets.QWidget(Form)
        self.layoutWidget.setGeometry(QtCore.QRect(0, 0, 531, 25))
        self.layoutWidget.setObjectName("layoutWidget")
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.layoutWidget)
        self.horizontalLayout.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.label_13 = QtWidgets.QLabel(self.layoutWidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_13.sizePolicy().hasHeightForWidth())
        self.label_13.setSizePolicy(sizePolicy)
        self.label_13.setMinimumSize(QtCore.QSize(300, 0))
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(12)
        self.label_13.setFont(font)
        self.label_13.setAlignment(QtCore.Qt.AlignCenter)
        self.label_13.setObjectName("label_13")
        self.horizontalLayout.addWidget(self.label_13)
        self.pad_select = QtWidgets.QComboBox(self.layoutWidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.pad_select.sizePolicy().hasHeightForWidth())
        self.pad_select.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(12)
        self.pad_select.setFont(font)
        self.pad_select.setObjectName("pad_select")
        self.pad_select.addItem("")
        self.pad_select.addItem("")
        self.pad_select.addItem("")
        self.pad_select.addItem("")
        self.pad_select.addItem("")
        self.pad_select.addItem("")
        self.pad_select.addItem("")
        self.pad_select.addItem("")
        self.pad_select.addItem("")
        self.pad_select.addItem("")
        self.pad_select.addItem("")
        self.horizontalLayout.addWidget(self.pad_select)

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Form"))
        self.plate_start.setText(_translate("Form", "開始生成"))
        self.label_13.setText(_translate("Form", "平板選項"))
        self.pad_select.setItemText(0, _translate("Form", "標準圓孔"))
        self.pad_select.setItemText(1, _translate("Form", "標準無孔"))
        self.pad_select.setItemText(2, _translate("Form", "標準方孔"))
        self.pad_select.setItemText(3, _translate("Form", "標準模墊型"))
        self.pad_select.setItemText(4, _translate("Form", "標準加大I型(無孔)"))
        self.pad_select.setItemText(5, _translate("Form", "標準加大I型(圓孔)"))
        self.pad_select.setItemText(6, _translate("Form", "標準加大I型(方孔)"))
        self.pad_select.setItemText(7, _translate("Form", "標準加大II型(無孔)"))
        self.pad_select.setItemText(8, _translate("Form", "標準加大II型(圓孔)"))
        self.pad_select.setItemText(9, _translate("Form", "標準加大II型(方孔)"))
        self.pad_select.setItemText(10, _translate("Form", "特殊平板"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Form = QtWidgets.QWidget()
    ui = Ui_Form()
    ui.setupUi(Form)
    Form.show()
    sys.exit(app.exec_())