# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'slide_main_simple.ui'
#
# Created by: PyQt5 UI code generator 5.15.4
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(665, 85)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Maximum, QtWidgets.QSizePolicy.Maximum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(Form.sizePolicy().hasHeightForWidth())
        Form.setSizePolicy(sizePolicy)
        Form.setMinimumSize(QtCore.QSize(665, 85))
        Form.setMaximumSize(QtCore.QSize(665, 85))
        self.layoutWidget = QtWidgets.QWidget(Form)
        self.layoutWidget.setGeometry(QtCore.QRect(0, 0, 664, 82))
        self.layoutWidget.setObjectName("layoutWidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.layoutWidget)
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout.setObjectName("verticalLayout")
        self.gridLayout_2 = QtWidgets.QGridLayout()
        self.gridLayout_2.setObjectName("gridLayout_2")
        self.label_13 = QtWidgets.QLabel(self.layoutWidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
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
        self.gridLayout_2.addWidget(self.label_13, 0, 0, 1, 1)
        self.slide_select = QtWidgets.QComboBox(self.layoutWidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.slide_select.sizePolicy().hasHeightForWidth())
        self.slide_select.setSizePolicy(sizePolicy)
        self.slide_select.setMinimumSize(QtCore.QSize(175, 0))
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(12)
        self.slide_select.setFont(font)
        self.slide_select.setObjectName("slide_select")
        self.slide_select.addItem("")
        self.slide_select.addItem("")
        self.slide_select.addItem("")
        self.slide_select.addItem("")
        self.gridLayout_2.addWidget(self.slide_select, 0, 2, 1, 2)
        self.slide_mold_locking = QtWidgets.QComboBox(self.layoutWidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.slide_mold_locking.sizePolicy().hasHeightForWidth())
        self.slide_mold_locking.setSizePolicy(sizePolicy)
        self.slide_mold_locking.setMinimumSize(QtCore.QSize(175, 0))
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(12)
        self.slide_mold_locking.setFont(font)
        self.slide_mold_locking.setObjectName("slide_mold_locking")
        self.slide_mold_locking.addItem("")
        self.slide_mold_locking.addItem("")
        self.gridLayout_2.addWidget(self.slide_mold_locking, 0, 1, 1, 1)
        self.verticalLayout.addLayout(self.gridLayout_2)
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.slide_create = QtWidgets.QPushButton(self.layoutWidget)
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(12)
        self.slide_create.setFont(font)
        self.slide_create.setObjectName("slide_create")
        self.horizontalLayout.addWidget(self.slide_create)
        self.slide_save = QtWidgets.QPushButton(self.layoutWidget)
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(12)
        self.slide_save.setFont(font)
        self.slide_save.setObjectName("slide_save")
        self.horizontalLayout.addWidget(self.slide_save)
        self.slide_finish = QtWidgets.QPushButton(self.layoutWidget)
        font = QtGui.QFont()
        font.setFamily("微軟正黑體")
        font.setPointSize(12)
        self.slide_finish.setFont(font)
        self.slide_finish.setObjectName("slide_finish")
        self.horizontalLayout.addWidget(self.slide_finish)
        self.verticalLayout.addLayout(self.horizontalLayout)

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Form"))
        self.label_13.setText(_translate("Form", "選項"))
        self.slide_select.setItemText(0, _translate("Form", "標準(330×250)"))
        self.slide_select.setItemText(1, _translate("Form", "標準加大I型(380×250)"))
        self.slide_select.setItemText(2, _translate("Form", "標準加大II型(430×250)"))
        self.slide_select.setItemText(3, _translate("Form", "特殊衝頭"))
        self.slide_mold_locking.setItemText(0, _translate("Form", "有鎖模"))
        self.slide_mold_locking.setItemText(1, _translate("Form", "無鎖模"))
        self.slide_create.setText(_translate("Form", "開始生成"))
        self.slide_save.setText(_translate("Form", "暫存按鈕"))
        self.slide_finish.setText(_translate("Form", "完成按鈕"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Form = QtWidgets.QWidget()
    ui = Ui_Form()
    ui.setupUi(Form)
    Form.show()
    sys.exit(app.exec_())