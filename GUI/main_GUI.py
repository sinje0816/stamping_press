from PyQt5 import QtCore, QtGui, QtWidgets
import sys, datetime, os
from GUI import Ui_Form
import para


class main(QtWidgets.QWidget, Ui_Form):
    def __init__(self):
        super(main, self).__init__()
        self.ui = Ui_Form()
        self.ui.setupUi(self)
        self.ui.pushButton.clicked.connect(self.start)
        self.add_item_for_comboBox()
        self.path = str()
        BASE_DIR = os.path.dirname(os.path.realpath(__file__))
        self.setWindowIcon(QtGui.QIcon(BASE_DIR + '\\ico.ico'))

    def start(self):
        # print('start')
        print(self.ui.comboBox.currentText())
        type = str(self.ui.comboBox.currentText())
        # print(type)
        self.create_dir(type)

    def add_item_for_comboBox(self):
        print('insert')
        # data = ['sd1', 'sd2']
        # for item in data:
        #     self.ui.comboBox.addItem(item)

    def create_dir(self, type):
        time_now = datetime.datetime.now()
        dir_name = '{}_{}_{}_{}_{}'.format(type, time_now.day, time_now.hour, time_now.minute, time_now.second)
        # dir_name = '%s_%s_%s_%s_%s' % (type, time_now.day, time_now.hour, time_now.minute, time_now.second)
        desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
        path = desktop + '\\' + dir_name
        os.mkdir(path)
        self.path = path


if __name__ == "__main__":
    app = QtWidgets.QApplication([])
    window = main()
    window.show()
    sys.exit(app.exec_())
