import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget, QPushButton, QLineEdit, QLabel, QTableWidget, QTableWidgetItem

class DataEntryApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("資料輸入程式")
        self.setGeometry(100, 100, 600, 400)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout()
        self.central_widget.setLayout(self.layout)

        self.data_entries = [None] * 4

        for i in range(4):
            label = QLabel(f"資料 {i+1}: ")
            entry = QLineEdit()
            self.layout.addWidget(label)
            self.layout.addWidget(entry)
            self.data_entries[i] = entry

        self.add_data_button = QPushButton("新增一筆資料")
        self.add_data_button.clicked.connect(self.add_data)
        self.layout.addWidget(self.add_data_button)

        self.table_widget = QTableWidget()
        self.table_widget.setColumnCount(4)
        self.table_widget.setHorizontalHeaderLabels(["資料 1", "資料 2", "資料 3", "資料 4"])
        self.layout.addWidget(self.table_widget)

    def add_data(self):
        row_position = self.table_widget.rowCount()
        self.table_widget.insertRow(row_position)

        for col, entry in enumerate(self.data_entries):
            value = entry.text()
            self.table_widget.setItem(row_position, col, QTableWidgetItem(value))
            entry.clear()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = DataEntryApp()
    window.show()
    sys.exit(app.exec_())
