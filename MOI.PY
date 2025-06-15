import sys
import openpyxl
from openpyxl import Workbook
import os
import re
from datetime import datetime
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QLabel, 
    QLineEdit, QPushButton, QMessageBox, QFormLayout,
    QTableWidget, QTableWidgetItem, QTabWidget, QHeaderView,
    QComboBox, QDateEdit
)
from PyQt5.QtCore import Qt, QDate


exel_file = "юзы_регнулию.lxsx"
shit_name = "юзеры"

class Registration(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ЗАРЕГАЙСЯ НАХУЙ")
        self.setFixedSize(1000, 700)

        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        self.табик_регистрации = QWidget()
        self.setup_register_tab()

        self.tabs.addTab(self.табик_регистрации, "Регистрация нахуй")

        self.init_exel_file()

    def setup_register_tab(self):
        layout = QVBoxLayout()


    def init_exel_file():
        if not os.path.exists(exel_file):
            wb = Workbook()
            ws = wb.active
            ws.title = shit_name
            ws.append([
                "ID", "Username", "Password", "Email", 
                "Lastname", "Firstname", "Middlename",
                "Birthdate", "Phone", "Gender", "Registration Date"
            ])
            wb.save(exel_file)






if __name__ == "__main__":
    app = QApplication(sys.argv)

window = Registration()
window.show()
sys.exit(app.exec())