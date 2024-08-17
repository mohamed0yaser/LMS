from datetime import datetime
import sys
import shutil
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTabWidget, QWidget, QVBoxLayout, QFormLayout,
    QLabel, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem, QMessageBox,
    QComboBox, QHBoxLayout, QDateEdit, QFileDialog
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPalette, QColor
from openpyxl import Workbook, load_workbook

class LibraryManagementSystem(QMainWindow):
    def __init__(self):
        super().__init__()
        self.excel_file = 'library_data.xlsx'
        self.backup_file = 'library_data_backup.xlsx'
        self.logged_in = False
        self.current_student_id = None
        self.wb = None
        self.initUI()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    system = LibraryManagementSystem()
    sys.exit(app.exec_())
