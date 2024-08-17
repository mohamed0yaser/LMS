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

    def initUI(self):
        self.setStyleSheet("background-color: #f5f5f5;")
        self.create_or_load_library_data_excel()
        self.show_login_window()

    def create_or_load_library_data_excel(self):
        try:
            self.wb = load_workbook(self.excel_file)
            print(f"Loaded existing Library Data Excel file: {self.excel_file}")
        except FileNotFoundError:
            print(f"Library Data Excel file not found at: {self.excel_file}")
            self.create_library_data_excel()

    def create_library_data_excel(self):
        self.wb = Workbook()
        ws_admin = self.wb.create_sheet(title='Admin')
        ws_students = self.wb.create_sheet(title='students')
        ws_books = self.wb.create_sheet(title='Books')
        ws_borrowing = self.wb.create_sheet(title='Borrowing')
        ws_returned = self.wb.create_sheet(title='Returned')

        ws_admin.append(['Username', 'Password', 'Role'])
        ws_admin.append(['admin', '123', 'admin'])
        ws_students.append(['ID', 'student ID', 'Position', 'student Name', 'Borrowed Count'])
        ws_books.append(['ID', 'Book ID', 'Book Title', 'Author', 'Copies Available', 'Borrowed Count'])
        ws_borrowing.append(['ID', 'Book ID', 'Book Title', 'student ID', 'student Name', 'Borrowed Date'])
        ws_returned.append(['ID', 'Book ID', 'Book Title', 'student ID', 'student Name', 'Borrowed Date', 'Returned Date'])

        self.wb.save(self.excel_file)
        print(f"Created new Library Data Excel file with schema: {self.excel_file}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    system = LibraryManagementSystem()
    sys.exit(app.exec_())
