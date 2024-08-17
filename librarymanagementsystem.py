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

    def show_login_window(self):
        self.login_window = QWidget()
        self.login_window.setWindowTitle("تسجيل الدخول القائد")
        self.login_window.setGeometry(100, 100, 300, 250)

        layout = QFormLayout()

        self.username_input = QLineEdit()
        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.Password)

        login_button = QPushButton("تسجيل الدخول")
        login_button.clicked.connect(self.login)

        layout.addRow(QLabel("اسم المستخدم:"), self.username_input)
        layout.addRow(QLabel("كلمة السر:"), self.password_input)
        layout.addWidget(login_button)

        self.login_window.setLayout(layout)
        self.login_window.show()

    def login(self):
        username = self.username_input.text()
        password = self.password_input.text()

        admin_sheet = self.wb['Admin']
        for row in admin_sheet.iter_rows(min_row=2, values_only=True):
            if username == row[0] and password == row[1]:
                self.logged_in = True
                self.current_student_id = username
                QMessageBox.information(self, "تم تسجيل الدخول", f"مرحبا, {username}!")
                self.create_main_window()
                self.login_window.close()
                return

        QMessageBox.critical(self, "فشل تسجيل الدخول", "خطاء فالاسم او كلمة السر.")

    def create_main_window(self):
        self.main_window = QMainWindow()
        self.main_window.setWindowTitle("نظام المكتبه الذكى")
        self.main_window.setGeometry(100, 100, 1200, 800)

        self.tabs = QTabWidget()
        self.tabs.addTab(self.create_students_tab(), "الطلاب")
        self.tabs.addTab(self.create_books_tab(), "الكتب")
        self.tabs.addTab(self.create_borrowing_tab(), "عمليات الاستعاره")
        self.tabs.addTab(self.create_returned_tab(), "عمليات الاعاده")
        self.tabs.addTab(self.create_reports_tab(), "التقارير")

        backup_button = QPushButton("نسخة احتياطيه من البيانات")
        backup_button.clicked.connect(self.backup_data)

        restore_button = QPushButton("استعادة اخر نسخه احتياطيه")
        restore_button.clicked.connect(self.restore_data)

        layout = QVBoxLayout()
        layout.addWidget(self.tabs)
        layout.addWidget(backup_button)
        layout.addWidget(restore_button)

        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.main_window.setCentralWidget(central_widget)
        self.main_window.show()

    def create_students_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()

        self.students_table = QTableWidget()
        self.students_table.setColumnCount(5)
        self.students_table.setHorizontalHeaderLabels(["ID", "الرقم العسكرى", "الرتبه", "الاسم", "عدد مرات الاستعاره"])
        layout.addWidget(self.students_table)

        self.load_students_from_excel()

        add_button = QPushButton("اضافة طالب جديد")
        add_button.clicked.connect(self.add_student)
        layout.addWidget(add_button)

        search_layout = QHBoxLayout()
        self.search_student_input = QLineEdit()
        self.search_student_input.setPlaceholderText("بحث طلاب...")
        self.search_student_input.textChanged.connect(self.search_students)
        search_layout.addWidget(self.search_student_input)
        layout.addLayout(search_layout)

        tab.setLayout(layout)
        return tab

    def create_books_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()

        self.books_table = QTableWidget()
        self.books_table.setColumnCount(6)
        self.books_table.setHorizontalHeaderLabels(["ID", "رقم الكتاب", "اسم الكتاب", "الكاتب", "عدد النسخ المتاحه", "عدد مرات الاستعاره"])
        layout.addWidget(self.books_table)

        self.load_books_from_excel()

        add_button = QPushButton("اضافة كتاب")
        add_button.clicked.connect(self.add_book)
        layout.addWidget(add_button)

        search_layout = QHBoxLayout()
        self.search_book_input = QLineEdit()
        self.search_book_input.setPlaceholderText("بحث كتب...")
        self.search_book_input.textChanged.connect(self.search_books)
        search_layout.addWidget(self.search_book_input)
        layout.addLayout(search_layout)

        tab.setLayout(layout)
        return tab

    def create_borrowing_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()

        self.borrowing_table = QTableWidget()
        self.borrowing_table.setColumnCount(6)
        self.borrowing_table.setHorizontalHeaderLabels(["ID", "رقم الكتاب", "اسم الكتاب", "الرقم العسكرى", "اسم الطالب", "تاريخ الاستعاره"])
        layout.addWidget(self.borrowing_table)

        self.load_borrowing_from_excel()

        borrow_button = QPushButton("استعاره جديده")
        borrow_button.clicked.connect(self.borrow_book)
        layout.addWidget(borrow_button)

        tab.setLayout(layout)
        return tab

    def create_returned_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()

        self.returned_table = QTableWidget()
        self.returned_table.setColumnCount(7)
        self.returned_table.setHorizontalHeaderLabels(["ID", "رقم الكتاب", "اسم الكتاب", "الرقم العسكرى", "اسم الطالب", "تاريخ الاستعاره", "تاريخ الاستعاده"])
        layout.addWidget(self.returned_table)

        self.load_returned_from_excel()

        return_button = QPushButton("استعادة كتاب")
        return_button.clicked.connect(self.return_book)
        layout.addWidget(return_button)

        tab.setLayout(layout)
        return tab

    def create_reports_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()

        self.reports_table = QTableWidget()
        self.reports_table.setColumnCount(7)
        self.reports_table.setHorizontalHeaderLabels(["ID", "رقم الكتاب", "اسم الكتاب", "الرقم العسكرى", "اسم الطالب", "تاريخ الاستعاره", "تاريخ الاستعاده"])
        layout.addWidget(self.reports_table)

        generate_report_button = QPushButton("انشاء تقرير")
        generate_report_button.clicked.connect(self.generate_report)
        layout.addWidget(generate_report_button)

        tab.setLayout(layout)
        return tab

if __name__ == "__main__":
    app = QApplication(sys.argv)
    system = LibraryManagementSystem()
    sys.exit(app.exec_())
