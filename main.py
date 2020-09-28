#! python3
import datetime
import sys

from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QApplication, QMainWindow

from templategenerator import *


# TODO Make it so one check box can be selected at a time
# TODO Checking select box makes line edit widgets appear
# TODO Connect functions etc to each Line edit, date, etc.


class MyWindow(QMainWindow):

    def __init__(self):
        super(MyWindow, self).__init__()
        self.setGeometry(200, 200, 800, 600)
        self.setWindowTitle('HLP Controls Pty Ltd - Certificate Generator - By Louis Adams')
        _translate = QtCore.QCoreApplication.translate

        self.printer_docx = True

        # Initialising Widgets
        self.central_widget = QtWidgets.QWidget(self)

        # Creating groups
        self.certificate = QtWidgets.QButtonGroup()
        self.months = QtWidgets.QButtonGroup()

        # Toolbar
        self.menu_bar = QtWidgets.QMenuBar(self)
        self.menu_file = QtWidgets.QMenu(self.menu_bar)
        self.status_bar = QtWidgets.QStatusBar(self)
        self.action_exit = QtWidgets.QAction(self)
        self.init_toolbar(_translate)

        # Certificate checkboxes
        self.check_box = QtWidgets.QCheckBox(self.central_widget)
        self.check_box_2 = QtWidgets.QCheckBox(self.central_widget)
        self.check_box_5 = QtWidgets.QCheckBox(self.central_widget)
        self.certificate_checkboxes(_translate)

        # Months checkboxes
        self.check_box_3 = QtWidgets.QCheckBox(self.central_widget)
        self.check_box_4 = QtWidgets.QCheckBox(self.central_widget)
        self.month_checkboxes(_translate)

        # Line edit boxes - Model, Serial Number, Name & Date
        self.text_edit = QtWidgets.QLineEdit(self.central_widget)
        self.text_edit_2 = QtWidgets.QLineEdit(self.central_widget)
        self.text_edit_3 = QtWidgets.QLineEdit(self.central_widget)
        self.date_edit = QtWidgets.QDateEdit(self.central_widget)
        self.text_date_widgets(_translate)

        # Generate buttons
        self.button1 = QtWidgets.QPushButton(self.central_widget)
        self.button1_2 = QtWidgets.QPushButton(self.central_widget)
        self.generate_buttons(_translate)

        # Labels
        self.label = QtWidgets.QLabel(self.central_widget)
        self.label1 = QtWidgets.QLabel(self.central_widget)
        self.label_2 = QtWidgets.QLabel(self.central_widget)
        self.label_3 = QtWidgets.QLabel(self.central_widget)
        self.label_4 = QtWidgets.QLabel(self.central_widget)
        self.label_5 = QtWidgets.QLabel(self.central_widget)
        self.label_6 = QtWidgets.QLabel(self.central_widget)
        self.labels(_translate)

        self.init_ui()

    def init_toolbar(self, _translate):
        self.menu_file.setTitle(_translate("MainWindow", "File"))
        self.action_exit.setText(_translate("MainWindow", "Exit"))
        self.action_exit.setStatusTip(_translate("MainWindow", "Exit the program"))
        self.action_exit.setShortcut(_translate("MainWindow", "Ctrl+E"))

        self.menu_bar.setObjectName("menu_bar")
        self.menu_file.setObjectName("menu_file")
        self.setMenuBar(self.menu_bar)
        self.status_bar.setObjectName("status_bar")
        self.setStatusBar(self.status_bar)
        self.menu_file.addAction(self.action_exit)
        self.menu_bar.addAction(self.menu_file.menuAction())

    def certificate_checkboxes(self, _translate):
        # Adding checkboxes to a group
        self.certificate.addButton(self.check_box_2)
        self.certificate.addButton(self.check_box)
        self.certificate.addButton(self.check_box_5)
        self.check_box.isChecked = True

        # Setting name & text
        self.check_box.setObjectName("check_box")
        self.check_box_2.setObjectName("check_box_2")
        self.check_box_5.setObjectName("check_box_5")
        self.check_box.setText(_translate("MainWindow", "Single Conformance Certificate"))
        self.check_box_2.setText(_translate("MainWindow", "Calibration Certificate"))
        self.check_box_5.setText(_translate("MainWindow", "Multiple Conformance Certificates"))

    def month_checkboxes(self, _translate):
        # Adding checkboxes to a group
        self.months.addButton(self.check_box_3)
        self.months.addButton(self.check_box_4)
        self.check_box_3.isChecked = True

        # Setting name & text
        self.check_box_3.setText(_translate("MainWindow", "12 Months"))
        self.check_box_4.setText(_translate("MainWindow", "24 Months"))
        self.check_box_3.setObjectName("check_box_3")
        self.check_box_4.setObjectName("check_box_4")

    def text_date_widgets(self, _translate):
        self.text_edit.setToolTip(_translate("MainWindow", "Enter your name here"))
        self.text_edit.setPlaceholderText(_translate("MainWindow", "Enter your name here E.g L.Adams"))
        self.text_edit_2.setToolTip(_translate("MainWindow", "Enter the Device\'s Serial Number"))
        self.text_edit_2.setPlaceholderText(_translate("MainWindow", "Enter the Device Serial Number Here"))
        self.text_edit_3.setToolTip(_translate("MainWindow", "Enter your Device Model Here"))
        self.text_edit_3.setPlaceholderText(_translate("MainWindow", "Enter the Device Model Here"))
        self.date_edit.setToolTip(_translate("MainWindow", "The date on the certificate - normally the current date"))
        self.date_edit.setDisplayFormat(_translate("MainWindow", "dd/MM/yyyy"))

        self.text_edit.setObjectName("text_edit")
        self.text_edit.setMaxLength(20)
        self.text_edit_2.setObjectName("text_edit_2")
        self.text_edit_2.setMaxLength(20)
        self.text_edit_3.setObjectName("text_edit_3")
        self.text_edit_3.setMaxLength(20)
        self.date_edit.setObjectName("date_edit")

    def generate_buttons(self, _translate):
        self.button1.setToolTip(_translate("MainWindow", "Click to generate your certificate/s"))
        self.button1.setText(_translate("MainWindow", "Generate and Print"))
        self.button1_2.setToolTip(_translate("MainWindow", "Click to generate your certificate/s"))
        self.button1_2.setText(_translate("MainWindow", "Generate"))

        self.button1.setObjectName("button1")
        self.button1_2.setObjectName("button1_2")

    def labels(self, _translate):
        self.label.setText(_translate("MainWindow", "Date:"))
        self.label1.setText(_translate("MainWindow", "Select a Template"))
        self.label_2.setText(_translate("MainWindow", "Enter your name below"))
        self.label_3.setText(_translate("MainWindow", "Device Serial Number"))
        self.label_4.setText(_translate("MainWindow", "Certificate Validation Period"))
        self.label_5.setText(_translate("MainWindow", "Device Model"))

    def init_ui(self):
        self.resize(900, 700)
        self.setMinimumSize(QtCore.QSize(900, 700))
        self.setCentralWidget(self.central_widget)

        # Initialising font sizes
        small_font = QtGui.QFont()
        small_font.setPointSize(11)
        font = QtGui.QFont()
        font.setPointSize(14)

        self.central_widget.setObjectName("central_widget")
        self.menu_bar.setGeometry(QtCore.QRect(0, 0, 900, 22))

        # Certificate checkboxes locations
        self.check_box.setGeometry(QtCore.QRect(80, 250, 241, 31))
        self.check_box_2.setGeometry(QtCore.QRect(80, 350, 171, 23))
        self.check_box_5.setGeometry(QtCore.QRect(80, 300, 261, 31))

        # Month checkboxes locations
        self.check_box_3.setGeometry(QtCore.QRect(80, 440, 121, 31))
        self.check_box_4.setGeometry(QtCore.QRect(80, 480, 111, 31))

        # LineEdit Widgets & Date locations
        self.date_edit.setGeometry(QtCore.QRect(124, 534, 111, 31))
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDate(QtCore.QDate(2020, 9, 12))
        self.text_edit.setGeometry(QtCore.QRect(500, 420, 321, 31))
        self.text_edit_2.setGeometry(QtCore.QRect(500, 319, 321, 31))
        self.text_edit_3.setGeometry(QtCore.QRect(500, 230, 321, 31))

        # Button locations
        self.button1.setGeometry(QtCore.QRect(650, 500, 171, 61))
        self.button1_2.setGeometry(QtCore.QRect(490, 500, 151, 61))

        # Label locations & font
        self.label.setGeometry(QtCore.QRect(80, 540, 71, 21))
        self.label.setFont(small_font)
        self.label.setObjectName("label")
        self.label1.setGeometry(QtCore.QRect(80, 210, 171, 31))
        self.label1.setFont(font)
        self.label1.setObjectName("label1")
        self.label_2.setGeometry(QtCore.QRect(500, 390, 321, 31))
        self.label_2.setFont(small_font)
        self.label_2.setAlignment(QtCore.Qt.AlignCenter)
        self.label_2.setObjectName("label_2")
        self.label_3.setGeometry(QtCore.QRect(500, 290, 321, 20))
        self.label_3.setAlignment(QtCore.Qt.AlignCenter)
        self.label_3.setObjectName("label_3")
        self.label_4.setGeometry(QtCore.QRect(80, 390, 251, 41))
        self.label_4.setFont(font)
        self.label_4.setObjectName("label_4")
        self.label_5.setGeometry(QtCore.QRect(500, 200, 321, 31))
        self.label_5.setAlignment(QtCore.Qt.AlignCenter)
        self.label_5.setObjectName("label_5")

        # Logo Image
        self.label_6.setGeometry(QtCore.QRect(40, 10, 801, 151))
        self.label_6.setText("")
        self.label_6.setPixmap(QtGui.QPixmap("HLP-Logo-Aus.jpg"))
        self.label_6.setScaledContents(True)
        self.label_6.setObjectName("label_6")

        self.button1.clicked.connect(self.generate)
        # TODO write function to print the generated docx files
        QtCore.QMetaObject.connectSlotsByName(self)
        self.show()

    @pyqtSlot()
    def generate(self):
        convert_date = self.date_edit.date().toPyDate()
        set_date = datetime.datetime.strftime(convert_date, '%d/%M/%YYYY ')
        set_name = self.text_edit.text()
        set_serial_number = self.text_edit_2.text()
        set_model = self.text_edit_3.text()
        set_checked_at = ''
        set_temperature = ''
        file_name = 'generated.docx'

        if self.check_box_3.isChecked:
            set_months = ' 12 '
        else:
            set_months = ' 24 '

        if self.check_box.isChecked:
            single_conformance = SingleConformance(file_name)
            single_conformance.model = set_model
            single_conformance.date = set_date
            single_conformance.name = set_name
            single_conformance.months = set_months
            single_conformance.serial_number = set_serial_number
            single_conformance.prefix = ''
            single_conformance.create_docx()

        elif self.check_box_2.isChecked:
            calibration_certificate = CalibrationCertificate(file_name)
            calibration_certificate.model = set_model
            calibration_certificate.date = set_date
            calibration_certificate.name = set_name
            calibration_certificate.months = set_months
            calibration_certificate.serial_number = set_serial_number
            calibration_certificate.prefix = ''
            calibration_certificate.check = set_checked_at
            calibration_certificate.temp = set_temperature
            calibration_certificate.create_docx()

        else:
            self.generate_multiple(set_date, set_serial_number, set_model, set_months, set_name)

    def generate_multiple(self, set_date, set_serial_number, set_model, set_months, set_name):
        print_multiple = True
        for i in range(0, 7):
            print(f'Generating docx {i}')
            multiple_conformance = MultipleConformance('generated' + str(i) + '.docx')
            multiple_conformance.model = set_model
            multiple_conformance.date = set_date
            multiple_conformance.name = set_name
            multiple_conformance.months = set_months
            multiple_conformance.serial_number = set_serial_number
            multiple_conformance.prefix = ''
            multiple_conformance.create_docx()
            set_serial_number += 4
            if self.printer_docx and print_multiple:
                self.printer('generated' + str(i) + '.docx', True)

    @staticmethod
    def printer(file, multiple_files: bool):
        if multiple_files:
            for file_number in range(0, 7):
                print(f'{file_number + 1}. Printing docx, {file}')
        else:
            print(f'Printing docx, {file}')


def window():
    app = QApplication(sys.argv)
    win = MyWindow()

    win.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    window()
