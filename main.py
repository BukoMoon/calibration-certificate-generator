#! python3
from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtCore import pyqtSlot, QDate
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox, qApp
from templates import *


class MyWindow(QMainWindow):

    def __init__(self):
        super(MyWindow, self).__init__()
        self.setGeometry(200, 200, 800, 600)
        self.setWindowTitle('HLP Controls Pty Ltd - Certificate Generator - By Louis Adams')
        self.setWindowIcon(QtGui.QIcon('favicon.ico'))
        _translate = QtCore.QCoreApplication.translate

        self._print = None

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

        self.label_checked_at = QtWidgets.QLabel(self.central_widget)
        self.label_temp = QtWidgets.QLabel(self.central_widget)
        self.temperature = QtWidgets.QLineEdit(self.central_widget)
        self.checked = QtWidgets.QLineEdit(self.central_widget)
        self.calibration_lines_and_labels(_translate)

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

    def calibration_lines_and_labels(self, _translate):
        self.label_temp.setText(_translate("MainWindow", "Temperature Read"))
        self.label_checked_at.setText(_translate("MainWindow", "@"))

        self.temperature.setPlaceholderText(_translate("MainWindow", "Enter temperature here on thermometer"))
        self.temperature.setObjectName("temperature")
        self.temperature.setMaxLength(5)

        self.checked.setPlaceholderText(_translate("MainWindow", "Enter temperature from master thermometer here"))
        self.temperature.setObjectName("checked against temperature")
        self.temperature.setMaxLength(5)

    def certificate_checkboxes(self, _translate):
        # Adding checkboxes to a group
        self.certificate.addButton(self.check_box_2)
        self.certificate.addButton(self.check_box)
        self.certificate.addButton(self.check_box_5)
        self.check_box.setChecked(True)

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
        self.check_box_3.setChecked(True)

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
        self.date_edit.setDate(QDate.currentDate())
        self.text_edit.setGeometry(QtCore.QRect(500, 420, 321, 31))
        self.text_edit_2.setGeometry(QtCore.QRect(500, 319, 321, 31))
        self.text_edit_3.setGeometry(QtCore.QRect(500, 230, 321, 31))

        self.temperature.setGeometry(QtCore.QRect(280, 500, 120, 20))
        self.checked.setGeometry(QtCore.QRect(280, 540, 120, 20))
        self.label_temp.setFont(small_font)
        self.label_temp.setObjectName("label_temp")
        self.label_temp.setAlignment(QtCore.Qt.AlignCenter)
        self.label_temp.setGeometry(QtCore.QRect(260, 480, 151, 16))
        self.label_checked_at.setFont(small_font)
        self.label_checked_at.setObjectName("label_checked_at")
        self.label_checked_at.setAlignment(QtCore.Qt.AlignCenter)
        self.label_checked_at.setGeometry(QtCore.QRect(320, 520, 31, 16))

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
        self.label_6.setPixmap(QtGui.QPixmap(resource_path("HLP-Logo-Aus.jpg")))
        self.label_6.setScaledContents(True)
        self.label_6.setObjectName("label_6")

        self.button1.clicked.connect(self.generate)
        if self.button1_2.clicked:
            self.print_files = True
        self.button1_2.clicked.connect(self.generate)
        QtCore.QMetaObject.connectSlotsByName(self)
        self.check_box.stateChanged.connect(self.state_change)
        self.check_box_2.stateChanged.connect(self.state_change)
        self.check_box_5.stateChanged.connect(self.state_change)
        self.action_exit.triggered.connect(qApp.quit)
        self.show()
        self.temperature.hide()
        self.checked.hide()
        self.label_temp.hide()
        self.label_checked_at.hide()

    def state_change(self):
        # Calibration certificates
        if self.check_box_2.isChecked():
            self.show()
            self.label_5.setText('Device Model')
            self.check_box_3.setChecked(True)
            self.temperature.show()
            self.checked.show()
            self.label_temp.show()
            self.label_checked_at.show()
            self.label_2.show()
            self.text_edit.show()

        # Single conformance certificates
        elif self.check_box.isChecked():
            self.show()
            self.label_5.setText('Device Model')
            self.temperature.hide()
            self.checked.hide()
            self.label_temp.hide()
            self.label_checked_at.hide()
            self.label_2.hide()
            self.text_edit.hide()

        # Multiple conformance certificates
        elif self.check_box_5.isChecked():
            self.show()
            self.label_5.setText('Pocket Temp Model')
            self.temperature.hide()
            self.checked.hide()
            self.label_temp.hide()
            self.label_checked_at.hide()
            self.label_2.show()
            self.text_edit.show()

    def keyPressEvent(self, event):
        if event.key() == QtCore.Qt.Key_Control + QtCore.Qt.Key_E:
            self.close()

    def closeEvent(self, event):
        close = QMessageBox()
        close.setWindowTitle('Confirm')
        close.setWindowIcon(QtGui.QIcon('favicon.ico'))
        close.setText("You sure you want to quit?")
        close.setStandardButtons(QMessageBox.Yes | QMessageBox.Cancel)
        close = close.exec()

        if close == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()

    @pyqtSlot()
    def generate(self):
        convert_date = self.date_edit.date().toPyDate()
        set_date = datetime.datetime.strftime(convert_date, '%d/%m/%Y   ')
        set_name = self.text_edit.text()
        set_serial_number = str(self.text_edit_2.text())
        set_model = self.text_edit_3.text()
        if self.label_temp.text or self.label_checked_at.text == "":
            set_checked_at = ""
            set_temperature = ""
        else:
            set_checked_at = self.label_temp.text()
            set_temperature = self.label_checked_at.text()
        file_name = 'generated.docx'

        if self.check_box_3.isChecked:
            set_months = ' 12 '
        else:
            set_months = ' 24 '

        if self.check_box.isChecked:
            SingleConformance(file_name, set_model, set_date, set_name, set_months, set_serial_number)
            if self.print_files:
                self.printer(file_name, False)

        elif self.check_box_2.isChecked:
            CalibrationCertificate(file_name, set_model, set_date, set_name, set_months,
                                   set_serial_number, float(set_checked_at), float(set_temperature))
            if self.print_files:
                self.printer(file_name, False)

        else:
            self.generate_multiple(set_date, set_serial_number, set_months, set_name)

    def generate_multiple(self, date: str, serial_number: str, months: str, name: str) -> None:
        print_multiple = True
        for i in range(0, 7):
            print(f'Generating docx {i}')
            MultipleConformance('generated' + str(i) + '.docx', date, name, months, serial_number)
            serial_number += 4
            if print_multiple and self.print_files:
                self.printer('generated' + str(i) + '.docx', True)

    @staticmethod
    def printer(file: str, multiple_files: bool) -> None:
        if multiple_files:
            for file_number in range(0, 7):
                GeneratorPrinter('generated' + str(file_number) + '.docx')
        else:
            GeneratorPrinter(file)

    @property
    def print_files(self) -> bool:
        return self._print

    @print_files.setter
    def print_files(self, print_docx: bool = False):
        self._print = print_docx


def window():
    app = QApplication(sys.argv)
    win = MyWindow()

    win.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    window()
