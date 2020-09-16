#! python3
#
# Docs used to create
# https://www.youtube.com/watch?v=Vde5SH8e1OQ
# https://stackoverflow.com/questions/51308772/pyqt5-from-qlineedit-to-variable
# https://www.learnpyqt.com/blog/adding-images-to-pyqt5-applications/


import sys
from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtWidgets import QApplication, QMainWindow


# TODO Make it so one check box can be selected at a time
# TODO Checking select box makes line edit widgets appear
# TODO Connect functions etc to each Line edit, date, etc.

class MyWindow(QMainWindow):

    def __init__(self):
        super(MyWindow, self).__init__()
        self.setGeometry(200, 200, 800, 600)
        self.setWindowTitle('HLP Controls Pty Ltd - Certificate Generator - By Louis Adams')

        # Initialising Widgets
        self.central_widget = QtWidgets.QWidget(self)
        self.button1 = QtWidgets.QPushButton(self.central_widget)
        self.button1_2 = QtWidgets.QPushButton(self.central_widget)
        self.label1 = QtWidgets.QLabel(self.central_widget)
        self.check_box = QtWidgets.QCheckBox(self.central_widget)
        self.check_box_2 = QtWidgets.QCheckBox(self.central_widget)
        self.date_edit = QtWidgets.QDateEdit(self.central_widget)
        self.text_edit = QtWidgets.QLineEdit(self.central_widget)
        self.label = QtWidgets.QLabel(self.central_widget)
        self.label1 = QtWidgets.QLabel(self.central_widget)
        self.label_2 = QtWidgets.QLabel(self.central_widget)
        self.text_edit_2 = QtWidgets.QLineEdit(self.central_widget)
        self.label_3 = QtWidgets.QLabel(self.central_widget)
        self.check_box_3 = QtWidgets.QCheckBox(self.central_widget)
        self.check_box_4 = QtWidgets.QCheckBox(self.central_widget)
        self.check_box_5 = QtWidgets.QCheckBox(self.central_widget)
        self.label_4 = QtWidgets.QLabel(self.central_widget)
        self.text_edit_3 = QtWidgets.QLineEdit(self.central_widget)
        self.label_5 = QtWidgets.QLabel(self.central_widget)
        self.label_6 = QtWidgets.QLabel(self.central_widget)
        self.menu_bar = QtWidgets.QMenuBar(self)
        self.menu_file = QtWidgets.QMenu(self.menu_bar)
        self.status_bar = QtWidgets.QStatusBar(self)
        self.action_exit = QtWidgets.QAction(self)

        self.init_ui()

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
        self.button1.setGeometry(QtCore.QRect(650, 500, 171, 61))
        self.button1.setObjectName("button1")
        self.label1.setGeometry(QtCore.QRect(80, 210, 171, 31))
        self.label1.setFont(font)
        self.label1.setObjectName("label1")
        self.check_box.setGeometry(QtCore.QRect(80, 250, 241, 31))
        self.check_box.setObjectName("check_box")
        self.check_box_2.setGeometry(QtCore.QRect(80, 350, 171, 23))
        self.check_box_2.setObjectName("check_box_2")
        self.date_edit.setGeometry(QtCore.QRect(124, 534, 111, 31))
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDate(QtCore.QDate(2020, 9, 12))
        self.date_edit.setObjectName("date_edit")
        self.text_edit.setGeometry(QtCore.QRect(500, 420, 321, 31))
        self.text_edit.setObjectName("text_edit")
        self.text_edit.setMaxLength(20)
        self.label.setGeometry(QtCore.QRect(80, 540, 71, 21))
        self.label.setFont(small_font)
        self.label.setObjectName("label")
        self.label_2.setGeometry(QtCore.QRect(500, 390, 321, 31))
        self.label_2.setFont(small_font)
        self.label_2.setAlignment(QtCore.Qt.AlignCenter)
        self.label_2.setObjectName("label_2")
        self.text_edit_2.setGeometry(QtCore.QRect(500, 319, 321, 31))
        self.text_edit_2.setObjectName("text_edit_2")
        self.text_edit_2.setMaxLength(20)
        self.label_3 = QtWidgets.QLabel(self.central_widget)
        self.label_3.setGeometry(QtCore.QRect(500, 290, 321, 20))
        self.label_3.setAlignment(QtCore.Qt.AlignCenter)
        self.label_3.setObjectName("label_3")
        self.check_box_3.setGeometry(QtCore.QRect(80, 440, 121, 31))
        self.check_box_3.setObjectName("check_box_3")
        self.check_box_4.setGeometry(QtCore.QRect(80, 480, 111, 31))
        self.check_box_4.setObjectName("check_box_4")
        self.label_4.setGeometry(QtCore.QRect(80, 390, 251, 41))
        self.label_4.setFont(font)
        self.label_4.setObjectName("label_4")
        self.text_edit_3.setGeometry(QtCore.QRect(500, 230, 321, 31))
        self.text_edit_3.setObjectName("text_edit_3")
        self.text_edit_3.setMaxLength(20)
        self.label_5.setGeometry(QtCore.QRect(500, 200, 321, 31))
        self.label_5.setAlignment(QtCore.Qt.AlignCenter)
        self.label_5.setObjectName("label_5")

        self.check_box_5.setGeometry(QtCore.QRect(80, 300, 261, 31))
        self.check_box_5.setObjectName("check_box_5")
        self.button1_2.setGeometry(QtCore.QRect(490, 500, 151, 61))
        self.button1_2.setObjectName("button1_2")
        self.menu_bar.setGeometry(QtCore.QRect(0, 0, 900, 22))
        self.menu_bar.setObjectName("menu_bar")
        self.menu_file.setObjectName("menu_file")
        self.setMenuBar(self.menu_bar)
        self.status_bar.setObjectName("status_bar")
        self.setStatusBar(self.status_bar)
        self.menu_file.addAction(self.action_exit)
        self.menu_bar.addAction(self.menu_file.menuAction())

        # Logo Image
        self.label_6.setGeometry(QtCore.QRect(40, 10, 801, 151))
        self.label_6.setText("")
        self.label_6.setPixmap(QtGui.QPixmap("HLP-Logo-Aus.jpg"))
        self.label_6.setScaledContents(True)
        self.label_6.setObjectName("label_6")

        self.retranslate_ui()
        QtCore.QMetaObject.connectSlotsByName(self)

    def retranslate_ui(self):
        _translate = QtCore.QCoreApplication.translate
        self.button1.setToolTip(_translate("MainWindow", "Click to generate your certificate/s"))
        self.button1.setText(_translate("MainWindow", "Generate and Print"))
        self.label1.setText(_translate("MainWindow", "Select a Template"))
        self.check_box.setText(_translate("MainWindow", "Single Conformance Certificate"))
        self.check_box_2.setText(_translate("MainWindow", "Calibration Certificate"))
        self.date_edit.setToolTip(_translate("MainWindow", "The date on the certificate - normally the current date"))
        self.date_edit.setDisplayFormat(_translate("MainWindow", "dd/MM/yyyy"))
        self.text_edit.setToolTip(_translate("MainWindow", "Enter your name here"))
        self.text_edit.setPlaceholderText(_translate("MainWindow", "Enter your name here E.g L.Adams"))
        self.label.setText(_translate("MainWindow", "Date:"))
        self.label_2.setText(_translate("MainWindow", "Enter your name below"))
        self.text_edit_2.setToolTip(_translate("MainWindow", "Enter the Device\'s Serial Number"))
        self.text_edit_2.setPlaceholderText(_translate("MainWindow", "Enter the Device Serial Number Here"))
        self.label_3.setText(_translate("MainWindow", "Device Serial Number"))
        self.check_box_3.setText(_translate("MainWindow", "12 Months"))
        self.check_box_4.setText(_translate("MainWindow", "24 Months"))
        self.label_4.setText(_translate("MainWindow", "Certificate Validation Period"))
        self.text_edit_3.setToolTip(_translate("MainWindow", "Enter your Device Model Here"))
        self.text_edit_3.setPlaceholderText(_translate("MainWindow", "Enter the Device Model Here"))
        self.label_5.setText(_translate("MainWindow", "Device Model"))
        self.check_box_5.setText(_translate("MainWindow", "Multiple Conformance Certificates"))
        self.button1_2.setToolTip(_translate("MainWindow", "Click to generate your certificate/s"))
        self.button1_2.setText(_translate("MainWindow", "Generate"))
        self.menu_file.setTitle(_translate("MainWindow", "File"))
        self.action_exit.setText(_translate("MainWindow", "Exit"))
        self.action_exit.setStatusTip(_translate("MainWindow", "Exit the program"))
        self.action_exit.setShortcut(_translate("MainWindow", "Ctrl+E"))


def window():
    app = QApplication(sys.argv)
    win = MyWindow()

    win.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    window()
