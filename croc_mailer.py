from PyQt5 import QtCore, QtWidgets
from PyQt5.QtWidgets import QMessageBox

import sys
import smtplib
from email.mime.text import MIMEText
from dotenv import load_dotenv
import pandas as pd
import os

PATH_TO_DF = "Customer Emails.xlsx"


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(800, 600)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(10, 440, 131, 23))
        self.pushButton.setObjectName("pushButton")
        self.selection_box = QtWidgets.QComboBox(self.centralwidget)
        self.selection_box.setGeometry(QtCore.QRect(11, 11, 131, 20))
        self.selection_box.setObjectName("selection_box")
        self.selection_box.addItem("")
        self.subject_edit = QtWidgets.QLineEdit(self.centralwidget)
        self.subject_edit.setGeometry(QtCore.QRect(10, 69, 133, 20))
        self.subject_edit.setObjectName("subject_edit")
        self.subject_label = QtWidgets.QLabel(self.centralwidget)
        self.subject_label.setGeometry(QtCore.QRect(10, 50, 40, 16))
        self.subject_label.setObjectName("subject_label")
        self.message_label = QtWidgets.QLabel(self.centralwidget)
        self.message_label.setGeometry(QtCore.QRect(10, 111, 46, 16))
        self.message_label.setObjectName("message_label")
        self.message_field = QtWidgets.QTextEdit(self.centralwidget)
        self.message_field.setGeometry(QtCore.QRect(10, 130, 371, 291))
        self.message_field.setObjectName("message_field")
        self.sheet_button = QtWidgets.QPushButton(self.centralwidget)
        self.sheet_button.setGeometry(QtCore.QRect(150, 10, 75, 23))
        self.sheet_button.setObjectName("sheet_button")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 800, 21))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)


        # Options for selection
        self.sheets = pd.ExcelFile(PATH_TO_DF).sheet_names
        self.add_sheets()


    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Croc Mailer"))
        self.pushButton.setText(_translate("MainWindow", "Send Mail"))
        self.selection_box.setItemText(0, _translate("MainWindow", "Choose Excel Sheet"))
        self.subject_label.setText(_translate("MainWindow", "Subject:"))
        self.message_label.setText(_translate("MainWindow", "Message:"))
        self.sheet_button.setText(_translate("MainWindow", "Select"))

        # adding actions to buttons
        self.sheet_button.pressed.connect(self.find)
        self.pushButton.pressed.connect(self.send)

    def send(self):
        # get text from message_field and subject_edit
        global message, subject
        message = self.message_field.toPlainText()
        subject = self.subject_edit.text()

        send_email()


    def find(self):
        global content
        # finding the content of current item in selection box
        content = self.selection_box.currentText()
        open_df()
        self.sheet_button.setText(f"{content} Selected!")
        self.sheet_button.adjustSize()

    def add_sheets(self):
        for sheet in self.sheets:
            self.selection_box.addItem(sheet)

    def show_popup(self):
        # Shows popup on success
        msg = QMessageBox()
        msg.setWindowTitle("Email Send Status")
        msg.setText("Emails sent successfully!")

        x = msg.exec_()


if __name__ == "__main__":


    def configure_env():
        load_dotenv()

    configure_env()

    my_email = os.getenv('admin_email')
    password = os.getenv('admin_password')

    def open_df():
        global df
        df = pd.read_excel(PATH_TO_DF, sheet_name=content)


    def send_email():

        recipients = df["Email"].tolist()
        # names = df['Name'].tolist()

        s = smtplib.SMTP(os.getenv('smtp'), port=587)

        s.ehlo()
        s.starttls()
        s.ehlo()
        s.login(my_email, password)

        for index in range(0, len(recipients)):
            msg = MIMEText(message)
            msg['From'] = my_email
            msg['Subject'] = subject
            msg['To'] = recipients[index]


            s.sendmail(my_email, recipients[index], msg.as_string())

        s.close()

        ui.show_popup()



    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()

    sys.exit(app.exec_())
