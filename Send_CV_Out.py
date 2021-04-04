'''
Goal: To send an email with a PDF file attached via the Yahoo server to a list of e-mails.

Requirement: Set up a Yahoo account to send an email by an app.
             Fill in the email.csv file with the email address into the DB folder under the app folder.

Source: https://stackoverflow.com/questions/3362600/how-to-send-email-attachments
        https://realpython.com/python-send-email/#option-1-setting-up-a-gmail-account-for-development
        https://myaccount.google.com/lesssecureapps?pli=1&rapt=AEjHL4O2ErNpA0DRf56drLkiOqls-6kBncD8kEy2Xi__1OfVCv2MbgY_Jiax-fUNJ_upVccqInp0hAwD82et0FcPOklngFyFhg

Author: Yanko Sarzedas da Costa (yscosta@yahoo.com or yscosta62@gmail.com)
'''
import sys
from os import getcwd, getenv
import json

# Packages to make the data entry.
from PyQt5.QtWidgets import (QApplication, QTableWidget, QGridLayout, QGroupBox, QVBoxLayout, \
                             QHBoxLayout, QTabWidget, QWidget, QPushButton, QDialog, \
                             QSizePolicy, QTableWidgetItem, QLabel, QLineEdit, QTextEdit, \
                             QComboBox, QFileDialog, QMessageBox)
from PyQt5.QtGui import (QIcon, QColor)
from PyQt5.QtCore import (Qt, QRect)

from lib import EmailManager

class SendCVOut(QWidget):
    def __init__(self):
        super().__init__()
        # set the database directory
        self.db_path = getcwd() + '\\DB\\'
        self.loadEmail()
        self.initLayout()
        self.email_manager = EmailManager()
        return

    def loadEmail(self):
        try:
            with open(self.db_path + 'e-mail-struct.json', 'r') as f:
                self.email_struct = json.load(f)
                f.close()
        except Exception as e:
            self.email_struct = {'subject': '', 'body': '', 'files': [], 'username':'', 'smtp_choice':''}
        return

    def saveEmail(self):
        files = self.getFileList()
        self.email_struct = {'subject': self.subject.text(),
                             'body': self.body.toPlainText(),
                             'files': files,
                             'username': self.username.text(),
                             'smtp_choice': self.smtp_choice.currentText()}
        try:
            with open(self.db_path + 'e-mail-struct.json', 'w') as f:
                json.dump(self.email_struct, f)
                f.close()
        except Exception as e:
            QMessageBox.about(self, 'Alerta', e)
        return

    def getFileList(self):
        files = []
        for row in range(self.files_list.rowCount()):
            file = self.files_list.model().data(self.files_list.model().index(row, 0))
            if file != None:
                files.append(file)
        return files

    def initLayout(self):
        self.setWindowTitle('Envio de CV em bateladas')
        self.setWindowIcon(QIcon('e-mail-icon.jpg'))
        self.setGeometry(QRect(350, 150, 670, 360))

        object_layout = QVBoxLayout()
        object_layout.addWidget(self.initTabLayout())

        self.setLayout(object_layout)
        return

    def initTabLayout(self):
        self.tab_layout = QTabWidget()
        self.tab_layout.resize(150, 100)

        # set tab e-mail text.
        self.tab_layout.addTab(self.setEmailTextLayout(), 'Corpo do e-mail')

        # set tab e-mail user.
        self.tab_layout.addTab(self.setUserEMailLayout(), 'Usuário do e-mail')

        # set tab e-mails have sent.
        self.tab_layout.addTab(self.setEmailsHaveSentLayout(), 'E-mails enviados')

        return self.tab_layout

    def setUserEMailLayout(self):
        groupBox = QGroupBox(self)
        groupBox.setObjectName(u"groupBox")
        groupBox.setTitle('Conexão')
        groupBox.setGeometry(QRect(110, 10, 321, 191))

        # set username field.
        label_1 = QLabel(groupBox)
        label_1.setObjectName(u"label_1")
        label_1.setText('Usuário:')
        label_1.setGeometry(QRect(230, 30, 47, 13))
        self.username = QLineEdit(groupBox)
        self.username.setObjectName(u"username")
        self.username.setGeometry(QRect(280, 30, 125, 20))
        self.username.setText(self.email_struct['username'])

        # set password field.
        label_2 = QLabel(groupBox)
        label_2.setObjectName(u"label_2")
        label_2.setText('Senha:')
        label_2.setGeometry(QRect(230, 70, 47, 13))
        self.password = QLineEdit(groupBox)
        self.password.setObjectName(u"password")
        self.password.setGeometry(QRect(280, 70, 125, 20))
        self.password.setEchoMode(QLineEdit.Password)

        # set SMTP Choice filed.
        label_3 = QLabel(groupBox)
        label_3.setObjectName(u"label_3")
        label_3.setText('Servidor:')
        label_3.setGeometry(QRect(230, 110, 47, 13))
        self.smtp_choice = QComboBox(groupBox)
        self.smtp_choice.addItem("Yahoo")
        self.smtp_choice.addItem("GMail")
        self.smtp_choice.setObjectName(u"smtp_choice")
        self.smtp_choice.setGeometry(QRect(280, 110, 75, 22))
        self.smtp_choice.setCurrentText(self.email_struct['smtp_choice'])

        # set button to send the e-mails
        send_button = QPushButton(groupBox)
        send_button.setObjectName(u"send_button")
        send_button.setText('Enviar')
        send_button.setGeometry(QRect(280, 150, 75, 23))
        send_button.clicked.connect(self.clickOnSend)

        return groupBox

    # def setUserEMailLayout(self):
    #     hlayout1 = QHBoxLayout()
    #     # set user field.
    #     hlayout1.addWidget(QLabel('Usuário:'), alignment = Qt.AlignRight)
    #     self.username = QLineEdit()
    #     self.username.setText(self.email_struct['username'])
    #     #self.username.resize(50, 10)
    #     #self.username.setText('')
    #     hlayout1.addWidget(self.username, alignment = Qt.AlignLeft)
    #
    #     # set password field.
    #     hlayout2 = QHBoxLayout()
    #     hlayout2.addWidget(QLabel('Senha:'), alignment = Qt.AlignRight)
    #     self.password = QLineEdit()
    #     self.password.setEchoMode(QLineEdit.Password)
    #     #self.password.resize(5, 100)
    #     #self.password.setText('')
    #     hlayout2.addWidget(self.password, alignment = Qt.AlignLeft)
    #
    #     # set SMTP Choice filed.
    #     hlayout3 = QHBoxLayout()
    #     hlayout3.addWidget(QLabel('Servidor:'), alignment=Qt.AlignHCenter)
    #     self.smtp_choice = QComboBox()
    #     self.smtp_choice.addItems(["Yahoo","GMail"])
    #     self.smtp_choice.resize(1,50)
    #     hlayout3.addWidget(self.smtp_choice, Qt.AlignHCenter)
    #
    #     # set button to send the e-mails
    #     hlayout4 = QHBoxLayout()
    #     send_button = QPushButton("Envia")
    #     send_button.clicked.connect(self.clickOnSend)
    #     hlayout4.addWidget(send_button, alignment = Qt.AlignCenter)
    #
    #     # set the user email layout.
    #     vlayout = QVBoxLayout()
    #     vlayout.addStretch()
    #     vlayout.addLayout(hlayout1)
    #     vlayout.addLayout(hlayout2)
    #     vlayout.addLayout(hlayout3)
    #     vlayout.addLayout(hlayout4)
    #     vlayout.addStretch()
    #     group = QGroupBox(self)
    #     group.setTitle('Conexão')
    #     group.setLayout(vlayout)
    #     group.setGeometry(QRect(0, 0, 0, 0))
    #     #group.setGeometry(QRect(200, 200, 300, 300))
    #
    #     return group

    def setEmailTextLayout(self):
        object_layout = QGridLayout()

        # set subject field.
        object_layout.addWidget(QLabel('Assunto:'), 0, 0)
        self.subject = QLineEdit()
        # self.subject.resize(5, 10)
        self.subject.setText(self.email_struct['subject'])
        object_layout.addWidget(self.subject, 0, 1)

        # set text body field.
        object_layout.addWidget(QLabel('Corpo do e-mail:'), 1, 0)
        self.body = QTextEdit()
        # self.body.resize(5, 200)
        self.body.setText(self.email_struct['body'])
        object_layout.addWidget(self.body, 1, 1, 15, 1)

        # set attach button.
        attach_button = QPushButton("Anexa")
        attach_button.clicked.connect(self.clickOnAttach)
        object_layout.addWidget(attach_button, 20, 0)

        # set detach button.
        detach_button = QPushButton("Desanexa")
        detach_button.clicked.connect(self.clickOnDetach)
        object_layout.addWidget(detach_button, 21, 0)

        # set save button.
        save_button = QPushButton("Salva")
        save_button.clicked.connect(self.clickOnSave)
        object_layout.addWidget(save_button, 22, 0)

        # set list files field.
        self.files_list = QTableWidget()
        self.files_list.setColumnCount(1)
        self.files_list.setColumnWidth(0, 550)
        header = ['Arquivos a anexar']
        self.files_list.setHorizontalHeaderLabels(header)
        self.files_list.insertRow(0)
        for file in self.email_struct['files']:
            self.insertRowFileList(file)
        object_layout.addWidget(self.files_list, 20, 1, 15, 1)

        # set the text email layout.
        group = QGroupBox(self)
        group.setLayout(object_layout)

        return group

    def setEmailsHaveSentLayout(self):
        object_layout = QGridLayout()

        # set list emails have sent field.
        self.emails_have_sent = QTableWidget()
        self.emails_have_sent.setColumnCount(1)
        self.emails_have_sent.setColumnWidth(0, 550)
        header = ['Emails eviados']
        self.emails_have_sent.setHorizontalHeaderLabels(header)
        object_layout.addWidget(self.emails_have_sent, 0, 0)

        # set the text email layout.
        group = QGroupBox(self)
        group.setLayout(object_layout)
        return group

    def clickOnDetach(self):
        indexes = self.files_list.selectionModel().selectedRows()
        for index in sorted(indexes):
            self.files_list.removeRow(index.row())
        return

    def clickOnAttach(self):
        path = QFileDialog.getOpenFileName(self, 'Abrir', getenv('HOME'), '*.*')
        file = path[0]
        self.insertRowFileList(file)
        return

    def insertRowFileList(self, file):
        rows = self.files_list.rowCount()
        row = rows - 1
        self.files_list.insertRow(row)
        item = QTableWidgetItem()
        item.setData(Qt.DisplayRole, file)
        self.files_list.setItem(row, 0, item)
        return

    def clickOnSave(self):
        self.saveEmail()
        return

    def clickOnSend(self):
        self.saveEmail()
        if len(self.username.text().strip()) == 0:
            QMessageBox.about(self, 'Alerta!', 'Entrar com o usuário!')
        elif len(self.password.text().strip()) == 0:
            QMessageBox.about(self, 'Alerta!', 'Entrar com a senha!')
        else:
            self.tab_layout.setCurrentIndex(2)  # Change to visualizae the e-mails is sending
            try:
                self.email_manager.setUsername(self.username.text())
                self.email_manager.setPassword(self.password.text())
                self.email_manager.setSubject(self.subject.text())
                self.email_manager.setBody(self.body.toPlainText())
                self.email_manager.setFiles(self.getFileList())
                self.email_manager.setEmailsHaveSentList(self.emails_have_sent)
                self.email_manager.openEmailList()
                self.email_manager.openEmailServer(self.smtp_choice.currentText())
                self.email_manager.sendEmails()
                msg = '{} e-mails enviados em {:.2f} minutos.'.format(self.email_manager.total_hast_sent, self.email_manager.total_time)
                QMessageBox.about(self, 'Alerta de Fim de Processamento', msg)
            except Exception as e:
                QMessageBox.about(self, 'Alerta de Erro', str(e))
        return

if __name__ == "__main__":
    app = QApplication(sys.argv)
    send_cv_out = SendCVOut()
    send_cv_out.show()
    app.exit(app.exec_())
