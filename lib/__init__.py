from os import getcwd
import time
import pandas as pd
import numpy as np
from datetime import datetime

# Package to connect to an email server.
import smtplib, ssl
# Packages to send an email
from os.path import basename
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.utils import formatdate

# Packages to make the data entry.
from PyQt5.QtWidgets import (QTableWidgetItem, QMessageBox)
from PyQt5.QtCore import (Qt)

class EmailManager():
    def __init__(self):
        self.total_time = 0
        self.total_hast_sent = 0
        self.db_path = getcwd() + '\\DB\\'
        self.emails_have_sent = None
        return

    def setUsername(self, username):
        self.username = username
        return

    def setPassword(self, password):
        self.password = password
        return

    def setSubject(self, subject):
        self.subject = subject
        return

    def setBody(self, body):
        self.body = body
        return

    def setPath(self, path):
        self.path = path
        return

    def setFiles(self, files):
        self.files = files
        return

    def setEmailsHaveSentList(self, emails_have_sent_list):
        self.emails_have_sent_list = emails_have_sent_list
        return

    def showEmailHasSent(self, email):
        if self.emails_have_sent_list != None:
            rows = self.emails_have_sent_list.rowCount()
            row = 0 if (rows - 1) < 0 else (rows - 1)
            self.emails_have_sent_list.insertRow(row)
            item = QTableWidgetItem()
            item.setData(Qt.DisplayRole, email)
            self.emails_have_sent_list.setItem(row, 0, item)
        return

    def connectToGMail(self):
        try:
            context = ssl.create_default_context()
            server = smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context)
            server.ehlo()
            server.login(self.username, self.password)
        except Exception as e:
            if str(e).find('Password not accepted') != -1:
                raise Exception('Usuário ou Senha Inválido!')
            else:
                raise Exception(e)
        return server

    def connectToYahoo(self):
        # Set up the e-mails server
        try:
            server = smtplib.SMTP("smtp.mail.yahoo.com", 587)
            server.ehlo()
            server.starttls()
            server.login(self.username, self.password)
        except Exception as e:
            if str(e).find('Connection unexpectedly closed') != -1:
                raise Exception('Usuário ou Senha inválido!')
            else:
                raise Exception(e)
        return server

    def openEmailServer(self, smtp_choice):
        try:
            self.smtp_choice = smtp_choice
            if self.smtp_choice == 'Yahoo':
                self.server = self.connectToYahoo()
            elif self.smtp_choice == 'GMail':
                self.server = self.connectToGMail()
            else:
                self.server = None
        except Exception as e:
            raise Exception(e)
        return

    def formatEmail(self):
        # Set up the email
        message = MIMEMultipart()
        message['From'] = self.username
        message['Date'] = formatdate(localtime=True)
        identifier = datetime.now().strftime('%Y%m%d%H%M%S') + str(np.random.randint(100, 200))
        subject_identified = f"{self.subject} #{identifier}"
        message['Subject'] = subject_identified

        message.attach(MIMEText(self.body, 'plain'))

        # Attache the CV in the email.
        for file in self.files:
            with open(file, "rb") as fil:
                part = MIMEApplication(fil.read(), Name=basename(file))
            # After the file is closed
            part['Content-Disposition'] = 'attachment; filename="%s"' % basename(file)
            message.attach(part)
        return message

    def openEmailList(self):
        # Open the file that contains all emails to be sent.
        try:
            email_list = pd.read_csv(self.db_path + 'emails.csv', sep=';')
        except Exception as e:
            raise Exception(e)

        # The file that control the email has sent.
        try:
            self.emails_have_sent = pd.read_csv(self.db_path + 'emails_have_sent.csv', sep=';')
        except:
            self.emails_have_sent = pd.DataFrame(columns=['email', 'date'])


        # The file that control the email has not sent.
        try:
            self.emails_have_not_sent = pd.read_csv(self.db_path + 'emails_have_not_sent.csv', sep=';')
        except:
            self.emails_have_not_sent = pd.DataFrame(columns=['email', 'date', 'erro'])

        self.emails = email_list['email'].tolist()
        return

    def sendEmails(self):
        start = time.time()
        MAX = 10
        self.total_hast_sent = 0
        total_control = 0
        for To in self.emails:
            is_there_the_email = len(self.emails_have_sent.loc[self.emails_have_sent['email'] == To]) == 0
            if is_there_the_email:
                try:
                    if total_control >= MAX:
                        t = 240
                        total_control = 0
                        self.server.close()
                        time.sleep(t)  # seconds
                        self.server = self.openEmailServer(self.smtp_choice)
                    # Send the e-mail.
                    message = self.formatEmail()
                    message['To'] = To
                    text = message.as_string()
                    self.server.sendmail(self.username, To, text)
                    # Input the email has sent in a control file.
                    new_row = {'email': To, 'date': datetime.now().strftime('%d/%m/%Y')}
                    self.emails_have_sent = self.emails_have_sent.append(new_row, ignore_index=True)
                    # Save the email has sent in a control file.
                    self.emails_have_sent.to_csv(self.db_path + 'emails_have_sent.csv', sep=';', index=False)
                    total_control += 1
                    self.total_hast_sent += 1
                    # Show the email has sent.
                    self.showEmailHasSent(To)
                except OSError as e:
                    if str(e).find('please run connect() first') != -1:
                        raise Exception('Erro de conexão!')
                    # Input the email has not sent in a control file.
                    new_row = {'email': To, 'date': datetime.now().strftime('%d/%m/%Y'), 'erro': e}
                    self.emails_have_not_sent = self.emails_have_not_sent.append(new_row, ignore_index=True)
                    # Save the email has not sent in a control file.
                    self.emails_have_not_sent.to_csv(self.db_path + 'emails_have_not_sent.csv', sep=';', index=False)
                except Exception as e:
                    if str(e).find('please run connect() first') != -1:
                        raise Exception('Erro de conexão!!')
                    # Input the email has not sent in a control file.
                    new_row = {'email': To, 'date': datetime.now().strftime('%d/%m/%Y'), 'erro': e}
                    self.emails_have_not_sent = self.emails_have_not_sent.append(new_row, ignore_index=True)
                    # Save the email has not sent in a control file.
                    self.emails_have_not_sent.to_csv(self.db_path + 'emails_have_not_sent.csv', sep=';', index=False)
        end = time.time()
        self.total_time = (end - start) // 60
        return
