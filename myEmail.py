# -*- coding: utf-8 -*-
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr


class MyEmail():
    mail_host = "smtp.163.com"
    mail_user = "iotprint@163.com"
    mail_pass = "139169qw"

    receiver = '1584224172@qq.com'

    def send_email(self, filepath):
        try:
            message = MIMEMultipart()
            message['From'] = formataddr(["小老虎", self.mail_user])
            message['To'] = formataddr(["ZhengZheng", self.receiver])
            message['Subject'] = "公司股票信息"

            files = MIMEApplication(open(filepath, 'rb').read(),'utf-8')
            files.add_header('Content-Disposition', 'attachment', filename=('gbk', '', 'A股公司及同业公司股价表现.docx'))
            message.attach(files)

            server = smtplib.SMTP_SSL("smtp.163.com", 465)
            server.login(self.mail_user, self.mail_pass)
            server.sendmail(self.mail_user, self.receiver, message.as_string())
            print("邮件发送成功")
            return "邮件发送成功"
        except smtplib.SMTPException as e:
            print("Error:无法发送邮件")
            print(e)
            return "Error:无法发送邮件"

email = MyEmail()
email.send_email("./files/A股公司及同业公司股价表现-201872620927.docx")
