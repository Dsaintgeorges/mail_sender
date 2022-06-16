import pandas as pd
from email.mime.text import MIMEText
from datetime import date
import smtplib
from bs4 import BeautifulSoup

SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587
SMTP_USERNAME = ""
SMTP_PASSWORD = ""


class client(object):
    def __init__(self, name, firstname, email):
        self.name = name
        self.firstnam = firstname
        self.email = email


df = pd.read_excel('test_script_pytho.xlsx')
clients = df.values.tolist()

with open('template_color.html', 'r') as f:
    contents = f.read()

    soup = BeautifulSoup(contents, 'lxml')
    title = soup.find('h3')
    for c in clients:
        title.string = c[0] + ' ' + c[1]
        print(c[2])
        EMAIL_TO = c[2]
        EMAIL_FROM = ""
        EMAIL_SUBJECT = "Demo Email :"

        DATE_FORMAT = "%d/%m/%Y"
        EMAIL_SPACE = ", "

        DATA = str(soup)


        def send_email():
            msg = MIMEText(DATA, 'html')
            msg['Subject'] = EMAIL_SUBJECT + " %s" % (date.today().strftime(DATE_FORMAT))
            msg['To'] = EMAIL_TO
            msg['From'] = EMAIL_FROM
            mail = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
            mail.starttls()
            mail.login(SMTP_USERNAME, SMTP_PASSWORD)
            mail.sendmail(EMAIL_FROM, EMAIL_TO, msg.as_string())
            mail.quit()


        if __name__ == '__main__':
            send_email()

    print(soup.h3)
