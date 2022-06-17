import pandas as pd
from email.mime.text import MIMEText
import smtplib
from bs4 import BeautifulSoup
import PySimpleGUI as sg

pad_right = (10, 10)

image_column = [
    [sg.Image(filename='LogoWhite.png', key='-LOGO-')],
]

left_column = [
    [sg.Text('SMTP server:', size=(15, 1), pad=(0, 15)), sg.InputText(key="-SMTP_SERVER-", pad=(0, 15))],
    [sg.Text('PORT:', size=(15, 1), pad=(0, 15)), sg.InputText(key="-SMTP_PORT-", pad=(0, 15))],
    [sg.Text('Email:', size=(15, 1), pad=(0, 15)), sg.InputText(key="-EMAIL-", pad=(0, 15))],
    [sg.Text('Mot de passe:', size=(15, 1), pad=(0, 15)), sg.InputText(key="-PASSWORD-", pad=(0, 15))]
]
top_right_column = [
    [sg.Text('Fichier excel:', size=(10, 1), pad=pad_right), sg.FileBrowse('Parcourir', key='-EXCEL-', pad=pad_right)],
    [sg.Text('Template HTML:', size=(10, 1), pad=pad_right), sg.FileBrowse('Parcourir', key='-HTML-', pad=pad_right)],
    [sg.Text('Titre', size=(10, 1), pad=pad_right), sg.InputText(key='-GENDER-', pad=pad_right)]
]
middle_right_column = [
    [sg.Checkbox('Ajout du nom', default=False, key='-IS_NAME-', pad=pad_right)]
]
bottom_right_column = [
    [sg.Text('Email subject:', size=(10, 1), pad=pad_right), sg.InputText(key='-SUBJECT-', pad=pad_right)],
]
submit_right_column = [
    [sg.Submit('Valider'), sg.Cancel('Annuler')]
]
right_column = [
    [sg.Column(top_right_column)],
    [sg.Column(middle_right_column, justification='center')],
    [sg.Column(bottom_right_column)],
    [sg.Column(submit_right_column, justification='center')]
]

layout = [
   [sg.Image(filename='LogoWhite.png', key='-LOGO-', subsample=10,pad=(340,0))],
   sg.Column(left_column, vertical_alignment='center'),
   sg.VSeparator(),
   sg.Column(right_column, vertical_alignment='center')],

window = sg.Window('Mail sender', layout,size=(800,400))


def send_email(email_to, email_from, email_subject, data):
    msg = MIMEText(data, 'html')
    msg['Subject'] = email_subject
    msg['To'] = email_to
    msg['From'] = email_from
    mail = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    mail.starttls()
    mail.login(SMTP_USERNAME, SMTP_PASSWORD)
    mail.sendmail(email_from, email_to, msg.as_string())
    mail.quit()


def open_html_and_update_data_for_mail(values):
    is_name = values['-IS_NAME-']
    df = pd.read_excel(values['-EXCEL-'])
    clients = df.values.tolist()
    with open(values['-HTML-'], 'r') as f:
        contents = f.read()
        soup = BeautifulSoup(contents, 'lxml')
        title = soup.find('h3')

        for c in clients:
            name_and_firstname = c[0] + ' ' + c[1]
            if is_name:
                title.string = values['-GENDER-'] + ' ' + name_and_firstname + ','
            else:
                title.string = values['-GENDER-'] + ','
            email_to = c[2]
            email_from = "doryan@dlsg.be"
            email_subject = values['-SUBJECT-']
            data = str(soup)

            if __name__ == '__main__':
                send_email(email_to, email_from, email_subject, data)


while True:
    event, values = window.read()
    SMTP_SERVER = values['-SMTP_SERVER-']
    SMTP_PORT = values['-SMTP_PORT-']
    SMTP_USERNAME = values['-EMAIL-']
    SMTP_PASSWORD = values['-PASSWORD-']
    if event in (None, 'Cancel'):
        break
    if event == 'Valider':
        open_html_and_update_data_for_mail(values)
