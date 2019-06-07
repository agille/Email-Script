# Residual Email Script
# Alexander Gille 2019

import csv
import datetime
import sys

import smtplib
import yagmail

import tkinter as tk
from tkinter import filedialog
from tkinter import simpledialog

root = tk.Tk()
root.withdraw()

password = simpledialog.askstring("Password", "Enter password:", show='*')

current_date = datetime.datetime.now()

subject_line = f'Pineapple Payments {current_date.strftime("%B")} {current_date.year} Residual File'

LNAMES = []
FNAMES = []
EMAILS = []
FILES = []

yag = yagmail.SMTP('agille@pineapplepayments.com', password)

email_data = filedialog.askopenfilename(filetypes=[('.csv', '.csv')],
                                        title='Select the Email Data file')


def send_email():

    with open(email_data) as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        line_count = 0

        for row in csv_reader:
            last_name = row[0]
            first_name = row[1]
            email = row[2]
            file = row[3]

            LNAMES.append(last_name)
            FNAMES.append(first_name)
            EMAILS.append(email)
            FILES.append(file)

            line_count += 1

    try:
        for first_name, last_name, email, file in zip(FNAMES, LNAMES, EMAILS, FILES):
            txt = f'Dear {first_name} {last_name},\n\n'\
                  'This email is a test of the python script that will be used for sending residual files.\n\n'\
                  'Thank you,\n\n'\
                  'Alex Gille\n\n' \
                  '<font face = "sans-serif">'\
                  '<p style="color:rgb(7, 55, 99);">'\
                  '<b>Pineapple Payments</b>\n'\
                  '11 Stanwix Street, Suite 1202\n'\
                  'Pittsburgh, PA 15222\n'\
                  'www.pineapplepayments.com'\
                  '</p>'\
                  '</font>'\
                  '<img src="https://drive.google.com/uc?export=view&id=1xjVA6vA1JZfLmkEq1KonISKxbrmUudH9"'\
                  'width="250" height="150"></img>'\
                  '<font face = "verdana">'\
                  '<p style="color:rgb(102, 102, 102);">'\
                  '<b>Disclaimer</b>\n'\
                  '<small>'\
                  'The information contained in this communication from '\
                  'the sender is confidential. It is intended solely for use '\
                  'by the recipient and others authorized to receive it. If '\
                  'you are not the recipient, you are hereby notified that '\
                  'any disclosure, copying, distribution or taking action in '\
                  'relation of the contents of this information is strictly '\
                  'prohibited and may be unlawful.'\
                  '</font>'\
                  '</p>'\
                  '</small>'

            yag.send(to=email,
                     subject=subject_line,
                     contents=[txt, file])

            print("Email(s) sent successfully")
            input("Press Enter to exit")
            sys.exit(1)

    except smtplib.SMTPAuthenticationError:
        print("Incorrect Email password entered")
        input("Press Enter to exit")
        sys.exit(1)


send_email()
