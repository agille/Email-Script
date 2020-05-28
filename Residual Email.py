# Residual Email Script
# Alexander Gille 2020

import csv
import datetime
import dateutil.relativedelta
import os

import smtplib
import yagmail

import tkinter as tk
from tkinter import *
from tkinter import simpledialog
from tkinter import filedialog

root = tk.Tk()
root.withdraw()

#Prompt to enter users email credentials
your_email = simpledialog.askstring("Email", "Enter your Email")
your_password = simpledialog.askstring("Password", "Enter your Password", show="*")

#Date that will be used in the email subject line and body. Can be adjusted depending on month sending.
#Ex. If current month is June and (months-1), Email will display May.
current_date = datetime.datetime.now()
one_month_prior = current_date+dateutil.relativedelta.relativedelta(months=-1)

#Subject line for the email
subject_line = f'Pineapple Payments {one_month_prior.strftime("%B")} {one_month_prior.year} Residual File'

LNAMES = []
FNAMES = []
EMAILS = []
FILES = []

yag = yagmail.SMTP(your_email, your_password)

#user prompt to select csv file containing recipient name, email, and file name
email_data = filedialog.askopenfilename(filetypes=[('.csv', '.csv')],
                                        title='Select the Email Data file')

#user prompt to select text file containing the body of the email
txt_file = filedialog.askopenfilename(filetypes=[('.txt', '.txt')],
                                      title='Select the EMail Template')

#user prompt to select the folder containing the files read in from the csv
dir_name = filedialog.askdirectory(title='Select Folder Containing Reports')
os.chdir(dir_name)


class EmailAttachmentNotFoundException(Exception):
    pass


def send_email():

    with open(email_data) as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        line_count = 0

        try:
            for row in csv_reader:
                last_name = row[0]
                first_name = row[1]
                email = row[2]
                file = row[3]

                if not os.path.isfile(file):
                    raise EmailAttachmentNotFoundException('The attachment file "{}" was not found in the directory'
                                                           .format(file))

                LNAMES.append(last_name)
                FNAMES.append(first_name)
                EMAILS.append(email)
                FILES.append(file)

                line_count += 1

        except EmailAttachmentNotFoundException as a:
            print(str(a))
            input('Press Enter to exit')
            sys.exit(1)

    with open(txt_file) as f:
        email_template = f.read()

    try:
        for first_name, last_name, email, file in zip(FNAMES, LNAMES, EMAILS, FILES):
            txt = email_template.format(first_name=first_name,
                                        last_name=last_name,
                                        month=one_month_prior.strftime("%B"),
                                        year=one_month_prior.year)

            yag.send(to=email,
                     subject=subject_line,
                     contents=[txt, file])

    #Lets user know when an authentication error relating to the email credentials occurs
    except smtplib.SMTPAuthenticationError:
        print('Incorrect Email or Password entered')
        input('Press Enter to exit')
        sys.exit(1)

#Lets user know the emails were sent successfully
try:
    send_email()
    print('Email(s) sent successfully')
    input('Press Enter to exit')
    sys.exit(1)

#If an error occurs, displace the exception as well as the files that were read in for debugging purposes
except Exception as e:
    print(e, FILES)

