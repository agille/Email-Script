# Residual Email Script
# Alexander Gille 2019

import csv
import yagmail

# yagmail.register('username', 'password')
# register the email to your computer's keyring

LNAMES = []
FNAMES = []
EMAILS = []
FILES = []


def send_email():

    with open('email_data.csv') as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        line_count = 0

        for row in csv_reader:
            lname = row[0]
            fname = row[1]
            email = row[2]
            file = row[3]

            LNAMES.append(lname)
            FNAMES.append(fname)
            EMAILS.append(email)
            FILES.append(file)

            line_count += 1

    with open('email_template.txt') as f:
        email_template = f.read()

    yag = yagmail.SMTP('agille@pineapplepayments.com')

    for fname, lname, email, file in zip(FNAMES, LNAMES, EMAILS, FILES):
        txt = email_template.format(first_name=fname,
                                    last_name=lname)
        print(txt)
        yag.send(to=email,
                 subject='April Residuals',
                 contents=[txt, file])

        print(fname, lname, email, file)


send_email()
