# -*- coding: utf-8 -*-
"""
Sends email with attachment to distribution list
"""

import smtplib
import pathlib
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.header import Header
from email.utils import formataddr

import pandas as pd
from decouple import config

EMAIL_USER = config('EMAIL_USER')
EMAIL_PASSWORD = config('EMAIL_PASSWORD')
BASE_PATH = pathlib.Path.cwd()

ATTACHMENT_NAME = 'attachment.pdf'
SUBJECT = 'My Subject'
DISTRIBUTION_LIST = 'distribution_list.csv'
FROM_NAME = 'Cristobal Aguirre'


def main():
    """ parse distribution list and send email to everyone """
    mailing_list = parse_distribution_list()
    for entry in mailing_list:
        print(f'Sending email to {entry.email}')
        send_email(entry)


def parse_distribution_list(filename=DISTRIBUTION_LIST):
    """ parses distribution csv to list of namedtuples """
    dist_list = BASE_PATH / filename
    df = pd.read_csv(dist_list)

    # Assumes this is the given order
    df.columns = ['email', 'first_name', 'last_name']
    return df.itertuples(index=False)


def add_attachment(filename, msg):
    """
    Add statement attachment to email
    """
    file_dir = BASE_PATH / filename

    with open(file_dir, 'rb') as fp:
        att = MIMEApplication(fp.read(), _subtype="pdf")
    att.add_header('Content-Disposition', 'attachment', filename=filename)
    msg.attach(att)


def send_email(to):
    """
    send email
    :param to: namedtuple with keys: email, first_name, last_name
    """
    server = smtplib.SMTP_SSL('smtp.gmail.com')

    msgRoot = MIMEMultipart('related')
    email_account = EMAIL_USER
    # Send from accounts@deetken.com alias
    msgRoot['From'] = formataddr((str(Header(FROM_NAME, 'utf-8')), EMAIL_USER))
    msgRoot['To'] = to.email
    # For testing
    # msgRoot['To'] = 'compliance@deetken.com'
    msgRoot['Subject'] = SUBJECT
    msgRoot.preamble = ''

    # Encapsulate the plain and HTML versions of the message body in an
    # 'alternative' part, so message agents can decide which they want to display.
    msgAlternative = MIMEMultipart('alternative')
    msgRoot.attach(msgAlternative)

    alternative_text = f"""
        Dear {to.first_name},
        
        This is some text\n
        And some more text\n
        \n
        Here's some more text\n
        \n
        Best regards,\n
        \n
        Cristobal
    """

    msgText = MIMEText(alternative_text, 'plain')
    msgAlternative.attach(msgText)

    html = email_body(to)

    # Record the MIME types of both parts - text/plain and text/html.
    body = MIMEText(html, 'html')

    msgAlternative.attach(body)

    # attach pdf
    add_attachment(ATTACHMENT_NAME, msgRoot)

    server.ehlo()
    server.login(email_account, EMAIL_PASSWORD)
    server.send_message(msgRoot)
    server.quit()


def email_body(to):
    """

    :param to: namedtuple with keys: email, first_name, last_name
    :return: HTML email body
    """
    body = """
    <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
    <html xmlns="http://www.w3.org/1999/xhtml">
      <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
        <title>Deetken Impact Quarterly Statements</title>
            <style type="text/css">
                body{width:100% !important; -webkit-text-size-adjust:100%; -ms-text-size-adjust:100%; margin:0; padding:0;}
                .ExternalClass {width:100%;}
                .ExternalClass, .ExternalClass p, .ExternalClass span, .ExternalClass font, .ExternalClass td, .ExternalClass div {line-height: 100%;}
                #backgroundTable {margin:0; padding:0; width:100% !important; line-height: 100% !important;}
                h1, h2, h3, h4, h5, h6 {color: black !important;}
                h1 a, h2 a, h3 a, h4 a, h5 a, h6 a {color: blue !important;}
                table td {border-collapse: collapse;}
                table { border-collapse:collapse; mso-table-lspace:0pt; mso-table-rspace:0pt; }

            </style>
        </head>
        <body>
                <p>Dear """ + to.first_name + """,</p>
                <p>
                    This is some text
                </p>
                
                <p>
                    And some more text
                </p>
                
                <p>
                    Here's even more text
                </p>
                
                <p>Best regards,</p>
                <p>Cristobal</p>
                
        </body>
    </html>"""
    return body

if __name__ == '__main__':
    main()