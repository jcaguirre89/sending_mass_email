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

ATTACHMENT_NAME = 'Deetken Impact Fund Quarterly Report 31 Mar 2019.pdf'
SUBJECT = 'Deetken Impact 2019 Q1 Report on Financial and Impact Performance'
DISTRIBUTION_LIST = 'distribution_list.csv'


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
    msgRoot['From'] = formataddr((str(Header('Alexa Blain', 'utf-8')), 'ablain@deetken.com'))
    msgRoot['To'] = to.email
    # For testing
    # msgRoot['To'] = 'compliance@deetken.com'
    msgRoot['Subject'] = SUBJECT
    msgRoot.preamble = 'Deetken Impact Quarterly Report'

    # Encapsulate the plain and HTML versions of the message body in an
    # 'alternative' part, so message agents can decide which they want to display.
    msgAlternative = MIMEMultipart('alternative')
    msgRoot.attach(msgAlternative)

    alternative_text = f"""
        Dear {to.first_name},
        
        We are pleased to share our latest report on the financial, social and environmental performance of the Deetken Impact Fund.
        
        We would love to see you at our Wednesday, June 5 #LatinasDeliver side event at the Women Deliver conference in Vancouver, which we are hosting together with Pro Mujer and MEDA. It will be a wonderful opportunity to connect with others from around the world interested in gender lens investing and advancing gender equality globally. You can register here.
        
        For more news and updates from the Deetken Impact team, follow us on Twitter.   
        
        As always, thank you for investing with us. 
        
        Best regards,
        
        Alexa
    """

    msgText = MIMEText(alternative_text, 'plain')
    msgAlternative.attach(msgText)

    html = email_body(to)

    # Record the MIME types of both parts - text/plain and text/html.
    body = MIMEText(html, 'html')

    msgAlternative.attach(body)

    logoDir = BASE_PATH / 'logo_email.png'
    with open(logoDir, 'rb') as fp:
        msgImage = MIMEImage(fp.read())

    # Define the image's ID as referenced in the HTML body
    msgImage.add_header('Content-ID', '<logo>')
    msgRoot.attach(msgImage)

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
        .signature td {
            font-size: 12px;
            border-left: solid 2px black;
    padding-left: 5px;
            }
        .signature {
            width: 60%;
            }

            </style>
        </head>
        <body>
                <p>Dear """ + to.first_name + """,</p>
                <p>
                    We are pleased to share our latest report on the financial, social and environmental performance
                    of the Deetken Impact Fund.
                </p>
                
                <p>
                    We would love to see you at our Wednesday, June 5 #LatinasDeliver side event at the Women
                    Deliver conference in Vancouver, which we are hosting together with Pro Mujer and MEDA.
                    It will be a wonderful opportunity to connect with others from around the world
                    interested in gender lens investing and advancing gender equality globally.
                    You can register
                    <a href="https://www.eventbrite.com/e/latinasdeliver-networking-cocktail-tickets-61037484760">here</a>.
                </p>
                
                <p>
                    For more news and updates from the Deetken Impact team, follow us on
                    <a href="https://twitter.com/DeetkenImpact">Twitter</a>.
                </p>
                
                <p>As always, thank you for investing with us.</p>
                <p>Best regards,</p>
                <p>Alexa</p>
                

                    <table class="signature">
                    <tr>
                        <td style="width: 50%; text-align:center; border-left: 0;" rowspan= "4"><img src="cid:logo" style="width: 150px; height: 75px;outline:none; text-decoration:none; -ms-interpolation-mode: bicubic;"></td>
                                <td style="width: 50%;">Suite 501 - 1755 W. Broadway</td>
                    </tr>
                    <tr>

                                <td style="width: 50%;">Vancouver, BC V6J 4S2</td>
                    </tr>
                    <tr>

                                <td style="width: 50%;">ablain@deetken.com</td>
                    </tr>
                    <tr>

                                <td style="width: 50%;">Office: +1 (604) 731-4424</td>
                    <tr>
                        <td style="width: 50%; text-align:center; border-left: 0;" rowspan= "3"><strong>Make an Impact with your investment.</strong></td>                           
                                <td style="width: 50%;">Fax: +1 (604) 736-2246</td>
                    </tr>
                    <tr>

                                <td style="width: 50%;">Web: <a href="www.deetkenimpact.com">www.deetkenimpact.com</a></td>
                    </tr>
                    <tr>
                        <td style="width: 50%; text-align:center; border-left: 0;" rowspan= "3"></td>
                        <td style="width: 50%;"> </td>
                    </tr>
                </table>
                <p style="font-size: 12px;">The information in this email is confidential and may be
                legally privileged. Access to this email by anyone other than the intended
                addressee is unauthorized. If you are not the intended recipient of this
                message, any review, disclosure, copying, distribution, retention, or any
                action taken or omitted to be taken in reliance on it is prohibited and may
                be unlawful. If you are not the intended recipient, please reply to or forward
                a copy of this message to the sender and delete the message, any attachments,
                and any copies thereof from your system.</p>
        </body>
    </html>"""
    return body

if __name__ == '__main__':
    main()