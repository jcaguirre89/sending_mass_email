# Send mass email
grabs a distribution list CSV file structured as `email,first_name,last_name` and sends an email to each person in the list, also attaching a pdf file. It's a generic boilerplate I use when I need to send emails to distribution lists and a) I don't care too much about formatting and/or b) I need to add an attachment and can't do that with mailchimp. 

## Usage
requirements are `pandas` and [`python-decouple`](https://github.com/henriquebastos/python-decouple), which requires a `.env` file present in the working directory with the email credentials to use. Example:
```
EMAIL_USER=example@domain.com
EMAIL_PASSWORD=password1234
```
The HTML is handled within the `send_email.py` module, and perhaps a simple enhacement would be creating a separate html file and reading it in, if formatting was an issue. to run, simply collect the required files (the distribution list and the attachement), adjust the text to suit your needs and run `python send_email.py`.
