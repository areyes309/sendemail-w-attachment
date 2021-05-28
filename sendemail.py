import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.header import Header
from email.utils import formataddr
from email import encoders
from string import Template

# Path for Attachment and Email Template
path = r'C:\Users\Desktop\data.csv'
msgT = r'C:\Users\Desktop\msg.txt'

def read_template(filename):
    with open(filename, 'r', encoding='utf-8') as template_file: template_file_content = template_file.read()
    return Template(template_file_content)

# Set up who to send email to and from
emailfrom = formataddr((str(Header('FrontDesk', 'utf-8')), 'office@test.com'))
emailto = ['info@test.com']
emailcc = ['admin@test.com']
toAddresses = emailto + emailcc

# Continue setup for email
msg = MIMEMultipart()
msg['From'] = emailfrom
msg['To'] = ", ".join(emailto)
msg['Cc'] = ", ".join(emailcc)
msg["Subject"] = 'Report for Han Solo'

# Use a simple text template for the email
message_template = read_template(msgT)
html = message_template.substitute(PERSON_NAME='HanSolo')
msg.attach(MIMEText(html,"html","utf-8"))

# Email Attachment
part = MIMEBase('application', "octet-stream")
part.set_payload(open(path, "rb").read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', 'attachment; filename="data.csv"')
msg.attach(part)

# Send the email
s = smtplib.SMTP("domain.com",25)
s.sendmail(msg['From'], toAddresses, msg.as_string())
s.quit()
