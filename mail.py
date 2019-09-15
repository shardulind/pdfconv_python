import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

fromAddress = ""  # tuza gmailcha email id
toAddress = "asthorat@mitaoe.ac.in"

msg = MIMEMultipart()
msg['From'] = fromAddress
msg['To'] = toAddress
msg['Subject'] = "Mail throuh Python"
body="here comes body of mail"
msg.attach(MIMEText(body, 'plain'))

filename="test.pdf"
attachment = open(filename, "rb")
p = MIMEBase('application', 'octet-stream')
p.set_payload((attachment).read())
encoders.encode_base64(p)
p.add_header('Content-Disposition', "attachment; filename= %s" % filename)
msg.attach(p)
s = smtplib.SMTP('smtp.gmail.com', 587)
s.starttls()
s.login(fromAddress, "")  # gmailcha tuzha password
text = msg.as_string()
s.sendmail(fromAddress, toAddress, text)
s.quit()
