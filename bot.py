import email.message, email.policy, email.utils, sys, smtplib, os, ssl
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email import encoders
import pandas as pd
import secret


# 1. Read the data
xl = pd.ExcelFile("./data.xlsx")
df = xl.parse("Sheet1")

for i in range(df.shape[0]):
    print(df.iat[i, 0], df.iat[i, 1])
    text = f"""Tere {df.iat[i, 0]}
    See on test email, mis saadetaks automaatselt.
    Lisasin oma CV ka kaasa.
    Lugupidamisega,
    Axel-Martin Tammekänd
    """

    message = MIMEMultipart()
    message['To'] = df.iat[i, 1]
    message['From'] = secret.email
    message['Subject'] = 'Test email'
    message['Date'] = email.utils.formatdate(localtime=True)
    message['Message-ID'] = email.utils.make_msgid()

    message.attach(MIMEText(text, 'plain'))

    # Open PDF file in binary mode
    with open("./cv.pdf", "rb") as attachment:
        # Add file as application/octet-stream
        # Email client can usually download this automatically as attachment
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

    # Encode file in ASCII characters to send by email    
    encoders.encode_base64(part)

    # Add header as pdf attachment
    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {'cv.pdf'}",
    )

    # Add attachment to message and convert message to string
    message.attach(part)




    with smtplib.SMTP('smtp.gmail.com', 587) as s:
        s.starttls()
        s.login(secret.email, secret.password)
        s.send_message(message)
        s.quit()
    print("Sent")