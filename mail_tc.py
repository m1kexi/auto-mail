import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
# from email.header import Header
# from email.utils import formataddr


sender = 'xzr@jc-edu.com'
key = 'JCjc2188'

def send_mail(receiver, subject, body, xlFile, xlFileName):
    error_msg = 'success'
    text_part = MIMEText(body)

    excel_part = MIMEApplication(open(xlFile,'rb').read())
    excel_part.add_header('Content-Disposition', 'attachment', filename=xlFileName)

    msg = MIMEMultipart()
    msg.attach(text_part)
    msg.attach(excel_part)

    try:
        msg['Subject'] = subject
        EMAIL_HOST, EMAIL_PORT = 'smtp.exmail.qq.com', 465
        client = smtplib.SMTP_SSL(EMAIL_HOST,EMAIL_PORT)
        # client.ehlo()
        # client.starttls()
        # client.ehlo()
        client.login(sender, key)
        client.sendmail(sender, receiver, msg.as_string())
        client.quit()
    except smtplib.SMTPConnectError as e:
        error_msg = 'connection error'
    except smtplib.SMTPAuthenticationError as e:
        error_msg = 'authentication error'
    except smtplib.SMTPSenderRefused as e:
        error_msg = 'refused to send'
    except smtplib.SMTPRecipientsRefused as e:
        error_msg = 'refused to receive'
    except smtplib.SMTPDataError as e:
        error_msg = 'message was rejected'
    except smtplib.SMTPException as e:
        error_msg = str(e)
    except Exception as e:
        error_msg = str(e)
    return error_msg


# if __name__ == '__main__':
#     send_mail('mikexi001@126.com','测试标题','测试邮件','files/dummy_file1.xlsx','test111.xlsx')