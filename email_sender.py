from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from smtplib import SMTP

DEFAULT_SUBJECT = "Merged cd test summary in xlsx"
DEFAULT_MESSAGE = "See the attached excel file"

MIME_TYPE_XLSX = {
    "maintype": "application",
    "subtype": "vnd.openxmlformats-officedocument.spreadsheetml.sheet",
}


def send_email(file,
               filename,
               to_email,
               subject=DEFAULT_SUBJECT,
               message=DEFAULT_MESSAGE):
    from_email = "QA Dashboard <noreply@nuance.com>"
    domain = "smtp.nuance.com"

    with SMTP(domain) as smtp:
        msg = MIMEMultipart()
        msg["Subject"] = subject
        msg["From"] = from_email
        msg["To"] = to_email
        msg.attach(MIMEText(DEFAULT_MESSAGE))
        with open(file, 'rb') as f:
            xlsxpart = MIMEApplication(f.read(),
                                       _subtype=MIME_TYPE_XLSX["subtype"])
            xlsxpart.add_header('Content-Disposition',
                                f'attachment; filename={filename}.xlsx')
            msg.attach(xlsxpart)
        smtp.send_message(msg)
