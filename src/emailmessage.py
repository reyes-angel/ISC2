from email import encoders
from email.message import EmailMessage
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from googleapiclient.errors import HttpError
from config import Google
import base64
import mimetypes
import os
from config import Config

class messaging:
    # def send_mail(self):
    #     g = Google()
    #     service = g.mail_service()

    #     try:
    #         message = EmailMessage()
    #         message.set_content('This is automated email')
    #         message['To'] = 'reyes.angel@gmail.com'
    #         message['from'] = '(ISC)² Silicon Valley Membership Director <membership@isc2-siliconvalley-chapter.org>'
    #         message['Subject'] = 'Automated draft'

    #         encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode()

    #         create_message = {'raw': encoded_message}
            
    #         send_message = (service.users().messages().send(userId="me", body=create_message).execute())
  
    #     except HttpError as error:
    #         print(f'An error occurred: {error}')
    #         send_message = None
        

    def send_message_and_attachments(self, member_name, member_email, attachment_file, meeting_date):
        c = Config()
        g = Google(c)
        mail_service = g.mail_service()
        
        file_attachments = [attachment_file]

        message_body_text = "Hello %s, Thank you for having joined the (ISC)² Silicon Valley Chapter Meeting on %s."\
            "<br><br> Your CPE certificate is attached.<br><br>"\
            "Note that the chapter will submit chapter meeting CPE's on your behalf." % (member_name, meeting_date.strftime("%B %d, %Y"))

        message = MIMEMultipart()
        message['to'] = member_email
        message['from'] = '(ISC)² Silicon Valley Membership Director <membership@isc2-siliconvalley-chapter.org>'
        message['subject'] = "CPE Certificate - %s" % (meeting_date.strftime("%B %Y"))
        message.attach(MIMEText(message_body_text, 'html'))

        for attachment in file_attachments:
            content_type, encoding = mimetypes.guess_type(attachment)
            main_type, sub_type = content_type.split("/", 1)
            file_name = os.path.basename(attachment)

            f = open(attachment, "rb")
            
            myfile = MIMEBase(main_type, sub_type)
            myfile.set_payload(f.read())
            myfile.add_header("Content-Disposition", "attachment", filename=file_name)
            encoders.encode_base64(myfile)

            f.close()

            message.attach(myfile)
        
        raw_string = base64.urlsafe_b64encode(message.as_bytes()).decode()

        message = mail_service.users().messages().send(userId="me", body={"raw": raw_string}).execute()
        print(message)