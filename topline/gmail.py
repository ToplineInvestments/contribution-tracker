from pathlib import Path
import logging
from oauth2client import file, client, tools
from googleapiclient.discovery import build
from httplib2 import Http
import base64
import mimetypes
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

logger = logging.getLogger(__name__)


class Gmail:
    # If modifying these scopes, delete any existing credentials files
    SCOPE = 'https://www.googleapis.com/auth/gmail.send'

    def __init__(self, cred_path='~/.credentials'):
        self.credential_path = Path(cred_path).expanduser()
        if not self.credential_path.exists():
            logger.warning("Credential path does not exists. Using default path!")
            self.credential_path = Path('~/.credentials').expanduser()
            if not self.credential_path.exists():
                self.credential_path.mkdir()
        self.secret_path = Path('credentials.json')
        self.service = None

    def authenticate(self):
        # Check for existing credentials
        send_cred_path = self.credential_path.joinpath('gmail-send.json')
        store = file.Storage(send_cred_path)
        credentials = store.get()
        if not credentials or credentials.invalid:
            # No existing credentials found. Authenticate using OAuth2
            # credentials.json file required in root program directory
            if not self.secret_path.exists():
                logger.warning('Credentials json file not found: %s', self.secret_path)
                return False
            flow = client.flow_from_clientsecrets(str(self.secret_path), self.SCOPE)
            credentials = tools.run_flow(flow, store)
            logger.info('Storing credentials to %s', send_cred_path)
        logger.info('Credentials verified!')
        self.service = build('gmail', 'v1', http=credentials.authorize(Http()), cache_discovery=False)
        return True

    @staticmethod
    def create_message(to, subject, message_text, attachment_file=None):
        if attachment_file:
            # Make sure attachment file exists
            if not Path(attachment_file).exists():
                logger.warning('Attachment file not found: %s', attachment_file)
                return None
            message = MIMEMultipart()
            message.attach(MIMEText(message_text))
            content_type, encoding = mimetypes.guess_type(attachment_file)

            if content_type is None or encoding is not None:
                content_type = 'application/octet-stream'
            main_type, sub_type = content_type.split('/', 1)
            if main_type == 'text':
                with open(attachment_file, 'rb') as fp:
                    msg = MIMEText(fp.read(), _subtype=sub_type)
            else:
                msg = MIMEBase(main_type, sub_type)
                with open(attachment_file, 'rb') as fp:
                    msg.set_payload(fp.read())
            encoders.encode_base64(msg)
            filename = Path(attachment_file).name
            msg.add_header('Content-Disposition', 'attachment', filename=filename)
            message.attach(msg)
        else:
            message = MIMEText(message_text)

        # Don't populate 'From' field. Causes suspicious email warning
        # in gmail.
        # message['from'] = sender
        message['to'] = to
        message['subject'] = subject
        return {'raw': base64.urlsafe_b64encode(message.as_bytes()).decode()}

    def send_message(self, user_id, message):
        if message is not None:
            try:
                message = self.service.users().messages().send(userId=user_id, body=message).execute()
                logger.debug('Message Id: %s', message['id'])
                return message
            except Exception as error:
                logger.error('An error occurred trying to send email: %s' % error)
        else:
            logger.error("Message is empty. Unable to send email!")
