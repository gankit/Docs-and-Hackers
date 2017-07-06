
from __future__ import print_function
import httplib2
import os
from email.mime.text import MIMEText
import base64
import string
from urllib2 import HTTPError
from email_validator import validate_email, EmailNotValidError

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

try:
    import argparse
    parser = argparse.ArgumentParser(parents=[tools.argparser])
    parser.add_argument("email_type", help="type of email that should be sent.", choices=['invite', 'welcome'])
    flags = parser.parse_args()
    print(flags.email_type)

except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/gmail-python-quickstart.json
SCOPES_GMAIL = 'https://www.googleapis.com/auth/gmail.send'
# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES_SHEETS = 'https://www.googleapis.com/auth/spreadsheets'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Docs and Hackers Invitation'

my_email = 'ankitgupta00@gmail.com'
invite_email_subject = "Invitation from <NAME/> - Docs and Hackers"
invite_email_body = "Hello,<br/><br/>\
Your friend <NAME/> just joined <a href=\"http://www.docsandhackers.com\">Docs and Hackers</a>, and thought you might be interested to join as well. Docs and Hackers is a group of medical practitioners and software builders who aim to improve the practice of medicine with technology.<br/><br/>\
One of the reasons why health-tech problems aren't getting solved is because builders and physicians aren't connected to each other. Builders who want to dive into healthcare don't know where to start. And physicians who face challenges every day don't have a partner to build a solution with.<br/><br/>\
We would like to accelerate the pace of innovation in digital health by bringing doctors and builders together, regularly. Whether you are looking for collaborators to bring a new digital health idea to life, or are simply curious about what's going on in the space - <a href=\"http://www.docsandhackers.com\">Docs and Hackers</a> could be a helpful group.<br/><br/>\
Please check it out, and register if you are interested -<br/>\
<a href=\"http://www.docsandhackers.com\">http://www.docsandhackers.com</a><br/><br/>\
Happy to answer any questions. Hope to see you at the next meetup!<br/><br/>\
Best,<br/>\
Ankit Gupta.<br/>\
<a href=\"https://www.linkedin.com/in/ankitgupta00\">https://www.linkedin.com/in/ankitgupta00</a><br/>"


welcome_email_subject = "Docs and Hackers - Welcome and Next Meetup!"
welcome_email_body = "Hi <NAME/>,<br/><br/>\
Welcome to <a href=\"http://www.docsandhackers.com\">Docs and Hackers</a>, a community of medical practitioners and software builders who aim to improve the practice of medicine with technology.\
<br/><br/>300+ people have joined the group so far from all across the world. 20% of the members are doctors. 28% of the members have a project and want help whereas 48% of members want a new project. Lots of great connections, waiting to happen!\
<br/><br/>\
Our next meetup is on July 19th, 2017 from 6pm - 8pm at WeWork Soma (156 2nd Street, San Francisco). <a href=\"https://www.eventbrite.com/e/docs-and-hackers-tickets-35800481203\">Get tickets</a> to join in person or tune in on <a href=\"https://www.facebook.com/docsandhackers/\">Facebook Live</a>. We can't wait to see you there!\
<br/><br/>\
Please share this email to any friends who might be interested in joining the group. They could either be physicians with new ideas for digital health tools looking for partners to build, or tech builders looking to start something in digital health.\
<br/><br/>\
Finally, we are on <a href=\"https://www.facebook.com/docsandhackers/\">Facebook</a> and <a href=\"https://www.twitter.com/docsandhackers\">Twitter</a>. Like, follow and show your support!\
<br/><br/>\
Happy to answer any questions. Hope to see you at the next meetup!<br/><br/>\
Best,<br/>\
Ankit Gupta.<br/>\
<a href=\"https://www.linkedin.com/in/ankitgupta00\">https://www.linkedin.com/in/ankitgupta00</a><br/>"

def create_invite_message(inviter_name, inviter_email, to):
  """Create a message for an email.

  Args:
    sender: Email address of the sender.
    to: Email address of the receiver.
    subject: The subject of the email message.
    message_text: The text of the email message.

  Returns:
    An object containing a base64url encoded email object.
  """
  message_text = string.replace(invite_email_body, '<NAME/>', inviter_name)
  subject_text = string.replace(invite_email_subject, '<NAME/>', inviter_name)
  message = MIMEText(message_text, 'html')
  message['to'] = to
  message['from'] = "Ankit Gupta (Docs and Hackers) <"+my_email+">"
  message['cc'] = inviter_email
  message['subject'] = subject_text
  return {'raw': base64.urlsafe_b64encode(message.as_string())}

def create_welcome_message(name, email):
  message_text = string.replace(welcome_email_body, '<NAME/>', name)
  subject_text = string.replace(welcome_email_subject, '<NAME/>', name)
  message = MIMEText(message_text, 'html')
  message['to'] = email
  message['from'] = "Ankit Gupta (Docs and Hackers) <"+my_email+">"
  message['subject'] = subject_text
  return {'raw': base64.urlsafe_b64encode(message.as_string())}

def send_message(service, user_id, message):
  """Send an email message.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    message: Message to be sent.

  Returns:
    Sent Message.
  """
  try:
    message = (service.users().messages().send(userId=user_id, body=message)
               .execute())
    # print 'Message Id: %s' % message['id']
    return message
  except HTTPError, error:
    print('An error occurred: %s', error)
    return


def get_credentials_gmail():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'gmail-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES_GMAIL)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def get_credentials_sheets():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'sheets.googleapis.com-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES_SHEETS)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def main():
    """
    Creates a Gmail API service object
    """
    credentials_gmail = get_credentials_gmail()
    http_gmail = credentials_gmail.authorize(httplib2.Http())
    service_gmail = discovery.build('gmail', 'v1', http=http_gmail)

    """
    Creates a Sheets API service object
    """
    credentials_sheets = get_credentials_sheets()
    http_sheets = credentials_sheets.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service_sheets = discovery.build('sheets', 'v4', http=http_sheets,
                              discoveryServiceUrl=discoveryUrl)

    spreadsheetId = '1Jdx2N4sXuAdJyccBeBK0SNY_G9RjQhkx3L0LwvylZEM'
    rangeName = 'Sheet1!A2:AA'
    result = service_sheets.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
    values = result.get('values', [])
    if not values:
        print('No data found.')
    else:
      if flags.email_type == 'invite':        
        invite_sent_column_number = 25
        invite_sent_column_values = []
        invite_sent_range='Sheet1!Z2:Z'
        i = 0
        for row in values:
            if len(row) >= (invite_sent_column_number + 1) and len(row[invite_sent_column_number]) > 0:
              # invite has been processed
              print('Already processed invite at row %d' % (i+2))
              invite_sent_column_values.append([row[invite_sent_column_number]])
            else:
              # invite hasn't been processed
              invite_success = ''
              try:
                if len(row) >= 8:
                  print('Processing invite at row %d' % (i+2))
                  inviter_name = row[0].strip().lower().title() + " " + row[1].strip().lower().title()
                  inviter_name = inviter_name.strip()
                  inviter_email = row[2].strip()
                  to_emails = [x.strip() for x in row[7].split(',')]
                  for to_email in to_emails:
                    if len(inviter_name) > 0 and len(inviter_email) > 0 and len(to_email) > 0:
                      print('%s %s %s' % (inviter_name, inviter_email, to_email))
                      validate_email(inviter_email)
                      validate_email(to_email)
                      message = create_invite_message(inviter_name=inviter_name, inviter_email=inviter_email, to=to_email)
                      try:
                        message_id = send_message(service=service_gmail, user_id="me", message=message)
                        if message_id:
                          invite_success = 'Y'
                        else:
                          invite_success = 'N'
                      except Exception as error:
                        print('Error sending invite at row %d' % (i+2))
                        print(error)
                        invite_success = 'N'

              except Exception as error:
                print('Could not process invite at row %d' % (i+2))
                print(error)
              invite_sent_column_values.append([invite_success])
            i += 1
        print('%s'%invite_sent_column_values)
        invite_sent_body = {
          'values': invite_sent_column_values
        }
        result = service_sheets.spreadsheets().values().update(
            spreadsheetId=spreadsheetId, range=invite_sent_range,
            valueInputOption='RAW', body=invite_sent_body).execute()
      elif flags.email_type == 'welcome':
        welcome_sent_column_number = 26
        welcome_sent_column_values = []
        welcome_sent_range='Sheet1!AA2:AA'
        i = 0
        for row in values:
            if len(row) >= (welcome_sent_column_number + 1) and row[welcome_sent_column_number] == 'Y':
              # welcome email has been processed
              print('Already processed welcome email at row %d' % (i+2))
              welcome_sent_column_values.append([row[welcome_sent_column_number]])
            else:
              # welcome email hasn't been processed
              welcome_success = ''
              try:
                if len(row) >= 2:
                  print('Processing welcome email at row %d' % (i+2))
                  name = row[0].strip().lower().title()
                  name = name.strip()
                  email = row[2].strip()
                  validate_email(email)
                  if len(name) > 0 and len(email) > 0:
                    print('%s %s' % (name, email))
                    message = create_welcome_message(name=name, email=email)
                    try:
                      message_id = send_message(service=service_gmail, user_id="me", message=message)
                      if message_id:
                        welcome_success = 'Y'
                      else:
                        welcome_success = 'N'
                    except Exception as error:
                      print('Error sending invite at row %d' % (i+2))
                      print(error)
                      welcome_success = 'N'

              except Exception as error:
                print('Could not process invite at row %d' % (i+2))
                print(error)
              welcome_sent_column_values.append([welcome_success])
            i += 1
        print('%s'%welcome_sent_column_values)
        welcome_sent_body = {
          'values': welcome_sent_column_values
        }
        result = service_sheets.spreadsheets().values().update(
            spreadsheetId=spreadsheetId, range=welcome_sent_range,
            valueInputOption='RAW', body=welcome_sent_body).execute()





if __name__ == '__main__':
    main()