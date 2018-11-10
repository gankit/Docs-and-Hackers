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
    parser.add_argument("email_type", help="type of email that should be sent.", choices=['invite', 'welcome', 'reminder', 'pitch_request', 'followup', 'event_invite', 'event_invite_past_attendees', 'event_invite_founders', 'challenges_invite', 'challenges_invite_past_attendees', 'volunteer_request', 'add_founder_notification'])
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

my_email = 'ankit@docsandhackers.com'
# my_email = 'ankitgupta00@gmail.com'
invite_email_subject = "Invitation from <NAME/> - Docs and Hackers"
invite_email_body = "Hello,<br/><br/>\
Your friend <NAME/> just joined <a href=\"http://www.docsandhackers.com\">Docs and Hackers</a>, and thought you might be interested to join as well. Docs and Hackers is a group of medical practitioners and software builders who aim to improve the practice of medicine with technology.<br/><br/>\
One of the reasons why health-tech problems aren't getting solved is because builders and physicians aren't connected to each other. Builders who want to dive into healthcare don't know where to start. And physicians who face challenges every day don't have a partner to build a solution with.<br/><br/>\
We would like to accelerate the pace of innovation in digital health by bringing doctors and builders together, regularly. Whether you are looking for collaborators to bring a new digital health idea to life, or are simply curious about what's going on in the space - <a href=\"http://www.docsandhackers.com\">Docs and Hackers</a> could be a helpful group.\
<br/><br/>1500+ people have joined the group so far from all across the world. 23% of the members are doctors. 34% of the members have a project and want help whereas 45% of members want a new project. Lots of great connections, waiting to happen!\
<br/><br/>\
Please check it out, and register if you are interested - <a href=\"http://www.docsandhackers.com\">http://www.docsandhackers.com</a>\
<br/><br/>\
Happy to answer any questions. Hope to see you at the next meetup!<br/><br/>\
Best,<br/>\
Ankit Gupta.<br/>\
<a href=\"https://www.linkedin.com/in/ankitgupta00\">https://www.linkedin.com/in/ankitgupta00</a><br/>"

event_invite_email_subject = "Docs and Hackers - A Healthcare Founder/CEO Happy Hour!"
event_invite_email_body = "Hi <NAME/>,<br/><br/>\
We are trying something new for the next Docs and Hackers meetup - a closed group gathering of healthcare Founders and CEOs! It's happening in Boston on Thursday August 23rd, 6p - 8p at Hurricane's at the Garden (<a href=\"https://www.google.com/maps/place/Hurricane's+at+the+Garden/@42.3648137,-71.0607941,15z/data=!4m5!3m4!1s0x0:0xe11ddd5e9c5ee75e!8m2!3d42.3648137!4d-71.0607941\">Map</a>).\
<br/><br/>\
<b>Get a ticket (free!)</b> if you are running a healthcare company - <a href=\"https://www.eventbrite.com/e/docs-and-hackers-a-healthcare-founderceo-happy-hour-tickets-48481477404\">https://www.eventbrite.com/e/docs-and-hackers-a-healthcare-founderceo-happy-hour-tickets-48481477404</a>.\
<br/><br/>\
Running a company up is difficult. The complexities of healthcare's incentive structures, closed networks, oligopolies and traditional thinking make this even more challenging. Healthcare founders can learn a lot from each other, and hopefully move faster as a community. This meetup aims to provide an open environment to freely share our learnings, problems and solutions.\
<br/><br/>\
The meetup is only open to healthcare founders or CEOs. Anything that gets discussed (or disclosed) within the meetup, stays confidential within the group.\
<br/><br/>\
If you aren't a healthcare startup founder yet, no sweat - there are many open meetups planned in the future.\
<br/><br/>\
Best,<br/>\
Team D&H Boston.\
<br/>"

event_invite_past_attendee_email_subject = "Docs and Hackers is back in Mountain View!"
event_invite_past_attendee_email_body = "Hello,<br/><br/>\
Thank you for your interest in the <a href=\"http://www.docsandhackers.com\">Docs and Hackers</a> community. The next meetup is happening on Wednesday June 6th, 6p - 8:30p at The Garage in Google HQ, Mountain View, CA (<a href=\"https://www.google.com/maps?hl=en&q=37.4190259,-122.08241299999997&sll=37.4190259,-122.08241299999997&z=13&markers=37.4190259,-122.08241299999997\">Map</a>). <b>Get tickets</b> (free!) if you are interested - <a href=\"https://www.eventbrite.com/e/docs-and-hackers-google-tickets-46213209955\">https://www.eventbrite.com/e/docs-and-hackers-google-tickets-46213209955</a>. Or join our <a href=\"https://www.facebook.com/groups/1905224313135407/\">Facebook Group</a> to watch it online.\
<br/><br/>\
This time we shall have a special conversation on digital therapeutic startups with <a href=\"https://www.linkedin.com/in/david-utley-m-d-1a3bb9116/\">Dr. David Utley</a>, President and CEO of <a href=\"https://pivot.co/about-us/\">Carrot Inc.</a>. Aiming to help 40 million US smokers and 1 billion smokers world wide, Carrot Inc. is a digital health company who's first offering, <a href=\"https://youtu.be/Bqgz3YN36ko\">Pivot</a> uses <a href=\"https://www.multivu.com/players/English/8185651-carrot-inc-carbon-monoxide-breath-sensor-smoking/\">FDA-cleared wearable tech</a>, personalized coaching, and clinical practice guidelines delivered at scale to engage and empower people to quit smoking (and save $300 billion annually for our health care system).\
<br/><br/>\
Being able to presecibe apps for both prevention and treatment is revolutionary. What kind of indications can an app treat better than a pill? How might one launch a digital therapeutic startup? What regulatory challenges would an entrepreneur need to overcome? What learnings do we have so far, in terms of what works and what doesn't? Come learn about the problems that digital therapeutics can solve, and hear from Dr. David Utley's fascinating personal journey that led to the founding of Carrot Inc.\
<br/><br/>\
In addition to this special topic, we will have lots of time to meet other healthcare enthusiasts, builders and dreamers. Look forward to seeing you at the event!\
<br/><br/>\
Best,<br/>\
Team D&H Bay Area.\
<br/>\
PS: Check out our website to find some <a href=\"http://www.docsandhackers.com/#find_physicians\">innovative physicians</a> for your next idea.<br/>"

welcome_email_subject = "Welcome to Docs and Hackers!"
welcome_email_body = "Hi <NAME/>,<br/><br/>\
Welcome to <a href=\"http://www.docsandhackers.com\">Docs and Hackers</a>, a community of medical practitioners and software builders who aim to improve the practice of medicine with technology.\
<br/><br/>1500+ people have joined the group so far from all across the world. 23% of the members are doctors. 34% of the members have a project and want help whereas 45% of members want a new project. Lots of great connections, waiting to happen!\
<br/><br/>\
We have frequent meetups in several cities around the world. Visit <a href=\"http://www.docsandhackers.com\">our website</a> to see our upcoming meetups. In addition, find some <a href=\"http://www.docsandhackers.com/#find_physicians\">innovative physicians</a> for your next idea.\
<br/><br/>\
Please share this email with any friends or colleagues who might be interested in joining the group. They could either be healthcare professionals or tech builders looking to build innovative products and services in digital health.\
<br/><br/>\
Finally, please join our highly curated private <a href=\"https://www.facebook.com/groups/1905224313135407/\">Facebook Group</a> to interact with the community on the interwebs.\
<br/><br/>\
Happy to answer any questions. Hope to see you at the next meetup!<br/><br/>\
Best,<br/>\
Team D&H.<br/>"

reminder_email_subject = "A Healthcare Founder/CEO Happy Hour - Reminder!"
reminder_email_body = "Hi <NAME/>,<br/><br/>\
Quick reminder - the next <a href=\"http://www.docsandhackers.com\">Docs and Hackers</a> meetup is a closed group gathering of Healthcare Founders and CEOs! It's happening in Boston on Thursday August 23rd, 6p - 8p at Hurricane's at the Garden (<a href=\"https://www.google.com/maps/place/Hurricane's+at+the+Garden/@42.3648137,-71.0607941,15z/data=!4m5!3m4!1s0x0:0xe11ddd5e9c5ee75e!8m2!3d42.3648137!4d-71.0607941\">Map</a>).\
<br/><br/>\
<b>Get tickets</b> (free!) to join in - <a href=\"https://www.eventbrite.com/e/docs-and-hackers-a-healthcare-founderceo-happy-hour-tickets-48481477404\">https://www.eventbrite.com/e/docs-and-hackers-a-healthcare-founderceo-happy-hour-tickets-48481477404</a>. Few tickets are still left!\
<br/><br/>\
Feel free to share this event with your friends who are running healthcare companies and might be interested. We can't wait to see you there!\
<br/><br/>\
Best,<br/>\
Team D&H Boston.\
<br/>\
PS: Check out our website to find some <a href=\"http://www.docsandhackers.com/#find_physicians\">innovative physicians</a> for your next idea.<br/>"

pitch_request_email_subject = "Docs and Hackers - Want to pitch?"
pitch_request_email_body = "Hi <NAME/>,<br/><br/>\
I noticed that you got a ticket to next Docs and Hackers meetup happening on Wednesday in Mountain View - looking forward to seeing you there. The audience at the meetup is full of fellow healthcare enthusiasts, engineers, designers and builders - great place to find tech collaborators!\
<br/><br/>\
At the beginning of the meetup, we've reserved time for ~4 doctors to share their idea/project with the audience. With a short talk that can last for a maximum of 4 mins; no slides, no demos, just you going up to the microphone and sharing your idea. At most 1 - 2 questions at the end to keep it moving fast. Very casual. It will generate interest for folks to reach out to you during the meetup.\
<br/><br/>\
We've tried this format in the past and it was quite succesful (<a href=\"https://medium.com/docs-and-hackers/2-ideas-shared-by-doctors-at-the-3-7-18-docs-and-hackers-meetup-ec352688f983\">Example</a>). Helped make lots of interesting connections.\
<br/><br/>\
Is there something you'd like to share and take advantage of this time at the meetup? 3 slots are still available. Please let me know by sending me an email with answers to these two questions.\
<br/><br/>\
<b>1. What would you like to talk about? (1 - 2 sentences)</b>\
<br/>\
<b>2. What kind of help are you looking for from the Docs and Hackers community?</b>\
<br/><br/>\
Happy to hop on a quick call if have any questions / want any help in figuring out the right pitch.\
<br/><br/>\
Best,<br/>\
Ankit.<br/>\
+1 650 384 9358<br/>\
<a href=\"https://www.linkedin.com/in/ankitgupta00\">https://www.linkedin.com/in/ankitgupta00</a><br/>"

followup_email_subject = "Docs and Hackers - New Facebook Group!"
followup_email_body = "Hi <NAME/>,<br/><br/>\
The first Docs and Hackers meetup in SF was a success. 40+ people attended, 4 new ideas pitched and lots of exciting healthcare conversations! Since then, we've been hard at work putting some background infrastructure in place to scale the group.\
<br/><br/>\
Some updates and asks -\
<br/><br/>\
1. You requested an online forum to stay connected with each other - here it is! <b>Please join</b> the <a href=\"https://www.facebook.com/groups/1905224313135407/\">Docs and Hackers Facebook Group</a>.\
<br/>\
2. The next meetup will be this month in Mountain View, CA, USA. Time/Date/Venue is still being figured out. <b>If you have access to a space where we can host the meetup, please let me know.</b>\
<br/>\
3. The Docs and Hackers community is geographically diverse - Boston, Seattle, New York, Delhi, Bangalore, London and more! If any of you would like to host a Docs and Hackers event in your city, please reach out!\
<br/><br/>\
Thank you once again for being a part of this amazing community. Together, we can accelerate the pace of digital innovation in healthcare. Please feel free to reach out if you have any feedback/ideas for the group.\
<br/><br/>\
Best,<br/>\
Ankit.<br/>\
<a href=\"https://www.linkedin.com/in/ankitgupta00\">https://www.linkedin.com/in/ankitgupta00</a><br/>"

challenges_email_subject = "Docs and Hackers - New challenge by Johnson & Johnson!"
challenges_email_body = "Hi<NAME/>,<br/><br/>\
Docs and Hackers is collaborating with payers, providers and other healthcare organizations to curate <a href=\"http://www.docsandhackers.com/#challenges\">challenges</a> - real world problems that you can take a shot at. If you want to start a healthcare startup, but don't have the right idea - this is your chance!\
<br/><br/>\
I'm excited to announce a new challenge by <a href=\"http://www.docsandhackers.com/?challenge=jnj_digital_beauty\">Johnson & Johnson on Digital Beauty</a>.\
<br/><br/>\
<i>Do you have a good idea on how to improve skincare? If so, we're looking to help you make that idea into a reality. Winning ideas will get up to $50,000 in grants, up to one year of residency at an available local JLABS incubator, access to a network of consumer experts, and admission to test the product with consumers at the Johnson & Johnson Consumer Inc. Consumer Experience Center (CxC).</i>\
<br/><br/>\
Also fyi - <b>just 4 days to go</b> to submit applications for the <a href=\"http://www.docsandhackers.com/?challenge=humana_challenge\">Humana Innovation Challenge - Patient Records for Humans</a>.\
<br/><br/>\
Please reach out if you have any questions and/or need any help in finding the right people to form a team.<br/><br/>\
Best,<br/>\
Ankit.<br/>\
<a href=\"https://www.linkedin.com/in/ankitgupta00\">https://www.linkedin.com/in/ankitgupta00</a><br/>"

volunteer_email_subject = "Help Wanted - Docs and Hackers"
volunteer_email_body = "Hi<NAME/>,<br/><br/>\
Thank you for being a part of the <a href=\"http://www.docsandhackers.com\">Docs and Hackers</a> community! Since we came into existence on July 19th 2017, we have hosted 12 meetups in 4 cities around the world and grown to a community of 1500+ members interested in health care innovation. This is just a start - we have a lot more planned in the future.\
<br/><br/>\
Hence, <b><u>we need your help!</u></b> Please reach out to me if you are interested in any of the following roles or know someone who could be a good fit (even if they are not a D&H member today):\
<br/><br/>\
<b>1. Pitch Curator</b>: Many good startups pitch their ideas at Docs and Hackers events. We want to publish one good pitch every month on our blog, website and social media. We are looking for someone who loves new ideas, likes writing blog posts, and has some journalism experience. Video editing skills are a plus! This would be a 6 month commitment of about 3-4 hours every month.\
<br/><br/>\
<b>2. Accountant</b>: Docs and Hackers is incorporated in California as a non profit entity. There are some annual state and federal filing obligations that need to be executed. If you know what these are, and are interested in helping us stay compliant, please reach out.\
<br/><br/>\
<b>3. Corporate Lawyer</b>: If you have experience in working with non profits and are open to spending time on legal things as and when they come up, please reach out. We expect the work to be minimal.\
<br/><br/>\
Looking forward to hearing from you!<br/><br/>\
Best,<br/>\
Ankit Gupta.<br/>"

add_founder_email_subject = "Founders @ Docs and Hackers"
add_founder_email_body = "Hi<NAME/>,<br/><br/>\
Thank you for being a part of the <a href=\"http://www.docsandhackers.com\">Docs and Hackers</a> community. You indicated that you are either a Founder or CEO of a health care startup. Hence, we have just added you to the <a href=\"https://groups.google.com/a/docsandhackers.com/d/forum/founders\">Founders @ Docs and Hackers</a> google group - a closed group of Docs and Hackers members that are also running health care startups.\
<br/><br/>\
Please utilize this group to ask questions, get advise, request introductions, or even just emotional support! Running a startup is difficult, but the challenges you face might be similar to the ones faced and/or overcome by others on this group. Simply send an email to <a href=\"mailto:founders@docsandhackers.com\">founders@docsandhackers.com</a> to post to the group.\
<br/><br/>\
<b>The best way to get started is to send an email introducing yourself and your company.</b> You can also browse all previously posted messages on the group's <a href=\"https://groups.google.com/a/docsandhackers.com/d/forum/founders\">website</a>. Hope this group helps your startup do well and make the world a better place.\
<br/><br/>\
Happy to answer any questions!<br/><br/>\
Best,<br/>\
Ankit Gupta.<br/>\
<br/>\
PS: If you would like to NOT be a part of this group, please reach out and I can remove you from the group."


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

def create_event_invite_message(name, email):
  message_text = string.replace(event_invite_email_body, '<NAME/>', name)
  subject_text = string.replace(event_invite_email_subject, '<NAME/>', name)
  message = MIMEText(message_text, 'html')
  message['to'] = email
  message['from'] = "Ankit Gupta (Docs and Hackers) <"+my_email+">"
  message['subject'] = subject_text
  return {'raw': base64.urlsafe_b64encode(message.as_string())}

def create_event_invite_past_attendee_message(name, email):
  message_text = string.replace(event_invite_past_attendee_email_body, '<NAME/>', name)
  subject_text = string.replace(event_invite_past_attendee_email_subject, '<NAME/>', name)
  message = MIMEText(message_text, 'html')
  message['to'] = email
  message['from'] = "Ankit Gupta (Docs and Hackers) <"+my_email+">"
  message['subject'] = subject_text
  return {'raw': base64.urlsafe_b64encode(message.as_string())}

def create_reminder_message(name, email):
  message_text = string.replace(reminder_email_body, '<NAME/>', name)
  subject_text = string.replace(reminder_email_subject, '<NAME/>', name)
  message = MIMEText(message_text, 'html')
  message['to'] = email
  message['from'] = "Ankit Gupta (Docs and Hackers) <"+my_email+">"
  message['subject'] = subject_text
  return {'raw': base64.urlsafe_b64encode(message.as_string())}

def create_pitch_request_message(name, email):
  message_text = string.replace(pitch_request_email_body, '<NAME/>', name)
  subject_text = string.replace(pitch_request_email_subject, '<NAME/>', name)
  message = MIMEText(message_text, 'html')
  message['to'] = email
  message['from'] = "Ankit Gupta (Docs and Hackers) <"+my_email+">"
  message['subject'] = subject_text
  return {'raw': base64.urlsafe_b64encode(message.as_string())}

def create_followup_message(name, email):
  message_text = string.replace(followup_email_body, '<NAME/>', name)
  subject_text = string.replace(followup_email_subject, '<NAME/>', name)
  message = MIMEText(message_text, 'html')
  message['to'] = email
  message['from'] = "Ankit Gupta (Docs and Hackers) <"+my_email+">"
  message['subject'] = subject_text
  return {'raw': base64.urlsafe_b64encode(message.as_string())}

def create_challenge_invite_message(name, email):
  if len(name) > 0:
    name = ' '+name
  message_text = string.replace(challenges_email_body, '<NAME/>', name)
  subject_text = string.replace(challenges_email_subject, '<NAME/>', name)
  message = MIMEText(message_text, 'html')
  message['to'] = email
  message['from'] = "Ankit Gupta (Docs and Hackers) <"+my_email+">"
  message['subject'] = subject_text
  return {'raw': base64.urlsafe_b64encode(message.as_string())}

def create_volunteer_invite_message(name, email):
  if len(name) > 0:
    name = ' '+name
  message_text = string.replace(volunteer_email_body, '<NAME/>', name)
  subject_text = string.replace(volunteer_email_subject, '<NAME/>', name)
  message = MIMEText(message_text, 'html')
  message['to'] = email
  message['from'] = "Ankit Gupta (Docs and Hackers) <"+my_email+">"
  message['subject'] = subject_text
  return {'raw': base64.urlsafe_b64encode(message.as_string())}

def create_add_founder_message(name, email):
  if len(name) > 0:
    name = ' '+name
  message_text = string.replace(add_founder_email_body, '<NAME/>', name)
  subject_text = string.replace(add_founder_email_subject, '<NAME/>', name)
  message = MIMEText(message_text, 'html')
  message['to'] = my_email #email
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
                                   'mail-docs-and-hackers.json')

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

def pitch_request():
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

    spreadsheetId = '15KVoFcdrC2ey-geN4YS49yOiVN8OeY4Opv6KkFAqRO0'
    rangeName = 'Sheet1!A2:D'
    result = service_sheets.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
    values = result.get('values', [])
    if not values:
        print('No data found.')
    else:
      request_sent_column_number = 3
      request_sent_column_values = []
      request_sent_range='Sheet1!D2:D'
      i = 0
      for row in values:
          if len(row) >= (request_sent_column_number + 1) and row[request_sent_column_number] == 'Y':
            # request has been processed
            print('Already processed pitch request at row %d' % (i+2))
            request_sent_column_values.append([row[request_sent_column_number]])
          else:
            # request hasn't been processed
            request_success = ''
            try:
              if len(row) >= 3:
                print('Processing pitch request email at row %d' % (i+2))
                name = row[0].strip().lower().title()
                name = name.strip()
                email = row[2].strip()
                validate_email(email)
                if len(name) > 0 and len(email) > 0:
                  print('%s %s' % (name, email))
                  message = create_pitch_request_message(name=name, email=email)
                  try:
                    message_id = send_message(service=service_gmail, user_id="me", message=message)
                    if message_id:
                      request_success = 'Y'
                    else:
                      request_success = 'N'
                  except Exception as error:
                    print('Error sending pitch request at row %d' % (i+2))
                    print(error)
                    request_success = 'N'

            except Exception as error:
              print('Could not process pitch request at row %d' % (i+2))
              print(error)
            request_sent_column_values.append([request_success])
          i += 1
          # if i > 1:
          #   break;
      print('%s'%request_sent_column_values)
      request_sent_body = {
        'values': request_sent_column_values
      }
      result = service_sheets.spreadsheets().values().update(
          spreadsheetId=spreadsheetId, range=request_sent_range,
          valueInputOption='RAW', body=request_sent_body).execute()      

def invite_past_attendees(flag):
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
    rangeName = 'Attendees!A2:D'
    result = service_sheets.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
    values = result.get('values', [])
    if not values:
        print('No data found.')
    else:
      if flag == 'event_invite_past_attendees':
        event_invite_sent_column_number = 4
        docs_and_hackers_member_column_number = 3
        location_column_number = 1
        event_invite_sent_column_values = []
        event_invite_sent_range='Attendees!C2:C'
        i = 0
        for row in values:
            if len(row) >= (event_invite_sent_column_number + 1) and row[event_invite_sent_column_number] == 'Y':
              # email has been processed
              print('Already processed event_invite email at row %d' % (i+2))
              event_invite_sent_column_values.append([row[event_invite_sent_column_number]])
            elif len(row) >= (docs_and_hackers_member_column_number + 1) and row[docs_and_hackers_member_column_number] == 'Y':
              # email present in docs and hackers main list
              print('Member already in Docs and Hackers, at row %d' % (i+2))
              event_invite_sent_column_values.append(['Y'])
            else:
              # email hasn't been processed
              event_invite_success = ''
              try:
                if len(row) >= 1:
                  print('Processing event_invite email at row %d' % (i+2))
                  email = row[0].strip()
                  validate_email(email)
                  if len(email) > 0:
                    print('%s' % (email))
                    message = create_event_invite_past_attendee_message(name='', email=email)
                    try:
                      message_id = send_message(service=service_gmail, user_id="me", message=message)
                      if message_id:
                        event_invite_success = 'Y'
                      else:
                        event_invite_success = 'N'
                    except Exception as error:
                      print('Error sending event_invite at row %d' % (i+2))
                      print(error)
                      event_invite_success = 'N'

              except Exception as error:
                print('Could not process event_invite at row %d' % (i+2))
                print(error)
              event_invite_sent_column_values.append([event_invite_success])
            i += 1
            if i >= 800:
              break
        print('%s'%event_invite_sent_column_values)
        event_invite_sent_body = {
          'values': event_invite_sent_column_values
        }
        result = service_sheets.spreadsheets().values().update(
            spreadsheetId=spreadsheetId, range=event_invite_sent_range,
            valueInputOption='RAW', body=event_invite_sent_body).execute()
      elif flag == 'challenges_invite_past_attendees':
        challenge_invite_sent_column_number = 4
        challenge_invite_sent_column_values = []
        challenge_invite_sent_range='Attendees!C2:C'
        i = 0
        for row in values:
            if len(row) >= (challenge_invite_sent_column_number + 1) and row[challenge_invite_sent_column_number] == 'Y':
              # email has been processed
              print('Already processed challenge_invite email at row %d' % (i+2))
              challenge_invite_sent_column_values.append([row[challenge_invite_sent_column_number]])
            elif len(row) >= 2 and row[1] == 'Y':
              # email present in docs and hackers main list
              print('Member already in Docs and Hackers, at row %d' % (i+2))
              challenge_invite_sent_column_values.append(['Y'])
            else:
              # email hasn't been processed
              challenge_invite_success = ''
              try:
                if len(row) >= 1:
                  print('Processing challenge_invite email at row %d' % (i+2))
                  email = row[0].strip()
                  validate_email(email)
                  if len(email) > 0:
                    print('%s' % (email))
                    message = create_challenge_invite_message(name='', email=email)
                    try:
                      message_id = send_message(service=service_gmail, user_id="me", message=message)
                      if message_id:
                        challenge_invite_success = 'Y'
                      else:
                        challenge_invite_success = 'N'
                    except Exception as error:
                      print('Error sending event_invite at row %d' % (i+2))
                      print(error)
                      challenge_invite_success = 'N'

              except Exception as error:
                print('Could not process event_invite at row %d' % (i+2))
                print(error)
              challenge_invite_sent_column_values.append([challenge_invite_success])
            i += 1
        print('%s'%challenge_invite_sent_column_values)
        challenge_invite_sent_body = {
          'values': challenge_invite_sent_column_values
        }
        result = service_sheets.spreadsheets().values().update(
            spreadsheetId=spreadsheetId, range=challenge_invite_sent_range,
            valueInputOption='RAW', body=challenge_invite_sent_body).execute()

def invite_founders(flag):
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
    rangeName = 'Healthcare Founders!A2:E'
    result = service_sheets.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
    values = result.get('values', [])
    if not values:
        print('No data found.')
    else:
      if flag == 'event_invite_founders':
        event_invite_sent_column_number = 4
        event_invite_sent_column_values = []
        event_invite_sent_range='Healthcare Founders!E2:E'
        i = 0
        j = 0
        for row in values:
            if len(row) >= (event_invite_sent_column_number + 1) and row[event_invite_sent_column_number] == 'Y':
              # email has been processed
              print('Already processed event_invite email at row %d' % (i+2))
              event_invite_sent_column_values.append([row[event_invite_sent_column_number]])
            # elif len(row) >= 4 and row[3] == 'Y':
            #   # email present in docs and hackers main list
            #   print('Member already in Docs and Hackers, at row %d' % (i+2))
            #   event_invite_sent_column_values.append(['Y'])
            else:
              # email hasn't been processed
              event_invite_success = ''
              try:
                if len(row) >= 1:
                  print('Processing event_invite email at row %d' % (i+2))
                  name = row[0].strip().lower().title()
                  name = name.strip()
                  email = row[2].strip()
                  validate_email(email)
                  if len(name) > 0 and len(email) > 0:
                    print('%s %s' % (name, email))
                    message = create_reminder_message(name=name, email=email)
                    try:
                      message_id = send_message(service=service_gmail, user_id="me", message=message)
                      if message_id:
                        j += 1
                        event_invite_success = 'Y'
                      else:
                        event_invite_success = 'N'
                    except Exception as error:
                      print('Error sending event_invite at row %d' % (i+2))
                      print(error)
                      event_invite_success = 'N'

              except Exception as error:
                print('Could not process event_invite at row %d' % (i+2))
                print(error)
              event_invite_sent_column_values.append([event_invite_success])
            i += 1
            # if i >= 1:
            #   break
        print('%s'%event_invite_sent_column_values)
        print('%d emails sent'%j)
        event_invite_sent_body = {
          'values': event_invite_sent_column_values
        }
        result = service_sheets.spreadsheets().values().update(
            spreadsheetId=spreadsheetId, range=event_invite_sent_range,
            valueInputOption='RAW', body=event_invite_sent_body).execute()
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
    rangeName = 'Members!A2:AL'
    result = service_sheets.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
    values = result.get('values', [])
    if not values:
        print('No data found.')
    else:
      if flags.email_type == 'invite':        
        invite_sent_column_number = 25
        invite_sent_column_values = []
        invite_sent_range='Members!Z2:Z'
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
        welcome_sent_range='Members!AA2:AA'
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
                if len(row) >= 3:
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
            # if i > 1:
            #   break;
        print('%s'%welcome_sent_column_values)
        welcome_sent_body = {
          'values': welcome_sent_column_values
        }
        result = service_sheets.spreadsheets().values().update(
            spreadsheetId=spreadsheetId, range=welcome_sent_range,
            valueInputOption='RAW', body=welcome_sent_body).execute()
      elif flags.email_type == 'reminder':
        reminder_sent_column_number = 30
        coming_column_number = 29
        location_column_number = 6
        doctor_or_hacker_column_number = 3
        founder_column_number = 10
        reminder_sent_column_values = []
        reminder_sent_range='Members!AE2:AE'
        i = 0
        j = 0
        for row in values:
          if len(row) >= (reminder_sent_column_number + 1) and len(row[reminder_sent_column_number]) > 0:
            # reminder email has been processed
            print('Already processed reminder email at row %d' % (i+2))
            reminder_sent_column_values.append([row[reminder_sent_column_number]])
          else:
            # reminder email hasn't been processed
            reminder_success = ''
            if len(row) >= (coming_column_number + 1) and len(row[coming_column_number]) > 0:
              # already coming. skip.
              print('Already coming at row %d' % (i+2))
              reminder_success = 'Y'
            else:
              try:
                if (len(row) >= (location_column_number + 1) and 'boston' in row[location_column_number].lower()) or (len(row) >= (founder_column_number + 1) and 'TRUE' in row[founder_column_number].lower()):                
                # if len(row) >= (location_column_number + 1):
                  # if ('san francisco' in row[location_column_number].lower() and ('hacker' in row[doctor_or_hacker_column_number].lower() or 'both' in row[doctor_or_hacker_column_number].lower())) :
                  # if (('san francisco' in row[location_column_number].lower()) or ('mountain view' in row[location_column_number].lower())) :
                  print('Processing reminder email at row %d' % (i+2))
                  name = row[0].strip().lower().title()
                  name = name.strip()
                  email = row[2].strip()
                  validate_email(email)
                  if len(name) > 0 and len(email) > 0:
                    print('%s %s' % (name, email))
                    message = create_reminder_message(name=name, email=email)
                    try:
                      message_id = send_message(service=service_gmail, user_id="me", message=message)
                      if message_id:
                        j += 1 
                        reminder_success = 'Y'
                      else:
                        reminder_success = 'N'
                    except Exception as error:
                      print('Error sending invite at row %d' % (i+2))
                      print(error)
                      reminder_success = ''
                else:
                  reminder_success = 'N'
                  # print('Not in san francisco or not a hacker email at row %d' % (i+2))
                  # print('Not in sf or mtv email at row %d' % (i+2))
                  print('Not in boston or a founder email at row %d' % (i+2))

              except Exception as error:
                print('Could not process invite at row %d' % (i+2))
                print(error)
            reminder_sent_column_values.append([reminder_success])
          i += 1
          # if i > 500:
          #   break;
        print('%s'%reminder_sent_column_values)
        print('%d emails sent'%j);
        reminder_sent_body = {
          'values': reminder_sent_column_values
        }
        result = service_sheets.spreadsheets().values().update(
            spreadsheetId=spreadsheetId, range=reminder_sent_range,
            valueInputOption='RAW', body=reminder_sent_body).execute()
      elif flags.email_type == 'followup':
        followup_sent_column_number = 31
        followup_sent_column_values = []
        followup_sent_range='Members!AF2:AF'
        i = 0
        for row in values:
            if len(row) >= (followup_sent_column_number + 1) and row[followup_sent_column_number] == 'Y':
              # followup email has been processed
              print('Already processed followup email at row %d' % (i+2))
              followup_sent_column_values.append([row[followup_sent_column_number]])
            else:
              # followup email hasn't been processed
              followup_success = ''
              try:
                if len(row) >= 3:
                  print('Processing followup email at row %d' % (i+2))
                  name = row[0].strip().lower().title()
                  name = name.strip()
                  email = row[2].strip()
                  validate_email(email)
                  if len(name) > 0 and len(email) > 0:
                    print('%s %s' % (name, email))
                    message = create_followup_message(name=name, email=email)
                    try:
                      message_id = send_message(service=service_gmail, user_id="me", message=message)
                      if message_id:
                        followup_success = 'Y'
                      else:
                        followup_success = 'N'
                    except Exception as error:
                      print('Error sending followup at row %d' % (i+2))
                      print(error)
                      followup_success = 'N'

              except Exception as error:
                print('Could not process followup at row %d' % (i+2))
                print(error)
              followup_sent_column_values.append([followup_success])
            i += 1
        print('%s'%followup_sent_column_values)
        followup_sent_body = {
          'values': followup_sent_column_values
        }
        result = service_sheets.spreadsheets().values().update(
            spreadsheetId=spreadsheetId, range=followup_sent_range,
            valueInputOption='RAW', body=followup_sent_body).execute()
      elif flags.email_type == 'event_invite':
        event_invite_sent_column_number = 32
        event_invite_sent_column_values = []
        location_column_number = 6
        founder_column_number = 10
        event_invite_sent_range='Members!AG2:AG'
        i = 0
        j = 0
        for row in values:
            if len(row) >= (event_invite_sent_column_number + 1) and row[event_invite_sent_column_number] == 'Y':
              # welcome email has been processed
              print('Already processed event_invite email at row %d' % (i+2))
              event_invite_sent_column_values.append([row[event_invite_sent_column_number]])
            else:
              # welcome email hasn't been processed
              event_invite_success = ''
              try:
                if (len(row) >= (location_column_number + 1) and 'boston' in row[location_column_number].lower()) or (len(row) >= (founder_column_number + 1) and 'TRUE' in row[founder_column_number].lower()):
                #   if 'boston' in row[location_column_number].lower():
                  print('Processing event_invite email at row %d' % (i+2))
                  name = row[0].strip().lower().title()
                  name = name.strip()
                  email = row[2].strip()
                  validate_email(email)
                  if len(name) > 0 and len(email) > 0:
                    print('%s %s' % (name, email))
                    message = create_event_invite_message(name=name, email=email)
                    try:
                      message_id = send_message(service=service_gmail, user_id="me", message=message)
                      if message_id:
                        event_invite_success = 'Y'
                        j += 1
                      else:
                        event_invite_success = 'N'
                    except Exception as error:
                      print('Error sending event_invite at row %d' % (i+2))
                      print(error)
                      event_invite_success = 'N'
                else:
                  event_invite_success = 'N'
                  print('Not in Boston or a founder email at row %d' % (i+2))

              except Exception as error:
                print('Could not process event_invite at row %d' % (i+2))
                print(error)
              event_invite_sent_column_values.append([event_invite_success])
            i += 1
            if j > 400:
              break;
        print('%s'%event_invite_sent_column_values)
        print('%d emails sent'%j);
        event_invite_sent_body = {
          'values': event_invite_sent_column_values
        }
        result = service_sheets.spreadsheets().values().update(
            spreadsheetId=spreadsheetId, range=event_invite_sent_range,
            valueInputOption='RAW', body=event_invite_sent_body).execute()
      elif flags.email_type == 'challenges_invite':
        challenge_sent_column_number = 33
        challenge_sent_column_values = []
        challenge_sent_range='Members!AH2:AH'
        i = 0
        for row in values:
            if len(row) >= (challenge_sent_column_number + 1) and row[challenge_sent_column_number] == 'Y':
              # welcome email has been processed
              print('Already processed challenge email at row %d' % (i+2))
              challenge_sent_column_values.append([row[challenge_sent_column_number]])
            else:
              # challenge email hasn't been processed
              challenge_success = ''
              try:
                if len(row) >= 3:
                  print('Processing challenge email at row %d' % (i+2))
                  name = row[0].strip().lower().title()
                  name = name.strip()
                  email = row[2].strip()
                  validate_email(email)
                  if len(name) > 0 and len(email) > 0:
                    print('%s %s' % (name, email))
                    message = create_challenge_invite_message(name=name, email=email)
                    try:
                      message_id = send_message(service=service_gmail, user_id="me", message=message)
                      if message_id:
                        challenge_success = 'Y'
                      else:
                        challenge_success = 'N'
                    except Exception as error:
                      print('Error sending invite at row %d' % (i+2))
                      print(error)
                      challenge_success = ''

              except Exception as error:
                print('Could not process invite at row %d' % (i+2))
                print(error)
              challenge_sent_column_values.append([challenge_success])
            i += 1
            # if i > 0:
            #   break;
        print('%s'%challenge_sent_column_values)
        challenge_sent_body = {
          'values': challenge_sent_column_values
        }
        result = service_sheets.spreadsheets().values().update(
            spreadsheetId=spreadsheetId, range=challenge_sent_range,
            valueInputOption='RAW', body=challenge_sent_body).execute()
      elif flags.email_type == 'volunteer_request':
        volunteer_invite_sent_column_number = 37
        volunteer_invite_sent_column_values = []
        volunteer_invite_sent_range='Members!AL2:AL'
        i = 0
        for row in values:
            if len(row) >= (volunteer_invite_sent_column_number + 1) and row[volunteer_invite_sent_column_number] == 'Y':
              # notification email has been processed
              print('Already processed notification email at row %d' % (i+2))
              volunteer_invite_sent_column_values.append([row[volunteer_invite_sent_column_number]])
            else:
              # notification email hasn't been processed
              volunteer_invite_success = ''
              try:
                if len(row) >= 3:
                  print('Processing volunteer invite email at row %d' % (i+2))
                  name = row[0].strip().lower().title()
                  name = name.strip()
                  email = row[2].strip()
                  validate_email(email)
                  if len(name) > 0 and len(email) > 0:
                    print('%s %s' % (name, email))
                    message = create_volunteer_invite_message(name=name, email=email)
                    try:
                      message_id = send_message(service=service_gmail, user_id="me", message=message)
                      if message_id:
                        volunteer_invite_success = 'Y'
                      else:
                        volunteer_invite_success = 'N'
                    except Exception as error:
                      print('Error sending invite at row %d' % (i+2))
                      print(error)
                      volunteer_invite_success = ''

              except Exception as error:
                print('Could not process invite at row %d' % (i+2))
                print(error)
              volunteer_invite_sent_column_values.append([volunteer_invite_success])
            i += 1
            if i > 800:
              break;
        print('%s'%volunteer_invite_sent_column_values)
        volunteer_invite_sent_body = {
          'values': volunteer_invite_sent_column_values
        }
        result = service_sheets.spreadsheets().values().update(
            spreadsheetId=spreadsheetId, range=volunteer_invite_sent_range,
            valueInputOption='RAW', body=volunteer_invite_sent_body).execute()
      elif flags.email_type == 'add_founder_notification':
        notification_sent_column_number = 36
        added_to_group_column_number = 35
        notification_sent_column_values = []
        notification_sent_range='Members!AK2:AK'
        i = 0
        for row in values:
            if len(row) >= (notification_sent_column_number + 1) and row[notification_sent_column_number] == 'Y':
              # notification email has been processed
              print('Already processed notification email at row %d' % (i+2))
              notification_sent_column_values.append([row[notification_sent_column_number]])
            else:
              # notification email hasn't been processed
              notification_success = ''
              if len(row) >= (added_to_group_column_number + 1) and row[added_to_group_column_number] == 'Y':
                try:
                  if len(row) >= 3:
                    print('Processing notification invite email at row %d' % (i+2))
                    name = row[0].strip().lower().title()
                    name = name.strip()
                    email = row[2].strip()
                    validate_email(email)
                    if len(name) > 0 and len(email) > 0:
                      print('%s %s' % (name, email))
                      message = create_add_founder_message(name=name, email=email)
                      try:
                        message_id = send_message(service=service_gmail, user_id="me", message=message)
                        if message_id:
                          notification_success = 'Y'
                        else:
                          notification_success = 'N'
                      except Exception as error:
                        print('Error sending invite at row %d' % (i+2))
                        print(error)
                        notification_success = ''

                except Exception as error:
                  print('Could not process notification at row %d' % (i+2))
                  print(error)
              else:
                print('Not added to the founder group at row %d' % (i+2))
              notification_sent_column_values.append([notification_success])            
            i += 1
            # if i > 1:
            #   break;
        print('%s'%notification_sent_column_values)
        notification_sent_body = {
          'values': notification_sent_column_values
        }
        # result = service_sheets.spreadsheets().values().update(
        #     spreadsheetId=spreadsheetId, range=notification_sent_range,
        #     valueInputOption='RAW', body=notification_sent_body).execute()



if __name__ == '__main__':
  if flags.email_type == 'pitch_request':
    pitch_request()
  elif flags.email_type == 'event_invite_past_attendees' or flags.email_type == 'challenges_invite_past_attendees':
    invite_past_attendees(flags.email_type)
  elif flags.email_type == 'event_invite_founders':
    invite_founders(flags.email_type)
  else:
    main()