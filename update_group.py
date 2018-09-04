from __future__ import print_function
from httplib2 import Http
import os
import httplib2
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
    parser.add_argument("group_type", help="the group that needs to be updated.", choices=['founders'])
    flags = parser.parse_args()
    print(flags.group_type)

except ImportError:
    flags = None

SCOPES = 'https://www.googleapis.com/auth/admin.directory.group'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Docs and Hackers'

def get_credentials_groups():
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
                                   'groups.googleapis.com-docs-and-hackers.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
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
    Creates a Groups API service object
    """

    credentials_groups = get_credentials_groups()
    if not credentials_groups or credentials_groups.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        credentials_groups = tools.run_flow(flow, store)
    service = discovery.build('admin', 'directory_v1', http=credentials_groups.authorize(Http()))

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
    rangeName = 'Members!A2:AJ'
    result = service_sheets.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
    values = result.get('values', [])
    if not values:
        print('No data found.')
    else:
      if flags.group_type == 'founders':        
        founder_column_number = 8
        added_column_number = 35
        column_values = []
        data_range='Members!AJ2:AJ'
        i = 0
        for row in values:
	        if len(row) >= (added_column_number + 1) and row[added_column_number] == 'Y':
	          # record has been processed
	          print('Already processed row %d' % (i+2))
	          column_values.append([row[added_column_number]])
	        else:
	          # record hasn't been processed
	          add_success = ''
	          try:
	            if len(row) >= (founder_column_number+1):
	              print('Processing row %d' % (i+2))
	              name = row[0].strip().lower().title()
	              name = name.strip()
	              email = row[2].strip()
	              validate_email(email)
	              if len(name) > 0 and len(email) > 0:
	              	if row[founder_column_number] == 'TRUE':
		                print('%s %s' % (name, email))
	              		# add to group
	              		try:	              			
											results = service.members().insert(groupKey='founders@docsandhackers.com', body={
												"email": email,
												"role": "MEMBER"
											}).execute()
											add_success = 'Y'
		                except Exception as error:
											print('Error adding founder row %d' % (i+2))
											print(error)
											add_success = ''
	              	else:
	              		add_success = ''

	          except Exception as error:
	            print('Could not process invite at row %d' % (i+2))
	            print(error)
	          column_values.append([add_success])
	        i += 1
	        # if i > 1:
	        #   break;
        print('%s'%column_values)
        body = {
          'values': column_values
        }
        result = service_sheets.spreadsheets().values().update(
            spreadsheetId=spreadsheetId, range=data_range,
            valueInputOption='RAW', body=body).execute()

    # Call the Admin SDK Directory API
    # print('Getting the first 10 users in the domain')
    # results = service.get('founders').users().list(customer='my_customer', maxResults=10,
    #                             orderBy='email').execute()
    # results = service.groups().get(groupKey='founders@docsandhackers.com').execute()
    # print(results)
    # users = results.get('users', [])

    # if not users:
    #     print('No users in the domain.')
    # else:
    #     print('Users:')
    #     for user in users:
    #         print('{0} ({1})'.format(user['primaryEmail'],
    #             user['name']['fullName']))


if __name__ == '__main__':
    main()