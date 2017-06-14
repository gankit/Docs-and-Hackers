from __future__ import print_function
import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage


# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES_SHEETS = 'https://www.googleapis.com/auth/spreadsheets'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Docs and Hackers Invitation'

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

import argparse
parser = argparse.ArgumentParser(parents=[tools.argparser])
parser.add_argument("input_file", help="json file with results from scraper")
args = parser.parse_args()
print(args.input_file)

import json

with open(args.input_file) as data_file:
	data = json.load(data_file)
	values = []
	for row in data:
		values.append([row['URL'], row['Name'], row['City'], row['State'], row['Phone'], row['Authorized Official']])
	print(len(values))
	credentials_sheets = get_credentials_sheets()
	http_sheets = credentials_sheets.authorize(httplib2.Http())
	discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
	service_sheets = discovery.build('sheets', 'v4', http=http_sheets,
                              discoveryServiceUrl=discoveryUrl)
	spreadsheetId = '12VehX3-z8wub9Ec1MTOifo9OuuqyE2YKjjKjJe-5N2g'
	rangeName = 'Sheet1!A2:Z'
	body = {
		'values': values
	}
	result = service_sheets.spreadsheets().values().update(
        spreadsheetId=spreadsheetId, range=rangeName,
        valueInputOption='RAW', body=body).execute()
	print(result)
