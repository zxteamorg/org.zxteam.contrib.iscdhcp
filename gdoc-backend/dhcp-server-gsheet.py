from __future__ import print_function
import httplib2
import os
import sys

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

try:
    import argparse
    parser = argparse.ArgumentParser(parents=[tools.argparser])
    parser.add_argument('--sheet-id', required=True)
    parser.add_argument('--range', required=True)
    parser.add_argument('--name-column-index', type=int, required=True)
    parser.add_argument('--ip-column-index', type=int, required=True)
    parser.add_argument('--mac-column-index', type=int, required=True)
    args =parser.parse_args()
except ImportError:
    args = None

SCOPES = 'https://www.googleapis.com/auth/spreadsheets.readonly'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Sheet to ISC DHCP configuration converter'


def get_credentials():
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir, 'dhcp-server-gsheet.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if args:
            credentials = tools.run_flow(flow, store, args)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path, file=sys.stderr)
    return credentials

def main():
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?version=v4')
    service = discovery.build('sheets', 'v4', http=http, discoveryServiceUrl=discoveryUrl)
    spreadsheetId = args.sheet_id
    rangeName = args.range
    nameColumnIndex = args.name_column_index
    ipColumnIndex = args.ip_column_index
    macColumnIndex = args.mac_column_index
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheetId, range=rangeName).execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')
    else:
        count = 0
        for row in values:
            count = count + 1
            recordName = ''
            try:
                recordName = row[nameColumnIndex]
            except IndexError:
                continue
            if recordName == '':
                continue
            ipAddress = ''
            macAddress = ''
            try:
                ipAddress = row[ipColumnIndex]
                macAddress = row[macColumnIndex]
            except IndexError:
                print('Skip "%s" due parse error' % (recordName))
                continue

#host 104.chufarm4 {
#  hardware ethernet 50:e5:49:da:62:2d;
#  fixed-address 192.168.254.104;
#}

            print('host %s {' % (recordName))
            print('  hardware ethernet %s;' % (macAddress))
            print('  fixed-address %s;' % (ipAddress))
            print('}')

if __name__ == '__main__':
    main()
