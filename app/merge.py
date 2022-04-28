from __future__ import print_function

import time
import os


from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

'''
DOCS_FILE_ID = The id of the Template file from Docs.
SHEETS_FILE_ID = The id of the Data Source file from Sheets
DRIVE_FOLDER_ID = The id of the folder that the merged Docs will be saved to in Drive.
'''
DOCS_FILE_ID = "DOCS_ID"
SHEETS_FILE_ID = "SHEETS_ID"
DRIVE_FOLDER_ID = "FOLDER_ID"

'''
This is the basic name of the Doc after merge. Each Doc will be named with this, then the first column value.
This can be changed below (in the merge function) by setting the COLUMNS[0] to the desired column number.
'''
MERGED_FILENAME = 'Name_Of_Doc'

'''
Name your columns. 
These do not have to match the column names on the Sheet. 
They should match the values on the Doc template. 
'''
COLUMNS = ['NAME', 'ADDRESS', 'CITY', 'STATE', 'ZIP']

# authorization constants

SCOPES = (  # iterable or space-delimited string
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/documents',
    'https://www.googleapis.com/auth/spreadsheets.readonly',
)

SOURCES = ('sheets', 'text')
SOURCE = 'sheets'



# creds, _ = google.auth.default()
creds = None
# The file token.json stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
if os.path.exists('token.json'):
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)
# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open('token.json', 'w') as token:
        token.write(creds.to_json())
# pylint: disable=maybe-no-member

# service endpoints to Google APIs

DRIVE = build('drive', 'v3', credentials=creds)
DOCS = build('docs', 'v1', credentials=creds)
SHEETS = build('sheets', 'v4', credentials=creds)


def get_data(source):
    """Gets mail merge data from chosen data source.
    """
    try:
        if source not in {'sheets', 'text'}:
            raise ValueError(f"ERROR: unsupported source {source}; "
                             f"choose from {SOURCES}")
        return SAFE_DISPATCH[source]()
    except HttpError as error:
        print(f"An error occurred: {error}")
        return error


def _get_sheets_data(service=SHEETS):
    """(private) Returns data from Google Sheets source. It gets all rows of
        'Sheet1' (the default Sheet in a new spreadsheet), but drops the first
        (header) row. Use any desired data range (in standard A1 notation).
    """
    return service.spreadsheets().values().get(spreadsheetId=SHEETS_FILE_ID,
                                               range='Sheet1').execute().get(
        'values')[1:]
    # skip header row


# data source dispatch table [better alternative vs. eval()]
SAFE_DISPATCH = {k: globals().get('_get_%s_data' % k) for k in SOURCES}


def _copy_template(tmpl_id, source, service, doc_name):
    """(private) Copies letter template document using Drive API then
        returns file ID of (new) copy.
    """
    try:
        body = {'name': doc_name, 'parents': [DRIVE_FOLDER_ID]}
        return service.files().copy(body=body, fileId=tmpl_id,
                                    fields='id').execute().get('id')
    except HttpError as error:
        print(f"An error occurred: {error}")
        return error

def _set_permissions(file_id, service, user):
    try:
        body = {'role': 'owner', 'type': 'user', 'emailAddress': user}
        return service.permissions().create(fileId=file_id, transferOwnership=True, body=body).execute()
    except HttpError as error:
        print(f"An error occurred: {error}")
        return error

def merge_template(tmpl_id, source, service, doc_name):
    """Copies template document and merges data into newly-minted copy then
        returns its file ID.
    """
    try:
        # copy template and set context data struct for merging template values
        
        copy_id = _copy_template(tmpl_id, source, service, doc_name)
        # perms = _set_permissions(copy_id, service, USER)
        context = merge.iteritems() if hasattr({},
                                               'iteritems') else merge.items()

        # "search & replace" API requests for mail merge substitutions
        reqs = [{'replaceAllText': {
            'containsText': {
                'text': '{{%s}}' % key.upper(),  # {{VARS}} are uppercase
                'matchCase': True,
            },
            'replaceText': value,
        }} for key, value in context]

        # send requests to Docs API to do actual merge
        DOCS.documents().batchUpdate(body={'requests': reqs},
                                     documentId=copy_id, fields='').execute()
        return copy_id
    except HttpError as error:
        print(f"An error occurred: {error}")
        return error


if __name__ == '__main__':
    # fill-in your data to merge into document template variables
    merge = {
        'NAME': None,
        'ADDRESS': None,
        'CITY': None,
        'STATE': None,
        'ZIP': None,
        # - - - - - - - - - - - - - - - - - - - - - - - - - -
        'date': time.strftime('%Y %B %d'),
        # - - - - - - - - - - - - - - - - - - - - - - - - - -
    }

    # get row data, then loop through & process each form letter
    data = get_data(SOURCE)  # get data from data source
    for i, row in enumerate(data):
        merge_fields = dict(zip(COLUMNS, row))
        doc_name = f"{MERGED_FILENAME}_{merge_fields['NAME'].replace(' ', '_')}"
        print(merge_fields)
        merge.update(merge_fields)
        print('Merged letter %d: docs.google.com/document/d/%s/edit' % (
            i + 1, merge_template(DOCS_FILE_ID, SOURCE, DRIVE, doc_name)))