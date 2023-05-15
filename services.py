'''
common client libraries for googleapi
'''
import os
import tempfile

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

DRIVE_CLIENT = None
SPREADSHEET_CLIENT = None
GMAIL_CLIENT = None
CRED_FILE_NAME = "credentials.json"


def get_creds(client_alias, scopes, re_grant=False):
    '''
    get credentials
    '''
    if not os.path.exists(CRED_FILE_NAME):
        raise FileNotFoundError(
            F"please refer https://developers.google.com/workspace/guides/create-credentials to get your credential file, and store as {CRED_FILE_NAME} in your dir")

    creds = None
    token_file = os.path.join(tempfile.gettempdir(), client_alias)
    if re_grant is False and os.path.exists(token_file):
        creds = Credentials.from_authorized_user_file(token_file, scopes)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                CRED_FILE_NAME, scopes)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(token_file, 'w', encoding="utf8") as token:
            token.write(creds.to_json())

    return creds


def get_spreadsheet_service(scopes=None):
    '''
    get client by scope, if scope is None, will using default
    '''
    global SPREADSHEET_CLIENT
    re_grant = False

    if scopes is None:
        scopes = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/spreadsheets.readonly',
        ]
    else:
        re_grant = True

    if SPREADSHEET_CLIENT is None:
        creds = get_creds("ss", scopes, re_grant)
        SPREADSHEET_CLIENT = build("sheets", "v4", credentials=creds)

    return SPREADSHEET_CLIENT


def get_drive_service(scopes=None):
    '''
    get client by scope, if scope is None, using default scope
    '''
    global DRIVE_CLIENT
    re_grant = False

    if scopes is None:
        scopes = [
            'https://www.googleapis.com/auth/drive.file',
            'https://www.googleapis.com/auth/drive',
            'https://www.googleapis.com/auth/drive.appdata',
        ]
    else:
        re_grant = True

    if DRIVE_CLIENT is None:
        creds = get_creds("drive", scopes, re_grant)
        DRIVE_CLIENT = build("drive", "v3", credentials=creds)

    return DRIVE_CLIENT


def get_gmail_service(scopes=None):
    '''
    get gmail client by scope, if scope is None, using default scope
    '''
    global GMAIL_CLIENT
    re_grant = False

    if scopes is None:
        scopes = [
            'https://www.googleapis.com/auth/gmail.compose'
        ]
    else:
        re_grant = True

    if GMAIL_CLIENT is None:
        creds = get_creds("gmail", scopes, re_grant)
        GMAIL_CLIENT = build("gmail", "v1", credentials=creds)

    return GMAIL_CLIENT
