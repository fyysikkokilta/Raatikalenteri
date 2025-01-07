import os
import datetime
import time
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ["https://www.googleapis.com/auth/calendar"]
CREDENTIALS_FILE = "client_secret.json"

pickle_file = "token.pickle"
expiration_time = datetime.timedelta(days=6)


def get_calendar_service():
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists(pickle_file):
        # file_age = datetime.timedelta(seconds=time.time() - os.path.getmtime(pickle_file))
        # if file_age >= expiration_time:
        #    print()
        #    print("Removing expired pickle")
        #    print()
        #    os.remove(pickle_file)
        # else:
        with open(pickle_file, "rb") as token:
            creds = pickle.load(token)

    # If there are no (valid) credentials available, let the user log in.
    if not creds:
        flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
        creds = flow.run_local_server(port=0)

        # Save the credentials for the next run
        with open(pickle_file, "wb") as token:
            pickle.dump(creds, token)

    service = build("calendar", "v3", credentials=creds)
    return service
