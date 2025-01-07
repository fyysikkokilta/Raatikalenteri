import time
import datetime
import pygsheets
from cal_setup import get_calendar_service
import pickle
import os.path
import subprocess
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from dotenv import load_dotenv

load_dotenv()

sheet_id = os.getenv("SHEET_ID")
calendar_id = os.getenv("CALENDAR_ID")
service = get_calendar_service()

SCOPES = ["https://www.googleapis.com/auth/calendar"]
CREDENTIALS_FILE = "client_secret.json"
SERVICE_ACCOUNT_FILE = "service_account_credentials.json"
PICKLE_FILE = "token.pickle"
CALENDAR_BOTTOM_RIGHT = "O53"
CALENDAR_LEFT = "B"

start_date = datetime.datetime(2023, 12, 25)  # First date on calendar
cur_week = datetime.date.today().isocalendar()[1]  # Current calendar week
calendar_range = (
    CALENDAR_LEFT + str(2 + cur_week),
    CALENDAR_BOTTOM_RIGHT,
)  # Ignore past weeks


def main():
    try:
        get_calendar_service()

        print("Getting events...")
        print()
        events = read_sheet()
        for d in range(len(events)):
            events_result = (
                service.events()
                .list(
                    calendarId=calendar_id,
                    timeMin=(
                        start_date.replace(hour=0)
                        + datetime.timedelta(days=d + cur_week * 7)
                    ).isoformat()
                    + "+02:00",
                    timeMax=(
                        start_date.replace(hour=23)
                        + datetime.timedelta(days=d + cur_week * 7)
                    ).isoformat()
                    + "+02:00",
                    singleEvents=True,
                    orderBy="startTime",
                )
                .execute()
            )
            existing_events = events_result.get("items", [])
            for existing_event in existing_events:
                try:
                    if existing_event["summary"] not in events[d]:
                        service.events().delete(
                            calendarId=calendar_id, eventId=existing_event["id"]
                        ).execute()
                        print(f"Deleted event '{existing_event['summary']}'")
                        print()
                    else:
                        events[d].remove(existing_event["summary"])
                except KeyError:
                    service.events().delete(
                        calendarId=calendar_id, eventId=existing_event["id"]
                    ).execute()
                    print(f"Deleted event '{existing_event}'")
                    print()
            if len(events[d]) > 0:
                for event in events[d]:
                    if event.strip() != "":
                        date = start_date + datetime.timedelta(days=d + cur_week * 7)
                        create_event(event, date)

        print("done")
        time.sleep(1)
    except BaseException as error:
        print("An exception occurred: {}".format(error))

        print("Removing pickle")
        print()
        os.remove(PICKLE_FILE)
        print("Retrying ...")
        print()
        time.sleep(3)
        subprocess.call([r"raatikalenteri.bat"])


def read_sheet():
    gc = pygsheets.authorize(service_file=SERVICE_ACCOUNT_FILE)
    sheet = gc.open_by_key(sheet_id)

    worksheet = sheet.worksheet()
    calendar_cells = worksheet.get_values(
        start=calendar_range[0], end=calendar_range[1], returnas="cells"
    )

    events = []
    for w in range(len(calendar_cells)):
        for d in range(1, len(calendar_cells[w]), 2):
            events.append(calendar_cells[w][d].value.splitlines())

    return events


def create_event(name, start, end=None):
    if end is None:
        end = start
    end = end + datetime.timedelta(days=1)
    event_result = (
        service.events()
        .insert(
            calendarId=calendar_id,
            body={
                "summary": name,
                "description": "",
                "start": {"date": start.isoformat().split("T")[0]},
                "end": {"date": end.isoformat().split("T")[0]},
                "transparency": "transparent",
            },
        )
        .execute()
    )
    print(
        f"Created event '{event_result['summary']}' [",
        start.isoformat().split("T")[0],
        "]",
    )
    print()


def get_calendar_service():
    creds = None
    if os.path.exists(PICKLE_FILE):
        with open(PICKLE_FILE, "rb") as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)

        with open(PICKLE_FILE, "wb") as token:
            pickle.dump(creds, token)

    service = build("calendar", "v3", credentials=creds)
    return service


if __name__ == "__main__":
    main()
