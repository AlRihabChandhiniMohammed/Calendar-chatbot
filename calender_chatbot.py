import datetime
import os.path
import logging
import pytz

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.interval import IntervalTrigger

try:
    from plyer import notification
except ImportError:
    notification = None
    print("Warning: Plyer not installed. Desktop notifications will not be available. Install with 'pip install plyer'")

SCOPES = ['https://www.googleapis.com/auth/calendar']
TOKEN_FILE = 'token.json'
CREDENTIALS_FILE = 'credentials.json'

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

LOCAL_TIMEZONE = pytz.timezone('Asia/Kolkata')

def get_calendar_service():
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())
    service = build('calendar', 'v3', credentials=creds)
    return service

def create_event(service, summary, start_time_local, end_time_local, description=None, location=None, attendees=None, minutes_before_reminder=15):
    start_time_utc = start_time_local.astimezone(pytz.utc)
    end_time_utc = end_time_local.astimezone(pytz.utc)

    event = {
        'summary': summary,
        'location': location,
        'description': description,
        'start': {
            'dateTime': start_time_utc.isoformat(),
            'timeZone': 'UTC',
        },
        'end': {
            'dateTime': end_time_utc.isoformat(),
            'timeZone': 'UTC',
        },
        'reminders': {
            'useDefault': False,
            'overrides': [
                {'method': 'email', 'minutes': minutes_before_reminder},
                {'method': 'popup', 'minutes': minutes_before_reminder},
            ],
        },
    }
    if attendees:
        event['attendees'] = [{'email': email} for email in attendees]

    try:
        event = service.events().insert(calendarId='primary', body=event).execute()
        logger.info(f"Event created: {event.get('htmlLink')}")
        return event
    except HttpError as error:
        logger.error(f"An error occurred while creating event: {error}")
        return None

def get_upcoming_events(service, max_results=10, time_min_local=None, time_max_local=None):
    if time_min_local is None:
        time_min_local = datetime.datetime.now(LOCAL_TIMEZONE)
    if time_max_local is None:
        time_max_local = time_min_local + datetime.timedelta(days=7)

    time_min_utc = time_min_local.astimezone(pytz.utc).isoformat()
    time_max_utc = time_max_local.astimezone(pytz.utc).isoformat()

    logger.info(f'Getting upcoming events from {time_min_local.strftime("%Y-%m-%d %H:%M")} to {time_max_local.strftime("%Y-%m-%d %H:%M")}')
    try:
        events_result = service.events().list(
            calendarId='primary',
            timeMin=time_min_utc,
            timeMax=time_max_utc,
            maxResults=max_results,
            singleEvents=True,
            orderBy='startTime'
        ).execute()
        events = events_result.get('items', [])
        return events
    except HttpError as error:
        logger.error(f"An error occurred while fetching events: {error}")
        return []

SENT_REMINDERS = {}

def send_notification(title, message):
    if notification:
        try:
            notification.notify(
                title=title,
                message=message,
                app_name='Virtual Event Scheduler',
                timeout=10
            )
            logger.info(f"Desktop Notification Sent: {title} - {message}")
        except Exception as e:
            logger.error(f"Failed to send desktop notification: {e}")
            logger.info(f"Console Reminder: {title} - {message}")
    else:
        logger.info(f"Console Reminder: {title} - {message}")

def check_and_remind(service):
    now_local = datetime.datetime.now(LOCAL_TIMEZONE)
    logger.info(f"Running reminder check at {now_local.strftime('%Y-%m-%d %H:%M:%S')}")

    fetch_start_time = now_local - datetime.timedelta(minutes=5)
    fetch_end_time = now_local + datetime.timedelta(minutes=60)

    events = get_upcoming_events(service, time_min_local=fetch_start_time, time_max_local=fetch_end_time, max_results=20)

    if not events:
        logger.info("No upcoming events found in the current check window.")
        return

    for event in events:
        event_id = event['id']
        event_summary = event.get('summary', 'No Title')
        event_html_link = event.get('htmlLink', '#')

        if event_id not in SENT_REMINDERS:
            SENT_REMINDERS[event_id] = {}

        start_data = event['start']
        if 'dateTime' in start_data:
            try:
                event_start_time_utc = datetime.datetime.fromisoformat(start_data['dateTime'].replace('Z', '+00:00'))
            except ValueError:
                logger.warning(f"Could not parse '{start_data['dateTime']}' with fromisoformat for event {event_summary}. Trying strptime.")
                try:
                    event_start_time_utc = datetime.datetime.strptime(start_data['dateTime'], "%Y-%m-%dT%H:%M:%S%z")
                except ValueError:
                    logger.error(f"Failed to parse event start time for {event_summary}: {start_data['dateTime']}")
                    continue

            event_start_time_local = event_start_time_utc.astimezone(LOCAL_TIMEZONE)

            time_to_event = event_start_time_local - now_local

            reminder_thresholds = [
                15,
                5,
                1,
            ]

            if datetime.timedelta(minutes=-5) < time_to_event <= datetime.timedelta(seconds=0):
                if not SENT_REMINDERS[event_id].get('started', False):
                    title = f"Event Started: {event_summary}"
                    message = f"It's happening now! Link: {event_html_link}"
                    send_notification(title, message)
                    SENT_REMINDERS[event_id]['started'] = True
                    logger.info(f"Notification sent for '{event_summary}' as it started.")
                continue

            for threshold_minutes in sorted(reminder_thresholds, reverse=True):
                threshold_timedelta = datetime.timedelta(minutes=threshold_minutes)

                if datetime.timedelta(seconds=-10) <= (time_to_event - threshold_timedelta) < datetime.timedelta(minutes=1) and \
                   time_to_event > datetime.timedelta(seconds=0) and \
                   not SENT_REMINDERS[event_id].get(threshold_minutes, False):

                    title = f"Upcoming Event: {event_summary}"
                    minutes_left = int(time_to_event.total_seconds() // 60)
                    if minutes_left == 0:
                        minutes_text = "less than a minute"
                    else:
                        minutes_text = f"{minutes_left} minutes"

                    message = (
                        f"Starts in approx. {minutes_text} "
                        f"at {event_start_time_local.strftime('%I:%M %p')} ({LOCAL_TIMEZONE.tzname(event_start_time_local)}).\n"
                        f"Link: {event_html_link}"
                    )
                    send_notification(title, message)
                    SENT_REMINDERS[event_id][threshold_minutes] = True
                    logger.info(f"Reminder sent for '{event_summary}' for {threshold_minutes} min threshold.")
                    break

        else:
            event_date_str = start_data['date']
            event_date = datetime.datetime.strptime(event_date_str, "%Y-%m-%d").date()
            if event_date == now_local.date() and not SENT_REMINDERS[event_id].get('all_day_today', False):
                title = f"All-Day Event Today: {event_summary}"
                message = f"This all-day event is happening today! Details: {event_html_link}"
                send_notification(title, message)
                SENT_REMINDERS[event_id]['all_day_today'] = True
                logger.info(f"All-day event notification sent for '{event_summary}'.")

if __name__ == '__main__':
    service = get_calendar_service()

    print("\n--- Starting Virtual Event Reminder Bot ---")
    print(f"Bot will check for events every 60 seconds (1 minute) in timezone: {LOCAL_TIMEZONE.tzname(datetime.datetime.now(LOCAL_TIMEZONE))}")
    print("Press Ctrl+C to stop the bot.")

    scheduler = BackgroundScheduler()
    scheduler.add_job(lambda: check_and_remind(service), IntervalTrigger(seconds=60))
    scheduler.start()

    try:
        while True:
            datetime.datetime.now()
            pass
    except (KeyboardInterrupt, SystemExit):
        logger.info("Bot stopped by user (KeyboardInterrupt/SystemExit).")
        scheduler.shutdown()
        print("Scheduler shut down. Exiting.")