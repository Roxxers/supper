
from O365 import Account, MSOffice365Protocol, FileSystemTokenBackend, Connection
import configparser
from datetime import datetime, timedelta


def read_config():
    """
    Reads config file and sets up variables foo
    :return: credentials: ('client_id', 'client_secret') and tenant: str
    """
    config = configparser.ConfigParser()
    config.read("./config")

    credentials = (config["client"]["id"], config["client"]["secret"])
    tenant = config["client"]["tenant"]

    return credentials, tenant


def create_session():
    """
    Create a session with the API and save the token for later use.
    :return: Account class
    """
    credentials, tenant = read_config()
    my_protocol = MSOffice365Protocol(api_version='v2.0')
    token_backend = FileSystemTokenBackend(token_filename='access_token.txt')
    return Account(credentials, protocol=my_protocol, tenant_id=tenant,
                   token_backend=token_backend)  # the default protocol will be Microsoft Graph


def authenticate_session(session: Account):
    """
    Authenticates account session object with oauth. Uses the default auth flow that comes with the library

    It could be merged with oauth wrapper but this works too.
    :param session: Account object
    :return:
    """
    try:
        session.con.oauth_request("https://outlook.office.com/api/v2.0/me/calendar", "get")
    except RuntimeError:  # Not authenticated. Need to ask user for url
        session.authenticate(scopes=['basic', 'address_book', 'users', 'calendar_shared'])
        session.con.oauth_request("https://outlook.office.com/api/v2.0/me/calendar", "get")


def oauth_request(connection: Connection, url: str):
    """
    Wrapper for Connection.oauth_request to provide some error handling
    :param connection:
    :param url:
    :return:
    """
    request = connection.oauth_request(url, "get")
    if request.status_code != 200:
        print(f"Request failed: GET request to {url} failed with {request.status_code} error")
    else:
        return request


def get_week_datetime():
    """
    Gets the current week's Monday and Friday to be used to filter a calendar.

    If this script is ran during the week, it will be the current week. If it is ran on the weekend, it will generate for next week.
    :return: Monday and Friday: Datetime object
    """
    today = datetime.now()
    weekday = today.weekday()
    # If this script is run during the week
    if weekday <= 4:  # 0 = Monday, 6 = Sunday
        monday = today - timedelta(days=weekday)  # Monday = 0-0, Friday = 4-4
        friday = today + timedelta(days=4 - weekday)  # Fri to Fri 4 + (4 - 4), Tues to Fri = 2 + (4 - 2)
        return monday, friday
    else:
        monday = today - timedelta(days=weekday) + timedelta(days=7)  # Monday = 0-0, Friday = 4-4
        friday = today + timedelta(days=4 - weekday) + timedelta(
            days=7)  # Fri to Fri 4 + (4 - 4), Tues to Fri = 2 + (4 - 2)
        return monday, friday


session = create_session()
authenticate_session(session)

r = oauth_request(session.con, "https://outlook.office.com/api/v2.0/users/EMAIL/")

# request = account.con.oauth_request("https://outlook.office.com/api/v2.0/users/EMAIL/calendar/events?$filter=Start/DateTime ge '2019-09-04T08:00' AND End/Datetime le '2019-09-05T08:00'&$top=50", "get")
print(r.text)

# setup csv
# For loop over all emails
# API request for user info of email
# Get display name
# API request for events this week
# Find sign of being in or not being in
# insert name in csv for the days they are in and at what seat (if regular)


monday, friday = get_week_datetime()
print(f"Monday is {monday}, Friday is {friday}")
