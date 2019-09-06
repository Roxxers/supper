#!/usr/bin/env python3


from O365 import Account, MSOffice365Protocol, FileSystemTokenBackend, Connection
import configparser
from datetime import datetime, timedelta
import argparse


parser = argparse.ArgumentParser(prog='seatingplan', description="Script to generate a seating plan via Office365 Calendars")
parser.add_argument("-c", "--client-id", type=str, required=True, dest='client_id',
                    help='Client ID for registered azure application')
parser.add_argument("-s", "--client-secret", type=str, required=True, dest='client_secret',
                    help='Client secret for registered azure application')
parser.add_argument("-t", "--tenant-id", type=str, required=True, dest='tenant_id',
                    help='Tenant ID for registered azure application')


def read_config():
    """
    Reads config file and sets up variables foo
    :return: credentials: ('client_id', 'client_secret') and tenant: str
    """
    # TODO: Deprecate this file so we use an arg version instead
    config = configparser.ConfigParser()
    config.read("./config")

    credentials = (config["client"]["id"], config["client"]["secret"])
    tenant = config["client"]["tenant"]
    email = config["client"]["out_of_office_email"]

    return credentials, tenant, email


def create_session():
    """
    Create a session with the API and save the token for later use.
    :return: Account class and email: str
    """
    config = parser.parse_args()
    credentials = (config.client_id, config.client_secret)
    tenant = config.tenant_id
    my_protocol = MSOffice365Protocol(api_version='v2.0') 
    token_backend = FileSystemTokenBackend(token_filename='access_token.txt')
    return Account(
        credentials, 
        protocol=my_protocol,
        tenant_id=tenant,
        token_backend=token_backend
        )

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

    If this script is ran during the work week (Monday-Friday), it will be the current week. If it is ran on the weekend, it will generate for next week.
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


def main():
    """
    Main function that is ran on start up. Script is ran from here.
    """
    
    session = create_session()
    authenticate_session(session)

    r = session.con.oauth_request("https://outlook.office.com/api/v2.0/users/{email}/calendar/events", "get")

    # request = account.con.oauth_request("https://outlook.office.com/api/v2.0/users/EMAIL/calendar/events?$filter=Start/DateTime ge '2019-09-04T08:00' AND End/Datetime le '2019-09-05T08:00'&$top=50", "get")
    print(r.text)

    # setup csv
    # API request events from out of office email
    # loop over events
        # find out who the event is talking about
        # insert into csv
    # output csv


    monday, friday = get_week_datetime()
    print(f"Monday is {monday}, Friday is {friday}")

if __name__ == "__main__":
    main()
