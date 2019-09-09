#!/usr/bin/env python3

import csv
import argparse
import configparser
from datetime import datetime, timedelta
from O365 import (Account, Connection, FileSystemTokenBackend,
                  MSOffice365Protocol)


parser = argparse.ArgumentParser(prog='seatingplan', description="Script to generate a seating plan via Office365 Calendars")
parser.add_argument("-i", "--client-id", type=str, dest='client_id', default="",
                    help='Client ID for registered azure application')
parser.add_argument("-s", "--client-secret", type=str, dest='client_secret', default="",
                    help='Client secret for registered azure application')
parser.add_argument("-t", "--tenant-id", type=str, dest='tenant_id', default="",
                    help='Tenant ID for registered azure application')
parser.add_argument("-c", "--config", type=str, dest="config_path", default="",
                    help="Path to a config file to read settings from.")
# TODO: DATETIME FORMATTING, MAKE THIS AVALIABLE FOR THE OUTPUT PATH SO THAT YOU CAN DO THE COOL SCRIPT THINGS
parser.add_argument("-o", "--output", type=str, dest="output_path", default="Seating Plan.csv",
                    help="Path of the outputted csv file.")

# TODO: Setup providing config file via commandline args


def read_config_file(config_path):
    """
    Reads config file and sets up variables foo
    :return: credentials: ('client_id', 'client_secret') and tenant: str
    """
    config = configparser.ConfigParser()
    config.read(config_path)

    credentials = (config["client"]["id"], config["client"]["secret"])
    tenant = config["client"]["tenant"]

    return credentials, tenant


def parse_args():
    """
    Parses arguments from the commandline. 
    :return: exits or returns (client_id, client_secret) and tenant_id
    """
    # TODO: Maybe the prints before exit could be a lil better at explaining the error.
    args = parser.parse_args()
    using_args = False
    using_conf = False
    if args.client_id and args.tenant_id and args.client_secret:
        using_args = True
    if args.config_path:
        using_conf = True
    
    if using_args and using_conf: 
        # Using both types of config input. Confusing so dump it
        print("Cannot use both command-line arguments and config file. Please only use one.")
        exit(1)
    elif using_conf:
        # Read the file provided and return the required config
        return read_config_file(args.config_path), args.output_path
    elif using_args:
        # Just grab the config from the command line args
        return (args.client_id, args.client_secret), args.tenant_id, args.output_path
    else:
        # If the code has gotten here, then a config can't be parsed so we must close the program
        print("Cannot login. No config/partial was provided.")
        exit(1)


def create_session(credentials, tenant_id):
    """
    Create a session with the API and save the token for later use.
    :param credentials: tuple of (client_id, client_secret)
    :param tentant_id: str of tenant_id
    :return: Account class and email: str
    """
    my_protocol = MSOffice365Protocol(api_version='v2.0') 
    token_backend = FileSystemTokenBackend(token_filename='access_token')
    return Account(
        credentials, 
        protocol=my_protocol,
        tenant_id=tenant_id,
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


def create_csv(output_path):
    with open(output_path, 'w', newline='') as csvfile:
        fieldnames = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        pass
        # TODO: https://docs.python.org/3.7/library/csv.html#csv.DictWriter Need to finish this when we actually get some data to input.


def main():
    """
    Main function that is ran on start up. Script is ran from here.
    """
    credentials, tenant_id, output_path = parse_args()
    session = create_session(credentials, tenant_id)
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
    
    create_csv(output_path)


if __name__ == "__main__":
    main()
