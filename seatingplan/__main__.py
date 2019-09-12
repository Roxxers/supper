#!/usr/bin/env python3

import csv
import json
import yaml
import argparse
import configparser
from datetime import datetime, timedelta
from O365 import (Account, Connection, FileSystemTokenBackend,
                  MSOffice365Protocol)


parser = argparse.ArgumentParser(prog='seatingplan', description="Script to generate a seating plan via Office365 Calendars")
parser.add_argument("-c", "--config", type=str, dest="config_path", default="",
                    help="Path to a config file to read settings from.")
# TODO: DATETIME FORMATTING, MAKE THIS AVALIABLE FOR THE OUTPUT PATH SO THAT YOU CAN DO THE COOL SCRIPT THINGS
parser.add_argument("-o", "--output", type=str, dest="output_path", default="Seating Plan.csv",
                    help="Path of the outputted csv file.")

strftime_pattern = "%Y-%m-%dT%H:%M:%S"
strptime_pattern = "%Y-%m-%dT%H:%M:%S.%f"

# TODO: Input for staff emails

def read_config_file(config_path):
    """
    Reads config file and sets up variables foo
    :return: config as a dict
    """
    with open(config_path, "r") as fp:
        config = yaml.load(fp)
    return config


def parse_args():
    """
    Parses arguments from the commandline. 
    :return: config yaml file as a dict
    """
    # TODO: Maybe the prints before exit could be a lil better at explaining the error.
    args = parser.parse_args()

    # Read the file provided and return the required config
    
    if args.config_path:
        config = read_config_file(args.config_path)
        config["output_path"] = args.output_path or "./Seating Plan.csv"  # Create a default seating plan | I could do this in argparser tbh
        config["users"] = sorted([x.lower() for x in config["users"]])  # make all names lowercase and sort alphabetically
        return config

    else:
        # If the code has gotten here, then a config can't be parsed so we must close the program
        print("Cannot login. No config file was provided.")
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
        # - the hour and minutes to get start of the day instead of when the
        extra_time = timedelta(hours=today.hour, minutes=today.minute, seconds=today.second, microseconds=today.microsecond)
        monday = today - timedelta(days=weekday) - extra_time  # Monday = 0-0, Friday = 4-4
        friday = (today + timedelta(days=4 - weekday)) - extra_time + timedelta(hours=23, minutes=59)
        return monday, friday
    
    # If the date the script is ran on is the weekend, do next week instead
    monday = monday + timedelta(days=7)  # Monday = 0-0, Friday = 4-4
    friday = friday + timedelta(days=7)  # Fri to Fri 4 + (4 - 4), Tues to Fri = 2 + (4 - 2)
    return monday, friday


def get_event_range(beginning_of_week: datetime, connection: Connection, email: str):
    base_url = "https://outlook.office.com/api/v2.0/"
    scope = f"users/{email}/CalendarView"
    
    # Create a range of dates for this week so we can catch long events within the search
    bottom_range = beginning_of_week - timedelta(days=14)
    top_range = beginning_of_week + timedelta(days=21)
    
    # Setup url
    date_range = "startDateTime={}&endDateTime={}".format(bottom_range.strftime(strftime_pattern), top_range.strftime(strftime_pattern))
    limit = "$top=150"
    select = "$select=Subject,Organizer,Start,End,Attendees"
    
    url = f"{base_url}{scope}?{date_range}&{select}&{limit}"
    r = connection.oauth_request(url, "get")
    
    return r.json()


def add_attendees_to_ooo_list(attendees: list, ooo_list: list):
    for attendee in attendees:
        attendee_name = attendee["EmailAddress"]["Name"].split(" ")[0]  # Get first name
        if attendee_name not in ooo_list.copy():
            ooo_list.append(attendee_name.lower())
    return ooo_list


def get_ooo_list(email: str, connection: Connection):
    # TODO: Docstrings of new functions
    # Get month
    # check for any events that are longer than 1 day that has a start or end point in the month
    # TODO: Update readme to explain that if an event is longer than a month, it won't be picked up by the script
    # check if those events happen within the week by checking if the day we are checking is in between the two points of the event
    # add this to a list of just events happening in this week that are defo a day long or less.
    
    monday, friday = get_week_datetime()
    events = get_event_range(monday, connection, email)
    events = events["value"]
    outofoffice = [[], [], [], [], []]
    # Using a list to take advantage of datetime.weekday instead of dealing with trying to figure out what key to use

    for event in events:
        # removes last char due to microsofts datetime using 7 sigfigs for microseconds, python uses 6
        start = datetime.strptime(event["Start"]["DateTime"][:-1], strptime_pattern) 
        end = datetime.strptime(event["End"]["DateTime"][:-1], strptime_pattern)
        attendees = event["Attendees"]
        # excluding outofoffice account
        attendees = [x for x in attendees if x["EmailAddress"]["Address"] != email]
        organizer = event["Organizer"]
        
        if not attendees and organizer["EmailAddress"]["Address"] != email:
            # Sometimes user will be the one who makes the event, not the outofoffice account. Get the organizer.
            attendees = [event["Organizer"]]
            
        if (end - start) <= timedelta(days=1):
            # Event is for one day only, check if it starts
            if monday <= start <= friday:
                # Event is within the week we are looking at, add all attendees +
                weekday = outofoffice[start.weekday()]
                weekday = add_attendees_to_ooo_list(attendees, weekday)

        else:
            # Check if long events cover the days of this week
            for x, day_array in enumerate(outofoffice.copy()):
                current_day = monday + timedelta(days=x)
                if start <= current_day <= end:
                    # if day is inside of the long event
                    outofoffice[x] = add_attendees_to_ooo_list(attendees, day_array)
    
    return outofoffice


def create_csv(ooo: list, users: list, output_path: str):
    # maybe input all the names then selectivly remove them based on the events?
    with open(output_path, 'w', newline='', encoding='utf-8') as fp:
        fieldnames = ("Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
        writer = csv.DictWriter(fp, fieldnames=fieldnames)
        writer.writeheader()
        for user in users:
            row = {}
            for i, day in enumerate(fieldnames):
                if user not in ooo[i]:
                    row[day] = user
                else:
                    row[day] = ""
            writer.writerow(row)


def main():
    """
    Main function that is ran on start up. Script is ran from here.
    """
    config = parse_args()
    session = create_session((config["client_id"], config["client_secret"]), config["tenant_id"])
    authenticate_session(session)
    
    ooo = get_ooo_list(config["ooo_email"], session.con)
    create_csv(ooo, config["users"], config["output_path"])


if __name__ == "__main__":
    # This is for running the file in testing, rather than installing via pip
    main()
