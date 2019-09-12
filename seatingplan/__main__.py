#!/usr/bin/env python3

import csv
import json
import yaml
import argparse
import configparser
from os.path import abspath, dirname, realpath
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


def read_config_file(config_path):
    """
    Reads config file and sets up variables foo
    
    :return: config as a dict
    """
    with open(config_path, "r") as fp:
        config = yaml.load(fp, Loader=yaml.FullLoader)
    return config


def parse_args():
    """
    Parses arguments from the commandline. 
    
    :return: config yaml file as a dict
    """
    args = parser.parse_args()

    if args.config_path:
        # Read the file provided and return the required config
        config = read_config_file(args.config_path)
        config["config_path"] = args.config_path
        config["output_path"] = args.output_path
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
    token_backend = FileSystemTokenBackend(token_filename=f"{dirname(realpath(__file__))}/access_token") # Put access token where the file is so it can always access it.
    return Account(
        credentials, 
        protocol=my_protocol,
        tenant_id=tenant_id,
        token_backend=token_backend
        )


def authenticate_session(session: Account):
    """
    Authenticates account session object with oauth. Uses the default auth flow that comes with the library
    
    :param session: Account object
    :return:
    """
    try:
        session.con.oauth_request("https://outlook.office.com/api/v2.0/me/calendar", "get")
    except RuntimeError:  # Not authenticated. Need to ask user for url
        session.authenticate(scopes=['basic', 'address_book', 'users', 'calendar_shared'])
        session.con.oauth_request("https://outlook.office.com/api/v2.0/me/calendar", "get")


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
    """
    Makes api call to grab a calender view within a 2 week window either side of the current week.
    
    :param beginning_of_week: datetime object for this weeks monday
    :param connection: a connection to the office365 api
    :return: dict of json response
    """
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
    """
    Function to aid the adding of attendees in a list to the list of people who will be out of offce.
    
    :param attendees: list of attendees to event
    :param ooo_list: list of the current days out of office users
    :return: ooo_list once appended
    """
    for attendee in attendees:
        attendee_name = attendee["EmailAddress"]["Name"].split(" ")[0]  # Get first name
        if attendee_name not in ooo_list.copy():
            ooo_list.append(attendee_name.lower())  # Converted to lowercase so program is case insensitive
    return ooo_list


def get_ooo_list(email: str, connection: Connection):
    """
    Makes request and parses data into a list of users who will not be in the office
    
    :param email: string of the outofoffice email where the out of office calender is located
    :param connection: a connection to the office365 api
    :return: list of 5 lists representing a 5 day list. Each list contains the lowercase names of who is not in the office.
    """
    monday, friday = get_week_datetime()
    events = get_event_range(monday, connection, email)
    
    events = events["value"]
    outofoffice = [[], [], [], [], []]

    for event in events:
        # removes last char due to microsoft's datetime using 7 sigfigs for microseconds, python uses 6
        start = datetime.strptime(event["Start"]["DateTime"][:-1], strptime_pattern) 
        end = datetime.strptime(event["End"]["DateTime"][:-1], strptime_pattern)
        attendees = event["Attendees"]
        # remove outofoffice account by list comprehension 
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


def create_ooo_csv(ooo: list, users: list, output_path: str):
    """
    Creates a csv of who is in the office and on what day.
    
    :param ooo: a list of lists representing each day of a 5 day week. Each day's list has users who are not in that day
    :param users: a list of names of people in the office
    :param output_path: a str representing the output path of the csv file
    """
    with open(output_path, 'w', newline='', encoding='utf-8') as fp:
        fieldnames = ("Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
        writer = csv.DictWriter(fp, fieldnames=fieldnames)
        writer.writeheader()
        for user in users:
            row = {}
            for i, day in enumerate(fieldnames):
                # for each day, check if the user is in that day, if not do not write their name into that day.
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
    
    client_id = config["client_id"]
    client_secret = config["client_secret"]
    tenant_id = config["tenant_id"]
    output_path = config["output_path"]
    users = config["users"]
    email = config["ooo_email"]
    
    session = create_session((client_id, client_secret), tenant_id)
    authenticate_session(session)
    
    ooo = get_ooo_list(email, session.con)
    create_ooo_csv(ooo, users, output_path)
    print("Created {}".format(abspath(output_path)))


if __name__ == "__main__":
    # This is for running the file in testing, rather than installing via pip
    main()
