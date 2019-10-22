# Copyright (C) 2019  Campaign Against Arms Trade
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.

"""Handles all api interaction with Office365"""

import os
from os.path import dirname, isfile, realpath
from datetime import datetime, timedelta

import O365
from requests import HTTPError

from . import ACCESS_TOKEN, STRFTIME, STRPTIME, LOG, dates


class Account(O365.Account):
    """Wrapper for the O365 Account class to add our api interactions"""
    @classmethod
    def create_session(cls, credentials, tenant_id):
        """
        Create a session with the API and save the token for later use.

        :param credentials: tuple of (client_id, client_secret)
        :param tentant_id: str of tenant_id
        :return: Account class and email: str
        """
        my_protocol = O365.MSOffice365Protocol(api_version='v2.0')
        token_backend = O365.FileSystemTokenBackend(token_filename=ACCESS_TOKEN)
        return cls(
            credentials,
            protocol=my_protocol,
            tenant_id=tenant_id,
            token_backend=token_backend,
            raise_http_error=False
        )

    def authenticate_session(self):
        """
        Authenticates account session object with oauth.
        Uses the default auth flow that comes with the library

        :param session: Account object
        :return: Bool for if the client is authenticated
        """
        if not isfile(f"{dirname(realpath(__file__))}/access_token"):
            try:
                self.authenticate(scopes=['basic', 'address_book', 'users', 'calendar_shared'])
                self.con.oauth_request("https://outlook.office.com/api/v2.0/me/", "get")
                LOG.debug("Successfully tested new access_token")
                return True
            except:
                os.remove(ACCESS_TOKEN)
                LOG.error("Could not authenticate. Make sure config has the correct client_id, client_secret, etc. Exiting...")
                exit(1)
        else:
            try:
                self.con.oauth_request("https://outlook.office.com/api/v2.0/me/", "get")
                LOG.debug("Successfully tested current access_token")
                return True
            except:
                os.remove(ACCESS_TOKEN)
                LOG.warning("Failed to authenticate with current access_token. Deleting access_token and running authentication again.")
                return False

    def get_event_range(self, beginning_of_week: datetime, email: str):
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
        date_range = "startDateTime={}&endDateTime={}".format(bottom_range.strftime(STRFTIME), top_range.strftime(STRPTIME))
        limit = "$top=150"
        select = "$select=Subject,Organizer,Start,End,Attendees"

        url = f"{base_url}{scope}?{date_range}&{select}&{limit}"
        resp = self.con.oauth_request(url, "get")
        return resp.json()

    def get_ooo_list(self, email: str):
        """
        Makes request and parses data into a list of users who will not be in the office

        :param email: string of the outofoffice email where the out of office calender is located
        :param connection: a connection to the office365 api
        :return: list of 5 lists representing 5 days. Each contains lowercase names of who is not in the office.
        """
        monday, friday = dates.get_week_datetime()
        try:
            events = self.get_event_range(monday, email)
            LOG.debug("Received response for two week range from week beginning with {:%Y-%m-%d} from outofoffice account with email: {}".format(monday, email))
        except HTTPError as error:
            LOG.error("Could not request CalendarView | %s", error.response)
            exit(1)

        events = events["value"]
        outofoffice = [[], [], [], [], []]

        for event in events:
            # removes last char due to microsoft's datetime using 7 sigfigs for microseconds, python uses 6
            start = datetime.strptime(event["Start"]["DateTime"][:-1], STRPTIME)
            end = datetime.strptime(event["End"]["DateTime"][:-1], STRPTIME)
            attendees = event["Attendees"]
            # remove outofoffice account by list comprehension
            attendees = [x for x in attendees if x["EmailAddress"]["Address"] != email]
            organizer = event["Organizer"]

            if not attendees and organizer["EmailAddress"]["Address"] != email:
                # Sometimes user will be the one who makes the event, not the outofoffice account. Get the organizer.
                attendees = [event["Organizer"]]

            if (end - start) <= timedelta(days=1):
                # Event is for one day only, check if it starts within the week
                if monday <= start <= friday:
                    # Event is within the week we are looking at, add all attendees
                    weekday = outofoffice[start.weekday()]
                    if not attendees:
                        LOG.warning("Event '%s' has no attendees. Cannot add to outofoffice list.", event["Subject"])
                    weekday = self.add_attendees_to_ooo_list(attendees, weekday)
            else:
                # Check if long events cover the days of this week
                for i, day_array in enumerate(outofoffice.copy()):
                    current_day = monday + timedelta(days=i)
                    if start <= current_day <= end:
                        # if day is inside of the long event
                        outofoffice[i] = self.add_attendees_to_ooo_list(attendees, day_array)
        LOG.debug("Parsed events and successfully created out of office list.")
        return outofoffice

    @staticmethod
    def add_attendees_to_ooo_list(attendees: list, ooo_list: list):
        """
        Adds attendees to ooo_list from api list of attendees

        :param attendees: list of attendees to event
        :param ooo_list: list of the current days out of office users
        :return: ooo_list once appended
        """
        for attendee in attendees:
            attendee_name = attendee["EmailAddress"]["Name"].split(" ")[0]  # Get first name
            if attendee_name not in ooo_list.copy():
                # Converted to lowercase so program is case insensitive
                ooo_list.append(attendee_name.lower())
        return ooo_list
