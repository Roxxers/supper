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

"""Collection of functions to deal with datetime manipulation"""

from datetime import datetime, timedelta

def get_week_datetime(start: datetime = None):
    """
    Gets the current week's Monday and Friday to be used to filter a calendar.

    If this script is ran during the work week (Monday-Friday), it will be the current week.
    If it is ran on the weekend, it will generate for next week.

    :param start: Specifies the today variable, used for make future weeks if given.
    :return: Monday and Friday: Datetime objects
    """
    if start:
        today = start
    else:
        today = datetime.now()

    weekday = today.weekday()
    extra_time = timedelta(
        hours=today.hour,
        minutes=today.minute,
        seconds=today.second,
        microseconds=today.microsecond
    )
    monday = today - timedelta(days=weekday) - extra_time  # Monday = 0-0, Friday = 4-4
    friday = (today + timedelta(days=4 - weekday)) - extra_time + timedelta(hours=23, minutes=59)

    # If this script is run during the week
    if weekday <= 4:  # 0 = Monday, 6 = Sunday
        # - the hour and minutes to get start of the day instead of when the
        return monday, friday
    # If the date the script is ran on is the weekend, do next week instead
    monday = monday + timedelta(days=7)  # Monday = 0-0, Friday = 4-4
    friday = friday + timedelta(days=7)  # Fri to Fri 4 + (4 - 4), Tues to Fri = 2 + (4 - 2)
    return monday, friday
