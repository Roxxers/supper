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

"""Script to generate a seating plan via calendar events in an organisation's Office365."""

import os
import logging

STRFTIME = "%Y-%m-%dT%H:%M:%S"
STRPTIME = "%Y-%m-%dT%H:%M:%S.%f"
# Put access token where the file is so it can always access it.
ACCESS_TOKEN = f"{os.path.dirname(os.path.realpath(__file__))}/access_token"

LOG = logging.getLogger('supper')
HANDLER = logging.StreamHandler()
LOG.addHandler(HANDLER)

HANDLER.setFormatter(logging.Formatter('%(levelname)s: %(asctime)s - %(message)s'))
