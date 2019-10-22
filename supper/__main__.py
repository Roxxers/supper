#!/usr/bin/env python3

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

"""Entry point for Supper Script"""

import argparse
import csv
import logging
from os.path import abspath
from datetime import datetime
from pathlib import Path

import yaml

from . import LOG, HANDLER, dates
from .api import Account

parser = argparse.ArgumentParser(
    prog='supper',
    description="Script to generate a seating plan via Office365 Calendars"
)
parser.add_argument(
    "-c", "--config",
    type=str,
    dest="config_path",
    default="{}/.config/supper.yaml".format(Path.home()),
    help="Path to a config file to read settings from. Default: ~/.config/supper.yaml"
)
parser.add_argument(
    "-o", "--output",
    type=str,
    dest="output_path",
    default="Seating Plan.csv",
    help="Path to save the output csv file to"
)
parser.add_argument("-d", "--debug", action="store_true", help="Enable debug output")


def read_config_file(config_path):
    """
    Reads config file and sets up variables foo

    :return: config as a dict
    """
    with open(config_path, "r") as file:
        config = yaml.load(file, Loader=yaml.FullLoader)
    return config


def format_output_path(output_path):
    """
    Checks the string for datetime formatting and formats it if possible.

    :param output_path: str of the output path
    :return: str of the new output path
    """
    try:
        new_path = output_path.format(datetime.now())
        if new_path.split(".")[-1] != "csv":
            new_path += ".csv"
            LOG.info("Output path does NOT have '.csv' file extension. Adding '.csv' to end of output_path.")
        LOG.debug("Formatted output file successfully as '%s'", new_path)
        return new_path
    except (SyntaxError, KeyError):
        # This is raised when formatting is incorrectly used in naming the output
        LOG.error("Invalid formatting pattern given for output_path. Cannot name output_path. Exiting...")
        exit(1)


def parse_args():
    """
    Parses arguments from the commandline.

    :return: config yaml file as a dict
    """
    args = parser.parse_args()

    if args.debug:
        LOG.setLevel(logging.DEBUG)
        HANDLER.setLevel(logging.DEBUG)
    else:
        LOG.setLevel(logging.WARNING)
        HANDLER.setLevel(logging.WARNING)

    output_path = format_output_path(args.output_path)

    if args.config_path:
        # Read the file provided and return the required config
        try:
            config = read_config_file(args.config_path)
            config["config_path"] = args.config_path
            config["output_path"] = output_path
            config["users"] = sorted([x.lower() for x in config["users"]])  # make all names lowercase and sort alphabetically
            LOG.debug("Loaded config successfully from '%s'", args.config_path)
            return config
        except FileNotFoundError:
            # Cannot open file
            LOG.error("Cannot find config file provided (%s). Maybe you mistyped it? Exiting...", args.config_path)
            exit(1)
        except (yaml.parser.ParserError, TypeError):
            # Cannot parse opened file
            # TypeError is sometimes raised if the metadata of the file is correct but the content doesn't parse
            LOG.error("Cannot parse config file. Make sure the provided config is a YAML file and that is is formatted correctly. Exiting...")
            exit(1)
    else:
        # No provided -c argument
        LOG.error("Cannot login. No config file was provided. Exiting...")
        exit(1)





def create_ooo_csv(ooo: list, users: list, output_path: str):
    """
    Creates a csv of who is in the office and on what day.

    :param ooo: a list of lists representing each day of a 5 day week. Each day's list has users who are not in that day
    :param users: a list of names of people in the office
    :param output_path: a str representing the output path of the csv file
    """
    with open(output_path, 'w', newline='', encoding='utf-8') as file:
        fieldnames = ("Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
        writer = csv.DictWriter(file, fieldnames=fieldnames)
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
    LOG.debug("Created csv file.")


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

    session = Account.create_session((client_id, client_secret), tenant_id)
    auth = False
    while not auth:
        auth = session.authenticate_session()

    LOG.debug("Session created and authenticated. %s", session)

    ooo = session.get_ooo_list(email)
    create_ooo_csv(ooo, users, output_path)
    monday, friday = dates.get_week_datetime()
    print("\nCreated CSV seating plan for week {:%a %d/%m/%Y} to {:%a %d/%m/%Y} at {}".format(monday, friday, abspath(output_path)))


if __name__ == "__main__":
    # This is for running the file in testing, rather than installing via pip
    main()
