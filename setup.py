
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

import sys
import setuptools


if sys.version_info < (3, 5):
    sys.exit('Python 3.5 is required to run Supper')


LONG_DESC = open('readme.md').read()

setuptools.setup(
    name="Supper",
    author="Roxanne Gibson",
    author_email="me@roxanne.dev",
    description="Script to generate a seating plan via Office365 Calendars",
    long_description_content_type="text/markdown",
    long_description=LONG_DESC,
    packages=["supper"],
    entry_points={"console_scripts": ["supper=supper.__main__:main"]},
    python_requires=">=3.5",
    install_requires=("o365==2.0.1", "pyyaml==5.1.1"),
    version="1.1.0",
    license="GPL-3"
)
