
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
    version="1.0",
    license="GPL-3"
)
