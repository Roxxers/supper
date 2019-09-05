import setuptools

# TODO: This needs to generate a config and make sure it is installed

setuptools.setup(
    name="Seating Plan Generator",
    author="Roxanne Gibson",
    author_email="me@roxanne.dev",
    description="Script to generate a seating plan via Office365 Calendars",
    packages=["seatingplan"],
    entry_points={"console_scripts": ["seatingplan=seatingplan.__main__:main"]},
    python_requires=">=3.5",
)