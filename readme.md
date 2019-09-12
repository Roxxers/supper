
# Seating Plan Generator

Script to generate a seating plan via calendar events in an organsation's Office365.

## What it does

The script looks at the current week and generates a seating plan for that week. By getting the calendar of a dedicated room or account, it can see who will be out of the office during the week. It then creates a csv of who will be in the office on the 5 days of the week.

> **Note:** Current week is defined during the normal work weeks. If the script is ran on the weekend (Saturday and Sunday) the script will generate next weeks and label it as such.

## Pre-Install

To setup the script, you will need to create an app in your organisation Azure Active Directory. You can find the app registration page [here](hhttps://portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/RegisteredApps).

For a guide on how to do this, see the guide provided by python-o365 below.

> To work with oauth you first need to register your application at [Azure App Registrations](https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade).

>    1. Login at [Azure Portal (App Registrations)](https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade)
>    1. Create an app. Set a name.
>    1. In Supported account types choose "Accounts in any organizational directory.
>    1. Set the redirect uri (Web) to: `https://login.microsoftonline.com/common/oauth2/nativeclient` and click register. This is the default redirect uri used by this library, but you can use any other if you want.
>    1. Write down the Application (client) ID AND Directory (tenant) ID. You will need these values.
>    1. Under "Certificates & secrets", generate a new client secret. Set the expiration preferably to never.
>    1. Write down the value of the client secret created now. It will be hidden later on.
>    1. Under Api Permissions add the delegated permissions for Microsoft Graph you want

The required Api Permissions are:

```
Calendars.Read.Shared
User.Read
User.ReadBasic.All
offline_access
```

## Installation

> **Warning:** This guide assumes you are using a UNIX based OS (Linux, Mac OS, etc.). If using Windows, god help you. (Ask me for help if you can't adapt this to Windows. Windows is weird and scary.)

Once the app has been created, git clone this repo, cd into it's folder and install it into your user's Python PATH.

```sh
git clone URL
cd seatplangen
python3 -m pip install . --user
```

This installs the script to your Python user bin.

## Configuration

### ID's and Secrets

Configuration like the `CLIENT_ID`, `CLIENT_SECRET`, etc. need to be inputted into a config file. Let's create one in the user config folder.

```sh
touch ~/.config/seatingplan.yaml
```

This should create an empty yaml file. Open up this file with your text editor of choice and copy and paste this example. 

```yaml
client_id: "CLIENT_ID"
client_secret: "CLIENT_SECRET"
tenant_id: "TENANT_ID"
ooo_email: "example@example.com"
users: ["Bob", "Alice"]
```

Replace `CLIENT_ID`, `CLIENT_SECRET`, and `TENANT_ID` with the values from the Azure website we wrote down earlier. Replace `ooo_email` with the email of the calendar that has the out of office events. Replace `users` with a list of all the first names of employee's in your organisation. This is case insensitive but has to be spelt the same as their Office365 accounts.

> **Note:** If you are trying to find this file in a file browser and cannot find it, ~/.config is a hidden directory and you will need to enable viewing hidden directories and files in your file browser.

### Output (Optional but recommended)

You can configure the output path too. Normally, the script will output a file called `Seating Plan.csv` in the directory you ran the script in. This can be edited with the `-o` flag. We can put the file in a different folder and have a different name like this:

```sh
seatingplan -c ~/.config/seatingplan.yaml -o "/path/to/file"
```

For example, we can generate a csv in our users Documents folder and name it "Who's in office?"

```sh
seatingplan -c ~/.config/seatingplan.yaml -o "~/Documents/Who's in office" # If you don't provide a .csv file extension, it will be added for you.
```

This also supports datetime formatting. This can be done using Python's formatting codes for datetime [which you can find the docs for here.](https://docs.python.org/3.7/library/datetime.html#strftime-and-strptime-behavior) The datetime provided to this function will be the datetime when the script it run. 

```sh
seatingplan -c ~/.config/seatingplan.yaml -o "Seating Plan {:%Y-%m-%d}"
```

This will output a file called `Seating Plan 2019-09-12.csv`

## Running the program

Now we have configured everything, we can now run the script. To run the script, enter this inside of the terminal.

```sh
seatingplan -c ~/.config/seatingplan.yaml
```

This will generate a `Seating Plan.csv` file in the directory you ran this program. Look at [Output](#Output) to see how to configure the file name of the output.

The first time the script is ran, it will ask you to visit a url. Open the url in your browser and allow the script access to the requested permissions. Once you have done that, you will be redirected to a blank page. Copy the URL and paste it into the console and press enter.

> **Note:** You should login and give permission to the app as a user with *full* permissions to the out of office calendar. This is to ensure the script has permissions to view this calendar in full. This will only need to be done every 90 days.

## Known Issues

- Long events (longer than a month) may not get picked up in the script as their start dates and end dates may not be in reach of the programs search range.
