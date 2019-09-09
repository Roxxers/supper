
# Seating Plan Generator

Script to generate a seating plan via calendar events in an organsation's Office365.

## What it does

The script looks at the current week and generates a seating plan for that week. It will then create a csv to represent this. This is done using the default calendar of the user who allows the script access to their account.

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

> **Warning** This guide assumes you are using a UNIX based OS (Linux, Mac OS, etc.). If using Windows, god help you. (Ask me for help if you can't adapt this to Windows. Windows is weird and scary.)

Once the app has been created, git clone this repo, cd into it's folder and install it into your user's Python PATH.

```sh
git clone URL
cd seatplangen
python3 -m pip install . --user
```

Once installed, you can run the script like this.

```sh
seatingplan -c "CLIENT_ID" -t "TENANT_ID" -s "CLIENT_SECRET"
```

The first time the script is ran, it will ask you to visit a url. Open the url in your browser and allow the script access to the requested permissions. Once you have done that, you will be redirected to a blank page. Copy the URL and paste it into the console and press enter.

> **Note:** You should login and give permission to the app *as* the account with the calendar you want to use. This calendar should be the one you are using to store who is out of office.

This will only need to be done every 90 days.

## Configuration

### ID's and Secrets

Configuration like the `CLIENT_ID`, `CLIENT_SECRET`, etc. can be inputted via the command line or via a config file. A config file might be better over a command line input as this does not expose sensitive information to the stdout of the tty. It also means not remembering this every time you run the script. Seeing as a client secret can only be viewed once and has to be stored, I recommend the config file for long term use. To create a config file, do the following:

```sh
touch ~/.config/seatingplan.conf
```

This should create an empty text file. Open up this file with your text editor of choice and copy and paste this example.

```ini
[client]
id=CLIENT_ID
secret=CLIENT_SECRET
tenant=TENANT_ID
```

> **Note:** If you are trying to find this file in a file browser and cannot find it, ~/.config is a hidden directory and you will need to enable viewing hidden directories and files in your file browser.

Make sure to replace `CLIENT_ID` etc. with your own client_id for the app you created earlier.

With this file, you now should be able to run the script like this:

```sh
seatingplan -c ~/.config/seatingplan.conf
```

### Output

***ADD ADVICE HERE ABOUT DATETIME FORMATTING AND HOW TO NAME THE OUTPUT CSV HERE WHEN YOU FINISH ALL THAT CODE***
