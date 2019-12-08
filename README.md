# Absence Management

Helps synchronize work holidays across your Google, Outlook and HR systems.

![image](https://user-images.githubusercontent.com/1783027/70386130-81622f00-198d-11ea-991b-616ad108d99c.png)

## Usage

Add absence records to the spreadsheet, marking them as Sick, Bank holiday, Holiday etc. 

Use the Calendar menu options to:

* Create/update/delete entries in your Google and Outlook calendars
* Check your HR system (copied in to the [year - HR] worksheet)

## Installation

1. Create a Google Spreadsheet
1. Import absences.sample.xlsx
1. Create a script project in your spreadsheet
1. Add the files in this repo (renaming to .gs extension)
1. Register the OAuth2 library [here](https://github.com/gsuitedevs/apps-script-oauth2)
1. Register your app in the [Azure Portal Active Directory](https://portal.azure.com)
1. Enter your tenant, client id and secret in the Settings.gs file
