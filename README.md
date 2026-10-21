# edxso-google-contacts-sync
Automated two-way contact creation between Google Sheets and Google Contacts using Apps Script.


# EDXSO Contact Synchronization

## Project Overview
This project demonstrates synchronization between Google Sheets and Google Contacts using Google Apps Script and the Google People API.

The goal is to reduce manual contact entry by transferring contact information between a spreadsheet and Google Contacts.

## Technologies Used
- Google Sheets
- Google Apps Script (JavaScript)
- Google People API
- Time-driven Apps Script trigger

## Features
- Reads contact names and phone numbers from Google Sheets.
- Creates Google Contacts for spreadsheet entries that do not already match by phone number.
- Imports contacts from Google Contacts into the spreadsheet when their phone numbers are not already present.
- Records the Google Contact resource name, last sync time, and sync status in the spreadsheet.
- Uses a scheduled trigger to run synchronization automatically.

## Spreadsheet Structure
The `Contacts` sheet uses these columns:

| Column | Description       |
|   A    | Name              |
|   B    | Phone Number      |
|   C    | Google Contact ID |
|   D    | Last Synced At    |
|   E    | Sync Status       |

## Setup Instructions
1. Create a Google Sheet and name the worksheet `Contacts`.
2. Open **Extensions → Apps Script**.
3. Add the project source code.
4. Enable the Advanced Google Service named **People API**. Enable the API in the associated Google Cloud project if prompted.
5. Run `setupSync()` and complete the Google authorization prompts.
6. Add test contacts to the spreadsheet.
7. Run `syncContacts()` and check the execution logs and sync status.
8. Verify that the contacts appear in Google Contacts.
9. Add a test contact directly in Google Contacts and check whether it is imported into the spreadsheet on a subsequent sync.
10. For phone verification, add the same Google account to an Android device and enable contact synchronization.

## Automation
The script installs a time-driven trigger intended to run every five minutes. Actual execution can be delayed by Google Apps Script scheduling.

## Current Limitations
- Matching is based on normalized phone numbers.
- The current implementation creates or imports missing contacts; it does not fully synchronize edits and deletions in both directions.
- Conflicting changes are not resolved automatically.
- Phone-number formats and duplicate contacts may require additional handling.
- Contacts synchronization on an Android phone depends on the device, account settings, and Google sync services.

## Testing and Results
Document the results of the tests you actually performed:
- Google Sheets → Google Contacts: Passed
- Google Contacts → Google Sheets: Passed

## Demo
Demo video: [Add your video link here]

## Author
Tanmayee Panat
