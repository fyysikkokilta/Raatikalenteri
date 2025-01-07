# Raatikalenteri (made by Severi Kasurinen)

# Setup
## Prerequisites
- Python 3.11

## Setup
1. Install required packages using `pip install -r .\requirements.txt`
2. Check that the constants and other variables are correct in `raatikalenteri.py` and `cal_setup.py`
3. Create a `.env` file using the example file `.env.example`
4. Create a Google service account with Google Sheets and Google Calendar permissions in Google Cloud Console. Export the service account info as a json and store it on the project root as `service_account_credentials.json`
5. Create a Google OAuth Client ID in Google Cloud console. Export the client id as a json and store it on the project root as `client_secret.json`
6. Share the calendar sheets to the created google service account using the account's email. The service account should have editor permissions.
7. Run `raatikalenteri.py` with Python

Additionally a script can be created to run the updates for example on startup. A bat file already exists in the repository for this.
