### Goodle Scholar Scraper
`scholar-scraper.py`

A simple Python script to crawl a Google Scholar profile page. The result is stored in a spreedsheet and the summary is sent via email. Together with `crontab` in Linux, the scraper can run periodically.
There are a number of FIXMEs in the script.

Dependencies: scholarly, pandas, openpyxl

### Upload Spreadsheet to Google Sheet
`write-to-google-doc.py`

A simple Python script to write the spreadsheet generated above (indeed, any local spreadsheet) to a Google spreadsheet file in your Google drive.

In order to do it, you will need to set up Google Sheets API on GCP first.
1. Go to Google Cloud Console
1. Create a new project
1. Enable Google Sheets API
1. Do OAuth consent screen (may need to refresh the browser afterwards to enable the next step)
1. Create credentials (OAuth 2.0 Client ID)
1. Download the credentials and save as `credentials.json` in the same directory as your script

At the first run of this script, the Google OAuth authentication process needs a browser for the first-time authentication, which will create a `token.pickle`. 
If running on a headless EC2 instance, you can run it locally first to generate `token.pickle` with a browser.
You then send it to the instance in the same directory as your script.

Note that this script assumes the Google Spreadsheet is created with a known spreadsheet ID (you can find this in the spreadsheet's URL). 
Replace the SPREADSHEET_ID in the script with your Google Spreadsheet ID.

Dependencies: google-auth-oauthlib google-auth-httplib2 google-api-python-client
