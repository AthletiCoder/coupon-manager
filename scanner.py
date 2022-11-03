from __future__ import print_function

import os.path
from asyncore import read
import cv2
import pandas as pd
import time

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

import smtplib, ssl

smtp_server = "smtp.gmail.com"
sender_email = "ashrayfoundation.edu@gmail.com"
password = "haphvuxwvhdxzzke"


# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
LAST_ROW = 2

# The ID and range of a sample spreadsheet.
FEEDBACKS_SPREADSHEET_ID = '1DaH9vXd1J_8WvkuHHR01Uix_16fIFSdohekGT3T9FdU'
SAMPLE_RANGE_NAME = 'Sheet1!A{row}:D'

def fetch_feedback_data():
    global LAST_ROW
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=FEEDBACKS_SPREADSHEET_ID,
                                    range=SAMPLE_RANGE_NAME.format(row=LAST_ROW)).execute()
        values = result.get('values', [])

        if not values:
            print('No data found!')
            return

        LAST_ROW += len(values)
        return values
    except HttpError as err:
        print("HTTP Error!")
        return None

def send_email(receiver_email):
    message = """\
    Subject: Hi there

    This message is sent from Python."""

    with smtplib.SMTP(smtp_server, 587) as server:
        server.starttls()
        server.login(sender_email, password)
        print(receiver_email)
        server.sendmail(sender_email, receiver_email, message)


while True:
    print(LAST_ROW)
    new_feedbacks = fetch_feedback_data()
    if new_feedbacks:
        for new_feedback in new_feedbacks:
            participant_email = new_feedback[2]
            send_email(participant_email)
    time.sleep(1)